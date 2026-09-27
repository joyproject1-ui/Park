"""FDA Warning Letter 수집: 목록 조회 → 무균·GMP 관련 건 선별 → 신규 건만 추림.

식약처 행정처분 메일과 같은 주기로 FDA 경고장을 받아보기 위한 모듈입니다.
fetcher와 마찬가지로 표준 라이브러리만 사용합니다.

FDA는 경고장 목록을 여러 경로로 제공하고 그 형식이 예고 없이 바뀝니다.
그래서 카탈로그의 direct_url/discover 방식과 같은 원리로 **소스를 순서대로
시도**하고, 먼저 성공한 결과를 씁니다.
"""

from __future__ import annotations

import json
import re
import xml.etree.ElementTree as ET
from dataclasses import asdict, dataclass, field
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Callable, Iterable, Sequence

from .fetcher import DEFAULT_RETRIES, DEFAULT_TIMEOUT, FetchError, fetch_text

LANDING_PAGE = (
    "https://www.fda.gov/inspections-compliance-enforcement-and-criminal-investigations"
    "/compliance-actions-and-activities/warning-letters"
)

# 경고장 목록 표의 공식 내보내기 주소. 표 위 'Export' 버튼이 쓰는 것과 같습니다.
# (FDA 는 경고장 전용 RSS 를 제공하지 않습니다 — 2026-09 확인.)
EXPORT_JSON_URL = LANDING_PAGE + "/datatables-data?_format=json"
EXPORT_CSV_URL = LANDING_PAGE + "/datatables-data?_format=csv"

# 목록 페이지가 표를 채울 때 부르는 내부 주소. 형식이 바뀔 수 있어 2순위로 둡니다.
DATATABLES_URL = (
    "https://www.fda.gov/datatables/views/data.json"
    "?total_count_needed=true&view_display_id=warning_letter_solr_block"
    "&view_name=warning_letter_solr_index&length=100&start=0"
)

STATE_NAME = "warning-letters-state.json"

# 1순위: 무균 제제와 직접 닿는 지적. 한 건이라도 걸리면 '무균 관련'으로 분류합니다.
STERILE_KEYWORDS = {
    "sterile": "무균",
    "sterility": "무균성",
    "aseptic": "무균조작",
    "ophthalmic": "점안제",
    "eye drop": "점안제",
    "injectable": "주사제",
    "injection": "주사제",
    "parenteral": "주사제",
    "media fill": "미디어필",
    "environmental monitoring": "환경모니터링",
    "endotoxin": "엔도톡신",
    "pyrogen": "발열성물질",
    "particulate": "이물",
    "isolator": "아이솔레이터",
    "rabs": "RABS",
    "smoke stud": "기류시험",  # smoke study/studies 모두 잡기 위해 어간으로 둡니다.
    "503b": "503B 조제시설",
    "outsourcing facility": "503B 조제시설",
}

# 2순위: 무균과 직접 닿지는 않지만 AQA가 추적해야 할 CGMP 주제.
CGMP_KEYWORDS = {
    "data integrity": "데이터 완전성",
    "out-of-specification": "OOS",
    "out of specification": "OOS",
    "cgmp": "CGMP",
    "quality unit": "품질부서",
    "stability": "안정성시험",
    "cleaning validation": "세척밸리데이션",
    "process validation": "공정밸리데이션",
    "contamination": "오염",
    "recall": "회수",
    "audit trail": "감사추적",
}

CFR_PATTERN = re.compile(r"21\s*CFR\s*(\d{3}\.\d+(?:\([a-z0-9]+\))*)", re.IGNORECASE)

_DATE_FORMATS = ("%m/%d/%Y", "%B %d, %Y", "%b %d, %Y", "%Y-%m-%d", "%d %B %Y")


class WarningLetterError(Exception):
    """경고장 목록을 어떤 소스에서도 가져오지 못했을 때."""


@dataclass
class WarningLetter:
    company: str
    url: str
    letter_date: str = ""       # 서한일자 (FDA가 회사에 보낸 날)
    posted_date: str = ""       # 게시일자 (FDA 사이트에 올라온 날)
    office: str = ""            # 발행 사무소
    subject: str = ""
    source: str = ""            # 어느 소스에서 얻었는지
    sterile_hits: list[str] = field(default_factory=list)
    cgmp_hits: list[str] = field(default_factory=list)
    cfr_citations: list[str] = field(default_factory=list)

    @property
    def key(self) -> str:
        """중복 판정 기준. URL이 가장 안정적입니다."""
        return self.url or f"{self.company}|{self.letter_date}"

    @property
    def is_sterile_related(self) -> bool:
        return bool(self.sterile_hits)

    @property
    def best_date(self) -> str:
        return self.posted_date or self.letter_date


# --------------------------------------------------------------------------- 날짜

def parse_date(value: str) -> str:
    """여러 표기의 날짜 문자열을 YYYY-MM-DD로. 실패하면 빈 문자열."""
    text = (value or "").strip()
    if not text:
        return ""
    match = re.search(r"\d{4}-\d{2}-\d{2}", text)
    if match:
        return match.group(0)
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    # 제목에 섞여 있는 09/16/2026 또는 September 16, 2026 형태를 찾아본다.
    match = re.search(r"\d{1,2}/\d{1,2}/\d{4}", text)
    if match:
        return parse_date(match.group(0))
    match = re.search(r"[A-Z][a-z]{2,9}\s+\d{1,2},\s*\d{4}", text)
    if match:
        return parse_date(match.group(0))
    try:
        return parsedate_to_datetime(text).strftime("%Y-%m-%d")
    except (TypeError, ValueError, IndexError):
        return ""


def _today() -> datetime:
    return datetime.now(timezone.utc)


# --------------------------------------------------------------------------- 파서

def parse_rss(xml_text: str) -> list[WarningLetter]:
    """FDA 경고장 RSS 피드를 파싱."""
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as exc:
        raise WarningLetterError(f"RSS 파싱 실패: {exc}") from exc

    letters: list[WarningLetter] = []
    for item in root.iter("item"):
        title = (item.findtext("title") or "").strip()
        link = (item.findtext("link") or "").strip()
        if not link:
            continue
        # "Company Name - MARCS-CMS 123456 - September 16, 2026" 형태가 많다.
        parts = [p.strip() for p in title.split(" - ") if p.strip()]
        company = parts[0] if parts else title
        letter_date = ""
        for part in reversed(parts[1:]):
            letter_date = parse_date(part)
            if letter_date:
                break
        letters.append(
            WarningLetter(
                company=company,
                url=link,
                letter_date=letter_date,
                posted_date=parse_date(item.findtext("pubDate") or ""),
                subject=(item.findtext("description") or "").strip()[:300],
                source="rss",
            )
        )
    return letters


def _pick(row: dict, *needles: str) -> str:
    """키 이름이 자주 바뀌므로 부분 일치로 값을 집어온다."""
    for needle in needles:
        for key, value in row.items():
            if needle in key.lower() and isinstance(value, str) and value.strip():
                return _strip_tags(value)
    return ""


def _strip_tags(value: str) -> str:
    return " ".join(re.sub(r"<[^>]+>", " ", value).split())


def _extract_href(value: str, base: str = "https://www.fda.gov") -> str:
    match = re.search(r'href="([^"]+)"', value or "")
    if not match:
        return ""
    href = match.group(1)
    return href if href.startswith("http") else base + href


def parse_datatables(json_text: str) -> list[WarningLetter]:
    """FDA 사이트가 목록 표를 채울 때 쓰는 JSON 응답을 파싱."""
    try:
        payload = json.loads(json_text)
    except json.JSONDecodeError as exc:
        raise WarningLetterError(f"JSON 파싱 실패: {exc}") from exc

    rows = payload.get("data") if isinstance(payload, dict) else payload
    if not isinstance(rows, list):
        raise WarningLetterError("JSON에 data 배열이 없습니다")
    return _rows_to_letters(rows, "datatables")


def parse_csv(text: str) -> list[WarningLetter]:
    """공식 내보내기의 CSV 형식. 열 이름은 JSON 과 같은 규칙으로 찾습니다."""
    import csv
    import io

    reader = csv.DictReader(io.StringIO(text.lstrip("\ufeff")))
    if not reader.fieldnames:
        raise WarningLetterError("CSV 머리글이 없습니다")
    return _rows_to_letters(list(reader), "csv")


def _rows_to_letters(rows: "Iterable[dict]", source: str) -> list[WarningLetter]:
    letters: list[WarningLetter] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        raw_company = _find_raw(row, "company", "legal_name", "title")
        url = _extract_href(raw_company) or _extract_href(_find_raw(row, "link", "url"))
        if not url:
            plain = _find_raw(row, "link", "url", "href")
            url = plain.strip() if plain.strip().startswith("http") else ""
        company = _pick(row, "company", "legal_name", "title")
        if not company:
            continue
        letters.append(
            WarningLetter(
                company=company,
                url=url,
                letter_date=parse_date(_pick(row, "issue_date", "letter_issue", "issue")),
                posted_date=parse_date(_pick(row, "change_date", "posted", "update")),
                office=_pick(row, "issuing_office", "office"),
                subject=_pick(row, "subject"),
                source=source,
            )
        )
    return letters


def _find_raw(row: dict, *needles: str) -> str:
    for needle in needles:
        for key, value in row.items():
            if needle in key.lower() and isinstance(value, str) and value.strip():
                return value
    return ""


def parse_landing_html(html: str) -> list[WarningLetter]:
    """마지막 수단: 목록 페이지의 <a> 링크에서 경고장 URL을 건져낸다."""
    from .fetcher import extract_links

    letters: list[WarningLetter] = []
    seen: set[str] = set()
    for url, text in extract_links(html, LANDING_PAGE):
        if "/warning-letters/" not in url or url.rstrip("/").endswith("warning-letters"):
            continue
        if url in seen or not text:
            continue
        seen.add(url)
        letters.append(
            WarningLetter(
                company=text.strip(),
                url=url,
                letter_date=parse_date(text),
                source="html",
            )
        )
    return letters


# --------------------------------------------------------------------------- 수집

@dataclass
class Source:
    name: str
    url: str
    parse: Callable[[str], "list[WarningLetter]"]


def default_sources(
    export_json_url: str = EXPORT_JSON_URL,
    export_csv_url: str = EXPORT_CSV_URL,
    datatables_url: str = DATATABLES_URL,
    landing_page: str = LANDING_PAGE,
) -> list[Source]:
    return [
        Source("export-json", export_json_url, parse_datatables),
        Source("export-csv", export_csv_url, parse_csv),
        Source("datatables", datatables_url, parse_datatables),
        Source("html", landing_page, parse_landing_html),
    ]


def collect(
    sources: "Sequence[Source] | None" = None,
    timeout: int = DEFAULT_TIMEOUT,
    retries: int = DEFAULT_RETRIES,
    on_source: "Callable[[str, str, int], None] | None" = None,
) -> list[WarningLetter]:
    """소스를 순서대로 시도해 첫 성공 결과를 반환합니다.

    on_source(name, outcome, count) 로 각 시도의 결과를 알려줍니다.
    """
    problems: list[str] = []
    for source in sources if sources is not None else default_sources():
        try:
            text = fetch_text(source.url, timeout=timeout, retries=retries)
            letters = source.parse(text)
        except (FetchError, WarningLetterError) as exc:
            problems.append(f"{source.name}: {exc}")
            if on_source:
                on_source(source.name, "실패", 0)
            continue
        if not letters:
            problems.append(f"{source.name}: 항목 0건")
            if on_source:
                on_source(source.name, "빈 응답", 0)
            continue
        if on_source:
            on_source(source.name, "성공", len(letters))
        return letters
    raise WarningLetterError("경고장 목록을 가져오지 못했습니다 — " + " / ".join(problems))


def classify(letter: WarningLetter, text: str) -> WarningLetter:
    """본문(또는 제목)에서 키워드와 21 CFR 인용을 뽑아 기록합니다."""
    lowered = text.lower()
    letter.sterile_hits = sorted({ko for en, ko in STERILE_KEYWORDS.items() if en in lowered})
    letter.cgmp_hits = sorted({ko for en, ko in CGMP_KEYWORDS.items() if en in lowered})
    letter.cfr_citations = sorted({m.group(1) for m in CFR_PATTERN.finditer(text)})
    return letter


def enrich(
    letters: "Iterable[WarningLetter]",
    timeout: int = DEFAULT_TIMEOUT,
    retries: int = 1,
    on_error: "Callable[[WarningLetter, Exception], None] | None" = None,
) -> list[WarningLetter]:
    """각 경고장 본문을 열어 키워드·CFR 인용을 채웁니다 (건당 1회 요청)."""
    result = []
    for letter in letters:
        try:
            body = fetch_text(letter.url, timeout=timeout, retries=retries)
        except FetchError as exc:
            if on_error:
                on_error(letter, exc)
            classify(letter, f"{letter.company} {letter.subject}")
        else:
            classify(letter, _strip_tags(body))
            if not letter.office:
                match = re.search(r"(?:Division of|Office of)[^.<\n]{3,80}", body)
                if match:
                    letter.office = _strip_tags(match.group(0))
        result.append(letter)
    return result


# --------------------------------------------------------------------------- 선별

def within_days(letters: "Iterable[WarningLetter]", days: int, now: "datetime | None" = None) -> list[WarningLetter]:
    """지난 N일 이내 건만. 날짜를 못 읽은 건은 판단할 수 없으므로 남깁니다."""
    cutoff = ((now or _today()) - timedelta(days=days)).strftime("%Y-%m-%d")
    kept = []
    for letter in letters:
        date = letter.best_date
        if not date or date >= cutoff:
            kept.append(letter)
    return kept


def only_relevant(letters: "Iterable[WarningLetter]", sterile_only: bool = False) -> list[WarningLetter]:
    if sterile_only:
        return [l for l in letters if l.is_sterile_related]
    return [l for l in letters if l.sterile_hits or l.cgmp_hits]


def sort_letters(letters: "Iterable[WarningLetter]") -> list[WarningLetter]:
    """무균 관련 건을 위로, 그 다음 최신순."""
    return sorted(letters, key=lambda l: (not l.is_sterile_related, _neg_date(l.best_date), l.company))


def _neg_date(date: str) -> str:
    # 내림차순 정렬을 문자열만으로 처리하기 위한 보수(補數) 키.
    return "".join(chr(ord("9") - int(c)) if c.isdigit() else c for c in date) if date else "~"


# --------------------------------------------------------------------------- 상태

def load_state(path: "str | Path") -> dict:
    file = Path(path)
    if not file.exists():
        return {"seen": {}}
    try:
        data = json.loads(file.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {"seen": {}}
    if not isinstance(data.get("seen"), dict):
        data["seen"] = {}
    return data


def new_letters(letters: "Iterable[WarningLetter]", state: dict) -> list[WarningLetter]:
    seen = state.get("seen", {})
    return [letter for letter in letters if letter.key not in seen]


def save_state(path: "str | Path", state: dict, letters: "Iterable[WarningLetter]") -> Path:
    file = Path(path)
    file.parent.mkdir(parents=True, exist_ok=True)
    seen = dict(state.get("seen", {}))
    stamp = _today().strftime("%Y-%m-%dT%H:%M:%SZ")
    for letter in letters:
        seen[letter.key] = {"company": letter.company, "date": letter.best_date, "notified_at": stamp}
    file.write_text(
        json.dumps({"generated_at": stamp, "seen": seen}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return file


def to_dicts(letters: "Iterable[WarningLetter]") -> list[dict]:
    return [asdict(letter) for letter in letters]
