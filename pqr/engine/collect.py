# -*- coding: utf-8 -*-
"""제품 폴더의 항 번호별 파일을 찾아 판독기에 넘기고, 보고서에 넣을 값(ProductData)으로 모은다.

담당자 폴더 규칙: 파일 이름이 보고서 항 번호로 시작한다 ("6. 제조내역 - ERP(수출용).pdf",
"9.2.3 충전 완료 후 - ERP.zip", "13 안정성 시험 - … 시험일지.pdf"). 압축은 풀어서 본다.
읽지 못한 것은 버리지 않고 `issues` 에 남겨 담당자 문의 목록으로 보여 준다.
"""
import json
import os
import re
import tempfile
import zipfile

from .. import build as build_module
from .readers import batch_record
from .readers import coa as coa_reader, erp, license as license_reader, deviation, change, \
    masters, yield_sheet, suppliers, trend as trend_reader
from .pdftext import is_scanned, PdfTextError

ITEM_RE = re.compile(r"^\s*(\d{1,2}(?:\.\d+)*)[.\s]")


class ProductData(object):
    def __init__(self):
        self.lots = []            # [(lot, mfg_date, export)]  제조일자는 성적서·제조내역에서
        self.yields = {}          # lot -> {"조제": "97.45", ...}
        self.yield_specs = {}     # {"조제": "94% 이상", ...} — 그해 수율현황표에 적힌 기준
        self.coa = {}             # lot -> {"922": {...}, "923": {...}, "924": {...}}
        self.raw_tests = []       # 8.2.1  [(코드, 시험번호, [lots])]
        self.pkg_tests = []       # 8.2.2  [(코드, 시험번호, [lots])]
        self.manufacturing = []   # 6항    [(lot, 품명, 제조일자, 사용기한)]
        self.batch = None         # 6항 공 기록서(제조·충전·포장 워드) — 제조단위·포장단위·수율 기준·원/자재 (batch_record.merge)
        self.batch_exp = None     # 이름에 '(수출용)' 이 붙은 공 기록서는 따로 — 수출 제조단위·포장단위·수율 기준
        self.license = {}
        self.deviations = []
        self.changes = []
        self.equipment = {}
        self.support = {}
        self.pv = None
        self.pv_exp = None
        self.suppliers_raw = []
        self.suppliers_mat = []
        self.api_chain = {}
        self.stability_files = []   # 스캔 PDF (손글씨) — 비전 판독 대상
        self.stability_logs = []    # 13항 시험일지 판독 결과 (비전 또는 담당자가 적은 .json)
        self.stability_cache = None   # (판독 파일 경로, reader_version) — 옛 판독기 결과면 애매한 칸을 다시 읽는다
        self.stability_all_read = False  # 판독 파일에 covers_all 이 있으면 시험일지 PDF 를 더 읽지 않는다
        self.stability_trend = []   # 이미 채워 둔 안정성 경향표(HLF-QC-126-06) 를 다시 읽은 것
        self.pv_reasons = {}        # {제조번호: 밸리데이션 실시 사유} — 전년도 결재본 10.1 에서
        self.prev_stability = {}    # 전년도 결재본의 13.1 · 13.3 표 (올해 시험일지를 못 읽었을 때)
        self.previous_name = ""     # 참고한 전년도 결재본 파일 이름
        self.prev_sections = {}     # 전년도 결재본의 항별 표 (10.1 등)
        self.folder = ""            # 이 자료를 읽은 제품 폴더
        self.previous_report = None
        self.files = {}           # item -> [paths]
        self.issues = []          # [(항, 파일, 설명)]

    @property
    def domestic(self):
        return [l for l, _, e in self.lots if not e]

    @property
    def export(self):
        return [l for l, _, e in self.lots if e]


_MATCHER = None


def _item_of(name):
    """파일 이름 → 평가항목 번호. 프로그램의 항목 규칙(item_matcher)을 그대로 쓴다 —
    '10.3, 10.4, 10.5 제조지원 …' 이 '10.3-5' 로 잡히는 것까지 같아야 한다."""
    global _MATCHER
    if _MATCHER is None:
        _MATCHER = build_module.item_matcher(build_module.load_config()["items"])
    item = _MATCHER(name)
    if item:
        return item
    m = ITEM_RE.match(name)
    return m.group(1) if m else None


# 만든 보고서를 넣는 폴더 — 자료가 아니므로 읽을 때는 지나친다 (정의는 build 에).
OUTPUT_DIR = build_module.OUTPUT_DIR


def is_output_dir(name):
    return os.path.basename(str(name).rstrip("/\\")) == OUTPUT_DIR


def _walk(folder):
    for root, dirs, files in os.walk(folder):
        dirs[:] = [d for d in dirs if not is_output_dir(d)]
        for name in files:
            if name.startswith("~$") or name.startswith("."):
                continue
            yield os.path.join(root, name)


def _member_name(info):
    """한글 이름이 CP437 로 깨져 오는 압축이 흔하다.

    파이썬은 UTF-8 표시가 없는 압축의 이름을 CP437 로 읽는다. 알집·탐색기가 만든 압축은
    실제로는 CP949 이고, 리눅스·맥이 만든 압축은 표시 없이 UTF-8 인 경우가 있다 —
    되돌린 바이트를 UTF-8 로 먼저, 안 되면 CP949 로 읽는다.
    """
    name = info.filename
    if not (info.flag_bits & 0x800):
        try:
            raw = name.encode("cp437")
        except UnicodeEncodeError:
            raw = None
        for enc in ("utf-8", "cp949"):
            if raw is None:
                break
            try:
                name = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
    return name.replace("\\", "/")


def _unzip(path, workdir):
    """압축을 작업 폴더에 푼다. 안의 폴더 구조는 그대로 둔다 — 폴더째 묶은 압축은
    '9.2.1 조제 완료 후/…' 처럼 폴더 이름이 항 번호를 들고 있기 때문이다."""
    out = os.path.join(workdir, re.sub(r"[^\w.-]+", "_", os.path.basename(path)))
    os.makedirs(out, exist_ok=True)
    with zipfile.ZipFile(path) as z:
        for info in z.infolist():
            if info.is_dir():
                continue
            parts = [p for p in _member_name(info).split("/") if p not in ("", ".", "..")]
            if not parts or parts[0] == "__MACOSX":
                continue
            target = os.path.join(out, *parts)
            os.makedirs(os.path.dirname(target), exist_ok=True)
            with open(target, "wb") as h:
                h.write(z.read(info.filename))
    return out


def discover(folder, workdir=None, depth=3):
    """{항: [파일 경로]} — 폴더 아래 파일·폴더·압축을 항 번호로 나눈다.

    번호가 붙은 폴더(`13. 안정성 시험`) 안의 파일은 이름에 번호가 없어도 그 항으로 친다.
    번호가 없는 중간 폴더(`필요 자료`)와 번호 없는 압축(폴더째 묶은 것)은 그냥 지나쳐
    안쪽을 계속 본다 — 담당자가 자료를 한 단계 더 접어 두는 일이 흔하다.
    """
    workdir = workdir or tempfile.mkdtemp(prefix="pqr-engine-")
    items = {}

    def scan(root, left):
        for name in sorted(os.listdir(root)):
            if name.startswith("~$") or name.startswith(".") or is_output_dir(name):
                continue                          # 우리가 만든 보고서는 자료가 아니다
            path = os.path.join(root, name)
            item = _item_of(name)
            if os.path.isdir(path):
                if item:
                    items.setdefault(item, []).extend(_walk(path))
                elif left > 0:                    # 번호 없는 중간 폴더는 지나쳐 들어간다
                    scan(path, left - 1)
                continue
            is_zip = name.lower().endswith(".zip")
            if not item:
                if is_zip and left > 0:           # 번호 없는 압축 = 폴더째 묶은 것
                    scan(_unzip(path, workdir), left - 1)
                continue
            paths = list(_walk(_unzip(path, workdir))) if is_zip else [path]
            items.setdefault(item, []).extend(paths)

    scan(folder, depth)
    root = os.path.abspath(folder)
    return {item: _dedupe(paths, root) for item, paths in items.items()}


def _dedupe(paths, root):
    """같은 파일을 두 번 세지 않는다 — 이름과 크기가 같으면 한 번만.

    담당자는 자료를 압축으로 올린 뒤 제품 폴더에서 압축을 풀어 두는 일이 흔하다. 그러면
    같은 파일이 '…필요 자료/' 와 '…필요 자료.zip' 양쪽에 있어 두 번 읽히고, 8.2 시험결과표에
    같은 줄이 두 벌 실린다(디겐타안연고 2026: 주원료 4줄이 8줄로).

    남기는 것은 **제품 폴더 안에 실제로 있는 파일**이다 — 담당자가 열어 고치는 것이 그쪽이고,
    압축 안의 것은 올릴 때 그대로 굳은 사본이라 손대면 어긋난다. 그래서 폴더와 압축을 둘 다
    두어도 되고, 어느 한쪽만 두어도 된다.
    """
    def outside(path):
        return not os.path.abspath(path).startswith(root)      # 압축에서 꺼낸 것은 밖에 있다

    best, order = {}, []
    for path in paths:
        try:
            key = (os.path.basename(path).lower(), os.path.getsize(path))
        except OSError:
            key = (path, None)
        if key not in best:
            best[key] = path
            order.append(key)
        elif outside(best[key]) and not outside(path):         # 폴더 안의 것이 나오면 그것으로
            best[key] = path
    return [best[key] for key in order]


def _lot_from_name(name):
    m = re.search(r"\b([A-Z]{2}[A-Z0-9]{4})\b", name)
    return m.group(1) if m else None


def same_files_once(paths, log=None):
    """내용이 같은 시험일지는 하나만 남긴다 — 담당자가 같은 일지를 이름만 바꿔 여러 번 올려 두는 일이 잦다.

    담당자 PC 2026-09-07: 13 폴더에 24개가 있었지만 실제 문서는 여덟 가지였다(같은 일지가
    '…시험일지 OEV301.pdf'·'[OG-22-1]…_QEV301_36M.pdf'·'…시험일지.pdf' 로 세 번). 한 장에 몇 분이라
    같은 것을 다시 읽으면 그만큼 그냥 기다리게 된다.
    """
    import hashlib
    out, seen, 같은것 = [], {}, 0
    for path in paths:
        try:
            with open(path, "rb") as handle:
                key = hashlib.md5(handle.read()).hexdigest()
        except OSError:
            out.append(path)
            continue
        if key in seen:
            같은것 += 1
            continue
        seen[key] = path
        out.append(path)
    if 같은것 and log:
        log("  [13] 내용이 같은 시험일지 %d장은 한 번만 읽습니다 (%d장 → %d장)"
            % (같은것, len(paths), len(out)))
    return out


def unread_scans(logs, scanned):
    """판독 파일(logs)에 아직 없는 시험일지 — 판독 결과는 읽은 일지 이름(source)을 들고 있다.
    어느 일지를 읽었는지 모르는 옛 판독 파일(source 없음)이면 아무것도 새로 읽지 않는다."""
    known = {one.get("source") for one in logs if one.get("source")}
    if not known:
        return []
    return [p for p in scanned if os.path.basename(p) not in known]


FORM_WORDS = ("점안액", "안연고", "점안제", "연고", "크림", "겔", "캡슐", "시럽", "주사", "정")


def _core_name(name):
    """제품 이름의 알맹이 — '한림포비돈점안액(내수용)' → '한림포비돈'."""
    name = re.sub(r"\(.*?\)|\s+", "", name or "")
    for word in FORM_WORDS:
        if word in name:
            return name.split(word)[0]
    return name


def other_product_record(file_name, product_name):
    """공 기록서 파일 이름이 다른 제품 것인가 — 이 제품 이름은 없고 다른 '…점안액' 꼴 이름이 있으면."""
    core = _core_name(product_name)
    if not core or core in re.sub(r"\s+", "", file_name or ""):
        return False
    others = re.findall(r"([가-힣A-Za-z]{2,})(?:%s)" % "|".join(FORM_WORDS), file_name or "")
    return any(o != core for o in others)


def handwriting_cache_name():
    from . import handwriting
    return handwriting.CACHE_NAME


# 항마다 프로그램이 읽을 수 있는 파일 꼴 — 이 밖의 파일은 조용히 지나가지 않고 문의 목록에 남긴다.
# 담당자 2026-09-07: 9·10.2·12·13항이 통째로 비었는데 까닭이 어디에도 없었다
# ("도대체 왜 작성이 안 되는 거야?", "뭐가 문제인지 정확히 나한테 요구해").
READABLE = {
    "3":       ((".pdf",),            "허가증 PDF"),
    "6":       ((".pdf", ".docx"),    "제조내역 ERP PDF 또는 공 기록서 .docx"),
    "7":       ((".xlsx", ".xls"),    "수율현황표 엑셀"),
    "8.1.1":   ((".xlsx",),           "공급업체 List 엑셀"),
    "8.1.2":   ((".xlsx",),           "주성분 공급망 엑셀"),
    "8.1.3":   ((".xlsx",),           "공급업체 List 엑셀"),
    "8.2.1":   ((".xlsx", ".xls"),    "원료 시험 ERP 엑셀"),
    "8.2.1.1": ((".xlsx", ".xls"),    "원료 시험 ERP 엑셀"),
    "8.2.2":   ((".xlsx", ".xls"),    "자재 시험 ERP 엑셀"),
    "9.2.1":   ((".pdf",),            "공정 시험성적서 PDF"),
    "9.2.2":   ((".pdf",),            "공정 시험성적서 PDF"),
    "9.2.3":   ((".pdf",),            "공정 시험성적서 PDF"),
    "9.2.4":   ((".pdf",),            "완제 시험성적서 PDF"),
    "10.1":    ((".xlsx",),           "PV 마스터파일 엑셀"),
    "10.2":    ((".xlsx",),           "적격성 마스터파일 엑셀"),
    "10.3-5":  ((".xlsx",),           "제조지원 설비 마스터파일 엑셀"),
    "11":      ((".pdf",),            "일탈·OOS 보고서 PDF"),
    "12":      ((".pdf",),            "변경요청서 PDF"),
    "13":      ((".pdf", ".json"),    "안정성 시험일지 PDF(스캔) 또는 판독 .json"),
    "16":      ((".doc", ".docx", ".xlsx"), "전년도 결재본 .docx 와 경향표 엑셀"),
}


def unread_files(got, log=None):
    """올렸는데 프로그램이 읽지 못한 파일 — [(항, 경로, 까닭)].

    '해당없음 확인' 마감 글과 우리가 남긴 글(PQR …, 판독 .json)은 뺀다.
    """
    rows = []
    for item, paths in sorted(got.items()):
        rule = READABLE.get(item)
        if rule is None:
            continue
        exts, want = rule
        살펴본것 = []
        for path in paths:
            name = os.path.basename(path)
            if name.startswith(("~$", ".", "PQR ")) or name == handwriting_cache_name():
                continue
            if name.lower().endswith(".txt"):        # '해당없음 확인' 마감 글
                continue
            살펴본것.append(path)
        읽을수있는것 = [p_ for p_ in 살펴본것 if exts and p_.lower().endswith(tuple(exts))]
        if 읽을수있는것 or not 살펴본것:
            continue                                 # 하나라도 읽을 수 있으면 조용히 — 잔소리가 되면 안 읽는다
        for path in 살펴본것:
            rows.append((item, path,
                         "★ 이 파일을 읽지 못해 %s항이 빈 칸으로 남았습니다 — 프로그램은 %s 를 읽습니다"
                         % (item, want)))
    if log and rows:
        log("  [문제] 읽지 못한 파일 %d개: %s"
            % (len(rows), ", ".join(os.path.basename(p) for _, p, _ in rows)))
    return rows


# 이름이 '12.' 로 시작하면 12항 자료로 셈해진다 — 우리가 남긴 글은 'PQR' 로 시작한다
CHANGE_CACHE = "PQR 변경요청서 판독.json"


def _change_cache_key(path):
    try:
        st = os.stat(path)
        return "%s|%d|%d" % (os.path.basename(path), st.st_size, int(st.st_mtime))
    except OSError:
        return os.path.basename(path)


def _change_cache(folder):
    try:
        with open(os.path.join(folder or "", CHANGE_CACHE), encoding="utf-8") as handle:
            got = json.load(handle)
        return got if isinstance(got, dict) else {}
    except (OSError, ValueError):
        return {}


def _save_change_cache(folder, got):
    if not folder or not os.path.isdir(folder):
        return
    try:
        with open(os.path.join(folder, CHANGE_CACHE), "w", encoding="utf-8") as handle:
            json.dump(got, handle, ensure_ascii=False, indent=2)
    except OSError:
        pass


def _change_by_claude(path, folder, log, note):
    """글자 없는 스캔 변경요청서를 Claude 로 읽는다 — 못 읽으면 None.

    담당자 2026-09-07: "못 읽으면 다른 방법을 사용해서라도 읽게 해야지, 공란으로 두면 안 돼."
    빠른 길부터: ① API 키(Claude) → ② 이 PC 의 Claude Code → ③ 이 PC 의 OCR(한글).

    한 번 읽은 것은 제품 폴더의 판독 파일에 남긴다 — 판독은 몇 분이 걸리므로 재작성 때마다
    다시 읽으면 안 된다 (담당자 2026-09-08: "보고서 작성 시간이 기존보다 오래 걸리는 이유가?").
    """
    say = log or (lambda *a: None)
    key, cache = _change_cache_key(path), _change_cache(folder)
    if key in cache:
        got = dict(cache[key])
        got["actions"] = [tuple(a) for a in got.get("actions") or []]
        say("    [12] %s — 지난 판독 결과를 그대로 씁니다 (%s)" % (os.path.basename(path), CHANGE_CACHE))
        return got
    갈래 = []
    try:
        from . import vision as vision_mod, vision_claude
        if vision_mod.available():
            갈래.append(("Claude(API 키)", lambda: vision_claude.read_change(path, say)))
    except Exception:
        pass
    try:
        from . import claude_cli
        if claude_cli.available():
            갈래.append(("이 PC 의 Claude Code", lambda: claude_cli.read_change(path, folder, say)))
    except Exception:
        pass
    try:
        # 마지막 보루 — 이 PC 의 OCR. 한글 인식 모델을 받아 둔 PC 에서만 쓴다
        # (담당자 2026-09-07: "OCR 로 변환해서 읽으면 안 되는 거야?" — 딸려 오는 모델은
        # 중국어·영어용이라 한글이 뭉개진다).
        from .readers import change_ocr
        # korean_engine() 은 처음 한 번 모델을 내려받으려 하므로 여기서 부르지 않는다 —
        # 앞 갈래가 이미 읽었으면 그 시도조차 하지 않게, 실제로 이 길을 탈 때 부른다.
        갈래.append(("이 PC 의 OCR(한글)", lambda: change_ocr.read(path, say)))
    except Exception:
        pass
    claude_있음 = any(이름 != "이 PC 의 OCR(한글)" for 이름, _ in 갈래)
    까닭 = []
    for 이름, 부르기 in 갈래:
        try:
            got = 부르기()
        except Exception as error:
            까닭.append("%s: %s" % (이름, error))
            continue
        if got.get("title") or got.get("description") or got.get("actions"):
            m = re.search(r"(CC-\d{6}-\d{2})", os.path.basename(path))
            if m and not got.get("doc_no"):
                got["doc_no"] = m.group(1)
            cache[key] = dict(got, actions=[list(a) for a in got.get("actions") or []])
            _save_change_cache(folder, cache)     # 다음 작성부터는 이 값을 그대로 쓴다
            return got
        까닭.append("%s: 읽어 낸 것이 없음" % 이름)
    도움 = ("" if claude_있음 else
            " 이 PC 에 Claude(API 키)도 Claude Code 도 없습니다 — 'PQR-Claude설치.bat' 을 실행하면 "
            "다음부터 읽습니다.")
    note("12", path, "스캔 변경요청서를 읽지 못했습니다 — %s%s" % (" / ".join(까닭) or "쓸 수 있는 판독기가 없습니다", 도움))
    return None


DEV_CACHE = "PQR 일탈보고서 판독.json"


def _deviation_by_claude(path, folder, log, note):
    """글자 없는 스캔 일탈 보고서를 Claude 로 읽는다 — 변경요청서 판독기를 그대로 쓴다.

    일탈 보고서도 '제목·내용·부서별 조치' 라는 짜임이 같다. 한 번 읽은 것은 되쓴다.
    """
    say = log or (lambda *a: None)
    key = _change_cache_key(path)
    box = os.path.join(folder or "", DEV_CACHE)
    try:
        with open(box, encoding="utf-8") as handle:
            cache = json.load(handle)
        cache = cache if isinstance(cache, dict) else {}
    except (OSError, ValueError):
        cache = {}
    if key not in cache:
        got = _change_by_claude(path, folder, log, lambda *a: None)
        if got is None:
            note("11", path, "글자 없는 스캔 일탈 보고서입니다 — Claude 로도 읽지 못했습니다. "
                             "일탈 내용을 직접 채우세요")
            return None
        cache[key] = dict(got, actions=[list(a) for a in got.get("actions") or []])
        try:
            with open(box, "w", encoding="utf-8") as handle:
                json.dump(cache, handle, ensure_ascii=False, indent=2)
        except OSError:
            pass
        say("    [11] %s — Claude 로 읽음: %s" % (os.path.basename(path), got.get("title") or ""))
    one = cache[key]
    m = re.search(r"(DR-\d{6}-\d{2})", os.path.basename(path))
    return {"doc_no": one.get("doc_no") or (m.group(1) if m else ""),
            "title": one.get("title") or "", "lot": "", "occurred": "", "planned": False,
            "actions": [tuple(a) for a in one.get("actions") or []], "completed": None,
            "description": one.get("description") or ""}


COA_CACHE = "PQR 시험성적서 판독.json"


def _coa_by_claude(path, folder, log, note, item):
    """글자 없는 스캔 시험성적서를 Claude 로 읽는다 — 못 읽으면 None.

    담당자 2026-09-08: "시험 성적을 제대로 못 읽는 것 같아." 아이퓨어점안액의 9.2.4 완제
    성적서(LIMS 연동 X)가 세 쪽 모두 글자 0자인 스캔본이라 한 칸도 못 읽었다. 변경요청서와
    같은 갈래로 읽고, 한 번 읽은 것은 되쓴다 — 성적서는 Lot 마다 있어 여러 장이다.
    """
    say = log or (lambda *a: None)
    key = _change_cache_key(path)
    box = os.path.join(folder or "", COA_CACHE)
    try:
        with open(box, encoding="utf-8") as handle:
            cache = json.load(handle)
        cache = cache if isinstance(cache, dict) else {}
    except (OSError, ValueError):
        cache = {}
    if key in cache:
        say("    [%s] %s — 지난 판독 결과를 그대로 씁니다" % (item, os.path.basename(path)))
        return dict(cache[key])
    갈래 = []
    try:
        from . import vision as vision_mod, vision_claude
        if vision_mod.available():
            갈래.append(("Claude(API 키)", lambda: vision_claude.read_coa(path, say)))
    except Exception:
        pass
    try:
        from . import claude_cli
        if claude_cli.available():
            갈래.append(("이 PC 의 Claude Code", lambda: claude_cli.read_coa(path, folder, say)))
    except Exception:
        pass
    까닭 = []
    for 이름, 부르기 in 갈래:
        try:
            got = 부르기()
        except Exception as error:
            까닭.append("%s: %s" % (이름, error))
            continue
        if got.get("assays") or got.get("appearance") or got.get("bioburden"):
            cache[key] = got
            try:
                with open(box, "w", encoding="utf-8") as handle:
                    json.dump(cache, handle, ensure_ascii=False, indent=2)
            except OSError:
                pass
            return got
        까닭.append("%s: 읽어 낸 것이 없음" % 이름)
    note(item, path, "글자 없는 스캔 시험성적서입니다 — %s"
         % (" / ".join(까닭) or "이 PC 에 Claude(API 키)도 Claude Code 도 없습니다. "
                                "'PQR-Claude설치.bat' 을 실행하면 다음부터 읽습니다"))
    return None


def collect(folder, product_name=None, log=None):
    log = log or (lambda *a: None)
    data = ProductData()
    data.folder = os.path.abspath(folder)
    data.files = discover(folder)
    got = data.files

    def note(item, path, why):
        data.issues.append((item, os.path.basename(path), why))
        log("  [%s] %s — %s" % (item, os.path.basename(path), why))

    # 6. 제조내역 (수출용 ERP)  · 7. 수율
    records = []
    for p in got.get("6", []):
        if p.lower().endswith(".pdf"):
            try:
                data.manufacturing += erp.read_manufacturing(p)
            except PdfTextError as e:
                note("6", p, str(e))
        elif p.lower().endswith(".docx"):
            # 공 기록서(제조·충전·포장) — 제조단위·포장단위·수율 기준·원/자재 코드 (담당자 2026-09-06)
            try:
                records.append(batch_record.read(p))
            except Exception as e:                        # 서식이 달라 못 읽어도 작성은 계속
                note("6", p, "공 기록서를 읽지 못함: %s" % e)
    # 다른 제품의 공 기록서가 섞여 오면(담당자 PC 2026-09-07: 한림포비돈점안액 폴더에 '올로원스점안액 3ml'
    # 기록서) 그 제품의 자재(PE병 5mL 등)와 포장단위가 표에 들어간다 — 이름이 다른 제품 것이면 쓰지 않고 알린다.
    wrong = [r for r in records if other_product_record(r["file"], product_name)]
    for r in wrong:
        note("6", r["file"], "★ 다른 제품의 공 기록서로 보여 쓰지 않았습니다 — 이 제품(%s)의 제조·충전·포장 "
                             "기록서를 올려 주세요" % (product_name or ""))
    records = [r for r in records if r not in wrong]
    if records:
        exp = [r for r in records if "수출용" in r["file"]]
        dom = [r for r in records if r not in exp] or records
        data.batch = batch_record.merge(dom)
        data.batch_exp = batch_record.merge(exp) if exp else None
        for label, rec in (("내수", data.batch), ("수출", data.batch_exp)):
            if not rec:
                continue
            mats = rec["materials"]
            log("  [6] 공 기록서(%s) %d개 — 제조단위 %s · 포장단위 %s · 수율 기준 %s · 주원료 %d·부원료 %d·포장자재 %d"
                % (label, len(rec["files"]), rec["batch_size"] or "못 읽음", rec["pack_unit"] or "못 읽음",
                   ", ".join("%s %s" % kv for kv in rec["yield_specs"].items()) or "못 읽음",
                   len(mats["주원료"]), len(mats["부원료"]), len(mats["포장자재"])))
    for p in got.get("7", []):
        if p.lower().endswith((".xlsx", ".xls")):
            for lot, vals in yield_sheet.read_yields(p):
                data.yields[lot] = vals
            data.yield_specs.update(yield_sheet.read_specs(p))
    # 8.x
    for p in got.get("8.1.1", []):
        if p.lower().endswith(".xlsx"):
            data.suppliers_raw = suppliers.read_supplier_list(p)
    for p in got.get("8.1.3", []):
        if p.lower().endswith(".xlsx"):
            data.suppliers_mat = suppliers.read_supplier_list(p)
    for p in got.get("8.1.2", []):
        if p.lower().endswith(".xlsx"):
            data.api_chain = suppliers.read_api_chain(p)
    # ERP 가 그대로 내려 준 표에는 그 원료를 쓴 모든 제품이 들어 있다 — 제품 이름으로 가려낸다.
    for p in got.get("8.2.1", []) + got.get("8.2.1.1", []):
        if p.lower().endswith((".xls", ".xlsx")):
            data.raw_tests += erp.group_by_test(erp.read_material_tests(p, product=product_name))
    for p in got.get("8.2.2", []):
        if p.lower().endswith((".xls", ".xlsx")):
            data.pkg_tests += erp.group_by_test(erp.read_material_tests(p, product=product_name))
    # 9.2.x 성적서
    # 9.2.1 도 조제 단계 성적서다 — 점안제는 9.2.1(점안제)·9.2.2(바이오버든)로 갈라 올린다
    # (담당자 2026-09-07: 한림포비돈점안액 폴더에 'ELYO01 조제 완료 후.pdf' 넉 장이 9.2.1 로 올라와 있었다).
    for item, key, fn in (("9.2.1", "921", coa_reader.read_ipc),
                          ("9.2.2", "922", coa_reader.read_ipc), ("9.2.3", "923", coa_reader.read_ipc),
                          ("9.2.4", "924", coa_reader.read_fp)):
        for p in got.get(item, []):
            if not p.lower().endswith(".pdf"):
                continue
            try:
                rec = fn(p)
            except PdfTextError as e:
                note(item, p, str(e)); continue
            # 글자가 없는 스캔 성적서는 한 칸도 못 읽는다 — Claude 에게 맡긴다
            # (담당자 2026-09-08: "시험 성적을 제대로 못 읽는 것 같아").
            읽힌것 = [k for k, v in rec.items() if k not in ("file",) and v]
            if len(읽힌것) <= 1 and is_scanned(p):
                읽음 = _coa_by_claude(p, folder, log, note, item)
                if 읽음 is not None:
                    rec = 읽음
            lot = rec.get("lot") or _lot_from_name(os.path.basename(p))
            if not lot:
                note(item, p, "제조번호를 읽지 못함"); continue
            rec["export"] = "수출용" in os.path.basename(p)
            data.coa.setdefault(lot, {})[key] = rec
    # 9.2.1 에서 읽은 조제 값은 9.2.2 에 없는 것만 보탠다 — 같은 조제 단계라 한 곳에서 보면 된다.
    for recs in data.coa.values():
        if recs.get("921"):
            recs["922"] = dict(recs["921"], **(recs.get("922") or {}))
    # Lot 목록: 성적서(조제)의 제조일자 > 제조내역
    lots = {}
    for lot, recs in data.coa.items():
        export = any(r.get("export") for r in recs.values())
        mfg = (recs.get("922") or recs.get("923") or {}).get("mfg_date")
        lots[lot] = ((mfg or "").replace("/", "."), export)
    for lot, _, mfg, _ in data.manufacturing:
        lots.setdefault(lot, (mfg, True))
    data.lots = sorted([(l, d, e) for l, (d, e) in lots.items()], key=lambda x: (x[2], x[1], x[0]))
    # 3. 허가증
    for p in got.get("3", []):
        if p.lower().endswith(".pdf"):
            try:
                # 스캔 허가증은 문의로 올리지 않는다 — 값은 결재본 것이 맞고, 해마다 같은 글이
                # 목록만 채운다 (담당자 2026-09-07: "허가증은 빼자, 의미가 없네").
                if not is_scanned(p):
                    data.license = license_reader.read_license(p)
                else:
                    log("  [3] %s — 글자 없는 스캔본이라 허가 정보는 결재본 값을 그대로 씁니다"
                        % os.path.basename(p))
            except PdfTextError as e:
                note("3", p, str(e))
    # 10.x 마스터
    for p in got.get("10.1", []):
        if p.lower().endswith(".xlsx"):
            try:
                from .. import master as master_module
                data.pv = master_module.read_pv_master(p, product_name or "")
            except Exception as e:
                note("10.1", p, "PV 마스터 판독 실패: %s" % e)
            try:                                   # 수출용은 마스터에 '(수출용)' 이름으로 따로 있다
                data.pv_exp = master_module.read_pv_master(p, (product_name or "") + "(수출용)")
            except Exception:
                data.pv_exp = None
    # 마스터파일이 여럿이면(‘…_OLD’ 와 최신본, IQ·OQ 파일과 PQ 파일) 옛것에 덮이지 않게 합친다.
    # 디겐타안연고 2026 폴더에서 실제로 옛 파일이 최신본을 덮어 PQ 가 2022년 것으로 나왔다.
    def _newest_first(paths):
        """새 마스터파일 먼저, 옛것(_OLD·구버전)은 뒤에 — 버리지는 않는다.

        옛 파일을 버리면 IQ·OQ 가 통째로 빈다. 디겐타 2026 폴더가 그렇다: 최신본
        (rev.033)에는 PQ 만 있고 IQ·OQ 는 '…(Rev.32)_OLD(IQ, OQ 정보 확인)' 에만 있다.
        먼저 온 파일의 값이 이기므로 새것이 옛것에 덮이지 않는다.
        """
        xlsx = [p_ for p_ in paths if p_.lower().endswith(".xlsx")]
        def 옛것(p_):
            base = os.path.basename(p_)
            return "old" in base.lower() or "구버전" in base
        by_time = lambda ps: sorted(ps, key=lambda p_: os.path.getmtime(p_), reverse=True)
        return by_time([p_ for p_ in xlsx if not 옛것(p_)]) + by_time([p_ for p_ in xlsx if 옛것(p_)])

    for p in _newest_first(got.get("10.2", [])):
        try:
            for key, entry in masters.equipment_docs(p).items():
                have = data.equipment.setdefault(key, entry)
                if have is entry:
                    continue
                # 같은 장비가 여러 파일에 있으면 문서를 합친다 — 새 파일에 없는 IQ·OQ 를
                # 옛 파일에서 이어받되, 이미 있는 문서번호는 새 파일 것을 그대로 둔다.
                본 = {d for d, _ in have["docs"]}
                have["docs"].extend([(d, dt) for d, dt in entry["docs"] if d not in 본])
                for key2 in ("name", "line"):
                    if not have.get(key2):
                        have[key2] = entry.get(key2, "")
        except Exception as e:
            note("10.2", p, str(e))
    # 10.3~10.5 는 파일이 두 벌이다 — IQ·OQ 는 '제조지원 설비 마스터파일(IQ, OQ 정보확인)',
    # PQ 는 '10.5 Qualification Master File …(PQ, 측정위치 타당성 정보 확인)' 에 있다
    # (담당자 2026-09). 종류마다 제 파일을 먼저 보고, 거기 없을 때만 다른 파일에서 받는다.
    def _kind_of_file(path):
        base = os.path.basename(path).upper()
        if "IQ" in base and "OQ" in base:
            return "IQOQ"
        return "PQ" if "PQ" in base else ""

    support_files = _newest_first(got.get("10.3-5", []) + got.get("10.3", []))
    읽음 = {}
    for p in support_files:
        try:
            읽음[p] = masters.support_docs(p)
        except Exception as e:
            note("10.3-5", p, str(e))

    def _entry(path, key):
        """PQ 마스터의 보고서는 QM… 으로도 적힌다(정기 성능적격성평가) — 모두 PQ 로 본다."""
        entry = 읽음[path][key]
        if _kind_of_file(path) != "PQ":
            return entry
        모두 = list(entry.get("docs") or
                    [d for kind in ("DQ", "IQ", "OQ", "PQ") for d in entry.get(kind, [])])
        return dict(entry, DQ=[], IQ=[], OQ=[], PQ=모두,
                    why=dict(entry.get("why") or {}, PQ=[""] * len(모두)))

    for kind, 제파일 in (("DQ", "IQOQ"), ("IQ", "IQOQ"), ("OQ", "IQOQ"), ("PQ", "PQ")):
        차례 = ([p for p in support_files if _kind_of_file(p) == 제파일]
                + [p for p in support_files if _kind_of_file(p) == ""]
                + [p for p in support_files if _kind_of_file(p) not in (제파일, "")])
        for p in 차례:
            for key in 읽음.get(p, {}):
                entry = _entry(p, key)
                have = data.support.setdefault(
                    key, {"name": entry.get("name", ""), "system": entry.get("system", ""),
                          "raw_id": entry.get("raw_id", key), "why": {}})
                if not have.get(kind) and entry.get(kind):
                    have[kind] = entry[kind]
                    have.setdefault("why", {})[kind] = (entry.get("why") or {}).get(kind, [])
                for 칸 in ("name", "system"):
                    if not have.get(칸):
                        have[칸] = entry.get(칸, "")
    # 11 · 12
    for p in got.get("11", []):
        if p.lower().endswith(".pdf"):
            got_one = None
            try:
                got_one = deviation.read_deviation(p)
            except PdfTextError as e:
                note("11", p, str(e))
            # 글자 없는 스캔 일탈 보고서도 12항 변경요청서와 같은 길로 읽는다
            # (담당자 2026-09-08: "다음 PQR 작성할 때 문제없도록 조치해 줘").
            비었나 = got_one is None or not (got_one.get("title") or got_one.get("doc_no"))
            if 비었나 and is_scanned(p):
                읽음 = _deviation_by_claude(p, folder, log, note)
                if 읽음 is not None:
                    got_one = 읽음
            if got_one is not None:
                data.deviations.append(got_one)
    # 12항에 올린 변경요청서는 하나도 빠뜨리지 않는다 — 읽지 못한 것도 파일 이름의 문서번호로 줄을
    # 세운다 (담당자 2026-09-06: "12항 변경관리가 5개인데 1개만 표시되어 있네").
    for p in got.get("12", []):
        if not p.lower().endswith(".pdf"):
            continue
        try:
            cc = change.read_change(p)
            m = re.search(r"(CC-\d{6}-\d{2})", os.path.basename(p))
            if not cc.get("doc_no"):                 # 문서번호를 못 읽으면 파일 이름의 것 — '[None]' 이 나가면 안 된다
                cc["doc_no"] = m.group(1) if m else os.path.splitext(os.path.basename(p))[0]
            if not (cc.get("title") or "").strip() and not cc.get("actions"):
                # 글자가 없는 스캔본이다 — 시험일지와 같은 길로 Claude 에게 읽힌다
                # (담당자 2026-09-07: "변경요청서 읽으면 돼 PDF 라서 못 읽는 거야?" — PDF 라서가 아니다).
                읽음 = _change_by_claude(p, folder, log, note)
                if 읽음 is not None:
                    읽음.setdefault("doc_no", cc.get("doc_no"))
                    cc = 읽음
                else:
                    cc["unread"] = True              # 못 읽은 것과 같다 — 노랑으로
                    note("12", p, "변경요청서에서 변경사항·조치사항을 읽지 못해 문서번호만 적었습니다 — 직접 채우세요")
            data.changes.append(cc)
        except Exception as e:                       # 서식이 아주 다르거나 파일이 깨진 것
            m = re.search(r"(CC-\d{6}-\d{2})", os.path.basename(p))
            읽음 = _change_by_claude(p, folder, log, note)      # 여기서도 Claude 에게 맡긴다
            if 읽음 is not None:
                if m and not 읽음.get("doc_no"):
                    읽음["doc_no"] = m.group(1)
                data.changes.append(읽음)
                continue
            data.changes.append({"doc_no": m.group(1) if m else os.path.splitext(os.path.basename(p))[0],
                                 "title": "", "unread": True, "actions": [], "products": ""})
            note("12", p, "변경요청서를 읽지 못해 문서번호만 적었습니다 — 변경사항·조치사항을 직접 채우세요 (%s)" % e)
    seen_cc = set()                                  # 같은 문서번호가 두 번 올라온 것은 한 줄만
    only = []
    for cc in data.changes:
        key = (cc.get("doc_no") or "").strip()
        if key and key in seen_cc:
            continue
        seen_cc.add(key)
        only.append(cc)
    data.changes = only
    if data.changes:
        log("  [12] 변경요청서 %d건: %s" % (len(data.changes),
                                        ", ".join((c.get("doc_no") or "?") + ("(못 읽음)" if c.get("unread") else "")
                                                  for c in data.changes)))
    # 13 안정성 — 스캔이면 비전 판독 대상. 판독 결과(.json)가 있으면 그것을 먼저 쓴다.
    for p in got.get("13", []):
        if p.lower().endswith(".pdf"):
            data.stability_files.append((p, is_scanned(p)))
        elif p.lower().endswith(".json"):
            try:
                with open(p, encoding="utf-8") as fh:
                    got_json = json.load(fh)
                data.stability_logs += got_json if isinstance(got_json, list) else got_json.get("logs", [])
                log("  [13] %s — 안정성 시험일지 판독 결과 %d Lot" % (os.path.basename(p), len(data.stability_logs)))
                data.stability_cache = (p, 0 if isinstance(got_json, list) else got_json.get("reader_version", 0))
                # 사람이(또는 이 대화의 Claude 가) 13 폴더 전체를 읽어 만든 판독 파일이면 PDF 를 더 읽지 않는다
                # (담당자 2026-09-07: "API 키를 발급받지 않고 Claude 로 확인할 수는 없나?")
                if not isinstance(got_json, list) and got_json.get("covers_all"):
                    data.stability_all_read = True
                    log("  [13] 이 판독 파일이 13항 자료 전부라고 적혀 있어(covers_all) 시험일지를 더 읽지 않습니다")
            except Exception as error:
                note("13", p, "안정성 판독 파일을 읽지 못했습니다 — %s" % error)
    # 판독 파일(.json)도 없고 Claude 비전(API 키)도 없으면 PC 에서 오프라인으로 손글씨를 읽는다
    # (담당자 2026-09: "API 안 하고 json 없이 최대한 판독, 애매한 것만 노랑"). 깨끗이 읽힌 값만
    # 쓰고 나머지는 '애매' 로 남겨 보고서가 노랑·주황으로 칠한다. 결과는 판독 파일로 저장해
    # 다음부터는 그것을 쓴다 — 담당자가 값을 손보면 그 값이 우선이다.
    scanned = same_files_once([p for p, is_scan in data.stability_files if is_scan], log)
    cache = getattr(data, "stability_cache", None)
    # 새로 읽을 일지가 있으면 옛 판독 파일 보완은 건너뛴다 — 어차피 새로 읽으면서 채워진다.
    # (담당자 PC 2026-09-07: 24장을 보완으로 한 번, 새로 한 번 — 두 번 읽어 한 시간이 넘었다)
    보완할것 = [] if getattr(data, "stability_all_read", False) else unread_scans(data.stability_logs, scanned)
    if scanned and not 보완할것 and data.stability_logs and cache and cache[0].endswith(handwriting_cache_name()):
        # 옛 판독기가 만든 판독 파일 — 애매한 칸만 새 판독기로 다시 읽어 예상값을 채운다
        # (담당자 2026-09-06: "주황색 부분에 너의 예상값을 기재하지 않았어" — 업데이트 전 판독 파일을 그대로 썼다)
        try:
            from . import handwriting
            specs = {}
            for recs in data.coa.values():
                for a in (recs.get("924") or {}).get("assays") or []:
                    try:
                        if a.get("part") and a["part"] not in specs:
                            specs[a["part"]] = (float(a["lo"]), float(a["hi"]))
                    except (TypeError, ValueError, KeyError):
                        pass
            if handwriting.refresh_cache(data.stability_logs, cache[1], scanned, specs or None, log):
                handwriting.save_cache(folder, data.stability_logs, log)
        except Exception as error:
            note("13", "", "옛 판독 파일을 새 판독기로 보완하지 못했습니다 — %s" % error)
    # 판독 파일이 있어도 거기에 없는 시험일지(뒤에 올린 장기·수출용 일지)는 읽는다 — 판독 파일은 읽은 일지의
    # 이름(source)을 들고 있다. 담당자 2026-09-06: "시험 기간은 안정성 시험일지 보면 확인 가능하잖아" —
    # 시판 후 일지만 읽은 판독 파일이 있어 장기 일지가 통째로 안 읽혔다.
    if getattr(data, "stability_all_read", False):
        scanned = []
    elif data.stability_logs:
        scanned = unread_scans(data.stability_logs, scanned)
        if scanned:
            log("  [13] 판독 파일에 없는 시험일지 %d장을 새로 읽습니다: %s"
                % (len(scanned), ", ".join(os.path.basename(p) for p in scanned)))
    if scanned:
        # 성분 규격 — 판독한 함량 값에 성분 이름을 붙이는 데 쓴다
        specs = {}
        for recs in data.coa.values():
            for a in (recs.get("924") or {}).get("assays") or []:
                try:
                    if a.get("part") and a["part"] not in specs:
                        specs[a["part"]] = (float(a["lo"]), float(a["hi"]))
                except (TypeError, ValueError, KeyError):
                    pass
        from . import claude_cli, handreq, handwriting
        try:
            from . import vision as vision_mod
            api_on = vision_mod.available()
        except Exception:
            api_on = False
        pc_on = handreq.pc_reading_on(folder)
        logs = []
        # ① API 키가 있으면 Claude 가 곧바로 읽는다 (한 장에 몇 초)
        if api_on:
            try:
                from . import vision_claude
                log("  [13] 손글씨 시험일지 %d장을 Claude 로 판독합니다" % len(scanned))
                logs = vision_claude.read_logs(scanned, specs or None, log)
            except Exception as error:
                data.issues.insert(0, ("13", "", "★ Claude(API 키) 판독에 실패해 13항을 채우지 못했습니다 — %s" % error))
                logs = []
        # ②' 키가 없어도 이 PC 에 Claude Code 가 깔려 있으면 그것으로 읽는다 — 담당자가 따로
        #    물어볼 것 없이 '보고서 작성' 한 번이면 된다 (담당자 2026-09-07: "이것을 자동화하면 안 돼?")
        if not logs and not api_on and claude_cli.available():
            try:
                log("  [13] 손글씨 시험일지 %d장을 이 PC 의 Claude Code 로 판독합니다 (몇 분 걸립니다)" % len(scanned))
                logs = claude_cli.read_logs(scanned, specs or None, log, folder)
            except Exception as error:
                data.issues.insert(0, ("13", "", "★ 이 PC 의 Claude Code 로 13항을 읽지 못했습니다 — %s" % error))
                logs = []
        # ② 키도 Claude Code 도 없으면 이 PC 로 읽지 않고, 대화에 올릴 판독 요청 묶음을 만든다.
        #    이 PC 의 판독기는 한 장에 몇 분이라 20장이면 한 시간을 넘긴다
        #    (담당자 2026-09-07: "PC 판독하지 말고 Claude 에서 확인해 주도록 해").
        #    그래도 이 PC 로 읽고 싶으면 제품 폴더에 'PC 판독 사용.txt' 를 두면 된다.
        if not logs and not pc_on:
            made = handreq.make_request(folder, scanned, product_name or "", log)
            묶음 = ", ".join(name for name, _ in made) or "만들지 못함"
            data.issues.insert(0, ("13", 묶음,
                                   "★ 손글씨 시험일지 %d장을 아직 읽지 못했습니다 — 13항이 '확인 필요' 로 남습니다. "
                                   "이 PC 에서 Cowork·Claude Code 를 쓰신다면 제품 폴더를 열고 "
                                   "'13 폴더의 안정성 시험일지를 읽어 13. 안정성시험일지 판독.json 을 만들어 줘' 라고 "
                                   "시키면 그 자리에서 만들어집니다. 웹 대화라면 제품 폴더의 '%s' 를 올리고 받은 json 을 "
                                   "그 폴더에 둔 뒤 '보고서 재작성' 을 누르세요. (이 PC 가 직접 읽게 하려면 '%s' 를 두세요)"
                                   % (len(scanned), 묶음, handreq.PC_OPT_IN)))
        # ③ 담당자가 이 PC 로 읽으라고 정해 두었으면 오프라인으로 읽는다
        if not logs and pc_on and handwriting.available():
            log("  [13] 손글씨 시험일지 %d장을 이 PC 로 판독합니다 ('%s' 가 있어서 — 한 장에 몇 분)"
                % (len(scanned), handreq.PC_OPT_IN))
            logs = handwriting.read_folder(scanned, specs or None, log)
        if logs:
            try:                                            # 지난 경향표가 있으면 그 값이 우선
                # (trend_reader 는 파일 맨 위에서 가져온 것 — 여기서 다시 import 하면 함수 전체에서
                #  지역 변수가 되어 아래 경향표 읽기가 'cannot access local variable' 로 넘어졌다,
                #  담당자 PC 작성 기록 2026-09-06 08:14)
                sheets = []
                for tp in got.get("16", []) + got.get("13", []) + got.get("첨부", []):
                    if (tp.lower().endswith(".xlsx") and not os.path.basename(tp).startswith("~$")
                            and trend_reader.is_trend_file(tp)):
                        try:
                            sheets += trend_reader.read_trend(tp)
                        except Exception:
                            pass
                handwriting.merge_known(logs, sheets, log)
            except Exception:
                pass
            # 같은 Lot 이 두 줄로 실리지 않게 합친다 — 이미 읽은 값(담당자가 고친 판독 파일)이 우선
            handwriting.merge_logs(data.stability_logs, logs, log)
            data.stability_logs.sort(key=lambda r: (r.get("year") or "", r.get("lot") or ""))
            겹침 = {}
            for one in data.stability_logs:
                # 구분이 없는 옛 판독 기록도 있다 — 글자로 맞춰 둔다(없으면 '' )
                겹침.setdefault(str(one.get("lot") or ""), set()).add(str(one.get("kind") or ""))
            for lot, kinds in sorted(겹침.items()):
                kinds = {k for k in kinds if k}
                if len(kinds) > 1:     # 같은 Lot 이 장기·시판 후 두 표에 들어간다 — 파일 이름을 확인해야 한다
                    note("13", lot, "같은 Lot 이 %s 두 가지로 읽혔습니다 — 시험일지 파일 이름에 "
                                    "'장기'·'시판 후' 를 바로 적어 주세요" % "·".join(sorted(kinds)))
            shaky = sum(len(p.get("unsure") or []) for one in logs for p in one["points"])
            어떻게 = "Claude(API 키)" if api_on else ("이 PC 의 Claude Code" if not pc_on else "이 PC 의 판독기")
            note("13", "", "손글씨 시험일지 %d장을 %s 로 판독했습니다 — 애매한 칸 %d개는 "
                           "노랑(워드)·주황(엑셀)으로 표시했으니 시험일지와 대조하세요"
                 % (len(logs), 어떻게, shaky))
            handwriting.save_cache(folder, data.stability_logs, log)   # 판독 파일에 새 일지를 보탠다
        elif pc_on and not handwriting.available():
            note("13", "", "이 PC 로 읽으라고 되어 있는데 판독기가 없습니다 — PQR-업데이트.bat 을 실행하거나 "
                           "'%s' 를 지우고 Claude 판독 묶음을 쓰세요" % handreq.PC_OPT_IN)
    # 담당자가 손으로 옮겨 적어 둔 값이라 스캔 판독보다 믿을 만하다.
    for item in ("13", "16", "첨부"):
        for p in got.get(item, []):
            if not p.lower().endswith(".xlsx") or os.path.basename(p).startswith("~$"):
                continue
            try:
                if trend_reader.is_trend_file(p):
                    sheets = trend_reader.read_trend(p)
                    if sheets:
                        data.stability_trend = sheets
                        log("  [%s] %s — 지난 경향표 %d 시트를 이어받습니다"
                            % (item, os.path.basename(p), len(sheets)))
            except Exception as error:
                note(item, p, "경향표를 읽지 못했습니다 — %s" % error)
    # 16 전년도 결재본
    for p in got.get("16", []):
        if p.lower().endswith((".doc", ".docx")):
            data.previous_report = p
    data.deviations.sort(key=lambda d: d.get("doc_no") or "")
    # 읽지 못한 파일은 반드시 알린다 — 조용히 지나가면 그 항이 왜 비었는지 아무도 모른다
    for item, path, why in unread_files(got, log):
        note(item, path, why)
    for item, why in empty_after_read(data, got, log):
        data.issues.insert(0, (item, "", why))
    return data


# 항마다 '읽혔다면 여기에 값이 있어야 한다' — 자료는 올렸는데 값이 없으면 그 항은 빈 채로 나간다.
# 담당자 2026-09-08: "다른 제품 작성할 때 동일한 문제가 발생 안 되도록 조치해 줘."
FILLED_BY = (
    ("6",       lambda d: d.manufacturing or d.batch,        "제조내역·공 기록서"),
    ("7",       lambda d: d.yields,                          "수율"),
    ("8.1.1",   lambda d: d.suppliers_raw,                   "주원료 공급업체"),
    ("8.1.3",   lambda d: d.suppliers_mat,                   "부원료·포장자재 공급업체"),
    ("8.2.1",   lambda d: d.raw_tests,                       "원료 시험번호"),
    ("8.2.2",   lambda d: d.pkg_tests,                       "자재 시험번호"),
    ("9.2.1",   lambda d: any(r.get("921") for r in d.coa.values()), "조제 성적서"),
    ("9.2.2",   lambda d: any(r.get("922") for r in d.coa.values()), "조제(바이오버든) 성적서"),
    ("9.2.3",   lambda d: any(r.get("923") for r in d.coa.values()), "충전 성적서"),
    ("9.2.4",   lambda d: any(r.get("924") for r in d.coa.values()), "완제 성적서"),
    ("10.1",    lambda d: d.pv,                              "공정밸리데이션"),
    ("10.2",    lambda d: d.equipment,                       "제조설비 적격성"),
    ("10.3-5",  lambda d: d.support,                         "제조지원 설비 적격성"),
    ("11",      lambda d: d.deviations,                      "일탈"),
    ("12",      lambda d: [c for c in d.changes if not c.get("unread")], "변경관리"),
    ("13",      lambda d: d.stability_logs or d.stability_files, "안정성 시험일지"),
)


def empty_after_read(data, got, log=None):
    """자료는 올렸는데 한 칸도 읽지 못한 항 — [(항, 알림 글)].

    형식이 맞는 파일이 있는데도 값이 비면 그 항은 보고서에서 사선으로 남는다. 조용히 넘어가면
    담당자가 다 만들고 나서야 알게 되므로, 작성이 끝나기 전에 문의 목록 맨 앞에 올린다.
    """
    rows = []
    for item, 채워졌나, 무엇 in FILLED_BY:
        paths = [p for p in got.get(item, [])
                 if not os.path.basename(p).startswith(("~$", ".", "PQR "))
                 and not os.path.basename(p).lower().endswith(".txt")]
        if not paths:
            continue                                  # 안 올린 항은 여기서 다루지 않는다
        try:
            if 채워졌나(data):
                continue
        except Exception:
            continue
        rows.append((item, "★ %s항에 파일 %d개를 올리셨는데 %s 값을 한 칸도 읽지 못했습니다 — "
                           "그 항은 빈 칸으로 남습니다. 파일을 열어 서식이 평소와 같은지 보시고, "
                           "그대로면 이 알림과 파일을 제작자에게 보내 주세요 (%s)"
                     % (item, len(paths), 무엇,
                        ", ".join(os.path.basename(p) for p in paths[:3])
                        + (" 외 %d개" % (len(paths) - 3) if len(paths) > 3 else ""))))
    if log and rows:
        log("  [문제] 자료는 있는데 값을 못 읽은 항: %s" % ", ".join(item for item, _ in rows))
    return rows
