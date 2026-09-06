# -*- coding: utf-8 -*-
"""손글씨 안정성 시험일지(HLF-QC-104-01) 오프라인 판독 — API 키 없이 PC 에서.

담당자 2026-09: "API 안 하고 json 없이 손글씨를 최대한 판독해서 안정성 내용을 작성하고,
애매한 것만 노랑마크로". 글자 인식은 pip 로 깔리는 RapidOCR(onnxruntime, 모델 내장) 을
쓴다 — 인터넷·API 키·별도 설치가 없다. 손글씨 숫자는 절반쯤만 깨끗이 읽히므로:

  * 인쇄된 글(제조번호·제조일자·초기 일자·규격)은 확실하다 — 표의 자리를 잡는 기준으로 쓴다.
  * 손글씨 값은 깨끗이 읽힌 것(숫자·소수점만 남고, 확신 ≥ 0.8, 규격 근처)만 값으로 쓴다.
  * 그 밖은 '애매' 로 표시한다 — 읽은 대로의 값이 있으면 그 값과 함께, 없으면 값 없이.
    보고서는 애매한 칸을 노랑(워드)·주황(엑셀)으로 칠하고 문의 목록에 적는다.
  * 전년도 경향표에 이미 있는 시점은 판독하지 않고 그 값을 쓴다(collect 에서 합친다).

판독 결과는 다른 판독기(Claude 비전)와 같은 꼴이다:
  {"lot", "year", "pack", "store", "why": "", "points": [{"period", "done", "assays": {성분: 값},
   "unsure": [성분 | "done", …]}, …], "notes": [...]}
"""
from __future__ import unicode_literals

import os
import re

PERIODS = ["3M", "6M", "9M", "12M", "18M", "24M", "36M", "48M"]   # 초기 다음 칸들의 표준 차례
DPI = 200


_WHY = ""


def available():
    """판독기를 쓸 수 있는가. 못 쓰면 why() 가 까닭(빠진 패키지·오류)을 돌려준다."""
    global _WHY
    try:
        import numpy  # noqa: F401
        import pypdfium2  # noqa: F401
        import onnxruntime  # noqa: F401 — 새 rapidocr 는 엔진을 따로 깔아야 한다
        _engine_class()
        _WHY = ""
        return True
    except Exception as error:
        _WHY = "%s: %s" % (type(error).__name__, error)
        return False


def _engine_class():
    """RapidOCR 클래스 — 옛 이름(rapidocr_onnxruntime)과 새 이름(rapidocr) 어느 쪽이든."""
    try:
        from rapidocr_onnxruntime import RapidOCR
        return RapidOCR
    except ImportError:
        from rapidocr import RapidOCR
        return RapidOCR


def why():
    return _WHY


_OCR = None


def _engine():
    global _OCR
    if _OCR is None:
        _OCR = _engine_class()()
    return _OCR


def render(pdf_path, dpi=DPI):
    import pypdfium2 as pdfium
    pdf = pdfium.PdfDocument(pdf_path)
    page = pdf[0]
    return page.render(scale=dpi / 72.0).to_pil().convert("RGB")


def ocr_image(image):
    """[(x0, y0, x1, y1, text, conf), …]"""
    import numpy as np
    got = _engine()(np.array(image))
    if isinstance(got, tuple):                                     # rapidocr_onnxruntime: (result, elapse)
        result = got[0] or []
    elif hasattr(got, "boxes"):                                    # rapidocr 2.x: 결과 객체(numpy 배열)
        def _l(v):
            return [] if v is None else list(v)
        result = list(zip(_l(got.boxes), _l(got.txts), _l(got.scores)))
    else:
        result = [] if got is None else list(got)
    out = []
    for box, text, conf in result:
        xs = [p[0] for p in box]
        ys = [p[1] for p in box]
        out.append((min(xs), min(ys), max(xs), max(ys), text, float(conf)))
    return out


# ---------- 표 자리 잡기 ----------
LOT = re.compile(r"^[A-Z]{3}[A-Z0-9]{3}$")
DATE = re.compile(r"(20\d{2})[.\-/ ]?\s?(\d{2})[.\-/ ]?\s?(\d{2})")
SPEC = re.compile(r"(\d{2,3}\.\d)\s*[~～\-]\s*(\d{2,3}\.\d)\s*%")


def _center(b):
    return ((b[0] + b[2]) / 2.0, (b[1] + b[3]) / 2.0)


def column_lines(image, y_center, half=12):
    """y_center 근처 가로 띠에서 세로 표선의 x 자리들 — 어두운 픽셀이 띠 높이의 대부분을 차지하는 x."""
    import numpy as np
    arr = np.asarray(image.convert("L"))
    y0, y1 = max(0, int(y_center - half)), min(arr.shape[0], int(y_center + half))
    band = arr[y0:y1, :] < 200                                    # 스캔 표선은 한 픽셀 폭의 옅은 회색
    # 살짝 기운 스캔이면 선이 한두 픽셀 옆으로 흐른다 — 이웃 3픽셀 중 하나라도 어두우면 친다
    wide = band.copy()
    wide[:, 1:] |= band[:, :-1]
    wide[:, :-1] |= band[:, 1:]
    dark = wide.sum(axis=0)
    hit = dark >= (y1 - y0) * 0.85
    xs, run = [], []
    for x, on in enumerate(hit):
        if on:
            run.append(x)
        elif run:
            xs.append(sum(run) / float(len(run)))
            run = []
    if run:
        xs.append(sum(run) / float(len(run)))
    merged = []
    for x in xs:
        if merged and x - merged[-1] < 8:
            merged[-1] = (merged[-1] + x) / 2.0
        else:
            merged.append(x)
    return merged


def table_columns(image, y_top, y_bottom, step=40, min_votes=0.5):
    """표 전체 높이에서 여러 띠를 보고, 대부분의 띠에 나타나는 x 만 세로 표선으로 친다.

    손글씨의 긴 사선(시험 안 한 칸을 지우는 빗금)도 한 띠에서는 선처럼 보이지만 표 높이
    전체에 걸쳐 같은 x 에 있지는 않다.
    """
    votes = []
    ys = list(range(int(y_top), int(y_bottom), step))
    for y in ys:
        for x in column_lines(image, y, half=12):
            for v in votes:
                if abs(v[0] - x) < 8:
                    v[0] = (v[0] * v[1] + x) / (v[1] + 1)
                    v[1] += 1
                    break
            else:
                votes.append([x, 1])
    need = max(2, int(len(ys) * min_votes))
    return sorted(x for x, n in votes if n >= need)


def _row_of(boxes, pattern, prefer=None):
    hits = [b for b in boxes if pattern.search(b[4])]
    if not hits:
        return None
    if prefer is not None:
        hits.sort(key=lambda b: abs(_center(b)[1] - prefer))
    return hits[0]


# ---------- 값 읽기 ----------
_FIX = {"O": "0", "o": "0", "D": "0", "Q": "0", "l": "1", "I": "1", "|": "1", "[": "1", "(": "1",
        "S": "5", "s": "5", "q": "9", "g": "9", "b": "6", "B": "8", "Z": "2", "z": "2", "T": "7",
        "，": ",", "。": ".", "·": ".", "'": "."}


# 예상값 추정 — 깨끗이 못 읽은 손글씨에서 그럴듯한 값을 만든다. 담당자 2026-09-06: "주황색 부분에
# 너의 예상값을 기재" — 여기서 나온 값은 반드시 '애매' 로 표시된다.
_GUESS = dict(_FIX, **{"a": "9", "A": "4", "n": "7", "y": "%", "Y": "%", "G": "6", "t": "7", "e": "9"})
READER_VERSION = 2          # 판독 방식이 바뀌면 올린다 — 옛 판독 파일은 애매한 칸만 다시 읽는다


def guess_assay(text, lo=None, hi=None):
    """읽기 애매한 글('9a.4-1.', 'qn.ay', '1013')에서 규격 근처의 값을 추정한다. 못 하면 None."""
    raw = (text or "").strip()
    if not raw:
        return None
    fixed = "".join(_GUESS.get(ch, ch) for ch in raw)
    fixed = re.sub(r"%.*$", "", fixed)                           # % 뒤는 흘린 획
    cands = []
    for m in re.finditer(r"(\d{2,3})\s*[.,\-]\s*(\d)", fixed):
        cands.append(float("%s.%s" % (m.group(1), m.group(2))))
    digits = re.sub(r"\D", "", fixed)
    if len(digits) in (3, 4):                                   # '987' → 98.7, '1013' → 101.3
        cands.append(float(digits[:-1] + "." + digits[-1]))
    for v in cands:
        if lo is None or hi is None or (lo - 5) <= v <= (hi + 5):
            return v
    return None


def parse_assay(text, lo=None, hi=None):
    """손글씨 함량 글 → (값, 깨끗한가). 값을 못 찾으면 (None, False)."""
    raw = (text or "").strip()
    if not raw:
        return None, False
    core = re.sub(r"[%％/\\,.。'’`·\s\-]+$", "", raw)          # 뒤에 붙은 % 나 흘린 획
    clean_match = re.match(r"^\d{2,3}[.,]\d$", core)
    fixed = "".join(_FIX.get(ch, ch) for ch in core)
    m = re.search(r"(\d{2,3})\s*[.,\-]\s*(\d)(?!\d)", fixed)
    if not m:
        m = re.search(r"(\d{2,3})\s*[.,\-]\s*(\d)", fixed)         # '96-57' — 뒤의 7 은 흘린 %
    if not m:
        return None, False
    value = float("%s.%s" % (m.group(1), m.group(2)))
    if lo is not None and hi is not None and not ((lo - 5) <= value <= (hi + 5)):
        return None, False                                          # 규격에서 한참 벗어나면 오독이다
    return value, bool(clean_match)


def parse_date(text, year_lo=None, year_hi=None):
    """손글씨 날짜 → ('YYYY.MM.DD', 깨끗한가)."""
    raw = (text or "").strip()
    digits = re.sub(r"\D", "", "".join(_FIX.get(ch, ch) for ch in raw))
    m = DATE.search(raw)
    clean = bool(m)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    elif len(digits) == 8:
        # 손글씨 '2025' 는 '2525'·'1225'·'202S' 로 읽히곤 한다 — 뒤 두 자리로 연도를 잡는다
        y, mo, d = 2000 + int(digits[2:4]), int(digits[4:6]), int(digits[6:8])
        if year_lo and not (year_lo - 1 <= y <= (year_hi or 9999) + 1):
            return None, False
        clean = False
    else:
        return None, False
    if not (1 <= mo <= 12 and 1 <= d <= 31):
        return None, False
    if year_lo and y < year_lo or year_hi and y > year_hi:
        clean = False
        if year_lo and (y < year_lo - 1 or y > (year_hi or 9999) + 1):
            return None, False
    return "%04d.%02d.%02d" % (y, mo, d), clean


def _inside_frac(b, x0, x1):
    """글상자의 가로 폭 가운데 칸 [x0, x1] 안에 든 비율."""
    w = max(1.0, b[2] - b[0])
    return max(0.0, min(b[2], x1) - max(b[0], x0)) / w


def _in_cell(boxes, x0, x1, y0, y1, strict=True):
    """칸 안의 글상자. strict 면 가로로 칸을 크게 벗어난(옆 칸까지 걸친) 상자는 뺀다 —
    '2025.7.02 2025.09.10' 처럼 두 칸이 한 상자로 읽히면 옆 칸 값이 섞이기 때문."""
    out = []
    for b in boxes:
        cx, cy = _center(b)
        if x0 <= cx <= x1 and y0 <= cy <= y1:
            if strict and _inside_frac(b, x0, x1) < 0.75:
                continue
            out.append(b)
    out.sort(key=lambda b: b[0])
    return out


def _spilled(boxes, x0, x1, y0, y1):
    """세로로는 칸 안이지만 가로로 옆 칸까지 걸친 글상자(칸이 비어 있지 않다는 증거)."""
    out = []
    for b in boxes:
        cy = _center(b)[1]
        if y0 <= cy <= y1 and _inside_frac(b, x0, x1) >= 0.2 and _inside_frac(b, x0, x1) < 0.75:
            out.append(b)
    return out


def _crop_ocr(image, x0, x1, y0, y1, pad=6):
    w, h = image.size
    crop = image.crop((max(0, int(x0 - pad)), max(0, int(y0 - pad)), min(w, int(x1 + pad)), min(h, int(y1 + pad))))
    if crop.size[0] < 20 or crop.size[1] < 20:
        return []
    scale = 2.0                                                    # 작은 칸은 키워서 읽는다
    crop = crop.resize((int(crop.size[0] * scale), int(crop.size[1] * scale)))
    return [b for b in ocr_image(crop)]


def read_log(pdf_path, specs=None, log=None):
    """시험일지 한 장 → 판독 결과 dict (실패하면 None).

    specs: {성분 이름: (하한, 상한)} — 값의 그럴듯함을 따질 때 쓴다. 없으면 시험일지의 인쇄된
    규격('90.0 ~ 110.0%')에서 읽는다.
    """
    def say(msg):
        if log:
            log("    " + msg)
    image = render(pdf_path)
    boxes = ocr_image(image)
    notes = []
    W, H = image.size
    s = DPI / 200.0                                                 # 200 dpi 기준 치수

    lot_box = next((b for b in boxes if LOT.match(b[4].strip().upper()) and b[1] < H * 0.2), None)
    lot = lot_box[4].strip().upper() if lot_box else None
    if not lot:
        m = re.search(r"_([A-Z]{3}[A-Z0-9]{3})_", os.path.basename(pdf_path))
        lot = m.group(1) if m else None
    if not lot:
        say("제조번호를 찾지 못함: %s" % os.path.basename(pdf_path))
        return None
    dates_printed = [b for b in boxes if DATE.search(b[4]) and b[5] >= 0.95 and b[1] < H * 0.3]
    mfg = None
    for b in sorted(dates_printed, key=lambda b: b[1]):
        if b[0] > W * 0.75:                                        # 오른쪽 위 '제조일자'
            mfg = DATE.search(b[4]); break
    mfg_year = int(mfg.group(1)) if mfg else None

    # 규격 줄(인쇄) — 함량 행의 자리와 성분 차례
    spec_boxes = sorted([b for b in boxes if SPEC.search(b[4]) and b[0] < W * 0.45], key=lambda b: b[1])
    if len(spec_boxes) < 1:
        say("인쇄된 함량 규격을 찾지 못함 — %s" % os.path.basename(pdf_path))
        return None
    parts = list(specs.keys()) if specs else []
    rows = []
    for i, b in enumerate(spec_boxes[:max(1, len(parts) or 2)]):
        m = SPEC.search(b[4])
        lo, hi = float(m.group(1)), float(m.group(2))
        name = None
        if specs:
            name = next((p for p, (plo, phi) in specs.items() if abs(plo - lo) < 0.6 and abs(phi - hi) < 0.6), None)
            if name is None and i < len(parts):
                name = parts[i]
        if name is None:
            name = "성분%d" % (i + 1)
        # 값은 규격 글 바로 위에 쓰인다 (규격 줄 위 60px ~ 아래 15px, 200dpi 기준)
        rows.append((name, lo, hi, b[1] - 62 * s, b[3] + 8 * s))

    # 초기(인쇄) 일자 → 열 0, 세로 표선으로 열 나누기
    init = next((b for b in dates_printed if W * 0.25 < b[0] < W * 0.45 and b[1] < H * 0.32), None)
    if init is None:
        say("초기 시험일자를 찾지 못함 — %s" % os.path.basename(pdf_path))
        return None
    header_y = _center(init)[1]
    lines = [x for x in table_columns(image, header_y - 20 * s, H * 0.9) if x > init[0] - 40 * s]
    cols = []
    for a, b in zip(lines, lines[1:]):
        if b - a > 60 * s:
            cols.append((a, b))
    if not cols:
        say("표의 세로선을 찾지 못함 — %s" % os.path.basename(pdf_path))
        return None
    # 열 0 = 초기 일자가 든 칸
    k0 = next((i for i, (a, b) in enumerate(cols) if a <= _center(init)[0] <= b), 0)
    cols = cols[k0:]

    # 확인자 줄(맨 아래 결재) — 완료 일자. 인쇄 날짜 아래 2/3 지점부터의 날짜 박스들
    sign_boxes = [b for b in boxes if b[1] > H * 0.72]
    year_lo, year_hi = mfg_year, (mfg_year + 6) if mfg_year else None

    points = []
    for k, (x0, x1) in enumerate(cols):
        period = "Initial" if k == 0 else (PERIODS[k - 1] if k - 1 < len(PERIODS) else "%dM" % (12 * k))
        assays, unsure, seen = {}, [], False
        for name, lo, hi, y0, y1 in rows:
            cell = _in_cell(boxes, x0, x1, y0, y1)
            text = " ".join(b[4] for b in cell)
            conf = min([b[5] for b in cell] or [0])
            value, clean = parse_assay(text, lo, hi)
            crops = []
            if value is None or not clean:
                # 칸만 오려 다시 읽어 본다
                crops = _crop_ocr(image, x0, x1, y0, y1)
                for cb in crops:
                    v2, c2 = parse_assay(cb[4], lo, hi)
                    if v2 is not None and (value is None or (c2 and not clean)):
                        value, clean, conf = v2, c2, cb[5]
                        if c2:
                            break
            if value is None and not text.strip() and not _spilled(boxes, x0, x1, y0, y1):
                continue                                         # 빈 칸(사선) — 시험 안 함
            seen = True
            if value is None:                                    # 글자가 섞여 못 읽은 칸 — 예상값이라도 낸다
                for t in [text] + [cb[4] for cb in crops]:
                    value = guess_assay(t, lo, hi)
                    if value is not None:
                        clean, conf = False, 0.0
                        break
            if value is not None:
                assays[name] = value                             # 읽은 값은 적는다 — 애매하면 아래서 표시
            if not (value is not None and clean and conf >= 0.8):
                unsure.append(name)                              # 담당자 2026-09-06: "판독 후 예상하는 값을
                                                                 # 우선 적어 주고 주황색으로 표시" — 값이 있으면
                                                                 # 그대로 두고 '애매' 로 표시해 노랑(워드)·주황(엑셀)
        # 완료 일자: 이 열의 맨 아래(확인자) 날짜, 없으면 담당자 날짜
        done, done_clean = None, False
        col_dates = sorted(_in_cell(sign_boxes, x0 - 10 * s, x1 + 10 * s, H * 0.72, H), key=lambda b: -b[1])
        for b in col_dates:
            d, c = parse_date(b[4], year_lo, year_hi)
            if d:
                done, done_clean = d, c and b == col_dates[0]
                break
        if done is None:
            for cb in _crop_ocr(image, x0, x1, H * 0.80, H * 0.92):
                d, c = parse_date(cb[4], year_lo, year_hi)
                if d:
                    done, done_clean = d, False
                    break
        if k == 0 and done is None and init is not None:
            done, done_clean = ".".join(DATE.search(init[4]).groups()), False  # 초기 인쇄 일자로 대신 — 확인자 일자와 다를 수 있다
        if not seen and done is None:
            continue
        expected = None
        due = None
        if mfg:
            months = 0 if period == "Initial" else int(period[:-1])
            y, mo = int(mfg.group(1)), int(mfg.group(2)) + months
            y, mo = y + (mo - 1) // 12, (mo - 1) % 12 + 1
            due = (y, mo)
            if done is None:
                expected = "%04d.%02d" % (y, mo)                    # 완료 일자를 못 읽었을 때의 예정 시기
        if done and done_clean and due:
            # 깨끗이 읽혔어도 예정 시기(제조일 + 기간)와 동떨어지면(한 달 이상 이르거나 넉 달 넘게 늦으면)
            # '2024.08.02' 를 '2024.04.02' 로 읽은 경우일 수 있다 — 확인 필요로 돌린다
            dy, dm = int(done[:4]), int(done[5:7])
            gap = (dy - due[0]) * 12 + (dm - due[1])
            if not (-1 <= gap <= 4):
                done_clean = False
        if done and done_clean and points and points[-1].get("done") and done < points[-1]["done"]:
            done_clean = False                                      # 앞 시점보다 이른 완료일 — 오독 가능
        for name, _, _, _, _ in rows:                              # 시험한 시점인데 값을 못 읽은 성분
            if name not in assays and name not in unsure:
                unsure.append(name)
        if done is None:
            unsure.append("done")
        elif not done_clean:
            unsure.append("done")
        point = {"period": period, "done": done or "", "assays": assays, "unsure": sorted(set(unsure))}
        if expected:
            point["expected"] = expected
        points.append(point)

    pack_box = next((b for b in boxes if re.search(r"\d+\s*g\s*[x×]?\s*Tube", b[4], re.I)), None)
    pack = ""
    if pack_box:
        m = re.search(r"(\d+)\s*g", pack_box[4])
        pack = "%sg/Tube" % m.group(1) if m else "4g/Tube"
    store = ""
    if any("25±2" in b[4] for b in boxes):
        store = "25±2°C,\n60±5%RH"
    say("%s: 시점 %d개 판독 (애매 %d칸)" % (lot, len(points), sum(len(p["unsure"]) for p in points)))
    return {"lot": lot, "year": str(mfg_year) if mfg_year else "", "pack": pack, "store": store,
            "why": "", "points": points, "notes": notes, "source": os.path.basename(pdf_path)}


def read_folder(paths, specs=None, log=None):
    out = []
    for i, p in enumerate(paths):
        if log:                                                    # 대시보드 진행 표시에 그대로 보인다
            log("    시험일지 판독 중 %d/%d: %s" % (i + 1, len(paths), os.path.basename(p)))
        try:
            rec = read_log(p, specs, log)
        except Exception as error:                                 # 한 장이 막혀도 나머지는 읽는다
            if log:
                log("    판독 실패 %s — %s" % (os.path.basename(p), error))
            rec = None
        if rec:
            out.append(rec)
    out.sort(key=lambda r: (r.get("year") or "", r.get("lot") or ""))   # 13.1 차례: 제조 연도 → 제조번호
    return out


def merge_known(logs, trend_sheets, log=None):
    """지난 경향표(HLF-QC-126-06)에 이미 적힌 값은 판독값 대신 쓴다 — 담당자가 옮겨 적은 값이 더 믿을 만하다.

    trend_sheets: readers.trend.read_trend() 결과. 성분 이름은 시트 item('함량 - 플루오로메톨론(%)')에서 찾는다.
    돌려주는 값: 바꿔 넣은 칸 수.
    """
    known = {}                                     # (lot, period) → {성분: 값}
    parts = set()
    for one in logs:
        for p in one.get("points", []):
            parts.update(p.get("assays", {}).keys())
            parts.update(u for u in p.get("unsure", []) if u != "done")
    for sheet in trend_sheets or []:
        item = sheet.get("item") or ""
        part = next((n for n in parts if n and n in item), None)
        if part is None:
            continue
        for lot, values in sheet.get("lots", []):
            for period, value in (values or {}).items():
                if value is not None:
                    known.setdefault((lot, period), {})[part] = value
    n = 0
    for one in logs:
        for p in one.get("points", []):
            got = known.get((one.get("lot"), p.get("period")))
            if not got:
                continue
            for part, value in got.items():
                if p["assays"].get(part) != value or part in p.get("unsure", []):
                    p["assays"][part] = value
                    n += 1
                if part in p.get("unsure", []):
                    p["unsure"] = [u for u in p["unsure"] if u != part]
    if log and n:
        log("    지난 경향표에 있는 값 %d칸은 판독값 대신 그 값을 씁니다" % n)
    return n


CACHE_NAME = "13. 안정성시험일지 판독.json"


def save_cache(folder, logs, log=None):
    """판독 결과를 제품 폴더에 남긴다 — 다음 실행은 이 파일을 쓰고, 담당자가 값을 고치면 그 값이 우선."""
    import json
    if not folder or not os.path.isdir(folder):
        return None
    path = os.path.join(folder, CACHE_NAME)
    payload = {"설명": "프로그램이 손글씨 안정성 시험일지를 오프라인(RapidOCR)으로 읽은 결과입니다. "
                     "unsure 에 적힌 성분·done 은 읽기 애매해 보고서에 노랑/주황으로 표시됩니다. "
                     "값을 확인해 고치고 unsure 에서 지우면 다음 재작성부터 그 값이 그대로 쓰입니다.",
               "reader_version": READER_VERSION,
               "logs": logs}
    try:
        with open(path, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2)
        if log:
            log("    판독 결과를 %s 에 저장 — 다음부터는 이 파일을 씁니다" % CACHE_NAME)
        return path
    except OSError:
        return None


def refresh_cache(logs, version, paths, specs=None, log=None):
    """옛 판독기가 만든 판독 파일(logs)을 새 판독기로 보완한다 — 애매한 칸(unsure)과 빈 값만 다시 읽고,
    깨끗이 읽혔거나 담당자가 고쳐 둔 값은 그대로 둔다. 바뀐 칸 수를 돌려준다."""
    try:
        version = int(version or 0)
    except (TypeError, ValueError):
        version = 0
    if version >= READER_VERSION or not paths or not available():
        return 0
    fresh = {one["lot"]: one for one in read_folder(paths, specs, log)}
    changed = 0
    for one in logs:
        new = fresh.get(one.get("lot"))
        if not new:
            continue
        by_period = {p["period"]: p for p in new.get("points", [])}
        for p in one.get("points", []):
            q = by_period.get(p.get("period"))
            if not q:
                continue
            for part in set(q.get("assays", {})) | set(p.get("assays", {})):
                shaky = part in (p.get("unsure") or []) or p["assays"].get(part) is None
                if shaky and q["assays"].get(part) is not None and p["assays"].get(part) != q["assays"][part]:
                    p["assays"][part] = q["assays"][part]
                    changed += 1
                    if part not in p.setdefault("unsure", []) and part in (q.get("unsure") or []):
                        p["unsure"].append(part)
            if not p.get("done") and q.get("done"):
                p["done"] = q["done"]
                changed += 1
    if log:
        log("    옛 판독 파일의 애매한 칸을 새 판독기로 다시 읽음 — %d칸 보완" % changed)
    return changed
