# -*- coding: utf-8 -*-
"""제품 폴더의 항 번호별 파일을 찾아 판독기에 넘기고, 보고서에 넣을 값(ProductData)으로 모은다.

담당자 폴더 규칙: 파일 이름이 보고서 항 번호로 시작한다 ("6. 제조내역 - ERP(수출용).pdf",
"9.2.3 충전 완료 후 - ERP.zip", "13 안정성 시험 - … 시험일지.pdf"). 압축은 풀어서 본다.
읽지 못한 것은 버리지 않고 `issues` 에 남겨 담당자 문의 목록으로 보여 준다.
"""
import os
import re
import tempfile
import zipfile

from .. import build as build_module
from .readers import coa as coa_reader, erp, license as license_reader, deviation, change, \
    masters, yield_sheet, suppliers
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


def _walk(folder):
    for root, _, files in os.walk(folder):
        for name in files:
            if name.startswith("~$") or name.startswith("."):
                continue
            yield os.path.join(root, name)


def _member_name(info):
    """한글 이름이 CP437 로 깨져 오는 압축이 흔하다 (Windows 가 만든 압축)."""
    name = info.filename
    if not (info.flag_bits & 0x800):
        try:
            name = name.encode("cp437").decode("cp949")
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
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
            if name.startswith("~$") or name.startswith("."):
                continue
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


def collect(folder, product_name=None, log=None):
    log = log or (lambda *a: None)
    data = ProductData()
    data.files = discover(folder)
    got = data.files

    def note(item, path, why):
        data.issues.append((item, os.path.basename(path), why))
        log("  [%s] %s — %s" % (item, os.path.basename(path), why))

    # 6. 제조내역 (수출용 ERP)  · 7. 수율
    for p in got.get("6", []):
        if p.lower().endswith(".pdf"):
            try:
                data.manufacturing += erp.read_manufacturing(p)
            except PdfTextError as e:
                note("6", p, str(e))
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
    for item, key, fn in (("9.2.2", "922", coa_reader.read_ipc), ("9.2.3", "923", coa_reader.read_ipc),
                          ("9.2.4", "924", coa_reader.read_fp)):
        for p in got.get(item, []):
            if not p.lower().endswith(".pdf"):
                continue
            try:
                rec = fn(p)
            except PdfTextError as e:
                note(item, p, str(e)); continue
            lot = rec.get("lot") or _lot_from_name(os.path.basename(p))
            if not lot:
                note(item, p, "제조번호를 읽지 못함"); continue
            rec["export"] = "수출용" in os.path.basename(p)
            data.coa.setdefault(lot, {})[key] = rec
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
                if is_scanned(p):
                    note("3", p, "글자 정보가 없는 PDF — 허가증 값은 결재본 값을 유지, 원본 확인 필요")
                else:
                    data.license = license_reader.read_license(p)
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
    for p in got.get("10.2", []):
        if p.lower().endswith(".xlsx"):
            try:
                data.equipment = masters.equipment_docs(p)
            except Exception as e:
                note("10.2", p, str(e))
    for p in got.get("10.3-5", []) + got.get("10.3", []):
        if p.lower().endswith(".xlsx"):
            try:
                data.support = masters.support_docs(p)
            except Exception as e:
                note("10.3-5", p, str(e))
    # 11 · 12
    for p in got.get("11", []):
        if p.lower().endswith(".pdf"):
            try:
                data.deviations.append(deviation.read_deviation(p))
            except PdfTextError as e:
                note("11", p, str(e))
    for p in got.get("12", []):
        if p.lower().endswith(".pdf"):
            try:
                data.changes.append(change.read_change(p))
            except PdfTextError as e:
                note("12", p, str(e))
    # 13 안정성 — 스캔이면 비전 판독 대상
    for p in got.get("13", []):
        if p.lower().endswith(".pdf"):
            data.stability_files.append((p, is_scanned(p)))
    # 16 전년도 결재본
    for p in got.get("16", []):
        if p.lower().endswith((".doc", ".docx")):
            data.previous_report = p
    data.deviations.sort(key=lambda d: d.get("doc_no") or "")
    return data
