# -*- coding: utf-8 -*-
"""'보고서 작성' 한 번에 결재본 양식의 제출용 보고서를 만든다.

절차: 전년도 결재본 → .docx 로 → 자료 판독(collect) → 항별 채움(recipe) → 조판(layout)
→ 저장 → 변환 흔적·OOXML 순서 정리(polish) → 순서 검사 0 건 확인.
헷갈린 값은 issues(문의 목록)로 남기고 보고서에는 '확인 필요' 표시를 둔다.
"""
import os
import re
import shutil
import tempfile
import zipfile

from . import collect as collect_module
from . import convert, layout, polish
from .polish import check_order


class EngineError(Exception):
    pass


def _year_of(text):
    m = re.search(r"PQR(\d{2})-", text or "")
    return (2000 + int(m.group(1))) if m else None


def find_previous(folder):
    """평가항목 16 의 전년도 결재본(.doc/.docx, 압축 안 포함)."""
    data_files = collect_module.discover(folder)
    for p in data_files.get("16", []):
        if p.lower().endswith((".docx", ".doc")) and not os.path.basename(p).startswith("~$"):
            return p
    return None


def write_report(folder, product, period, out_path, today=None, recipe=None, log=None, vision=None):
    """folder: 제품 폴더 · product: {"code","name","group"} · period: {"from","to"} · out_path: 저장할 .docx

    돌려주는 값: {"path", "issues": [(항, 파일, 설명)], "log": [...]}
    """
    lines = []

    def log_(msg):
        lines.append(msg)
        if log:
            log(msg)

    # 서식은 언제나 EDMS 결재본 서식(E-HLF-32). 전년도 결재본은 값·문안의 근거로만 쓴다.
    # 서식이 폴더에 없을 때만 전년도 결재본을 바탕으로 쓴다(예전 방식).
    from . import edms
    previous = find_previous(folder)
    form = edms.find_form(folder)
    source, why = edms.choose_base(form, previous)
    if not source:
        raise EngineError(why)
    work = tempfile.mkdtemp(prefix="pqr-report-")
    base = os.path.join(work, "base.docx")
    how = convert.to_docx(source, base)
    log_("바탕 문서: %s — %s (%s)" % (os.path.basename(source), why, how))
    if previous and previous != source:
        log_("전년도 결재본(값 근거): %s" % os.path.basename(previous))

    data = collect_module.collect(folder, product_name=product.get("name"), log=log_)
    data.previous_report = previous
    data.period = period
    if vision is not None:
        try:
            vision(data, log_)
        except Exception as error:               # 비전 판독 실패는 보고서를 막지 않는다
            data.issues.append(("13", "", "손글씨 판독 실패: %s" % error))

    import docx
    document = docx.Document(base)
    if recipe is None:
        from . import recipe_ointment as recipe_module
        recipe = recipe_module.fill
    ctx = recipe(document, data, product, period, today=today, log=log_)
    layout.apply(document, log=log_, product_title=(ctx or {}).get("cover_title"))
    document.save(out_path)
    polish.polish(out_path)
    with zipfile.ZipFile(out_path) as z:
        bad = sum(check_order(z.read(n).decode("utf-8")) for n in z.namelist()
                  if n.startswith("word/") and n.endswith(".xml"))
    if bad:
        raise EngineError("OOXML 순서 검사에 실패했습니다 (%d 곳)" % bad)
    log_("순서 검사 0 건 · 저장: %s" % os.path.basename(out_path))
    # 첨부 엑셀 — Cpk 계산 파일 4종(결재본 것을 물려받아 값 갱신) + 안정성 경향 분석
    attachments = []
    try:
        from . import excel_attach
        day = today.strftime("%Y.%m.%d") if hasattr(today, "strftime") else re.sub(r"-", ".", str(today or ""))[:10]
        attachments += excel_attach.write_cpk_files(os.path.dirname(out_path), data, previous, day)
        attachments += excel_attach.write_stability_workbook(
            os.path.dirname(out_path), data, product, day,
            input_dir=os.path.dirname(os.path.abspath(folder)), report_path=out_path)
        log_("첨부 엑셀: %s" % ", ".join(n for n, _ in attachments))
    except Exception as error:
        data.issues.append(("첨부", "", "첨부 엑셀 생성 실패: %s" % error))
    shutil.rmtree(work, ignore_errors=True)
    return {"path": out_path, "issues": data.issues + list((ctx or {}).get("issues", [])), "log": lines,
            "data": data, "attachments": attachments}
