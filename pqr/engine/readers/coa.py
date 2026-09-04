# -*- coding: utf-8 -*-
"""공정 시험성적서(HLF-QC-121-04) 와 완제 시험성적서(HLF-QC-327-06) 판독.

퀴노비드안연고 2026 PQR 에서 19 Lot × 3 공정 성적서를 이 규칙으로 읽어 담당자 확인을
거쳤다. 값은 성적서 원문 표기를 그대로 둔다("10CFU/g 미만", "5.04 ~ 5.20g").
"""
import os
import re

from ..pdftext import read_text, squash

LOT = re.compile(r"\b([A-Z]{2}[A-Z0-9]{4})\b")      # OEY101 · OZYD01 · LKY401 …


def norm(s):
    return re.sub(r"\s+", " ", s or "").strip()


def _first(pattern, text, group=1, flags=0):
    m = re.search(pattern, text, flags)
    return norm(m.group(group)) if m else None


def _result_tail(both):
    """기준과 결과가 이어 붙은 글에서 **결과**만 떼어 낸다.

    성적서는 '성상 <기준> <결과> <시험자> <적부>' 꼴이라 둘이 붙어 나온다. 기준이 결과보다 넓게
    적히므로(디겐타안연고: 기준 '백색~미황색의 …', 결과 '미황색의 …') 결과는 **앞쪽에도 나오는
    가장 긴 꼬리**다. 기준과 결과가 같은 글인 제품(퀴노비드안연고)도 같은 규칙으로 떨어진다.
    """
    words = both.split()
    for start in range(len(words)):
        tail = " ".join(words[start:])
        if len(tail) >= 6 and tail in " ".join(words[:start]):
            return tail
    return both


def _clean_result(text):
    """떼어 낸 결과가 미덥지 않으면 None — 기준의 '~'(범위 표시)가 남아 있으면 갈라지지 않은 것이다.

    두 칸짜리 PDF 는 줄이 엇갈려 나오는 것이 있어(조제·충전 성적서) 기준과 결과가 깨끗이
    갈라지지 않는다. 어설프게 적느니 비워 두고 다른 성적서(완제)의 값을 쓴다 — 값은 지어내지 않는다.
    """
    got = norm(text)
    return None if (not got or "~" in got) else got


def _appearance(text):
    """성적서의 성상 **결과**. 없으면 None."""
    m = re.search(r"성상\s+(\S+(?: \S+){0,8}?)\s+\1(?:\s|$)", text)
    if m:
        return m.group(1)
    m = re.search(r"성상\s+(.+?)\s+[가-힣]{2,4}\s+(?:적합|부적합)", text)
    if m:
        return _clean_result(_result_tail(norm(m.group(1))))
    m = re.search(r"성상\s+(.+?)\s+\S+\s+\d{4}\.\d{2}\.\d{2}", text)
    return _clean_result(_result_tail(norm(m.group(1)))) if m else None


def read_ipc(path):
    """공정 시험성적서(조제·충전 등) 한 장."""
    text = squash(read_text(path))
    out = {"file": os.path.basename(path)}
    out["lot"] = _first(r"제\s*조\s*번\s*호.*?\n.*?\b([A-Z]{2}[A-Z0-9]{4})\b", text, flags=re.S) \
        or _first(r"\b(O[A-Z]{2}[A-Z0-9]{3})\b", text)
    out["stage"] = _first(r"공\s*정\s*명\s+(\S+)", text)
    out["mfg_date"] = _first(r"제\s*조\s*일\s*자\s+(\d{4}/\d{2}/\d{2})", text)
    out["sampled"] = _first(r"검체채취일\s+(\d{4}\s*/\s*\d{2}\s*/\s*\d{2})", text)
    out["lot_size"] = _first(r"L\s*O\s*T\s*_?\s*수\s*량\s+([\d,]+)", text)
    out["done"] = _first(r"시험완료일자\s+(\d{4}/\s*\d{2}/\s*\d{2})", text)
    out["verdict"] = _first(r"판\s*정\s*결\s*과\s+(\S+)", text)
    out["appearance"] = _appearance(text)
    m = re.search(r"생균수\s+(\S+)\s*이하\s+(\S+(?:\s*\S+)?)\s+\S+\s+\d{4}\.\d{2}\.\d{2}", text)
    if m:
        out["bioburden_spec"], out["bioburden"] = m.group(1) + " 이하", norm(m.group(2))
    m = re.search(r"질량[·․\.]?\s*용량\s+평균\s*:\s*([\d\.]+\s*~\s*[\d\.]+g)\s+([\d\.]+\s*g?)", text)
    if m:
        out["mass_avg_spec"], out["mass_avg"] = norm(m.group(1)), norm(m.group(2))
    m = re.search(r"개개\s*:\s*([\d\.]+g 이상)\s+([\d\.]+\s*~\s*[\d\.]+\s*g?)", text)
    if m:
        out["mass_each_spec"], out["mass_each"] = norm(m.group(1)), norm(m.group(2))
    if "인쇄상태가 양호하며" in text:
        out["tube_print"] = "인쇄상태가 양호하며 제조번호 및 사용기한의 압인상태가 명확히 식별 가능함"
    if "메틸렌블루시액 침투 없음" in text:
        out["leak"] = "메틸렌블루시액 침투 없음"
    return out


ASSAY_HEAD = re.compile(r"함량\s+(?!시험)")
SPEC = re.compile(r"([\d]+(?:\.\d)?)\s*~\s*([\d]+(?:\.\d)?)\s*%")
VALUE = re.compile(r"([\d]+\.\d)\s*%")
PART = re.compile(r"^\s*([가-힣A-Za-z][가-힣A-Za-z0-9]{2,})\s*$", re.M)


def _assays(text):
    """함량 시험 결과 — [{"part": 성분, "value": 결과, "lo": 하한, "hi": 상한}, ...]

    성적서마다 기준과 결과의 차례가 다르다. 어느 쪽이든 읽어야 한다.

        퀴노비드안연고 : 함량  90.0 ~ 110.0%  107.3%  박소은  적합      (기준 먼저)
        디겐타안연고   : 함량  97.7%  김득규  적합 / 90.0 ~ 110.0% (…)  (결과 먼저)

    주성분이 둘 이상인 제품은 성분마다 함량 줄이 따로 있고, 성분 이름은 그 앞줄에 있다.
    기준을 먼저 찾아 짝짓던 예전 방식은 디겐타에서 성분을 엇갈리게 붙였다(97.7 을 놓치고
    99.3 을 110.0 기준에 붙임).
    """
    out = []
    heads = [m.end() for m in ASSAY_HEAD.finditer(text)]
    for i, at in enumerate(heads):
        stop = min(heads[i + 1] - 3 if i + 1 < len(heads) else len(text), at + 300)
        window = text[at:stop]
        spec = SPEC.search(window)
        if not spec:
            continue
        lo, hi = spec.group(1), spec.group(2)
        value = None
        for m in VALUE.finditer(window):
            if m.start() >= spec.end() or m.end() <= spec.start():      # 기준 숫자는 건너뛴다
                got = m.group(1)
                if float(lo) - 20 <= float(got) <= float(hi) + 20:
                    value = got
                    break
        if value is None or not float(lo) < float(hi):
            continue
        before = text[max(0, at - 300):at]
        names = PART.findall(before)
        out.append({"part": names[-1].strip() if names else "", "value": value, "lo": lo, "hi": hi})
    return out


def read_fp(path):
    """완제 시험성적서 한 장 (HLF-QC-327-06)."""
    text = squash(read_text(path))
    out = {"file": os.path.basename(path)}
    out["lot"] = _first(r"\b(O[A-Z]{2}[A-Z0-9]{3})\b", text) or _first(LOT.pattern, text)
    out["expiry"] = _first(r"\b(\d{4}\.\d{2}\.\d{2})\s+한림제약", text)
    out["appearance"] = _appearance(text)
    out["metal_total"] = _first(r"합계는\s*(\d+)\s*개\s*,\s*개개는 8개를 초과하는 것", text)
    out["metal_each"] = _first(r"이\s*(\d+)\s*매\s*$", text, flags=re.M)
    # 입자도: '입자도  75㎛ 이하  15㎛ 이하' — 앞이 기준, 뒤가 결과다. 단위는 ㎛(U+339B) 이거나
    # um 으로 나온다. 예전에는 um 만 보고 첫 숫자(=기준)를 집어 결과를 놓쳤다.
    m = re.search(r"입자도\s+([\d.]+)\s*(?:㎛|um|μm)\s*이하\s+([\d.]+)\s*(?:㎛|um|μm)\s*이하", text)
    if m:
        out["particle"], out["particle_spec"] = m.group(2), m.group(1)
    else:
        out["particle"] = _first(r"입자도.*?([\d.]+)\s*(?:㎛|um|μm)\s*이하", text, flags=re.S)
    # 함량: '<하한> ~ <상한>%  <결과>%' — 규격 숫자는 제품마다 다르므로 글에서 읽는다
    # (퀴노비드 90.0~110.0, 다른 제품은 95.0~105.0 등). 함량 항목이 둘 이상이면 차례로 assays 에 둔다.
    assays = _assays(text)
    if assays:
        out["assay"] = assays[0]["value"]
        out["assay_spec"] = "%s ~ %s%%" % (assays[0]["lo"], assays[0]["hi"])
        out["assays"] = assays
    m = re.search(r"평균\s*:\s*([\d\.]+)\s*g\s*,\s*개개\s*:\s*([\d\.]+)\s*g\s*이상", text)
    if m:
        out["mass_avg"], out["mass_each_min"] = m.group(1), m.group(2)
    if re.search(r"무균\s+음성\(불검출\)\s+음성\(불검출\)", text):
        out["sterility"] = "음성(불검출)"
    elif re.search(r"무균\s+음성\s+음성", text):
        out["sterility"] = "음성"
    if "메틸렌블루시액 침투 없이 양호" in text:
        out["leak"] = "메틸렌블루시액 침투 없이 양호"
    elif "누출 발생 없이 양호함" in text:
        out["leak"] = "누출 발생 없이 양호함"
    m = re.search(r"액은\s*(\S+?)(?:을|를)\s*나타(?:냄|낸다)", text)
    if m:
        out["ident_color"] = "액은 %s을 나타냄" % m.group(1)
    flat = re.sub(r"\s+", " ", text)
    if "침전이 생김" in flat:
        mm = re.search(r"((?:[가-힣]+ ){0,2}침전이 생김)", flat)
        out["ident_precip"] = mm.group(1).strip() if mm else "침전이 생김"
    m = re.search(r"(\d{2,3})\s*nm\s*,\s*(\d{2,3})\s*nm\s*에서\s*흡수극대", text)
    if m:
        out["uv_max"] = "%snm, %snm에서 흡수극대를 나타냄" % (m.group(1), m.group(2))
    m = re.search(r"합격\s+(\d{4}\.\d{2}\.\d{2})\s+\S+\s*\S*\s+(\d{4}\.\d{2}\.\d{2})", text)
    if m:
        out["checked"], out["judged"] = m.group(1), m.group(2)
    return out
