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
    out["appearance"] = _first(r"성상\s+(.+?)\s{2,}", text)
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


def read_fp(path):
    """완제 시험성적서 한 장 (HLF-QC-327-06)."""
    text = squash(read_text(path))
    out = {"file": os.path.basename(path)}
    out["lot"] = _first(r"\b(O[A-Z]{2}[A-Z0-9]{3})\b", text) or _first(LOT.pattern, text)
    out["expiry"] = _first(r"\b(\d{4}\.\d{2}\.\d{2})\s+한림제약", text)
    out["appearance"] = _first(r"성상\s+(.+?)\s{2,}", text)
    out["metal_total"] = _first(r"합계는\s*(\d+)\s*개\s*,\s*개개는 8개를 초과하는 것", text)
    out["metal_each"] = _first(r"이\s*(\d+)\s*매\s*$", text, flags=re.M)
    out["particle"] = _first(r"입자도.*?(\d+)\s*um\s*이하", text, flags=re.S)
    out["assay"] = _first(r"90\.0 ~ 110\.0%\s+.*?([\d]+\.\d)\s*%", text, flags=re.S)
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
    m = re.search(r"(\d{2,3})\s*nm\s*,\s*(\d{2,3})\s*nm\s*에서\s*흡수극대", text)
    if m:
        out["uv_max"] = "%snm, %snm에서 흡수극대를 나타냄" % (m.group(1), m.group(2))
    m = re.search(r"합격\s+(\d{4}\.\d{2}\.\d{2})\s+\S+\s*\S*\s+(\d{4}\.\d{2}\.\d{2})", text)
    if m:
        out["checked"], out["judged"] = m.group(1), m.group(2)
    return out
