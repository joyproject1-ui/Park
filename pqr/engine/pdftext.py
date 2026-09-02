# -*- coding: utf-8 -*-
"""PDF 글자를 표 배치 그대로 읽는다 (pdftotext -layout 과 같은 모양).

성적서·일탈보고서·변경요청서는 라벨과 값이 가로로 나란히 놓인 표라, 글자 순서만
뽑으면 어느 칸의 값인지 알 수 없다. pdfplumber 의 layout 모드가 칸 위치를 살려 준다.
"""
import re


class PdfTextError(Exception):
    pass


def _normalize(text):
    # layout 모드는 낱말 사이를 두 칸 이상 벌려 놓는다 — 한 칸으로 줄여 정규식이 단순해지게
    return "\n".join(re.sub(r"[ \t]{2,}", "  ", line.rstrip()) for line in text.split("\n"))


def read_layout(path, pages=None):
    """쪽마다 배치를 살린 글자를 돌려준다. [str, ...]"""
    try:
        import pdfplumber
    except ImportError:
        raise PdfTextError("PDF 를 읽으려면 pdfplumber 가 필요합니다: pip install pdfplumber")
    out = []
    try:
        with pdfplumber.open(path) as pdf:
            for i, page in enumerate(pdf.pages, 1):
                if pages and i not in pages:
                    continue
                try:
                    text = page.extract_text(layout=True, x_density=4.5, y_density=10) or ""
                except Exception:
                    text = page.extract_text() or ""
                out.append(_normalize(text))
    except PdfTextError:
        raise
    except Exception as error:
        raise PdfTextError("PDF 를 읽지 못했습니다: %s (%s)" % (path, error))
    return out


def read_text(path):
    """모든 쪽을 이어 붙인 글자."""
    return "\f".join(read_layout(path))


def squash(text):
    """정규식용: 낱말 사이 공백을 한 칸으로."""
    return re.sub(r"[ \t]+", " ", text or "")


def is_scanned(path):
    """글자 정보가 거의 없으면 스캔본으로 본다 (손글씨 시험일지 등)."""
    try:
        pages = read_layout(path, pages={1})
    except PdfTextError:
        return True
    return len((pages[0] if pages else "").strip()) < 40
