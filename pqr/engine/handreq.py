# -*- coding: utf-8 -*-
"""13항 손글씨 시험일지를 'Claude 판독 요청' 한 묶음(zip)으로 만든다.

담당자 2026-09-07: "PC 판독하지 말고 Claude 에서 확인해 주도록 해 / API 키 없이 Claude 로 판독해 주도록 해."

이 PC 의 판독기(RapidOCR)는 한 장에 몇 분이 걸린다(담당자 PC: 20장에 71분). API 키가 있으면
Claude 가 곧바로 읽지만(vision_claude), 키 없이 쓰려면 사람이 대화에 올려 주어야 한다. 그래서
프로그램은 판독을 붙들고 있지 않고 — 올리기만 하면 되는 묶음을 만들고 보고서는 바로 낸다.

받은 판독 결과(`13. 안정성시험일지 판독.json`)를 제품 폴더에 두고 다시 작성하면 13항이 채워진다.
"""
import os
import zipfile

REQUEST_NAME = "13. Claude 판독 요청%s.zip"
PART_LIMIT = 15 * 1024 * 1024          # 한 묶음이 너무 크면 대화에 올리기 어렵다 — 나눠 담는다

GUIDE = """13항 안정성 시험일지 판독 요청
=====================================

이 묶음 안의 시험일지를 Claude 대화창에 올리고 아래처럼 부탁하세요.

    이 안정성 시험일지를 읽어서 '13. 안정성시험일지 판독.json' 을 만들어 줘.

받은 json 파일을 아래 폴더에 그대로 두고(같은 이름 파일이 있으면 덮어쓰기),
PQR 관리 시스템에서 '보고서 재작성' 을 누르면 13항이 채워집니다.

    %(folder)s

이 묶음에 든 시험일지 (%(count)d개)
%(files)s

--------------------------------------------------------------------
Claude 가 만들 판독 파일의 꼴 (이 설명은 그대로 두셔도 됩니다)

{
  "reader_version": 3,
  "covers_all": true,               // 13 폴더 전체를 읽었다는 표시 — PC 판독을 건너뜁니다
  "logs": [
    {
      "lot": "OEX101",              // 제조번호
      "year": "2024",               // 제조 연도
      "kind": "장기",                // "장기" 또는 "시판후"
      "market": "내수",              // "내수" 또는 "수출"
      "pack": "5g tube/갑",
      "store": "25±2°C,\\n60±5%%RH",
      "mfg": "2024.01.10",
      "expiry": "2027.01.09",
      "source": "시험일지 파일 이름.pdf",
      "points": [
        {
          "period": "12M",          // Initial · 3M · 6M · 9M · 12M · 18M · 24M · 36M
          "done": "2025.03.10",     // 그 시점 시험(결재) 완료 일자
          "assays": {"오플록사신": 101.2},
          "unsure": []              // 읽기 애매한 것: "done" 또는 성분 이름
        }
      ]
    }
  ]
}

· unsure 에 적힌 칸은 보고서에서 노랑(워드)·주황(엑셀)으로 표시되어 원본과 대조할 수 있습니다.
· 값을 직접 고치셔도 됩니다 — 담당자가 고친 값이 언제나 먼저입니다.
· 값을 지어내지 않습니다. 읽지 못한 시점은 넣지 않거나 unsure 에 적습니다.
"""


def make_request(folder, paths, product="", log=None):
    """시험일지 paths 를 묶어 제품 폴더에 판독 요청 zip 을 만든다. [(이름, 경로)] 를 돌려준다."""
    if not folder or not os.path.isdir(folder) or not paths:
        return []
    say = log or (lambda *a: None)
    for old in os.listdir(folder):                    # 지난번 요청 묶음은 지운다
        if old.startswith("13. Claude 판독 요청") and old.endswith(".zip"):
            try:
                os.remove(os.path.join(folder, old))
            except OSError:
                pass
    groups, batch, size = [], [], 0
    for path in paths:
        try:
            one = os.path.getsize(path)
        except OSError:
            one = 0
        if batch and size + one > PART_LIMIT:
            groups.append(batch)
            batch, size = [], 0
        batch.append(path)
        size += one
    if batch:
        groups.append(batch)
    out = []
    for i, group in enumerate(groups, 1):
        tail = "" if len(groups) == 1 else " (%d of %d)" % (i, len(groups))
        name = REQUEST_NAME % ((" - " + product if product else "") + tail)
        dst = os.path.join(folder, name)
        guide = GUIDE % {"folder": os.path.abspath(folder), "count": len(group),
                         "files": "\n".join("  · " + os.path.basename(p) for p in group)}
        try:
            with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
                z.writestr("읽어주세요.txt", guide)
                seen = set()
                for path in group:
                    base = os.path.basename(path)
                    if base in seen:                  # 같은 이름이 두 곳에 있으면 뒤엣것에 번호를 붙인다
                        stem, ext = os.path.splitext(base)
                        base = "%s (%d)%s" % (stem, len(seen), ext)
                    seen.add(base)
                    z.write(path, base)
            out.append((name, dst))
            say("  [13] 판독 요청 묶음: %s (%d개)" % (name, len(group)))
        except OSError as error:
            say("  [13] 판독 요청 묶음을 만들지 못했습니다 — %s" % error)
    return out


PC_OPT_IN = "PC 판독 사용.txt"


def pc_reading_on(folder, root=None):
    """이 PC 로 직접 읽으라고 담당자가 정해 두었는가 — 제품 폴더나 프로그램 폴더에 표시 파일이 있으면.

    기본은 끔이다: 한 장에 몇 분이라 20장이면 한 시간을 넘긴다(담당자 2026-09-07).
    """
    here = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    for where in (folder, root or here):
        if where and os.path.isfile(os.path.join(where, PC_OPT_IN)):
            return True
    return False
