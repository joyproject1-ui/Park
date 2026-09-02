# -*- coding: utf-8 -*-
"""결재본 양식으로 제출용 PQR 보고서를 자동 작성하는 엔진.

담당자가 하는 방식 그대로 — 전년도 결재본을 열어 첨부 자료 값을 올해 것으로 갈아
끼우고, 회사 조판 규칙(굴림 10 · 표 안 간격 0 · 빈 칸 사선 · 행 분할 금지 · 각주
윗첨자 …)을 적용한 뒤, Word 가 속성을 무시하지 않도록 OOXML 순서를 검사한다.

python-docx · lxml · openpyxl · xlrd · pdfminer.six 가 필요하다 (requirements.txt).
"""
