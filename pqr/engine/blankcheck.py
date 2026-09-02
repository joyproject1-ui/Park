import re, subprocess, sys, glob
pdf = sys.argv[1] if len(sys.argv) > 1 else glob.glob('render/*.pdf')[0]
pages = subprocess.run(['pdftotext', '-layout', pdf, '-'], capture_output=True, text=True).stdout.split('\f')
SKIP = ("문서번호", "Rev. No.", "작성일자", "page", "제품코드번호", "제품 품질 평가 보고서",
        "Product Quality Review", "QUIO3", "퀴노비드안연고(오플록사신)", "한림제약", "HLF-QC-126-01")
blank = []
for i, p in enumerate(pages, 1):
    if i >= len(pages):
        break
    body = [ln for ln in p.split('\n')
            if ln.strip() and not any(k in ln for k in SKIP)]
    if len(re.sub(r'\s+', '', ' '.join(body))) < 20:
        blank.append(i)
print('총', len(pages) - 1, '쪽 / 내용 없는 쪽:', blank)
