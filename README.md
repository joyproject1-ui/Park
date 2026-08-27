# GMP 품질 도구 모음

이 저장소에는 두 가지 도구가 있습니다.

| 도구 | 하는 일 | 문서 |
| --- | --- | --- |
| `gmpai` | EU · FDA 의 GMP × AI 규정 원문을 공식 사이트에서 내려받고 개정을 추적 | 아래 |
| `pqr` | 담당자가 제품 폴더에 올린 자료로 제품품질평가(PQR) 집계 · 판정 · 보고서 초안 생성 | [`pqr/README.md`](pqr/README.md) · [대시보드](docs/pqr/README.md) |

```bash
python -m gmpai list                 # 규정 문서 목록
python -m pqr demo -o sample_input   # PQR 예시 자료 만들기
python -m pqr serve --in sample_input # 대시보드 + 화면에서 자료 올리기
```

---

# GMP × AI 규정 다운로더

EU(European Commission · EMA · EUR-Lex)와 미국 FDA가 발행한 **GMP 환경에서의 인공지능 관련 규정 원문**을
공식 사이트에서 직접 내려받아 로컬에 보관하고, 개정 여부를 추적하는 도구입니다.

외부 라이브러리 없이 **Python 3.9+ 표준 라이브러리만** 사용합니다. (사내망 · 프록시 환경 고려)

---

## 빠른 시작

```bash
# 등록된 문서 목록 보기
python -m gmpai list

# 전체 내려받기 (downloads/EU, downloads/FDA 에 저장)
python -m gmpai download

# EU 문서만
python -m gmpai download --authority EU

# 핵심 3건만 (EU Annex 22 + FDA AI 지침 + FDA 제조 AI 논의문서)
python -m gmpai download \
  --id eu-gmp-annex22-ai \
  --id fda-ai-regulatory-decision-making \
  --id fda-ai-drug-manufacturing
```

내려받지 않고 링크만 클릭해서 받고 싶다면 [`INDEX.md`](INDEX.md) 또는 브라우저에서
[`docs/index.html`](docs/index.html)을 여세요. 두 파일 모두 카탈로그에서 자동 생성됩니다.

## 명령어

| 명령 | 설명 |
| --- | --- |
| `list` | 카탈로그 목록 출력 (`--long` 으로 URL·비고 포함) |
| `download` | PDF 내려받기 + `manifest.json` 갱신 |
| `verify` | 등록된 링크가 아직 PDF를 반환하는지 점검 |
| `index` | `INDEX.md` 와 `docs/index.html` 재생성 |
| `status` | 이미 받은 파일과 해시 확인 |

공통 필터: `--id`(반복 가능) · `--authority` · `--category` · `--status`

`download` 옵션:

- `-o, --out DIR` 저장 위치 (기본 `downloads/`)
- `--force` 이미 받은 문서도 다시 받아 **개정 여부 확인** (해시가 달라지면 `[갱신]`으로 표시)
- `--dry-run` 실제 저장 없이 사용할 URL만 확인
- `--prefer-discovery` 직접 URL 대신 출처 페이지에서 찾은 링크를 우선 사용
- `--timeout`, `--retries` 네트워크 조정 (재시도는 2·4·8초 지수 백오프)

## 동작 방식

각 문서는 `direct_url`(알려진 PDF 주소)과 `discover` 규칙(출처 페이지에서 링크를 찾는 정규식)을 함께 갖습니다.

1. `direct_url` 로 먼저 시도합니다.
2. 실패하거나 PDF가 아니면 **출처 페이지를 열어 링크를 다시 찾아** 내려받습니다.
   규제기관이 파일을 새 주소로 재게시해도 계속 동작하도록 만든 장치입니다.
3. 응답이 실제 PDF인지(`%PDF-` 시그니처) 확인한 뒤에만 저장합니다.
   쿠키 동의 페이지나 오류 페이지가 PDF로 둔갑해 저장되는 일을 막습니다.
4. SHA-256 해시·크기·최종 URL·수신 시각을 `downloads/manifest.json` 에 기록합니다.

이미 받은 문서는 기본적으로 건너뜁니다. 정기적으로 `--force` 로 다시 받으면
해시 비교로 개정본 여부를 알 수 있습니다.

```bash
python -m gmpai download --force        # 분기별 개정 점검 용도
python -m gmpai verify                  # 링크만 빠르게 점검
```

### 프록시 환경

`HTTPS_PROXY` / `HTTP_PROXY` / `NO_PROXY` 환경변수를 그대로 따릅니다.

```bash
export HTTPS_PROXY=http://proxy.company.local:8080
python -m gmpai download
```

## 수록 문서

전체 목록과 설명은 [`INDEX.md`](INDEX.md)에 있습니다. 핵심만 추리면:

**EU**

- EudraLex Vol. 4 **Annex 22 인공지능** (2025-07-07 의견수렴 초안) — GMP 영역에서 AI를 정면으로 다룬 최초의 EU 문서
- EudraLex Vol. 4 **Annex 11 전산화 시스템** (개정 초안 + 현행 발효본)
- EudraLex Vol. 4 **Chapter 4 문서화** (개정 초안)
- EMA **AI 리플렉션 페이퍼** (EMA/CHMP/CVMP/83833/2023)
- EMA/FDA 공동 **Good AI Practice 지침 원칙** (2026-01)
- **EU AI Act** (Regulation (EU) 2024/1689)

**FDA**

- **Considerations for the Use of AI to Support Regulatory Decision-Making for Drug and Biological Products** (2025-01 초안)
- **Artificial Intelligence in Drug Manufacturing** 논의문서 (CDER FRAME, 2023-03)
- **Computer Software Assurance** 최종 지침 (2025-09)
- AI 기반 의료기기 소프트웨어 전주기 관리 초안 지침, **PCCP** 최종 지침
- **Data Integrity and CGMP Q&A**, **Part 11** 지침

> 상태 표기(초안/최종/발효)는 카탈로그 검토일(`catalog.json`의 `last_reviewed`) 기준입니다.
> Annex 22와 FDA AI 지침은 2026년 8월 기준으로 아직 초안 단계이므로, 규제 대응에 사용하기 전
> `verify` 또는 출처 페이지에서 최신 상태를 반드시 확인하세요.

## 문서 추가하기

`gmpai/data/catalog.json` 의 `documents` 배열에 항목을 추가하면 됩니다.

```json
{
  "id": "example-doc",
  "title": "Document title in English",
  "title_ko": "한국어 제목",
  "authority": "EU",
  "issuer": "EMA",
  "category": "gmp-ai",
  "status": "draft",
  "document_date": "2026-01-01",
  "reference": "EMA/1234/2026",
  "landing_page": "https://www.ema.europa.eu/...",
  "direct_url": "https://www.ema.europa.eu/....pdf",
  "discover": { "pattern": "(?i)[^\"']*example[^\"']*\\.pdf" },
  "filename": "EMA_Example_2026.pdf",
  "notes": "왜 필요한 문서인지",
  "language": "en"
}
```

`direct_url` 과 `discover` 중 최소 하나는 있어야 하며, 테스트가 URL이 공식 도메인(HTTPS)인지 검사합니다.
추가 후 `python -m gmpai index` 로 목록 파일을 다시 생성하세요.

## 테스트

```bash
python -m unittest discover -s tests -t .
```

두 도구를 합쳐 103건의 테스트가 네트워크 없이 동작합니다. `gmpai` 는 로컬 모의 서버로
다운로드·재개정 감지·링크 탐색 폴백·HTML 오응답 차단을 검증하고, `pqr` 은 엑셀·CSV 적재부터
공정능력·경향 판정·보고서 생성까지 검증합니다 (Claude API 호출은 가짜 클라이언트로 대체).

## 라이선스와 저작권

이 저장소의 코드는 MIT 라이선스입니다. 내려받는 규정 문서 자체의 권리는 각 발행기관(EC, EMA, FDA)에
있으며, 재배포 시 각 기관의 이용 조건을 따르세요. 편의를 위해 PDF 자체는 저장소에 커밋하지 않습니다
(`downloads/` 는 `.gitignore` 처리).
