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
| `letters` | **FDA Warning Letter 신규 건 수집 → 선별 → 메일 발송** |

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

## FDA Warning Letter 메일 (`letters`)

식약처 행정처분 메일과 같은 형식으로 FDA 경고장을 받아보기 위한 명령입니다.

```bash
python -m gmpai letters                      # 최근 7일 신규 건을 화면에 출력
python -m gmpai letters --sterile-only        # 무균·주사제·점안제 관련 건만
python -m gmpai letters --mail                # 메일로 발송
python -m gmpai letters --since 30 --no-state # 최근 30일, 발송 이력 무시하고 전부
```

동작 순서:

1. **목록 수집** — RSS → 목록 JSON → 목록 HTML 순으로 시도해 **먼저 성공한 소스**를 씁니다.
   FDA가 목록 제공 방식을 바꿔도 한 경로가 막히면 다음 경로로 넘어갑니다.
2. **기간 선별** — `--since` 일 이내. 날짜를 읽지 못한 건은 버리지 않고 남깁니다.
3. **본문 확인** — 경고장을 한 건씩 열어 키워드와 `21 CFR` 인용 조항을 뽑습니다
   (`--no-detail` 로 생략 가능).
4. **관련성 선별** — 무균 키워드(무균조작·점안제·미디어필·환경모니터링·기류시험 등)와
   CGMP 키워드(데이터 완전성·OOS·세척밸리데이션 등)로 거릅니다. `--all` 이면 선별하지 않습니다.
5. **중복 제거** — 이미 보낸 건은 상태 파일에 기록해 두고 다음 실행에서 제외합니다.
6. **발송** — 무균 관련 건을 위로, 그 다음 최신순으로 정렬해 HTML 표로 보냅니다.

### 메일 설정

비밀번호를 저장소에 두지 않기 위해 전부 환경변수로 받습니다.

| 환경변수 | 설명 |
| --- | --- |
| `GMPAI_SMTP_USER` | 보내는 계정 |
| `GMPAI_SMTP_PASSWORD` | Gmail은 2단계 인증 후 발급한 **앱 비밀번호** (계정 비밀번호 아님) |
| `GMPAI_MAIL_TO` | 수신자, 쉼표로 구분 |
| `GMPAI_SMTP_HOST` / `GMPAI_SMTP_PORT` | 기본 `smtp.gmail.com` / `587` |
| `GMPAI_MAIL_FROM` | 생략하면 `GMPAI_SMTP_USER` |
| `GMPAI_SMTP_STARTTLS` | `0` 이면 465 SSL 로 접속 |

### 자동 실행

`.github/workflows/fda-warning-letters.yml` 이 **매주 수요일 09:00 KST**(수 00:00 UTC)에 돌립니다.
FDA가 경고장 목록을 통상 **화요일(미 동부시간)** 에 갱신하기 때문입니다.

저장소 Settings → Secrets and variables → Actions 에 `GMPAI_SMTP_USER`,
`GMPAI_SMTP_PASSWORD`, `GMPAI_MAIL_TO` 를 등록하면 동작합니다. Actions 탭에서
**Run workflow** 로 수동 실행할 수 있고, `dry_run = true` 로 두면 메일을 보내지 않고
로그로만 결과를 확인합니다.

발송 이력은 `state/warning-letters-state.json` 에 남아 저장소에 커밋됩니다.
같은 경고장을 다음 주에 다시 보내지 않기 위한 것이며, 무엇을 언제 알렸는지에 대한 기록도 됩니다.

> 자동 분류는 키워드 판정입니다. 제형이나 지적의 경중을 단정하지 않으므로
> 메일에 실린 원문 링크로 반드시 확인하세요.

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

57건의 테스트가 로컬 모의 서버를 띄워 다운로드·재개정 감지·링크 탐색 폴백·HTML 오응답 차단과
경고장 파싱·소스 폴백·중복 제거·메일 생성까지 네트워크 없이 검증합니다. SMTP도 가짜 서버로 대체하므로
테스트가 실제로 메일을 보내는 일은 없습니다.

## 라이선스와 저작권

이 저장소의 코드는 MIT 라이선스입니다. 내려받는 규정 문서 자체의 권리는 각 발행기관(EC, EMA, FDA)에
있으며, 재배포 시 각 기관의 이용 조건을 따르세요. 편의를 위해 PDF 자체는 저장소에 커밋하지 않습니다
(`downloads/` 는 `.gitignore` 처리).
