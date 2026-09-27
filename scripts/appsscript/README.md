# Apps Script 판 — 식약처 메일과 같은 계정에서 보내기

식약처 행정처분 메일이 개인 구글 계정에서 발송되고 있으므로, FDA 경고장도 같은
Apps Script 프로젝트에 붙이면 발송처가 하나로 모입니다.

GitHub Actions 판(`.github/workflows/fda-warning-letters.yml`)과 **같은 일을 합니다.
둘 다 켜면 같은 메일이 두 번 옵니다 — 반드시 하나만 쓰세요.**

| | Apps Script | GitHub Actions |
| --- | --- | --- |
| 발송 계정 | 식약처 메일과 동일 | 동일하게 맞출 수 있음 |
| 앱 비밀번호 | **불필요** (MailApp 이 계정 권한으로 발송) | 필요 |
| 설정할 비밀값 | 없음 | Secrets 3개 |
| 실행 시간 제한 | 6분 (본문 30건까지만 열어봄) | 없음 |
| 코드 이력 관리 | 수동 (이 파일을 복사) | 저장소가 곧 이력 |
| 테스트 | 없음 | 57건 자동 실행 |

## 설치

1. 식약처 스크립트가 있는 Apps Script 프로젝트를 엽니다
   (script.google.com → 해당 프로젝트).
2. 왼쪽 **파일 +** → **스크립트** → 이름을 `fdaWarningLetters` 로 하고,
   `fdaWarningLetters.gs` 내용을 통째로 붙여 넣습니다.
3. **프로젝트 설정**에서 시간대가 `(GMT+09:00) 서울` 인지 확인합니다.
4. 함수 목록에서 `previewFdaWarningLetters` 를 골라 **실행**합니다.
   - 처음 실행하면 권한 승인 창이 뜹니다. 외부 사이트 접근(UrlFetchApp)과
     메일 발송(MailApp) 권한이 필요합니다.
   - **실행 기록**에 수집 결과가 찍힙니다. 메일은 나가지 않습니다.
5. 결과가 맞으면 `installWeeklyTrigger` 를 한 번 실행합니다.
   매주 **수요일 오전 9시**에 자동 발송됩니다.

## 설정 바꾸기

**프로젝트 설정 → 스크립트 속성**에서 추가합니다. 없으면 기본값을 씁니다.

| 속성 | 기본값 | 설명 |
| --- | --- | --- |
| `MAIL_TO` | `jhp5408@hanlim.com` | 수신자, 쉼표로 구분 |
| `SINCE_DAYS` | `7` | 최근 며칠 이내 건을 볼지 |
| `STERILE_ONLY` | `false` | `true` 면 무균·주사제·점안제 관련 건만 |

## 함수

| 함수 | 용도 |
| --- | --- |
| `previewFdaWarningLetters` | 메일 없이 실행 기록으로만 확인 |
| `sendFdaWarningLetters` | 실제 발송 (트리거가 부르는 함수) |
| `installWeeklyTrigger` | 매주 수요일 9시 트리거 설치 |
| `removeTriggers` | 트리거 제거 |
| `resetNotifiedState` | 발송 이력 삭제 — 지난 건을 다시 받고 싶을 때만 |

## 주의

- 키워드 표(`STERILE_KEYWORDS`, `CGMP_KEYWORDS`)는 파이썬 판
  `gmpai/warningletters.py` 와 같은 내용입니다. **한쪽만 고치면 두 판의 결과가 달라집니다.**
- 발송 이력은 스크립트 속성에 저장되며, 용량 한도(속성당 9KB) 때문에
  300건을 넘으면 오래된 것부터 버립니다.
- 실행 시간 제한이 6분이라 본문은 최대 30건까지만 열어봅니다.
  그보다 많으면 제목만으로 분류하므로 인용 조항이 비어 있을 수 있습니다.
