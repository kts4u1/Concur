# Concur 경비 자동 입력 — 작업 인계

사용자는 개발자가 아닙니다. 답은 한국어로, 쉬운 말로 합니다.

## 무엇인가

Windows PC에서 돌리는 Python 프로그램 하나(`concur_auto.py`)입니다.

1. iPhone 단축어가 iCloud 폴더에 올린 영수증 사진(`2026-10-22_123342.jpg`)을 지난달 것만 고름
2. (선택) 이미지 보정 → Windows OCR(한국어) → 날짜·금액·분류(주차/카페/식대)·병원 추출
3. `결과\영수증목록_YYYY-MM.csv`에 저장. 사용자가 엑셀로 고치면 다음 실행에도 유지
4. Playwright로 PC의 Edge(전용 프로필)를 열어 SAP Concur에서 지난달 리포트 생성 → 지난달 경비 추가 → 경비마다 Comment·Business Purpose·Receipt·LCRT·Expense Type·도시 입력, 영수증 첨부 → 저장
5. **제출은 하지 않음.** `결과\경비처리_YYYY-MM.csv`의 '확인 필요' 칸을 보고 사용자가 직접 제출

GUI(tkinter) 버튼 세 개: ① 영수증 분석만 / ② 전체 실행 / ③ 만든 리포트에 입력만. `--cli`로 명령 프롬프트 모드도 있음. `build_exe.bat`로 exe를 만듭니다.

## 내력

- 동료가 만든 Power Automate Desktop(PAD) 흐름과 `receipt_preprocess.py`(영수증 보정)가 출발점입니다.
- claude.ai 채팅 "Power Automate로 Concur 작업 자동화"에서 둘을 Python 하나로 합쳤고, 동료 설명서를 반영해 GUI 버전을 만들었습니다. 2026-11-12에 그 결과물(4개 파일)을 이 저장소로 가져왔습니다.
- PAD 흐름 원문(7.5MB 텍스트), `receipt_preprocess.py`, 동료 설명서는 저장소에 없습니다. 보정 로직은 `preprocess_receipt()`에 이미 들어 있습니다.

## 정해진 것과 이유

- **PAD 대신 Python**: 브라우저 내부에서 동작해 실행 중에도 PC를 쓸 수 있고, 파일 열기 창과 고정 대기가 없으며, exe로 나눠 줄 수 있습니다. Concur 공식 API는 회사 관리자가 앱을 등록해야 해서 브라우저 자동화를 씁니다.
- **Concur 요소 ID**는 PAD UI 요소 저장소에 있던 실제 값입니다: `#name`, `#custom11-date-input-field-input`/`#custom12-…`(리포트 기간), `#businessPurpose`, `#comment`, `#transactionAmount`, `#transactionDate-date-input-field-input`, `#location`, `[name=receiptType]`, `[name=expenseType]`, `#custom32`(LCRT), `#upload-receipt-button`.
- **경비 선택 버그 수정**: PAD는 날짜에 "10/"만 있으면 골라 11/10 같은 다른 달도 넣었습니다. `JS_SELECT_MONTH`가 월·연도를 정확히 비교합니다.
- **연도 고정(2026) 제거**: 처리할 달은 날짜로 계산합니다.
- **업로드는 원본 사진**: 보정본은 2배 확대 PNG라 4배 가까이 커지고 Concur는 1장 5MB 제한이 있습니다. 넘으면 JPEG로 줄인 사본을 올립니다.
- **로그인**: 전용 Edge 프로필(`브라우저_프로필_폴더`)에서 처음 한 번 SSO 로그인하면 유지됩니다.
- **동료 설명서와 일부러 다른 점** (모두 `concur_settings.json`에서 바꿀 수 있음):
  - 같은 날짜·금액 영수증: 건너뛰지 않고 순서대로 첨부하고 '확인 필요'로 표시
  - Undefined인 Expense Type: 주차는 `Parking & Toll`, 나머지는 `Other Staff Expenses`를 고르고 '확인 필요'로 표시
  - 식대 Comment: 설명서 문구 "Overtime meal fee for service"를 따름 (흐름 파일은 "Overtime meal during service")
- **한국어 화면 기준**: 버튼은 '리포트 생성', '경비 저장' 같은 글자로 찾습니다. Concur 언어가 바뀌면 설정의 `화면_문구`만 고칩니다.
- OCR 키워드의 이상한 글자('주자', '메리카', '중남대' 등)는 OCR이 잘못 읽은 결과에 맞춘 것이라 일부러 둔 것입니다.

## 검증 상태

- 이전 채팅: 로직 테스트, 가짜 Concur 페이지에서 처음부터 저장까지, GUI 뜨는지 확인. 그 테스트 코드와 가짜 페이지는 넘어오지 않았습니다.
- 이 저장소: `tests/`에 로직 테스트와 경비 선택 JavaScript 테스트(Chromium)가 있습니다.
- **실제 Concur와 실제 Windows OCR에서는 아직 한 번도 돌려 보지 않았습니다.**

## 다음 할 일

1. 사용자가 ①을 돌려 `영수증목록_YYYY-MM.csv`를 보내 주면, 식대 키워드(지금 '버터' 하나뿐이라 식당 영수증 대부분이 '분류 없음'이 될 수 있음)와 병원 표를 실제 OCR 결과에 맞춰 보강
2. 첫 실제 ②/③ 실행의 `실행기록` 폴더(로그·화면 캡처)를 받아 셀렉터·대기 수정
3. 회사 Edge 정책이 브라우저 자동화를 막으면 코드로는 해결할 수 없음 → IT 확인 안내
4. 기본 병원 목록은 대전·충청권이고 기본 도시는 Daejeon. 담당 지역이 바뀌면 `병원목록.csv`와 `기본_도시`를 바꿈

## 이 클라우드 환경(Linux)에서 작업할 때

- Windows에서만 되는 것: winrt OCR, Edge(`channel="msedge"`), `os.startfile`, exe 빌드. 여기서는 실행할 수 없습니다.
- `concur_auto.py`는 맨 위에서 표준 라이브러리만 import합니다(cv2·winrt·playwright·tkinter는 쓸 때 import). 이 구조를 유지해야 Linux에서 테스트가 돌아갑니다.
- 테스트:
  ```bash
  pip install -r requirements-dev.txt
  python -m pytest
  ```
  Chromium은 `/opt/pw-browsers/chromium`을 씁니다(`PW_CHROMIUM`으로 바꿀 수 있음). `playwright install`은 하지 않습니다.
- `requirements.txt`는 Windows용(winrt 포함)이라 Linux에서 설치하지 않습니다.
- **이 저장소는 공개(public)입니다.** 사번이 들어간 경로, 회사 계정, 실제 영수증 사진, 개인 설정 파일(`concur_settings.json`, `병원목록.csv`, `결과/`, `실행기록/`)은 커밋하지 않습니다.
- `build_exe.bat`는 CRLF를 유지합니다(`.gitattributes`).
