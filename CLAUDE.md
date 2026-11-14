# Concur 경비 자동 입력 — 작업 인계 메모

> 이 파일은 Claude가 이 저장소에서 작업을 시작할 때마다 자동으로 읽는 인계 메모입니다.
> 파일 이름(CLAUDE.md)은 바꾸면 안 됩니다. 내용은 사람이 읽어도 되도록 한국어로 씁니다.

사용자는 개발자가 아닙니다. 답은 한국어로, 쉬운 말로 합니다. 작업 단계 설명도 한국어로 씁니다.

## 무엇을 하는 프로그램인가

회사 Windows PC에서 돌리는 Python 프로그램 하나(`concur_auto.py`)입니다.

1. 영수증 폴더 여러 곳(iCloud 사진·iCloud Drive·OneDrive 갤럭시 사진·구글 드라이브·카카오톡 받은 파일·'가져온 영수증' 등, 하위 폴더 포함)에서 지난달 1일부터 그 달이 끝나고 60일 안에 찍은 사진을 고릅니다. 파일 이름·형식(jpg/png/heic)은 상관없습니다.
2. Windows에 들어 있는 한국어 글자 인식(OCR)으로 사진을 그대로 읽어 영수증이 아닌 사진은 뺍니다. 금액·날짜를 못 읽은 영수증만 (선택) 보정해서 한 번 더 읽고, 날짜·금액·분류(주차/카페/식대)·병원을 뽑습니다. **어느 달 영수증인지는 영수증에 적힌 날짜로 정합니다.**
3. 결과를 `결과\영수증목록_YYYY-MM.csv`에 저장합니다. 사용자가 엑셀로 고치면 다음 실행에도 고친 값이 쓰입니다.
4. 이 프로그램 전용 Edge 창을 열어 Concur에서 다음을 합니다.
   - 지난달 리포트 만들기
   - 지난달 경비를 리포트에 넣기
   - 경비마다 Concur 화면 위에서 아래 순서로 Expense Type, Business Purpose, 도시, LCRT, Receipt Status, (Personal Expense), Comment를 채우고 영수증 사진 첨부
   - 저장
5. **제출은 하지 않습니다.** `결과\경비처리_YYYY-MM.csv`의 '확인 필요' 칸을 보고 사용자가 직접 제출합니다.

프로그램 창은 [실행]·[영수증 불러오기]·[백업·되살리기] 세 탭입니다. [실행] 탭에 버튼이 세 개 있습니다: ① 영수증 분석만 / ② 전체 실행 / ③ 만든 리포트에 입력만.
창 없이 명령 프롬프트로 쓰는 방법(`--cli`)도 있고, `build_exe.bat`를 더블클릭하면 exe 파일로 만들어집니다.

## 지금까지의 경과

- 출발점은 동료가 만든 Power Automate Desktop 흐름과 영수증 보정 스크립트(`receipt_preprocess.py`)입니다.
- claude.ai 채팅 "Power Automate로 Concur 작업 자동화"에서 둘을 Python 프로그램 하나로 합쳤고, 동료 설명서를 반영해 버튼으로 쓰는 창 버전을 만들었습니다.
- 2026-11-12에 그 결과물(파일 4개)을 이 저장소로 가져왔습니다.
- Power Automate 흐름 원문(7.5MB 텍스트), `receipt_preprocess.py`, 동료 설명서는 저장소에 없습니다. 보정 기능은 이미 `concur_auto.py` 안에 들어 있어서 따로 필요하지 않습니다.

## 정해 둔 것과 그 이유

- **Power Automate 대신 Python을 쓰는 이유**
  - 마우스를 흉내 내지 않고 브라우저 안에서 직접 동작해서, 실행 중에도 PC를 쓸 수 있습니다.
  - 파일 열기 창을 누르거나 정해진 시간만큼 기다리는 동작이 없어 덜 깨집니다.
  - exe 파일로 만들어 동료에게 나눠 줄 수 있습니다.
  - Concur 공식 연결 방법(API)은 회사 관리자가 따로 등록해 줘야 해서 개인이 쓰기 어렵습니다.
- **프로그램 판 표시:** `APP_VERSION`이 실행 기록 첫 줄과 창 제목에 나옵니다. 사용자가 최신 파일을 받았는지 확인할 때 씁니다. 고칠 때마다 바꿉니다.
- **Concur 화면의 입력칸 이름**은 Power Automate 흐름에 저장돼 있던 실제 값을 그대로 씁니다.
  - 리포트: `#name`(이름), `#custom11-date-input-field-input`·`#custom12-date-input-field-input`(기간)
  - 경비: `#businessPurpose`, `#comment`, `#transactionAmount`(금액), `#transactionDate-date-input-field-input`(날짜), `#location`(도시), `[name=receiptType]`(Receipt), `[name=expenseType]`(Expense Type), `#custom32`(LCRT), `#upload-receipt-button`(영수증 첨부 버튼)
- **경비를 잘못 고르던 문제를 고쳤습니다.** Power Automate 흐름은 날짜에 "10/"만 들어 있으면 골라서, 11/10처럼 다른 달 경비까지 넣었습니다. 지금은 월과 연도를 정확히 비교합니다(`JS_SELECT_MONTH`).
- **연도가 2026으로 고정돼 있던 문제를 고쳤습니다.** 처리할 달은 오늘 날짜로 계산합니다.
- **영수증은 원본 사진을 올립니다.** 보정한 사진은 2배로 키운 PNG라 크기가 4배 가까이 커지는데, Concur는 1장에 5MB까지만 받습니다. 원본도 5MB를 넘으면 자동으로 줄인 사본을 올립니다.
- **로그인:** 전용 Edge 창에서 처음 한 번 회사 계정으로 로그인하면 다음부터는 로그인 상태가 유지됩니다.
- **동료 설명서와 일부러 다르게 한 부분** (모두 설정 파일 `concur_settings.json`에서 바꿀 수 있습니다)
  - 같은 날짜·같은 금액 영수증이 여러 장이면: 건너뛰지 않고 순서대로 첨부한 뒤 '확인 필요'로 표시합니다.
  - Expense Type이 Undefined이면: 자동 입력 편집의 항목에 정한 Expense Type을 고릅니다. 맞는 항목이 없으면 Undefined 그대로 두고 '확인 필요'(2026-11-14부터, 예전에는 '분류 없음'이면 `Other Staff Expenses`).
  - 식대 Comment 문구: 흐름 파일("Overtime meal during service")과 설명서("Overtime meal fee for service")가 달라서 설명서 문구를 따랐습니다.
- **Concur는 한국어 화면을 기준으로 하고 영어도 받습니다.** 버튼은 '리포트 생성', '경비 저장' 같은 글자로 찾는데, `화면_문구`에 `한국어|영어`처럼 여러 글자를 적어 두어 둘 중 하나가 보이면 됩니다(`_alts`, `_re`). 2026-11-12에 로그인 화면이 갑자기 영어(Sign In / Next)로 바뀌어 [다음]을 못 누른 일이 있어서, Edge를 한국어로 열고(`launch_options`: `locale=ko-KR`, `--lang=ko-KR`, 번역 창 끔) 영어 글자도 같이 찾게 했습니다. 예전 설정 파일에 저장된 한국어 하나짜리 문구는 `load_settings`가 새 기본값으로 바꿉니다(사용자가 직접 바꾼 값은 그대로). 영어 [이메일 링크로 로그인] 버튼 글자는 확인하지 못해 여러 후보와 'email…link' 정규식으로 찾습니다.
- **단축어 없이 쓰기 (2026-11-12 추가):** 사용자가 단축어 앱이 불편하다고 해서, 카메라로 찍기만 해도 되게 바꿨습니다.
  - 촬영 시각: 파일 이름 날짜(`2026-10-22_123342`) → 사진 속 촬영 날짜(EXIF) → 파일 수정 날짜 순서로 봅니다. 파일 수정 날짜가 처리할 달보다 앞인 파일은 열지 않습니다(사진이 수천 장인 iCloud 사진 폴더에서도 빠르게). 촬영 날짜는 `결과\photo_date_cache.json`에 기억해 둡니다.
  - 영수증 판단(`looks_like_receipt`): 금액이 있고 (영수증 글자 / OCR 날짜 / 분류 키워드) 중 하나가 있으면 영수증. 금액이 없어도 영수증 글자가 두 개 이상이면 영수증. 글자 목록은 설정 `영수증_판별_단어`.
  - 영수증이 아닌 사진은 목록 맨 아래에 `사용` X, "영수증 아닌 것 같음(자동 제외)"으로 남겨, 틀렸으면 사용자가 O로 바꿀 수 있게 했습니다. 이를 위해 목록에 `촬영 시각`, `[자동] 사용` 칸을 추가했습니다(예전 목록은 `[자동] 사용`을 O로 봄).
  - **보정 줄임 (2026-11-13, 사용자 지적 "png로 바뀌며 용량이 엄청 커지는데 도움이 되나"):** 예전에는 영수증마다 2배 확대 PNG(수십 MB)를 `보정본_폴더`(iCloud Drive 안 → iCloud 용량도 차지)에 남기고, OCR 임시 PNG도 `실행기록\임시`에 쌓였습니다. iPhone 사진(4032×3024)은 이미 커서 2배 확대는 효과가 작습니다. 이제 원본을 그대로 읽고, 영수증인데 금액이나 날짜를 못 읽었을 때만 `make_enhanced_copy`로 임시 JPG(작은 사진만 확대)를 만들어 다시 읽은 뒤 지웁니다. 보정본 글자가 정보를 더 많이 읽었을 때만(`info_score`) 씁니다. 결과는 `ocr_cache.json`에 `파일이름#보정`으로 기억합니다. OCR 임시 사본도 JPG(품질 95)로 만들고 읽은 뒤 지우며, `run_job`이 시작·끝에 `임시` 폴더를 비웁니다. 예전 보정본 폴더가 있으면 지워도 된다고 실행 기록에 안내합니다(`old_processed_note`, 자동으로 지우지는 않음).
  - HEIC(iPhone 기본 형식)는 `pillow-heif`로 읽고, OCR 전에는 PNG로, Concur에는 JPEG로 바꿔 올립니다. jpg도 OCR 전에 회전 정보를 반영한 PNG로 바꿉니다(세로 사진이 누운 채로 읽히지 않게).
  - iCloud 사진 폴더를 쓰면 개인 사진 전체가 회사 PC로 내려온다는 점을 사용자에게 알렸습니다(README 방법 1 주의).
- **영수증 날짜 최우선 (2026-11-12, 사용자 요청):** "영수증을 한 달 뒤에 찍을 수도 있으니 영수증에 있는 날짜를 최우선으로, 촬영 날짜는 참고만" 하라는 요청에 따라 바꿨습니다.
  - 읽는 사진 범위(`photo_window`): 처리할 달 1일 ~ 그 달이 끝나고 `영수증_촬영_허용_일수`(기본 60일). 처리할 달 전에 찍은 사진은 그 달 영수증일 수 없어서 뺍니다.
  - 이 달 영수증인지(`place_in_month`): OCR 날짜 중 처리할 달 날짜가 있으면 이 달(그 날짜를 대표 날짜로 앞에 둠). OCR 날짜가 있는데 모두 다른 달이면 "다른 달 영수증(날짜)(자동 제외)". OCR 날짜가 없으면 촬영 날짜가 처리할 달일 때만 이 달("날짜 못 읽음(촬영 날짜 사용)" 표시), 아니면 "날짜 못 읽음: 영수증 날짜를 적어 주세요(자동 제외)".
  - 사용자가 목록에서 날짜를 고치면 그 날짜로 다시 정합니다. 날짜를 적었으면 '영수증 아님'으로 빠졌던 사진도 영수증으로 봅니다. `사용` 칸을 직접 바꾼 것은 마지막에 그대로 따릅니다.
  - 경비 매칭(`find_receipt`): 영수증 대표 날짜가 같은 것을 먼저 보고, 촬영 날짜 같은 다른 후보 날짜는 나중에 봅니다.
  - 목록 순서: 이 달 영수증(영수증 날짜 순) → 날짜 못 읽음 → 다른 달 영수증 → 영수증 아님.
- **기본 영수증 폴더 이름:** 사용자 폴더 이름이 기본값(`Concur_Expenses`)과 달라서(`s` 없음), 기본 폴더가 없으면 iCloudDrive 안에서 Concur로 시작하는 폴더가 하나뿐일 때 그 폴더를 씁니다(`default_receipt_folder`, 설정 파일을 처음 만들 때만). 폴더가 없으면 오류 문구에 비슷한 이름의 폴더를 알려 줍니다.
- **이번 달도 처리 (2026-11-12, 사용자 요청):** 창의 처리할 달 목록 맨 위에 이번 달을 넣었습니다(`recent_months(include_current=True)`). 처음 선택은 계속 지난달입니다. 아직 끝나지 않은 달(`is_unfinished_month`)을 고르면 ②·③ 전에 확인 창을 띄우고 실행 기록에도 "아직 끝나지 않은 달" 안내를 남깁니다. 명령 프롬프트 메뉴에서도 처리할 달을 물어봅니다.
- **이메일 링크 자동 로그인 (2026-11-12, 사용자 요청):** 사용자의 실제 로그인 순서는 이메일 입력 → [다음] → 로그인 방법 화면(회사 SSO / SAP Concur 암호 / 이메일 링크로 로그인) → [이메일 링크로 로그인] → Outlook에 "SAP Concur 로그인 요청" 메일(약 10초, 한 시간 동안 한 번만 쓰는 링크) → 메일의 [SAP Concur 로그인] → us2 홈 화면("경비 리포트 생성" 카드).
  - `wait_login`이 1초마다 화면을 보고 처리합니다: 로그인됨(`is_logged_in`: concursolutions.com이고 로그인 주소·로그인 화면이 아니며 "리포트 생성" 글자가 있음) → 진행 / 이메일 화면(`_email_step`) → 칸이 비면 설정 `로그인_이메일`을 넣고 [다음] / 로그인 방법 화면(`_email_link_step`) → [이메일 링크로 로그인]은 **한 번만** 누르고 Outlook에서 링크를 찾아 **프로그램의 Edge 창에서** 엽니다.
  - 메일 링크를 사용자가 직접 누르면 평소 Edge에서 열려 프로그램 창(전용 프로필)은 로그인되지 않으므로, 링크를 꺼내 여는 방식이 필수입니다.
  - **로그인 화면이 영어면 메일도 영어로 옵니다**(제목 "Sign-In Request to SAP Concur", 버튼 "Sign in to SAP Concur"). 2026-11-12에 한국어 제목만 찾다가 120초 뒤 실패한 일이 있어서, `로그인_메일_제목`·`로그인_메일_버튼`을 `한국어|영어`로 바꾸고, 제목이 달라도 보낸 사람 주소에 concur가 있으면 보며(`_is_login_mail`), 버튼 글자가 달라도 'Concur'와 '로그인/Sign in/Log in'이 든 링크를 고릅니다(`SIGNIN_LINK_TEXT`). 예전 설정 파일에 저장된 한국어 하나짜리 값(최상위 설정 포함)은 `load_settings`가 새 기본값으로 바꿉니다.
  - Edge를 `locale=ko-KR`로 열어도 Concur 로그인 화면은 영어로 나왔고, `--disable-features=Translate`로도 Edge 번역 창이 떴습니다(효과 없음). 번역 창은 동작에 지장이 없고, 그 창에서 '영어 번역 안 함'을 한 번 고르면 전용 프로필에 기억됩니다.
  - Outlook 읽기(`find_login_link_in_outlook`): pywin32로 클래식 Outlook의 받은 편지함을 최근 순으로 30개까지 보고, 요청 시각 2분 전 이후에 온 제목 "SAP Concur 로그인 요청" 메일에서 글자가 "SAP Concur 로그인"인 링크(`extract_login_link`)를 찾습니다. 3초마다, 최대 `로그인_메일_대기_초`(120초). pywin32 시각은 PC 시각에 UTC 표시가 붙어 오므로 tzinfo를 떼고 비교합니다. 메일 보안이 링크를 감싸도(safelinks 등) 그대로 엽니다.
  - 못 찾으면 붙여넣기 창(`ui.ask_text`)을 띄우고, 취소하면 직접 로그인을 기다립니다. 창의 체크를 끄면 `로그인_방법`="직접"(예전 방식).
  - 실제 Concur 로그인 화면의 HTML 구조는 모릅니다. 첫 실제 실행에서 [다음]을 누르지 못했다는 보고를 받아(로그인 주소 `us2.concursolutions.com/nui/signin?targetURL=%2Fhome`, 이메일은 저장돼 있음), 버튼을 여러 방법으로 찾게 보강했습니다(`_button_like`: 역할 button/link, 글자가 든 button·div·a, `input[type=submit][value=다음]`, 글자). 이메일 칸도 라벨·type=email·username 등으로 찾습니다(`_email_box`).
  - [다음]을 누르지 못하거나 눌러도 화면이 안 바뀌면 이메일 칸에서 Enter를 번갈아 시도합니다. 세 번째에도 안 되거나 15초 동안 아무것도 못 찾으면 화면의 버튼 글자·입력칸 목록과 화면 캡처를 실행 기록에 남깁니다(`_diagnose_login`). 다음에 못 누르면 이 기록을 받아 고칩니다.
  - Playwright가 붙이던 '--no-sandbox' 경고줄을 없앴습니다(`chromium_sandbox=True`).
  - exe에는 pywin32 날짜 처리 때문에 `--hidden-import win32timezone`을 넣었습니다.
- **병원 찾기 (2026-11-12, 사용자 요청):** 사용자가 '병원 없음'이 많다고 해서 세 가지를 더했습니다. 우선순위: 영수증 글자(`HOSP_FROM_OCR`) = 직접 고침 > Concur 가게 이름(`HOSP_FROM_VENDOR`) > 같은 날 영수증(`HOSP_FROM_DAY`). 목록·결과 파일의 `병원 근거` 칸에 표시합니다.
  - **병원 정보 엑셀:** 사용자 회사 SharePoint 엑셀의 "MR System Information" 탭이 병원 정보의 출처입니다(영문 이름 없음, 한글만). 이 클라우드 환경에서는 회사 SharePoint에 들어갈 수 없고, 회사 자료라 **저장소에 올리면 안 됩니다.** 그래서 PC의 프로그램이 직접 읽습니다(`import_hospitals_from_excel`, openpyxl). 창의 [병원 엑셀 가져오기]로 파일을 고르면 설정 `병원_엑셀_파일`에 저장하고, 실행할 때마다 새 병원만 `병원목록.csv`에 더합니다(이미 있는 찾을 글자는 건너뜀 → 사용자가 고친 영문명 유지). 병원 이름 열은 병원처럼 보이는 값이 가장 많은 열을 자동으로 고르거나 `병원_엑셀_열`(열 제목/열 문자)을 씁니다.
  - **한글→영문:** `hospital_english_name`(접두 국립/도립/시립, 접미 대학교병원→University Hospital, 성모병원→St. Mary's Hospital, 의료원→Medical Center 등 + 나머지는 `romanize_korean` 국어의 로마자 표기 간단판). 공식 영문명과 다를 수 있어 비고 칸에 "자동 변환 … (영문 이름 확인)"을 남깁니다. 찾을 글자(`hospital_key`)는 '(학교)병원/의료원…'을 뗀 글자, 3글자 미만이면 전체 이름(지역명만 남는 것 방지).
  - **같은 날 영수증(`fill_same_day_hospitals`):** 같은 대표 날짜의 사용 영수증 중 영수증 글자/직접 고침으로 찾은 병원이 한 곳이면 채움. 여러 곳이면 주차 영수증 것만, 그래도 여러 곳이면 비움. 사용자가 병원을 지운 영수증은 채우지 않음.
  - **Concur 가게 이름(`apply_vendor_hospital`):** 경비 편집 화면의 가게 이름 칸(설정 `가게_이름_칸` 선택자 목록, `가게_이름_라벨` 라벨)을 읽고, 못 읽으면 리포트 경비 목록(`JS_READ_ROWS`)에서 같은 날짜·금액 줄 글자를 씁니다. 병원 찾을 글자를 로마자로 바꿔 `normalize_roman`(eo/o/u 같게, k/g·t/d·p/b·ch/j·r/l 같게 등)으로 비교하고 6글자 이상 맞아야 인정합니다. **실제 Concur의 가게 이름 칸 이름은 모릅니다(추측).** 결과 파일 `가게 이름` 칸에 읽은 값이 나오니 첫 실행에서 확인하세요.
- **도시·금액·병원 보강 (2026-11-12, 사용자 요청):**
  - 기본 도시는 `Gwangju (ex Kwangju), KOREA, REPUBLIC OF`(사용자 담당 지역). 예전 설정 파일의 대전 기본값은 `load_settings`가 한 번만 광주로 바꿉니다(`_기본_도시_광주로_바꿈`). Concur 도시 목록에서 확인한 이름: Daejeon (ex Taejon), Gwangju (ex Kwangju), Jeonju, Iksan, Mokpo. 나머지 `DEFAULT_CITIES`는 추측이라 `choose_city`가 정확히 못 찾으면 앞 단어로 검색해 그 단어와 KOREA가 든 항목을 고릅니다.
  - 도시 찾기(`detect_city`): 영수증 글자의 'OO광역시/시/군' 중 먼저 나온 것(경기도 광주 제외) → 병원 찾을 글자 앞의 도시(순천향 제외) → 같은 날 영수증(`fill_same_day_cities`) → 기본 도시. Concur에서는 `city_to_choose`: 칸이 비면 영수증 도시/기본 도시, 영수증 도시가 지금 값과 다르면 영수증 도시. 목록 `도시` 칸을 고치면(`resolve_city`) 그 도시.
  - 6,000원 영수증이 안 붙은 원인은 사용자 자료를 못 봐서 확정하지 못했습니다. 흔한 원인을 고쳤습니다: 쉼표 없는 금액(`AMOUNT_NEAR_WORD`: 합계/금액/요금 등 바로 뒤 숫자), OCR의 O→0·쉼표 띄어쓰기(`_fix_amount_text`), 두 자리 연도·날짜 글자 뒤 8자리(`DATE_SHORT_YEAR`, `DATE_COMPACT`), 금액 못 읽은 영수증이 그날 한 장이면 날짜로 첨부하고 '확인 필요'(`find_receipt_by_date_only`).
  - 병원이 목록에 없어도 영수증 글자의 'OO병원/OO의료원'(띄어쓰기 기준 낱말, 법인·재단 이름 앞부분은 버림, 동물병원 제외)을 쓰고(`find_hospital_name_in_text`), `build_receipts(..., hospital_csv)`가 병원목록.csv에 "영수증에서 찾음"으로 더합니다(`add_found_hospitals`). 영문 변환은 '대병원'→University Hospital, 도시로 시작하면 띄어 씀(목포한국 → Mokpo Hanguk).
- **영수증 첨부 보강 (2026-11-13, 사용자 보고):** 99,000원 주차 경비에서 Comment·Business Purpose는 들어갔는데(=영수증은 찾음) 영수증이 안 붙었습니다. 예전 코드는 `#upload-receipt-button`을 1.5초 안에 못 찾으면 '이미 첨부됨'으로 보고 건너뛰었습니다. 로그인 후 Concur 화면도 영어였고(Save Expense, 3 of 4, Upload New Receipt / Choose From Available Receipts) 영수증 칸이 늦게 뜰 수 있습니다. 이제 `attach_receipt`가 최대 `upload_wait_s`(15초) 동안 업로드 버튼(id 또는 `화면_문구` "영수증 업로드" 글자) 또는 '첨부된 영수증' 표시(Detach/Delete Receipt 등, 추측)를 기다립니다. 버튼이 없으면 숨은 `input[type=file]`에 바로 넣고, 그것도 없으면 '확인 필요'와 화면 캡처를 남깁니다. 올린 뒤 버튼이 30초 안에 안 사라지면 "영수증이 올라갔는지 확인".
- **병원 목록 = 영수증에서 확인한 병원만 (2026-11-13, 사용자 요청):** 사용자가 동료분 기본 목록(대전·충청권 20줄)을 "다 없애고 영수증에서 확인한 것만" 두라고 했습니다. `DEFAULT_HOSPITALS`는 원광대(Wonkwang University Hospital)·목포한국(Mokpo Hankook Hospital) 두 곳뿐입니다(잠깐 넣었던 광주·전라권 17곳도 뺌). `load_hospital_rows`가 예전 기본 줄(비고가 '기본 목록(동료분 흐름)'·'기본 목록(광주·전라권)'이거나 `OLD_DEFAULT_HOSPITALS`와 똑같은 줄)을 지우고, '영수증에서 찾음'으로 자동 만든 원광대·목포한국 영문명은 공식 영문명으로 바꿉니다. 사용자가 넣은 줄은 그대로. 테스트는 매칭 동작을 보던 곳에 `OLD_DEFAULT_HOSPITALS`를 씁니다.
- **병원 엑셀 여러 탭 (2026-11-13, 사용자 요청):** 사용자가 병원 정보가 더 많은 다른 회사 SharePoint 엑셀(MR·CT·XR 탭)을 알려 줬습니다(링크·파일 이름은 회사 자료라 저장소에 적지 않음). `병원_엑셀_탭` 기본값 `MR|CT|XR|MR System Information`(예전 저장값 'MR System Information'은 자동으로 바꿈). `_find_sheets`: 같은 이름 → 그 이름으로 시작 → 3글자 이상일 때만 포함(‘CT’가 ‘Contact’에 걸리지 않게). 탭마다 한글 병원 열을 고르고, 영문 병원 이름 열(`ENGLISH_HOSPITAL`: Hospital/University/Medical/Center…가 든 영어 값이 3개 이상)이 있으면 그 이름을 씁니다(`read_hospitals_from_excel`). 병원 열이 없는 탭은 안내만 남기고, 어느 탭에서도 못 찾으면 오류. 가져올 때 자동 영문명(비고 '자동 변환'·'영수증에서 찾음')은 엑셀 영문명으로 바꿉니다. 실제 엑셀의 열 구조는 모릅니다.
- **도시 맞추기 바꿈 (2026-11-13):** 99,000원 경비 도시가 Seoul로 들어가 있었습니다(카드 정보로 미리 들어갔거나 주차 회사 본사 주소). 이제 영수증이 있는 경비는 영수증 도시, 없으면 기본 광주로 맞춥니다(미리 들어간 값이 있어도). 영수증 없는 경비는 비어 있을 때만 기본 도시. `detect_city`는 `도시_낮은_우선순위`(기본 서울)를 다른 도시가 없을 때만 씁니다.
- **자동 입력표 (2026-11-13, 사용자 질문 "새 expense 항목의 자동 입력은? 저장된 자동 입력 값은 어떻게 고치나?"):** 분류별 찾을 글자·Comment·Business Purpose·Undefined일 때 Expense Type을 `자동입력표.csv`(엑셀, 한 줄=한 분류, 위쪽 줄이 먼저)로 옮겼습니다. 창의 [자동 입력표 편집]으로 엽니다. 처음에는 설정 파일 값(`category_rows_from_cfg`)으로 만들고, 실행할 때마다 `apply_category_table`이 읽어 `cfg`의 `키워드`·`분류_우선순위`·`입력_문구`·`Undefined일_때_선택할_유형`을 덮어씁니다(나머지 코드는 그대로). 예전에는 `CATEGORIES`가 코드에 고정돼 있어 새 분류를 목록 `분류` 칸에 적어도 무시됐는데, 이제 `apply_user_edits(..., categories=...)`로 표의 분류를 받습니다. `{병원}`도 `{hospital}`로 받고, 병원이 없으면 ' in/at {hospital}'을 뺍니다(`texts_for`). 표가 비어 있으면 설정 값을 씁니다. `.gitignore`에 넣었습니다(개인 파일).
- **Concur 화면이 바뀔 때 대비 (2026-11-13, 사용자 질문):** 입력칸·목록 위치를 코드에서 설정 `화면_요소`로 옮겼습니다(`DEFAULT_ELEMENTS`, 코드는 `self.sel(key)`). 텍스트 입력칸은 `self.field(key)`가 위치로 찾고, 못 찾으면 `화면_요소_라벨`(칸 이름, 한국어|영어)로 찾고, 끝내 못 찾으면 "'키' 칸을 찾지 못했습니다… 설정 '화면_요소'를 고치세요"라는 오류를 냅니다. 창의 [화면 조사](명령 프롬프트 `--inspect`, 메뉴 4)는 로그인 후 사용자가 연 화면마다 `실행기록\화면조사_*.txt`(설정 위치를 이 화면에서 찾는지, 보이는 입력칸의 종류·id·name·칸 이름·값, 버튼 글자)와 `.png`를 남깁니다(`inspect_screen`, `JS_INSPECT`).
  - **다음 Claude가 고칠 때:** 사용자에게 화면조사 txt·png와 실행 기록을 받아, '없음'인 키의 새 id/name을 `DEFAULT_ELEMENTS`에 반영하고(사용자 설정 파일에 예전 값이 저장돼 있으면 `load_settings`에서 옮겨 주기), 버튼 글자는 `화면_문구`에 `|`로 덧붙이고, `tests/test_screen_change.py`처럼 가짜 화면 테스트를 더합니다. 아직 설정으로 못 옮긴 것: `JS_SELECT_MONTH`의 체크박스 클래스, 진행 표시("3 of 4") 정규식, 로그인 확인용 `#cnqr-app-content`.
- **끝나도 Edge 열어 두기 (2026-11-13, 사용자 요청):** "다시 직접 보고 수정해야 하니 닫지 말아 달라". `작업_후_Edge_열어두기`(기본 true)면 ②·③이 끝나거나 오류가 나도 `run_job`의 finally에서 `bot.wait_until_closed(stop_flag)`로 사용자가 Edge 창을 닫을 때까지(또는 [중지]) 기다린 뒤 정리합니다. 그동안 작업 스레드가 살아 있어 실행 버튼은 눌리지 않습니다. 사용자가 취소(Cancelled)한 경우는 바로 닫습니다.
- **Concur Expense Type 목록 가져오기 (2026-11-13, 사용자 질문 "골라야 할 항목이 엄청 많은데 쉽게 넣는 방법"):** 이 클라우드 환경은 사용자 회사 계정으로 Concur에 들어갈 수 없어 직접 볼 수 없습니다. 대신 [화면 조사]가 경비 편집 화면에서 Expense Type 드롭다운을 열어 목록을 모두 읽고(`read_dropdown_options`: 화면에 보이는 만큼만 그리는 목록도 아래로 내리며 읽음, Recently Used 같은 제목은 뺌) 화면조사 파일에 적은 뒤, `add_expense_types_to_table`이 자동입력표에 없는 항목을 '(분류 없음)' 위에 줄로 더합니다(찾을 글자는 비워 둠 → 사용자가 적기 전에는 영수증에 안 맞춰짐, 이미 Expense Type 칸에 쓰는 이름·Undefined는 뺌). 확인한 항목 일부: Car mileage, Parking & Toll, Meals & Drinks, Other, Training. 식대를 지금처럼 Other Staff Expenses로 둘지 Meals & Drinks로 바꿀지는 사용자 결정.
- **자동 입력 편집 창 (2026-11-13, 사용자 요청 "엑셀로는 힘들다, Expense Type을 고르면 넣을 글을 직접 만드는 창"):** 창의 [자동 입력 편집] → `CategoryEditor`(tkinter Toplevel). 왼쪽 목록(● 설정됨 / ○ 안 함, 찾기·설정한 항목만), 오른쪽 입력칸(항목 이름·찾을 글자·Comment·Business Purpose·Expense Type 콤보), [새 항목]·[삭제]·▲▼·[저장]·[엑셀로 열기]·[닫기](저장 안 했으면 물어봄). 저장 전 `category_problems`(같은 이름, 같은 찾을 글자가 여러 항목, 찾을 글자는 있는데 문구 빈 항목). 파일 읽기·쓰기는 `load_category_rows`/`save_category_rows`(아래 기본 Expense Type 칸과 ‘(분류 없음)’ 줄은 2026-11-14에 없앰). 사용자는 실제로 10개 정도만 쓴다고 함.
- **넣을 경비가 없을 때 (2026-11-13, 사용자 요청):** "No Available Expenses"(Select from Available Expenses (0)) 화면에서 예전에는 체크박스를 60초 기다리다 오류로 끝났습니다. 이제 `add_available_expenses`가 체크박스와 `화면_문구` "사용 가능한 경비 없음" 중 먼저 뜨는 것을 보고, 없으면 `NoExpenses`(오류 아님)를 냅니다. 그 달 경비가 0건일 때, ③에서 연 리포트가 빈 리포트(빈 리포트 카드가 보임)일 때(`open_first_expense`)도 같습니다. `run_job`은 "넣을 경비 없음" 안내를 띄우고 Edge 창은 열어 둡니다(사용자가 닫을 때까지).
- **도시 찾기 보강 (2026-11-13, 사용자 요청 "영수증에서 지역이 보이면 그쪽으로"):** `DEFAULT_CITIES`에 광주·전라 시·군 전부와 주요 도시를 더했습니다(Concur 이름은 추측, 설정 파일의 `도시_목록`은 사전 합치기라 새 기본 도시가 자동으로 들어감). `detect_city`가 'OO시/군' 말고도 도 이름 뒤의 도시(`전북 익산…`, `PROVINCES`)와 광역시 뒤의 구(`대전 유성구…`, `METRO_CITIES`)도 봅니다(경기 광주 제외). `fill_expense`는 Receipt Status·LCRT·Expense Type·도시를 하나 못 골라도 `safely`로 '확인 필요'에 남기고 다음 칸과 **영수증 첨부까지 계속**합니다(예전에는 예외로 그 경비의 첨부가 빠졌음). 찾은 도시를 못 고르면 기본 도시로 다시 고릅니다(`_try_choose_city`).
- **항목 고르기·빠른 입력 (2026-11-13, 사용자 요청):** 편집 창의 Expense Type·기본 Expense Type은 Concur 항목 목록(`type_choices` = 설정 `Concur_경비_유형`(기본 `KNOWN_EXPENSE_TYPES`: 사용자 화면에서 본 Car mileage, Meals & Drinks, Other, Other Staff Expenses, Parking & Toll, Training) ∪ 표의 Expense Type)에서 고르는 콤보입니다. [화면 조사]가 읽은 Expense Type은 자동입력표에 줄을 늘리지 않고 설정 목록에 더합니다(`remember_expense_types`). 사용자가 "글을 천천히 한 글자씩 치는 게 원래 그런가" 물어서, `type_text`가 `fill`로 한꺼번에 넣고 값이 안 들어가면 한 글자씩(0.01초) 치게 바꿨습니다(설정 `글자_입력_방식`="한 글자씩"이면 예전 방식).
- **항목 이름은 한글로 마음대로 (2026-11-13, 사용자 요청 "항목 이름은 그냥 한글로 임의로 넣게"):** 잠깐 항목 이름도 Concur 목록에서 고르게 했다가 되돌렸습니다. 항목 이름은 일반 입력칸(`name_entry`)이고, [새 항목]은 '새 항목'(겹치면 '새 항목 2'…) 줄을 더하고 이름 칸을 골라 둡니다(`add` → `add_named`, 고르는 작은 창·`on_name_picked`는 없앰). Expense Type만 Concur 목록 콤보.
- **리포트 만들기 값 (2026-11-13, 사용자 화면 기준):** Report Name = `Monthly Expense for {month_abbr}`(Nov처럼 달 영어 세 글자, `TargetMonth.month_abbr`), From/To Date = 그 달 1일~말일, Business Purpose = `service`(설정 `리포트_Business_Purpose`), Comment = 리포트 이름(설정 `리포트_Comment_형식` 기본 `{report_name}`). Policy(Korea Policy)·Cost Object Type·Cost Object는 Concur가 미리 넣으므로 손대지 않습니다. 예전 기본값 `{month_name} Expenses`가 저장된 설정 파일은 `load_settings`가 새 형식으로 바꿉니다(`OLD_REPORT_NAME_FORMATS`, 사용자가 직접 바꾼 값은 그대로). 실제 화면의 날짜 칸 이름이 From Date/To Date라 `화면_요소_라벨`에 더했습니다. 테스트 `tests/test_report_create.py`.
- **Concur Expense Type 전체 목록 (2026-11-13, 사용자 캡처):** 사용자가 Expense Type 목록 전체를 캡처해 보내 줘서 `KNOWN_EXPENSE_TYPES`에 글자 그대로(대소문자 포함, 예: 'M Card Service Fee (Airfare Booking)'만 C가 대문자) 넣었습니다. 모두 47개(주차 기본값 `Parking & Toll`도 목록에 있는 것을 사용자가 확인). 'M card …' 항목은 법인카드가 아니라 일반 직원은 쓰지 않는 카드용이라고 사용자가 알려 줬습니다. 목록에만 두고 기본값이나 자동입력표에는 쓰지 않습니다. 이미 저장된 설정 파일의 `Concur_경비_유형`에는 `load_settings`가 새로 알게 된 항목을 뒤에 더합니다. `choose_option`은 이름이 완전히 같은 항목만 고르므로(`_exact`) 비슷한 이름끼리 헷갈리지 않습니다(테스트 `test_similar_expense_type_names_chosen_exactly`).
- **영수증 불러오기 탭 (2026-11-13, 사용자 요청 "아이폰이 아닌 핸드폰은 어딘가 저장된 곳에서 불러와야 한다. 아이폰도 그렇게"):**
  - 폴더 여러 곳: 기존 `영수증_원본_폴더`(첫 폴더) + 새 설정 `영수증_추가_폴더`(목록). `receipt_folders`/`set_receipt_folders`. 예전 설정 파일·테스트는 그대로 동작(추가 폴더 기본값 []). 첫 폴더를 빼면 다음 폴더가 `영수증_원본_폴더`가 됨.
  - `list_originals(cfg, tm, cache, log)`: 모든 폴더를 하위 폴더까지(`iter_photo_files`, os.walk, 숨김 폴더·`Concur_Expenses_Processed`·휴지통 제외) 봅니다. 없는 폴더는 기록만 남기고 건너뛰고, 모두 없으면 오류(`missing_folder_message`, "[영수증 불러오기] 탭에서"). 같은 사진(이름·크기 같음)은 먼저 넣은 폴더 것 하나만, 이름만 같은 다른 사진은 둘 다 쓰고 경고(영수증 목록·캐시가 파일 이름을 키로 쓰기 때문).
  - 파일 이름 날짜: `NAME_DATETIME_COMPACT`로 갤럭시 `20261103_123456`, 안드로이드 `IMG_…`, 픽셀 `PXL_…`, 카카오톡 `KakaoTalk_…`(보낸 시각), 스크린샷 `…-123456`도 읽음. 카카오톡은 EXIF가 지워지는 경우가 있어 이름 날짜가 중요.
  - 사진 직접 가져오기(`import_receipt_files`): 고른 파일을 프로그램 폴더의 `가져온 영수증`(`IMPORTED_FOLDER_NAME`, 개인 파일이라 .gitignore)에 `shutil.copy2`(날짜 유지)로 복사. 같은 내용은 건너뛰고, 이름만 같으면 `_2`를 붙임. 사진이 아닌 파일은 뺌. 이 폴더는 자동으로 목록에 들어감. tkinter 파일 창으로는 USB로 연결한 휴대폰(MTP)을 직접 못 열 수 있어서, 먼저 PC에 옮긴 뒤 고르도록 안내.
  - 이 PC에서 찾은 곳(`suggested_receipt_folders`): iCloud Drive 영수증 폴더, iCloud 사진, OneDrive(`OneDriveConsumer`/`OneDrive` 환경 변수, ~/OneDrive)의 `Pictures|사진\Samsung Gallery`·카메라 앨범, 구글 드라이브(`G:\내 드라이브` 등), `문서\카카오톡 받은 파일`, 다운로드 중 있는 폴더만. 실제 경로 이름은 일부 추측(갤럭시 OneDrive 동기화 폴더 이름 등).
  - 창: `GuiApp`을 `ttk.Notebook` 두 탭([실행], [영수증 불러오기])으로 나누고 실행 기록은 탭 아래에 항상 보임. [실행] 탭의 영수증 폴더 줄은 요약(`update_folder_summary`) + [영수증 불러오기 탭] 버튼. `ReceiptSourcesTab`: 폴더 목록(Treeview: 폴더, 이 달 사진), [폴더 추가…]·[빼기]·[폴더 열기]·[이 달 사진 세기](`count_photos_per_folder`, 작업 스레드 → `("call", 함수)` 메시지), 이 PC에서 찾은 곳(콤보+[추가]), [사진 고르기…]·[가져온 영수증 폴더 열기], 휴대폰별 안내(`PHONE_GUIDE`). 바꾸면 바로 설정에 저장. 테스트 `tests/test_receipt_sources.py`(창 테스트는 화면 있을 때만).
- **백업·되살리기 (2026-11-14, 사용자 요청 "규칙을 계속 넣을 텐데, 나중에 업데이트하면 백업 리스토어가 가능하게"):** 규칙은 개인 파일(`자동입력표.csv` 등)에 저장되어 다음 실행부터 바로 쓰이지만, 새 판을 다른 폴더에 받으면 따라가지 않습니다.
  - 백업 내용(`backup_members`): `concur_settings.json`, `자동입력표.csv`, `병원목록.csv`, `결과` 폴더의 .csv·.json(영수증 목록·경비처리·OCR/촬영 날짜 기억). 사진·실행기록·화면 캡처·`가져온 영수증`은 넣지 않음. zip 안에 `백업정보.json`(프로그램 판, 만든 시각(마이크로초까지, 정렬용), 원래 폴더, 메모, 파일 목록, 내용 해시).
  - 백업 폴더(`backup_dir`): 환경 변수 `CONCUR_BACKUP_DIR`(테스트용, `tests/conftest.py`가 자동으로 임시 폴더로 바꿈) → 설정 `백업_폴더` → `~\Documents\Concur 경비 자동 입력 백업`(프로그램 폴더와 따로라 새 판을 받아도 남음).
  - 자동 백업(`auto_backup`): 창을 켤 때·끌 때·자동 입력 편집 저장 때, 내용 해시가 최근 자동 백업과 다를 때만 `자동백업_YYYYMMDD_HHMMSS.zip`, 최근 `AUTO_BACKUP_KEEP`(20)개만 남김. 직접 백업은 `직접백업_…zip`(지우지 않음).
  - 되살리기(`restore_backup`): 이 프로그램이 만든 zip만(백업정보 필요), 알려진 개인 파일 이름과 `결과/이름.csv|json`만(`_restorable`, zip 경로 공격 방지). 되살리기 전 상태를 '되살리기 전' 자동 백업으로 남겨 되돌릴 수 있음. 엑셀에 열린 파일은 오류 안내. 예전 판 설정은 `load_settings`의 옮기기 규칙이 처리.
  - 창: 설정 파일이 없는 새 폴더에서 처음 켤 때(`offer_restore_on_first_run`, root를 먼저 만들고 load_settings 전에) 최근 백업을 되살릴지 물음. [백업·되살리기] 탭(`BackupTab`): 백업 폴더, 목록(만든 시각·종류·프로그램 판·파일 수), [지금 백업하기]·[고른 백업 되살리기]·[다른 백업 파일에서 되살리기…]·[새로 고침]. 작업 중이거나 자동 입력 편집 창이 열려 있으면 되살리지 않음. 되살린 뒤 `reload_from_settings`로 창에 반영. 테스트 `tests/test_backup.py`.
  - 테스트 주의: 여러 창 테스트를 이어서 돌리면 닫은 창의 tk 변수를 작업 스레드가 치우다 멈춰 수십 초 걸리는 일이 있어, `conftest.py`가 테스트마다 `gc.collect()`를 합니다.
- **자동 입력 항목 늘림·Concur 순서대로 (2026-11-14, 사용자 요청):**
  - "병원 이름이 없는 영수증이 많으니 무시하고 진행": '병원 없음'을 `receipt_issues`(영수증 목록 '확인 필요'·입력 전 확인 창)와 경비처리 결과 '확인 필요'("병원명 없음")에서 뺐습니다. Comment에서 ' in {hospital}'만 빠집니다.
  - "Concur 경비 화면의 칸을 전부 자동 입력 편집에서": 자동입력표에 칸 4개를 더했습니다(`CATEGORY_HEADER` 10칸, 위치 상수 `C_NAME`…`C_PERSONAL`): `City(비우면 영수증 주소)`, `LCRT`, `Receipt Status`, `Personal Expense(O면 체크)`. `apply_category_table`이 `cfg["항목_추가_입력"][분류] = item_extras(row)`({도시, LCRT, Receipt Status, Personal Expense})로 넣고 `fill_expense`가 씁니다. 빈칸이면 예전 동작(영수증 주소의 도시→기본 도시, NA, Receipt, 건드리지 않음). 도시 우선순위: 영수증 목록에서 직접 고친 도시 > 항목에 정한 도시 > 영수증 주소 > 기본 도시(`city_to_choose(..., fixed)`). Transaction Date·Amount·Currency는 카드 정보라 넣지 않음. Personal Expense 체크박스는 `화면_문구` "Personal Expense"(개인 경비|Personal Expense) 라벨로 찾습니다(실제 화면 구조는 모름).
  - "어느 항목에도 안 맞을 때 Expense Type은 삭제": 편집 창 아래 칸과 자동입력표의 `(분류 없음)` 줄, 설정의 `분류없음` 기본값을 없앴습니다. 맞는 항목이 없으면 Undefined 그대로 두고 '확인 필요'("Expense Type 직접 고르기(맞는 자동 입력 항목 없음)"). 예전 파일의 `(분류 없음)` 줄은 읽을 때 건너뛰고 저장하면 없어집니다. `load_category_rows`는 이제 줄 목록만 돌려주고 `save_category_rows(path, rows)`는 인자가 둘입니다. 자동입력표대로 고른 Expense Type은 더 이상 '확인 필요'로 남기지 않습니다(사용자가 직접 정한 규칙이라).
  - "입력 순서를 Concur 화면처럼 위에서 아래로": 편집 창 칸 순서 = 항목 이름, 찾을 글자, ─ Concur 경비 화면에 넣을 값 ─ Expense Type, Business Purpose, City(콤보: `(영수증 주소에서 찾기)` + 도시 목록), LCRT(콤보 NA), Receipt Status(콤보 Receipt/No Receipt), Personal Expense(체크), Comment. 미리보기는 사용자가 필요 없다고 해서 없앴습니다(2026-11-14). `fill_expense`도 같은 순서로 넣고 영수증 첨부는 맨 끝. Expense Type을 맨 먼저 고르는 것은 Concur가 유형에 따라 칸을 바꿀 수 있어서이기도 합니다. 이미 맞는 값이 들어 있는 LCRT·Receipt Status는 건드리지 않습니다. 테스트 `tests/test_item_fields.py`(사용자 화면과 같은 칸 순서의 가짜 경비 화면에서 입력 순서 확인).
- 분류 키워드와 병원 표에 있는 이상한 글자('주자', '메리카', '중남대' 등)는 OCR이 잘못 읽은 글자에 맞춘 것이라 일부러 둔 것입니다. 지우면 안 됩니다.

## 어디까지 확인했나

- 이전 채팅에서: 글자 분석·매칭 테스트, Concur를 흉내 낸 가짜 화면에서 처음부터 저장까지 실행, 프로그램 창이 뜨는지 확인했습니다. 다만 그 테스트 코드와 가짜 화면은 넘어오지 않았습니다.
- 이 저장소에서: `tests` 폴더에 테스트 260개가 있습니다(창 테스트 5개는 화면이 있을 때만). 날짜·금액·분류·병원 뽑기, 엑셀로 고친 내용 유지, 경비와 영수증 맞추기, 설정 파일, 지난달 경비만 고르기, 촬영 날짜로 사진 고르기, HEIC 읽기·JPEG 변환, 영수증 아닌 사진 빼기, 영수증 날짜로 달 정하기(늦게 찍은 영수증, 다른 달 영수증, 날짜 못 읽은 영수증에 날짜 적기), 이번 달 처리, 이메일 링크 자동 로그인(가짜 Concur 로그인 화면과 가짜 Outlook), 병원 찾기(엑셀 가져오기·한글→영문·같은 날·가게 이름·영수증 글자의 새 병원), 도시 자동 선택(가짜 도시 드롭다운), 쉼표 없는 금액·두 자리 연도, 리포트 만들기 값(가짜 리포트 생성 화면), 영수증 폴더 여러 곳·사진 가져오기·안드로이드 파일 이름 날짜, 백업·되살리기, 자동 입력 항목의 Concur 칸을 화면 위에서 아래 순서로 넣기를 확인합니다. 모두 통과합니다.
- **실제 Concur와 실제 Windows OCR에서는 아직 한 번도 돌려 보지 않았습니다.**
- 프로그램 창(tkinter): 2026-11-13에 이 환경에 `python3-tk`(시스템 Python 3.12용)와 한글 글꼴을 설치해 가상 화면(Xvfb)에서 처음 띄워 봤습니다. 메인 창과 자동 입력 편집 창이 제대로 뜨고, 편집·저장이 됩니다(`tests/test_category_editor.py`의 창 테스트는 화면이 있을 때만 돎). 다시 하려면 아래 '창 테스트'를 보세요.
- 아직 확인 못 한 것(로그인): 실제 Concur 로그인 화면에서 버튼·입력칸을 찾는지, 회사 PC의 Outlook이 클래식인지와 pywin32 접근이 막히지 않는지, 메일 링크가 메일 보안에 감싸져도 열리는지.
- 아직 확인 못 한 것: 실제 iCloud 사진 폴더(Windows)의 파일 날짜와 폴더 위치, 갤럭시 OneDrive 동기화 폴더·구글 드라이브·카카오톡 받은 파일의 실제 경로, Windows에서 `pillow-heif` 설치와 exe 안에 들어가는지(`--collect-all pillow_heif`), 영수증 판단이 실제 사진에서 얼마나 맞는지.

## 다음에 할 일

1. 사용자가 ①(영수증 분석만)을 돌려 `영수증목록_YYYY-MM.csv`를 보내 주면, 실제 OCR 결과에 맞춰 식대 키워드와 병원 표를 채웁니다. 지금 식대 키워드는 '버터' 하나뿐이라 식당 영수증 대부분이 '분류 없음'이 될 수 있습니다.
2. 처음 실제로 ②나 ③을 돌린 뒤 `실행기록` 폴더(기록 파일과 화면 캡처)를 받아서, 버튼을 못 찾거나 너무 빨리 넘어가는 곳을 고칩니다.
3. 회사 Edge 보안 정책이 자동 조작을 막으면 프로그램으로는 해결할 수 없으니 IT 부서에 확인하도록 안내합니다.
4. 병원: 사용자가 [병원 엑셀 가져오기]를 한 뒤 자동 변환 영문명을 확인했는지, 결과 파일의 `가게 이름` 칸이 비어 있지 않은지(가게 이름 칸 선택자가 맞는지) 확인합니다.
5. 영수증 목록 맨 아래(자동 제외)에 진짜 영수증이 섞여 있으면 `영수증_판별_단어`를 보강하고, 반대로 일반 사진이 영수증으로 잡히면 판단 기준을 좁힙니다.
6. 6,000원 영수증: 사용자에게서 그 영수증의 `영수증목록` 줄(OCR 앞부분·금액·날짜·확인 필요)과 `경비처리` 줄을 받으면 정확한 원인을 확인합니다.
7. 기본 도시는 광주, 병원 목록은 영수증에서 확인한 병원(원광대·목포한국)과 영수증·회사 엑셀(MR·CT·XR 탭)에서 더해지는 병원입니다. 사용자가 엑셀 탭의 열 제목을 알려 주면 `병원_엑셀_열`이나 영문 이름 열 찾기를 맞춥니다.
8. 99,000원 영수증: 사용자에게서 `경비처리_2026-11.csv`의 그 줄(`처리`, `확인 필요`)을 받아 첨부가 빠진 원인을 확인합니다(예전 코드라면 `처리`에 "이미 첨부됨(건너뜀)"이 있을 것).

## 이 클라우드 작업 환경(Linux)에서 작업할 때

- 여기서는 실행할 수 없는 것: Windows OCR, Edge 창 조작, 탐색기·엑셀로 파일 열기, exe 만들기. 모두 Windows PC에서만 됩니다.
- `concur_auto.py`는 맨 위에서 Python 기본 기능만 불러오고, 사진 보정·OCR·브라우저·창 기능은 쓸 때 불러옵니다. 이 구조를 지켜야 Linux에서도 테스트가 돌아갑니다.
- 테스트 실행:
  ```bash
  pip install -r requirements-dev.txt
  ```
  ```bash
  python -m pytest
  ```
  브라우저는 이 환경에 이미 설치된 Chromium(`/opt/pw-browsers/chromium`)을 씁니다. 다른 위치를 쓰려면 환경 변수 `PW_CHROMIUM`에 경로를 넣습니다. `playwright install`은 하지 않습니다.
- **창 테스트:** 기본 Python 3.13에는 tkinter가 없습니다. 시스템 Python 3.12로 돌립니다(컨테이너가 새로 뜨면 다시 설치).
  ```bash
  apt-get install -y python3-tk fonts-nanum python3-pip
  ```
  ```bash
  /usr/bin/python3.12 -m pip install --break-system-packages pytest
  ```
  ```bash
  xvfb-run -a -s "-screen 0 1280x800x24" /usr/bin/python3.12 -m pytest -q -p no:cacheprovider tests/test_category_editor.py
  ```
  화면 캡처는 `xvfb-run` 안에서 `import -window root 파일.png`(ImageMagick)로 찍을 수 있습니다.
- `requirements.txt`는 Windows용(Windows OCR 패키지 포함)이라 Linux에서는 설치하지 않습니다. 테스트용 `requirements-dev.txt`에는 사진 처리 패키지(opencv-python-headless, Pillow, pillow-heif)가 들어 있습니다.
- **이 저장소는 누구나 볼 수 있는 공개 저장소입니다.** 사번이 들어간 경로, 회사 계정, 실제 영수증 사진, 개인 설정 파일(`concur_settings.json`, `병원목록.csv`, `자동입력표.csv`, `결과`·`실행기록`·`가져온 영수증` 폴더)은 올리지 않습니다.
- `build_exe.bat`는 Windows 줄바꿈(CRLF)을 유지해야 합니다. `.gitattributes`에 설정해 두었습니다.
