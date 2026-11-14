# -*- coding: utf-8 -*-
"""
Concur 경비 자동 입력 (Power Automate Desktop 흐름의 Python 버전)

  영수증 사진(iPhone으로 찍기만 하면 됨. iCloud 사진 폴더 또는 iCloud Drive 폴더, 파일 이름·jpg/heic 상관없음)
   → 지난달 1일 이후에 찍은 사진을 읽음(Windows OCR) → 영수증이 아닌 사진은 뺌 → (선택) 이미지 보정
   → 영수증에 적힌 날짜로 지난달 영수증만 고름(날짜를 못 읽은 영수증만 촬영 날짜 참고)
   → 날짜·금액·분류·병원 추출
   → 영수증 목록(엑셀로 열어 고칠 수 있음, 고친 내용은 다음 실행에도 유지)
   → Concur: 지난달 리포트 생성 → 지난달 경비 추가 → 경비마다 영수증 첨부·필드 입력 → 저장
   제출은 하지 않습니다. 결과를 확인한 뒤 직접 제출하세요.

실행
  python concur_auto.py            → 창(GUI)이 열립니다
  python concur_auto.py --cli      → 명령 프롬프트 메뉴
  python concur_auto.py --cli --ocr-only | --edit-only | --full  [--month 2026-10]
"""

import argparse
import asyncio
import calendar
import csv
import hashlib
import html as html_lib
import json
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import time
import traceback
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path

# exe를 창 모드(--noconsole)로 만들면 stdout/stderr가 없어서 Playwright가 실패하므로 대체
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")


# ============================================================
# 기본 설정
# ============================================================

APP_VERSION = "2026-11-15 식대·숙박 찾을 글자"     # 고칠 때마다 바꿉니다. 실행 기록 첫 줄과 창 제목에 나옵니다.
SETTINGS_FILE_NAME = "concur_settings.json"
HOSPITAL_FILE_NAME = "병원목록.csv"
CATEGORY_FILE_NAME = "자동입력표.csv"
CATEGORIES = ("주차", "카페", "식대")

# 기본 병원 목록: 사용자 영수증에서 확인한 병원만 (2026-11-13 사용자 요청. 나머지는 영수증·회사 엑셀에서 더해짐)
DEFAULT_HOSPITALS = [
    ["원광대", "Wonkwang University Hospital"],
    ["목포한국", "Mokpo Hankook Hospital"],
]

# 예전 기본 목록(동료분 흐름의 대전·충청권). 예전 병원목록.csv에서 지울 줄을 알아보는 데 씁니다.
OLD_DEFAULT_HOSPITALS = [
    ["국립소방", "National Fire Hospital"],
    ["건양대", "KYUH"],
    ["속편", "Nonsan Sok Hospital"],
    ["남대", "CNUH"],
    ["북대", "CBNUH"],
    ["천안충무", "Cheonan CM Hospital"],
    ["주한국", "Cheongju Hankook Hospital"],
    ["순천", "SCHUH"],
    ["백제", "BaekJae Hospital"],
    ["단국대", "DKUH"],
    ["진천성모", "Junang Jeil Hospital"],
    ["세종충남대", "SCNUH"],
    ["아산충무", "Asan CM Hospital"],
    ["음성", "National Fire Hospital"],
    ["진천", "Junang Jeil Hospital"],
    ["중남대", "CNUH"],
    ["충주", "Chungju medical center"],
    ["전한국", "Daejeon Hankook Hospital"],
    ["세종중남대", "SCNUH"],
    ["예수", "Jeonju Jesus hospital"],
]


# Concur 화면의 입력칸·목록 위치(CSS 선택자). Power Automate 흐름에 저장돼 있던 값입니다.
# Concur가 바뀌면 설정 파일의 '화면_요소'만 고치면 됩니다(화면 조사 파일 참고: python concur_auto.py --inspect). 쉼표로 여러 개를 적으면 그중 하나를 찾습니다.
DEFAULT_ELEMENTS = {
    "리포트 이름": "#name",
    "리포트 시작일": "#custom11-date-input-field-input",
    "리포트 종료일": "#custom12-date-input-field-input",
    "리포트 Business Purpose": "#businessPurpose",
    "리포트 Comment": "#comment",
    "사용 가능한 경비 카드": ".report-empty-state-actions__card-body",
    "경비 체크박스": "input[type=checkbox].sapcnqr-checkbox__input",
    "경비 목록": ".entries-list",
    "경비 목록 체크박스": ".entries-list .sapcnqr-dropzone input.sapcnqr-checkbox__input",
    "도구 모음": ".sapcnqr-toolbar__toolbar",
    "경비 날짜": "#transactionDate-date-input-field-input",
    "경비 금액": "#transactionAmount",
    "Comment": "#comment",
    "Business Purpose": "#businessPurpose",
    "Receipt Status": "[name=receiptType]",
    "Expense Type": "[name=expenseType]",
    "LCRT": "#custom32 .sapcnqr-select-field__input",
    "도시": "#location",
    "영수증 업로드 버튼": "#upload-receipt-button",
    "드롭다운 항목": ".sapcnqr-listbox-item__wrapper",
}
# 위치(id)로 못 찾을 때 화면 글자(라벨)로 찾는 입력칸
DEFAULT_ELEMENT_LABELS = {
    "리포트 이름": "Report Name|리포트 이름",
    "리포트 시작일": "From Date|Start Date|시작일|시작 날짜",
    "리포트 종료일": "To Date|End Date|종료일|종료 날짜",
    "리포트 Business Purpose": "Business Purpose|비즈니스 목적|업무 목적",
    "리포트 Comment": "Comment|코멘트|의견",
    "경비 날짜": "Transaction Date|거래 날짜|거래일",
    "경비 금액": "Amount|금액",
    "Comment": "Comment|코멘트|의견",
    "Business Purpose": "Business Purpose|비즈니스 목적|업무 목적",
}

# 사용자 Concur의 Expense Type(사용자가 화면을 캡처해 보내 준 것, 글자 그대로). Concur에서 고를 때는 이름이 완전히 같은
# 항목만 고르므로 띄어쓰기·대소문자까지 같아야 합니다. 화면 조사(--inspect)가 Concur에서 전체 목록을 읽어 설정에 더하기도 합니다.
KNOWN_EXPENSE_TYPES = [
    # 숙박·접대·식사
    "Accommodation", "M card accommodation",
    "Accommodation Covered Recipient", "Airfare Covered Recipient",
    "M card accommodation - Covered recipient", "M card airfare - Covered recipient",
    "M card business meals & drinks - Covered recipient", "M card entertainment - Covered recipient",
    "M card gifts - Covered recipient", "M card taxi - Covered recipient",
    "Entertainment", "M card entertainment", "M card gifts",
    "Business Meal & Drinks (with clients)", "M card business meals and drinks (with clients)",
    "M card meals and drinks", "Meals & Drinks",
    # 회의·마케팅
    "M card business meeting", "M card internal meeting",
    "M card Advertising Costs", "M card Booth Rental Fees", "M card Donation", "M card Donation Covered Recipient",
    "M card Grant", "M card Grant Covered Recipient", "M card Proprietary Event", "M card sponsorship",
    "M card Sponsorship Covered Recipient", "M card Venue Rental Fees", "Other",
    "M card Marketing Costs", "Sales and Distribution", "Trade show and Marcom Expenses",
    "Other Staff Expenses", "Training",
    # 통신·교통
    "Phone & Fax",
    "Airfare", "Car mileage", "M card airfare", "M card parking & toll", "M card taxi", "M card transportation",
    "Parking & Toll", "Public Transport", "Rental Car", "Taxi",
    "M Card Service Fee (Airfare Booking)",
]

OLD_DEFAULT_CITY = "Daejeon (ex Taejon), KOREA, REPUBLIC OF"

# 영수증에서 찾을 글자(기본). 식대·숙박은 사용자가 자동 입력 편집에서 만든 글자(앞쪽)에 더 넣은 것입니다.
# 한 글자짜리나 지명과 겹치는 글자(예: 부산 해운대구 '우동')는 다른 영수증에도 맞을 수 있어 넣지 않았습니다.
MEAL_USER_WORDS = ["찌개", "한식", "식당"]
MEAL_MORE_WORDS = ["중식", "일식", "양식", "분식", "백반", "국밥", "김밥", "칼국수", "국수", "냉면", "짜장", "짬뽕", "반점",
                   "돈까스", "돈가스", "삼겹살", "갈비", "불고기", "제육", "비빔밥", "덮밥", "볶음밥", "공기밥", "해장국",
                   "설렁탕", "곰탕", "감자탕", "순대", "족발", "보쌈", "치킨", "피자", "버거", "도시락", "초밥", "라면",
                   "떡볶이", "스테이크", "한우", "횟집", "식사", "음식", "식권"]
ACCOM_USER_WORDS = ["모텔", "호텔"]
ACCOM_MORE_WORDS = ["숙박", "여관", "여인숙", "리조트", "펜션", "게스트하우스", "레지던스", "콘도", "객실", "입실", "퇴실",
                    "체크인", "체크아웃", "야놀자", "여기어때", "에어비앤비", "hotel", "motel", "resort"]
# 예전 자동입력표에 한 번만 더할 글자(사용자가 나중에 지운 글자는 다시 넣지 않음)
DEFAULT_WORDS_VERSION = 1
DEFAULT_WORDS_TO_ADD = {"식대": MEAL_USER_WORDS + MEAL_MORE_WORDS, "숙박": ACCOM_USER_WORDS + ACCOM_MORE_WORDS}
DEFAULT_CITIES = {
    # Concur 화면에서 확인한 이름
    "광주": "Gwangju (ex Kwangju), KOREA, REPUBLIC OF",
    "대전": "Daejeon (ex Taejon), KOREA, REPUBLIC OF",
    "전주": "Jeonju, KOREA, REPUBLIC OF",
    "익산": "Iksan, KOREA, REPUBLIC OF",
    "목포": "Mokpo, KOREA, REPUBLIC OF",
    # 아래는 Concur 이름을 확인하지 못함(정확히 안 맞으면 앞 단어로 검색해서 고름)
    "서울": "Seoul, KOREA, REPUBLIC OF", "군산": "Gunsan, KOREA, REPUBLIC OF",
    "정읍": "Jeongeup, KOREA, REPUBLIC OF", "남원": "Namwon, KOREA, REPUBLIC OF",
    "김제": "Gimje, KOREA, REPUBLIC OF", "순천": "Suncheon, KOREA, REPUBLIC OF",
    "여수": "Yeosu, KOREA, REPUBLIC OF", "나주": "Naju, KOREA, REPUBLIC OF",
    "광양": "Gwangyang, KOREA, REPUBLIC OF", "무안": "Muan, KOREA, REPUBLIC OF",
    "해남": "Haenam, KOREA, REPUBLIC OF", "화순": "Hwasun, KOREA, REPUBLIC OF",
    "영광": "Yeonggwang, KOREA, REPUBLIC OF", "청주": "Cheongju, KOREA, REPUBLIC OF",
    "천안": "Cheonan, KOREA, REPUBLIC OF", "세종": "Sejong, KOREA, REPUBLIC OF",
    "부산": "Busan, KOREA, REPUBLIC OF", "대구": "Daegu, KOREA, REPUBLIC OF",
    "인천": "Incheon, KOREA, REPUBLIC OF",
    # 광주·전라 시·군 (Concur 이름은 확인 못 함 → 정확히 안 맞으면 앞 단어로 검색)
    "완주": "Wanju, KOREA, REPUBLIC OF", "진안": "Jinan, KOREA, REPUBLIC OF", "무주": "Muju, KOREA, REPUBLIC OF",
    "장수": "Jangsu, KOREA, REPUBLIC OF", "임실": "Imsil, KOREA, REPUBLIC OF", "순창": "Sunchang, KOREA, REPUBLIC OF",
    "고창": "Gochang, KOREA, REPUBLIC OF", "부안": "Buan, KOREA, REPUBLIC OF", "담양": "Damyang, KOREA, REPUBLIC OF",
    "곡성": "Gokseong, KOREA, REPUBLIC OF", "구례": "Gurye, KOREA, REPUBLIC OF", "고흥": "Goheung, KOREA, REPUBLIC OF",
    "보성": "Boseong, KOREA, REPUBLIC OF", "장흥": "Jangheung, KOREA, REPUBLIC OF", "강진": "Gangjin, KOREA, REPUBLIC OF",
    "영암": "Yeongam, KOREA, REPUBLIC OF", "함평": "Hampyeong, KOREA, REPUBLIC OF", "장성": "Jangseong, KOREA, REPUBLIC OF",
    "완도": "Wando, KOREA, REPUBLIC OF", "진도": "Jindo, KOREA, REPUBLIC OF", "신안": "Sinan, KOREA, REPUBLIC OF",
    # 그 밖의 주요 도시
    "울산": "Ulsan, KOREA, REPUBLIC OF", "제주": "Jeju, KOREA, REPUBLIC OF", "수원": "Suwon, KOREA, REPUBLIC OF",
    "창원": "Changwon, KOREA, REPUBLIC OF", "포항": "Pohang, KOREA, REPUBLIC OF", "원주": "Wonju, KOREA, REPUBLIC OF",
    "춘천": "Chuncheon, KOREA, REPUBLIC OF", "강릉": "Gangneung, KOREA, REPUBLIC OF", "김해": "Gimhae, KOREA, REPUBLIC OF",
}


def default_receipt_folder(home: Path) -> Path:
    """기본 영수증 폴더 iCloudDrive\\Concur_Expenses. 없으면 iCloudDrive 안에서 Concur로 시작하는 폴더 하나를 씁니다."""
    drive = home / "iCloudDrive"
    preferred = drive / "Concur_Expenses"
    if preferred.is_dir():
        return preferred
    try:
        similar = sorted(p for p in drive.iterdir() if p.is_dir() and p.name.lower().startswith("concur"))
    except OSError:
        similar = []
    return similar[0] if len(similar) == 1 else preferred


def default_settings():
    home = Path.home()
    source = default_receipt_folder(home)
    return {
        "_안내": "메모장으로 고칠 수 있습니다. 경로의 \\ 는 \\\\ 처럼 두 번 씁니다. 병원 목록은 병원목록.csv 에서 고칩니다.",
        "_영수증_폴더_설명": "iPhone으로 찍기만 한다면 iCloud 사진 폴더(보통 사진\\iCloud Photos\\Photos)를 고릅니다. 파일 이름·형식은 상관없습니다.",
        "영수증_원본_폴더": str(source),
        "_영수증_추가_폴더_설명": ("영수증 사진이 있는 다른 폴더들(하위 폴더 포함). 아이폰이 아닌 휴대폰의 사진 폴더(OneDrive·구글 드라이브), "
                          "카카오톡 받은 파일, 프로그램이 사진을 복사해 두는 '가져온 영수증' 폴더 등. 창의 [영수증 불러오기] 탭에서 고칩니다."),
        "영수증_추가_폴더": [],
        "_백업_폴더_설명": ("자동 입력 규칙·병원 목록·설정을 자동으로 백업해 두는 폴더. 비워 두면 문서\\Concur 경비 자동 입력 백업. "
                       "프로그램 폴더와 따로 두어야 프로그램을 새로 받아도 남습니다."),
        "백업_폴더": "",
        "_영수증_날짜_설명": "어느 달 영수증인지는 영수증에 적힌 날짜로 정하고, 날짜를 못 읽은 영수증만 촬영 날짜를 참고합니다. 처리할 달이 끝나고 아래 일수 안에 찍은 사진까지 봅니다.",
        "영수증_촬영_허용_일수": 60,
        "보정본_폴더": str(source / "Concur_Expenses_Processed"),
        "Concur_주소": "https://us2.concursolutions.com",
        "브라우저_프로필_폴더": str(home / "AppData" / "Local" / "ConcurAuto" / "edge-profile"),
        "_로그인_설명": "'이메일 링크': 로그인 화면이 나오면 [다음] → [이메일 링크로 로그인]을 누르고, Outlook에 온 로그인 메일의 링크를 이 프로그램의 Edge 창에서 엽니다. '직접': Edge 창에서 직접 로그인합니다.",
        "로그인_방법": "이메일 링크",
        "_로그인_이메일_설명": "로그인 화면의 이메일 칸이 비어 있을 때 넣을 회사 이메일. 비워 두면 Edge 창에서 직접 입력합니다.",
        "로그인_이메일": "",
        "_로그인_메일_설명": "로그인 메일의 제목과 버튼 글자. '|'로 나눠 한국어·영어를 같이 적습니다. 보낸 사람이 concur이면 제목이 달라도 봅니다.",
        "로그인_메일_제목": "SAP Concur 로그인 요청|Sign-In Request to SAP Concur",
        "로그인_메일_버튼": "SAP Concur 로그인|Sign in to SAP Concur",
        "로그인_메일_대기_초": 120,
        "_보정_설명": "영수증 사진은 그대로 읽습니다. 금액이나 날짜를 못 읽은 영수증만 보정(대비·선명하게)해서 한 번 더 읽고, 보정한 사진은 바로 지웁니다. 확대는 작은 사진(긴 변 2000픽셀 미만)에만 합니다.",
        "영수증_보정_사용": True,
        "보정_확대_배율": 2.0,
        "입력_전_확인": True,
        "결과_자동_열기": True,
        "_Edge_열어두기_설명": "true면 ②·③이 끝나도 Edge 창을 닫지 않습니다. Concur에서 확인·수정한 뒤 Edge 창을 직접 닫으면 끝납니다(닫아야 다음 작업을 할 수 있음).",
        "작업_후_Edge_열어두기": True,
        "_업로드_설명": "Concur에는 원본 사진을 올립니다(보정본은 더 이상 만들지 않음). Concur는 1장당 5MB 제한이 있어 넘으면 자동으로 줄여서 올립니다.",
        "업로드_이미지": "원본",
        "업로드_최대_MB": 5,
        "기본_도시": "Gwangju (ex Kwangju), KOREA, REPUBLIC OF",
        "_도시_설명": "영수증 주소(예: 익산시, 광주광역시)나 병원 이름(예: 목포한국병원)에서 도시를 찾아 Concur 도시를 고릅니다. 못 찾으면 기본_도시. 왼쪽은 영수증에서 찾을 한글, 오른쪽은 Concur 도시 목록에 보이는 이름 그대로.",
        "도시_목록": dict(DEFAULT_CITIES),
        "_도시_낮은_우선순위_설명": "주차 회사·카드사 본사 주소처럼 영수증에 자주 찍히지만 실제 장소가 아닌 도시. 다른 도시가 없을 때만 씁니다.",
        "도시_낮은_우선순위": ["서울"],
        "_영수증_병원_설명": "병원 목록에 없어도 영수증에 'OO병원'이 보이면 그 이름을 쓰고 병원 목록에 더합니다(영문 이름은 자동 변환).",
        "영수증에서_병원_찾기": True,
        "_리포트_설명": ("리포트 만들기 화면에 넣는 값. {month_abbr}=Nov처럼 달 영어 줄임말, {month_name}=November, "
                       "{month}=11, {year}=2026, {report_name}=리포트 이름. Policy·Cost Object는 Concur가 미리 넣은 값을 그대로 둡니다."),
        "리포트_이름_형식": "Monthly Expense for {month_abbr}",
        "리포트_Business_Purpose": "service",
        "리포트_Comment_형식": "{report_name}",
        "_리포트_찾기_설명": ("② 전체 실행 때 Concur 리포트 목록(리포트_목록_주소)에서 같은 이름의 리포트를 찾아, 있으면 이어서 넣을지 묻습니다. "
                         "리포트_찾기_대기_초 동안 못 찾으면 새로 만듭니다."),
        "리포트_목록_주소": "/nui/expense",
        "리포트_찾기_대기_초": 8,
        "_키워드_설명": "아래 키워드·분류_우선순위·입력_문구·Undefined일_때_선택할_유형은 자동입력표.csv를 처음 만들 때만 씁니다. 그 뒤로는 자동입력표.csv(창의 [자동 입력표 편집])에서 고치세요.",
        "키워드": {
            "주차": ["주차", "입차", "출차", "주자"],
            "카페": ["메리카", "리카노", "테이크", "카페"],
            "식대": MEAL_USER_WORDS + MEAL_MORE_WORDS + ["버터"],
            "숙박": ACCOM_USER_WORDS + ACCOM_MORE_WORDS,
        },
        "분류_우선순위": ["식대", "숙박", "카페", "주차"],
        "_영수증_판별_설명": "폴더에 일반 사진이 섞여 있어도 됩니다. 금액이 보이고 아래 글자·OCR 날짜·분류 키워드 중 하나가 있으면 영수증으로 봅니다. 아니면 영수증 목록 맨 아래에 '사용 X'로 넣습니다.",
        "영수증_판별_단어": ["합계", "결제", "승인", "카드", "영수증", "금액", "부가세", "사업자", "매출", "가맹점", "거래", "주차"],
        "_병원_엑셀_설명": "병원 정보 엑셀(MR System Information 탭)을 고르면 실행할 때마다 새 병원을 병원목록.csv에 더합니다. 영문 이름은 한글을 자동으로 바꾼 것이라 확인이 필요합니다. 열을 비워 두면 병원 이름처럼 보이는 열을 자동으로 고르고, 열 제목(예: 병원명)이나 열 문자(예: C)를 적어도 됩니다.",
        "병원_엑셀_파일": "",
        "병원_엑셀_탭": "MR|CT|XR|MR System Information",
        "병원_엑셀_열": "",
        "_가게_이름_설명": "Concur 경비의 가게(공급업체) 이름을 읽어 병원을 찾을 때 쓰는 칸. 영수증 글자에서 병원을 못 찾았을 때만 씁니다.",
        "가게_이름_칸": ["#vendorDescription", "[name=vendorDescription]", "#vendor", "[name=vendor]",
                      "#vendorName", "[name=vendorName]", "#merchantName", "[name=merchantName]"],
        "가게_이름_라벨": "공급업체|가맹점|판매처|상호|Vendor|Merchant",
        "_입력_문구_설명": "자동입력표.csv를 처음 만들 때만 씁니다. {hospital} 자리에 병원 영문명. 병원을 못 찾으면 ' in {hospital}' 부분은 빠집니다.",
        "입력_문구": {
            "주차": {"Comment": "Parking fee for service in {hospital}", "Business Purpose": "For service"},
            "카페": {"Comment": "Cafe for wifi in {hospital}", "Business Purpose": "For wifi"},
            "식대": {"Comment": "Overtime meal fee for service in {hospital}", "Business Purpose": "For Overtime meal"},
            "숙박": {"Comment": "Accommodation fee for service in {hospital}", "Business Purpose": "For service in hospital"},
        },
        "_Undefined_설명": ("자동입력표.csv를 처음 만들 때만 씁니다. Expense Type이 Undefined인 경비에 고를 유형. "
                          "어느 항목에도 맞지 않는 경비는 Undefined 그대로 두고 결과 파일 '확인 필요'에 남깁니다."),
        "Undefined일_때_선택할_유형": {
            "주차": "Parking & Toll",
            "카페": "Other Staff Expenses",
            "식대": "Other Staff Expenses",
            "숙박": "Accommodation",
        },
        "_화면_문구_설명": "Concur 화면에 보이는 글자. '|'로 나눠 여러 개를 적으면 그중 하나가 보이면 됩니다(예: 다음|Next). 화면이 영어로 바뀌어도 되도록 영어도 같이 적어 두었습니다.",
        "화면_문구": {
            "리포트 생성": "리포트 생성|Create Report",
            "다음": "다음|Next|Continue",
            "이메일 링크로 로그인": "이메일 링크로 로그인|Sign in with email link|Email me a sign in link|Email sign in link|Email link",
            "사용 가능한 경비": "사용 가능한 경비|Available Expenses",
            "사용 가능한 경비 없음": "사용 가능한 경비 없음|사용 가능한 경비가 없습니다|No Available Expenses",
            "리포트에 추가": "리포트에 추가|Add To Report",
            "편집": "편집|Edit",
            "다음 경비": "다음 경비|Next Expense",
            "경비 저장": "경비 저장|Save Expense",
            "영수증 업로드": "새 영수증 업로드|영수증 업로드|Upload New Receipt|Upload Receipt",
            "첨부된 영수증": "영수증 분리|영수증 삭제|영수증 제거|Detach Receipt|Delete Receipt|Remove Receipt",
            "No Receipt": "No Receipt",
            "Receipt": "Receipt",
            "NA": "NA",
            "Undefined": "Undefined",
            "Personal Expense": "개인 경비|Personal Expense",
        },
        "_Concur_경비_유형_설명": "자동 입력 편집 창에서 고를 Concur Expense Type 목록. 화면 조사(python concur_auto.py --inspect)를 경비 편집 화면에서 하면 Concur의 전체 목록을 읽어 여기에 더합니다.",
        "Concur_경비_유형": list(KNOWN_EXPENSE_TYPES),
        "_글자_입력_설명": "'한꺼번에': Comment 같은 글을 한 번에 넣습니다(칸이 안 받으면 한 글자씩). '한 글자씩': 예전처럼 한 글자씩 칩니다(느리지만 Concur가 글을 못 받을 때).",
        "글자_입력_방식": "한꺼번에",
        "_화면_요소_설명": "Concur 화면의 입력칸 위치(CSS 선택자). Concur가 바뀌어 칸을 못 찾으면 화면 조사(python concur_auto.py --inspect)로 만든 파일의 id·name을 보고 여기를 고칩니다. 예: 입력칸 id가 txtComment면 \"#txtComment\". 쉼표로 여러 개를 적으면 그중 하나를 찾습니다.",
        "화면_요소": dict(DEFAULT_ELEMENTS),
        "_화면_요소_라벨_설명": "위치로 못 찾을 때 화면에 보이는 칸 이름(라벨)으로 찾습니다. '|'로 여러 개.",
        "화면_요소_라벨": dict(DEFAULT_ELEMENT_LABELS),
        "대기시간_초": 60,
    }


# ============================================================
# 공통 도우미
# ============================================================

def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


# ---------- 백업·되살리기 ----------
BACKUP_INFO_NAME = "백업정보.json"
AUTO_BACKUP_PREFIX = "자동백업_"
AUTO_BACKUP_KEEP = 20
RESULT_BACKUP_SUFFIXES = (".csv", ".json")


def backup_dir(cfg=None) -> Path:
    """백업 폴더: 설정 '백업_폴더' → 문서\\Concur 경비 자동 입력 백업 (프로그램 폴더와 따로 있어 새 판을 받아도 남음)"""
    env = os.environ.get("CONCUR_BACKUP_DIR")
    if env:
        return Path(env)
    custom = str((cfg or {}).get("백업_폴더") or "").strip()
    return Path(custom) if custom else Path.home() / "Documents" / "Concur 경비 자동 입력 백업"


def backup_members(base: Path) -> list:
    """백업할 개인 파일 [(백업 안의 이름, 경로), ...]: 설정, 자동입력표, 병원목록, 결과 폴더의 목록·기억 파일"""
    out = []
    for name in (SETTINGS_FILE_NAME, CATEGORY_FILE_NAME, HOSPITAL_FILE_NAME):
        if (base / name).is_file():
            out.append((name, base / name))
    results = base / "결과"
    if results.is_dir():
        for f in sorted(results.iterdir()):
            if f.is_file() and f.suffix.lower() in RESULT_BACKUP_SUFFIXES:
                out.append((f"결과/{f.name}", f))
    return out


def _members_hash(members) -> str:
    h = hashlib.sha256()
    for name, path in members:
        h.update(name.encode("utf-8") + b"\0" + path.read_bytes() + b"\0")
    return h.hexdigest()


def make_backup(base: Path, dest: Path, note: str = "") -> dict:
    """개인 파일을 zip 하나로 묶습니다. → 백업 정보"""
    members = backup_members(base)
    if not members:
        raise RuntimeError("백업할 것이 없습니다(아직 설정·자동 입력표·병원 목록이 없음).")
    dest.parent.mkdir(parents=True, exist_ok=True)
    info = {"프로그램": APP_VERSION, "만든_시각": datetime.now().isoformat(timespec="microseconds"),
            "프로그램_폴더": str(base), "메모": note, "파일": [n for n, _ in members], "내용_해시": _members_hash(members)}
    tmp = dest.with_name(dest.name + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(BACKUP_INFO_NAME, json.dumps(info, ensure_ascii=False, indent=1))
        for name, path in members:
            z.write(path, name)
    os.replace(tmp, dest)
    return info


def read_backup_info(path: Path) -> dict:
    try:
        with zipfile.ZipFile(path) as z:
            return json.loads(z.read(BACKUP_INFO_NAME).decode("utf-8"))
    except Exception:
        return {}


def list_backups(folder: Path) -> list:
    """이 프로그램이 만든 백업 [(경로, 정보), ...] — 최근 것부터"""
    try:
        items = [(f, read_backup_info(f)) for f in folder.glob("*.zip")] if folder.is_dir() else []
    except OSError:
        items = []
    items = [(f, i) for f, i in items if i]
    return sorted(items, key=lambda x: (str(x[1].get("만든_시각", "")), x[0].name), reverse=True)


def _unique_path(folder: Path, stem: str) -> Path:
    path, n = folder / f"{stem}.zip", 2
    while path.exists():
        path, n = folder / f"{stem}_{n}.zip", n + 1
    return path


def auto_backup(base: Path, folder: Path, note: str = "자동 백업"):
    """바뀐 것이 있을 때만 자동 백업(최근 AUTO_BACKUP_KEEP개만 남김). → 만든 백업 경로 또는 None"""
    members = backup_members(base)
    if not members:
        return None
    autos = [(f, i) for f, i in list_backups(folder) if f.name.startswith(AUTO_BACKUP_PREFIX)]
    if autos and autos[0][1].get("내용_해시") == _members_hash(members):
        return None
    path = _unique_path(folder, f"{AUTO_BACKUP_PREFIX}{datetime.now():%Y%m%d_%H%M%S}")
    make_backup(base, path, note)
    for old, _ in autos[AUTO_BACKUP_KEEP - 1:]:
        try:
            old.unlink()
        except OSError:
            pass
    return path


def _restorable(name: str) -> bool:
    if name in (SETTINGS_FILE_NAME, CATEGORY_FILE_NAME, HOSPITAL_FILE_NAME):
        return True
    parts = name.split("/")
    return (len(parts) == 2 and parts[0] == "결과" and parts[1] not in ("", ".", "..")
            and "\\" not in parts[1] and Path(parts[1]).suffix.lower() in RESULT_BACKUP_SUFFIXES)


def restore_backup(path: Path, base: Path, folder: Path) -> list:
    """백업을 프로그램 폴더에 되살립니다. 되살리기 전 상태도 자동 백업해 두어 되돌릴 수 있습니다. → 되살린 파일 이름들"""
    info = read_backup_info(path)
    if not info:
        raise RuntimeError(f"이 프로그램이 만든 백업 파일이 아닙니다: {path.name}")
    with zipfile.ZipFile(path) as z:
        listed = set(info.get("파일") or [])
        names = [n for n in z.namelist() if n in listed and _restorable(n)]
        if not names:
            raise RuntimeError("백업 파일에 되살릴 내용이 없습니다.")
        auto_backup(base, folder, "되살리기 전")
        for name in names:
            target = base.joinpath(*name.split("/"))
            target.parent.mkdir(parents=True, exist_ok=True)
            tmp = target.with_name(target.name + ".tmp")
            tmp.write_bytes(z.read(name))
            try:
                os.replace(tmp, target)
            except PermissionError:
                tmp.unlink()
                raise RuntimeError(f"{name} 파일이 엑셀 등에 열려 있습니다. 닫고 다시 되살리세요.")
    return names


def backup_label(path: Path, info: dict) -> tuple:
    """백업 목록에 보일 (만든 시각, 종류, 프로그램 판, 파일 수)"""
    when = str(info.get("만든_시각", "")).replace("T", " ")[:16]
    kind = "자동" if path.name.startswith(AUTO_BACKUP_PREFIX) else "직접"
    if info.get("메모") and info["메모"] not in ("자동 백업", "직접 백업"):
        kind += f" ({info['메모']})"
    return when, kind, str(info.get("프로그램", "")), str(len(info.get("파일") or []))


def open_path(path: Path, notepad: bool = False):
    try:
        if sys.platform.startswith("win"):
            if notepad:
                subprocess.Popen(["notepad.exe", str(path)])
            else:
                os.startfile(str(path))
    except Exception:
        pass


def read_csv_dicts(path: Path) -> list:
    """엑셀이 UTF-8 또는 CP949(ANSI)로 저장해도 읽습니다."""
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "cp949"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = raw.decode("utf-8", errors="replace")
    return list(csv.DictReader(text.splitlines()))


def write_csv(path: Path, header: list, rows: list):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


class Log:
    def __init__(self, folder: Path, sink=print):
        folder.mkdir(parents=True, exist_ok=True)
        self.folder = folder
        self.sink = sink
        self.path = folder / f"실행기록_{datetime.now():%Y%m%d_%H%M%S}.txt"
        self._f = open(self.path, "a", encoding="utf-8")

    def __call__(self, msg: str = ""):
        line = f"[{datetime.now():%H:%M:%S}] {msg}" if msg else ""
        try:
            self.sink(line)
        except Exception:
            pass
        self._f.write(line + "\n")
        self._f.flush()

    def close(self):
        try:
            self._f.close()
        except Exception:
            pass


class Cancelled(Exception):
    pass


class NoExpenses(Exception):
    """넣을 경비가 없음(오류가 아님): Concur 창은 그대로 두고 작업만 끝냅니다."""


# ============================================================
# 설정 파일 / 병원 목록
# ============================================================

def load_settings(path: Path) -> tuple:
    """(설정, 새로 만들었는지)"""
    defaults = default_settings()
    if not path.exists():
        save_settings(path, defaults)
        return defaults, True
    try:
        user = json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as e:
        raise RuntimeError(f"설정 파일 형식 오류: {path}\n{e.lineno}번째 줄 {e.colno}번째 글자 근처를 확인하세요. ({e.msg})")
    merged = dict(defaults)
    for key, value in user.items():
        if isinstance(value, dict) and isinstance(defaults.get(key), dict):
            sub = dict(defaults[key])
            sub.update(value)
            merged[key] = sub
        else:
            merged[key] = value
    # 예전 기본값(한국어 하나)이 저장돼 있으면 새 기본값(한국어|영어)으로 (사용자가 직접 바꾼 값은 그대로)
    def upgraded(old, new):
        return new if isinstance(old, str) and isinstance(new, str) and "|" in new and old == new.split("|")[0] else old
    for k, v in user.items():
        if isinstance(v, str):
            merged[k] = upgraded(v, defaults.get(k))
    if user.get("병원_엑셀_탭") in ("MR System Information",):      # 예전 기본값 → 여러 탭
        merged["병원_엑셀_탭"] = defaults["병원_엑셀_탭"]
    if user.get("리포트_이름_형식") in OLD_REPORT_NAME_FORMATS:       # 예전 기본값 → Monthly Expense for Nov
        merged["리포트_이름_형식"] = defaults["리포트_이름_형식"]
    saved_types = user.get("Concur_경비_유형")                     # 저장된 목록에 새로 알게 된 항목을 더함
    if isinstance(saved_types, list):
        merged["Concur_경비_유형"] = saved_types + [t for t in KNOWN_EXPENSE_TYPES if t not in saved_types]
    screen, old_screen = merged.get("화면_문구"), user.get("화면_문구")
    if isinstance(screen, dict) and isinstance(old_screen, dict):
        for k, v in old_screen.items():
            screen[k] = upgraded(v, defaults["화면_문구"].get(k, ""))
    # 기본 도시를 대전 → 광주로 바꾼 것을 예전 설정 파일에도 한 번 반영
    if user.get("기본_도시") == OLD_DEFAULT_CITY and not user.get("_기본_도시_광주로_바꿈"):
        merged["기본_도시"] = defaults["기본_도시"]
        merged["_기본_도시_광주로_바꿈"] = True
        try:
            save_settings(path, merged)
        except OSError:
            pass
    return merged, False


def save_settings(path: Path, cfg: dict):
    path.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


HOSPITAL_HEADER = ["OCR에서 찾을 글자", "Concur에 쓸 영문명", "비고"]
DEFAULT_HOSPITAL_NOTE = "영수증에서 확인"
OLD_DEFAULT_NOTES = ("기본 목록(동료분 흐름)", "기본 목록(광주·전라권)")


def _is_old_default_row(row: list) -> bool:
    return row[2] in OLD_DEFAULT_NOTES or (row[0], row[1]) in {(k, e) for k, e in OLD_DEFAULT_HOSPITALS}


def load_hospital_rows(path: Path) -> list:
    """[[찾을 글자, 영문명, 비고], ...].
    예전 기본 목록(동료분 대전·충청권, 잠깐 넣었던 광주·전라권)은 지우고 영수증에서 확인한 병원만 남깁니다."""
    if not path.exists():
        write_csv(path, HOSPITAL_HEADER, [pair + [DEFAULT_HOSPITAL_NOTE] for pair in DEFAULT_HOSPITALS])
    rows = []
    for row in read_csv_dicts(path):
        vals = [str(v or "").strip() for v in row.values()] + ["", ""]
        if vals[0] and vals[1]:
            rows.append(vals[:3])
    if any(_is_old_default_row(r) for r in rows):
        rows = [r for r in rows if not _is_old_default_row(r)]
        official = {compact(k): e for k, e in DEFAULT_HOSPITALS}
        for r in rows:          # 영수증에서 찾아 자동으로 만든 영문명은 확인한 공식 영문명으로
            if compact(r[0]) in official and r[2].startswith("영수증에서 찾음"):
                r[1], r[2] = official[compact(r[0])], DEFAULT_HOSPITAL_NOTE
        keys = {compact(r[0]) for r in rows}
        rows = [k_e + [DEFAULT_HOSPITAL_NOTE] for k_e in DEFAULT_HOSPITALS if compact(k_e[0]) not in keys] + rows
        try:
            write_csv(path, HOSPITAL_HEADER, rows)
        except PermissionError:
            pass
    return rows


def load_hospitals(path: Path) -> list:
    if not path.exists():
        load_hospital_rows(path)
    pairs = []
    for row in read_csv_dicts(path):
        vals = [str(v or "").strip() for v in row.values()]
        if len(vals) >= 2 and vals[0] and vals[1]:
            pairs.append([vals[0], vals[1]])
    return pairs


# ---------- 한글 → 영문 ----------

_CHO = ["g", "kk", "n", "d", "tt", "r", "m", "b", "pp", "s", "ss", "", "j", "jj", "ch", "k", "t", "p", "h"]
_JUNG = ["a", "ae", "ya", "yae", "eo", "e", "yeo", "ye", "o", "wa", "wae", "oe", "yo", "u", "wo", "we", "wi",
         "yu", "eu", "ui", "i"]
_JONG = ["", "k", "k", "k", "n", "n", "n", "t", "l", "k", "m", "l", "l", "l", "p", "l", "m", "p", "p", "t", "t",
         "ng", "t", "t", "k", "t", "p", "t"]
_JONG_LINK = {1: "g", 2: "kk", 4: "n", 7: "d", 8: "r", 16: "m", 17: "b", 19: "s", 20: "ss", 22: "j", 23: "ch",
              24: "k", 25: "t", 26: "p", 27: ""}     # 받침이 다음 글자 'ㅇ'으로 넘어갈 때 소리(연음)


def romanize_korean(text: str) -> str:
    """국어의 로마자 표기법(간단판). 예: 원광 → wongwang, 한림 → hallim"""
    out, carry = [], None
    codes = [ord(c) - 0xAC00 for c in text]
    for i, ch in enumerate(text):
        code = codes[i]
        if not 0 <= code < 11172:
            out.append(ch)
            carry = None
            continue
        cho, jung, jong = code // 588, (code % 588) // 28, code % 28
        nxt = codes[i + 1] if i + 1 < len(text) else -1
        nxt_cho = nxt // 588 if 0 <= nxt < 11172 else None
        initial = carry if carry is not None else _CHO[cho]
        carry, final = None, _JONG[jong]
        if jong and nxt_cho is not None:
            if nxt_cho == 11 and jong in _JONG_LINK:          # 받침 + ㅇ → 연음
                carry, final = _JONG_LINK[jong], ""
            elif (nxt_cho == 5 and jong in (4, 8)) or (nxt_cho == 2 and jong == 8):   # ㄴ/ㄹ + ㄹ, ㄹ + ㄴ → ll
                carry, final = "l", "l"
        out.append(initial + _JUNG[jung] + final)
    return "".join(out)


HOSPITAL_SUFFIX_RE = r"(병원|의료원|의료센터|메디컬센터|의원)"
_PREFIX_EN = [("국립", "National"), ("도립", "Provincial"), ("시립", "Municipal")]
_SUFFIX_EN = [("대학교병원", "University Hospital"), ("대학병원", "University Hospital"),
              ("대학교의료원", "University Medical Center"), ("대병원", "University Hospital"),
              ("성모병원", "St. Mary's Hospital"),
              ("기독병원", "Christian Hospital"), ("한방병원", "Korean Medicine Hospital"),
              ("요양병원", "Care Hospital"), ("메디컬센터", "Medical Center"), ("의료센터", "Medical Center"),
              ("의료원", "Medical Center"), ("병원", "Hospital"), ("의원", "Clinic"), ("센터", "Center")]
_HEADER_WORDS = {"병원", "병원명", "병원이름", "고객", "고객명", "기관", "기관명", "사이트", "site", "customer", "hospital"}


def clean_hospital_name(raw) -> str:
    """'원광대학교병원 1.5T (익산)' → '원광대학교병원'"""
    t = re.sub(r"\(.*?\)|\[.*?\]", " ", str(raw or "")).strip()
    t = t.splitlines()[0].strip() if t else ""
    m = re.match(r"^(.*?" + HOSPITAL_SUFFIX_RE + ")", t)
    return re.sub(r"\s+", " ", (m.group(1) if m else t)).strip()


def looks_like_hospital(value) -> bool:
    t = compact(clean_hospital_name(value))
    return len(t) >= 4 and bool(re.search(HOSPITAL_SUFFIX_RE + "$", t))


def hospital_key(name) -> str:
    """영수증에서 찾을 글자: '원광대학교병원' → '원광대', 너무 짧으면(지역 이름 등) 전체 이름"""
    n = compact(clean_hospital_name(name))
    short = re.sub(r"(학교)?" + HOSPITAL_SUFFIX_RE + "$", "", n)
    return short if len(short) >= 3 else n


def hospital_english_name(name) -> str:
    """'원광대학교병원' → 'Wongwang University Hospital' (자동 변환이라 공식 영문명과 다를 수 있음)"""
    n = compact(clean_hospital_name(name))
    head = tail = ""
    for ko, en in _PREFIX_EN:
        if n.startswith(ko) and len(n) > len(ko) + 1:
            head, n = en, n[len(ko):]
            break
    for ko, en in _SUFFIX_EN:
        if n.endswith(ko) and len(n) > len(ko):
            tail, n = en, n[:-len(ko)]
            break
    words = [head] if head else []
    city = next((c for c in sorted(DEFAULT_CITIES, key=len, reverse=True)
                 if n.startswith(c) and len(n) > len(c) + 1), "")
    parts = ([city] if city else []) + re.split(r"(대학교|대학)", n[len(city):])
    for part in parts:
        if part in ("대학교", "대학"):
            words.append("University")
        elif part:
            r = romanize_korean(part)
            words.append(r[:1].upper() + r[1:])
    if tail:
        words.append(tail)
    return " ".join(w for w in words if w)


# ---------- 병원 정보 엑셀 (MR System Information 탭) ----------

def _column_index(letters: str) -> int:
    idx = 0
    for ch in letters.upper():
        idx = idx * 26 + (ord(ch) - 64)
    return idx - 1


def _pick_hospital_column(rows: list, column: str) -> tuple:
    """(열 번호, 데이터 시작 행). column: 열 제목, 열 문자(C), 또는 빈칸(자동)"""
    column = (column or "").strip()
    if column and re.fullmatch(r"[A-Za-z]{1,3}", column):
        return _column_index(column), 0
    if column:
        want = compact(column).lower()
        for ri, row in enumerate(rows[:30]):
            for ci, v in enumerate(row):
                if compact(str(v or "")).lower() == want:
                    return ci, ri + 1
        raise RuntimeError(f"엑셀에서 '{column}' 열 제목을 찾지 못했습니다.")
    scores = {}
    for row in rows[:3000]:
        for ci, v in enumerate(row):
            if v is not None and looks_like_hospital(v):
                scores[ci] = scores.get(ci, 0) + 1
    if not scores:
        raise RuntimeError("엑셀에서 병원 이름 열을 찾지 못했습니다. 설정 '병원_엑셀_열'에 열 제목이나 열 문자를 적어 주세요.")
    ci = max(scores, key=scores.get)
    start = next(ri for ri, row in enumerate(rows) if ci < len(row) and row[ci] is not None and looks_like_hospital(row[ci]))
    return ci, start


ENGLISH_HOSPITAL = re.compile(r"\b(hospital|medical|university|univ|center|centre|clinic|health)\b", re.I)


def _find_sheets(sheetnames: list, spec: str) -> list:
    """'MR|CT|XR'처럼 적은 탭 이름 → 실제 탭 이름들. 같은 이름 → 그 이름으로 시작 → (3글자 이상이면) 포함 순서로 찾습니다."""
    found = []
    names = [(compact(n).lower(), n) for n in sheetnames]
    for want in (compact(w).lower() for w in str(spec or "").split("|")):
        if not want:
            continue
        hit = (next((n for k, n in names if k == want), None)
               or next((n for k, n in names if k.startswith(want)), None)
               or (next((n for k, n in names if want in k), None) if len(want) >= 3 else None))
        if hit and hit not in found:
            found.append(hit)
    return found


def _english_column(rows: list, start: int, ko_col: int):
    """병원 영문 이름처럼 보이는 열(Hospital/Medical/University… 가 든 영어)"""
    scores = {}
    for row in rows[start:start + 3000]:
        for ci, v in enumerate(row):
            if ci == ko_col or not isinstance(v, str):
                continue
            t = v.strip()
            if t and not re.search(r"[가-힣]", t) and ENGLISH_HOSPITAL.search(t):
                scores[ci] = scores.get(ci, 0) + 1
    if not scores:
        return None
    ci = max(scores, key=scores.get)
    return ci if scores[ci] >= 3 else None


def read_hospitals_from_excel(path: Path, sheets: str, column: str = "") -> tuple:
    """병원 정보 엑셀의 여러 탭에서 ([(한글 병원 이름, 영문 이름 또는 ''), ...], [안내 문구, ...])"""
    try:
        import openpyxl
    except ImportError:
        raise RuntimeError("병원 엑셀을 읽으려면 openpyxl 이 필요합니다:  pip install -r requirements.txt")
    try:
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    except Exception as e:
        raise RuntimeError(f"병원 엑셀을 열지 못했습니다(.xlsx 파일인지 확인): {path.name} ({e})")
    out, notes, seen = [], [], {}
    try:
        tabs = _find_sheets(wb.sheetnames, sheets)
        if not tabs:
            raise RuntimeError(f"엑셀에 '{sheets}' 탭이 없습니다. 있는 탭: {', '.join(wb.sheetnames)}")
        for tab in tabs:
            rows = [list(r) for r in wb[tab].iter_rows(values_only=True)]
            try:
                ci, start = _pick_hospital_column(rows, column)
            except RuntimeError as e:
                notes.append(f"{tab} 탭: {e}")
                continue
            en_col = _english_column(rows, start, ci)
            count = 0
            for row in rows[start:]:
                if ci >= len(row) or row[ci] is None:
                    continue
                name = clean_hospital_name(row[ci])
                key = compact(name).lower()
                if len(re.findall(r"[가-힣]", name)) < 2 or key in _HEADER_WORDS:
                    continue
                en = ""
                if en_col is not None and en_col < len(row) and isinstance(row[en_col], str):
                    cand = " ".join(row[en_col].split())
                    if cand and not re.search(r"[가-힣]", cand):
                        en = cand
                hk = compact(hospital_key(name))
                if hk in seen:
                    if en and not out[seen[hk]][1]:
                        out[seen[hk]] = (out[seen[hk]][0], en)       # 영문명이 있는 줄로 채움
                    continue
                seen[hk] = len(out)
                out.append((name, en))
                count += 1
            notes.append(f"{tab} 탭: 병원 {count}곳" + (" (영문 이름 열 있음)" if en_col is not None else ""))
    finally:
        wb.close()
    if not out and any("탭: 병원" not in n for n in notes):
        raise RuntimeError(" / ".join(notes))              # 어느 탭에서도 병원 이름 열을 못 찾음
    return out, notes


def read_hospital_names_from_excel(path: Path, sheet: str, column: str = "") -> list:
    return [name for name, _ in read_hospitals_from_excel(path, sheet, column)[0]]


def import_hospitals_from_excel(cfg: dict, csv_path: Path, log=print) -> tuple:
    """병원 엑셀의 병원 중 병원목록.csv에 없는 것을 더합니다. → (읽은 수, 더한 수)
    엑셀에 영문 이름 열이 있으면 그 이름을 쓰고, 자동으로 만든 영문명(자동 변환·영수증에서 찾음)도 엑셀 영문명으로 바꿉니다."""
    path = str(cfg.get("병원_엑셀_파일") or "").strip()
    if not path:
        return 0, 0
    p = Path(path)
    if not p.exists():
        log(f"병원 엑셀 파일이 없습니다: {p}")
        return 0, 0
    pairs, notes = read_hospitals_from_excel(p, str(cfg.get("병원_엑셀_탭") or "MR|CT|XR|MR System Information"),
                                             str(cfg.get("병원_엑셀_열") or ""))
    rows = load_hospital_rows(csv_path)
    by_key = {compact(r[0]): r for r in rows}
    added, updated = [], 0
    for name, en in pairs:
        key = hospital_key(name)
        row = by_key.get(compact(key))
        if row is not None:
            if en and ("자동 변환" in row[2] or row[2].startswith("영수증에서 찾음")) and row[1] != en:
                row[1], row[2] = en, f"엑셀 영문명: {name}"
                updated += 1
            continue
        new = [key, en, f"엑셀 영문명: {name}"] if en else \
              [key, hospital_english_name(name), f"엑셀에서 자동 변환: {name} (영문 이름 확인)"]
        by_key[compact(key)] = new
        added.append(new)
    if added or updated:
        write_csv(csv_path, HOSPITAL_HEADER, rows + added)
    auto = sum(1 for a in added if "자동 변환" in a[2])
    log(f"병원 엑셀: {len(pairs)}곳 읽음, 병원 목록에 새로 {len(added)}곳 추가"
        + (f", 영문명 {updated}곳 바꿈" if updated else "")
        + (f" (영문 이름 자동 변환 {auto}곳은 [병원 목록 편집]에서 확인하세요)" if auto else "")
        + (f" | {' / '.join(notes)}" if notes else ""))
    return len(pairs), len(added)


# ---------- Concur 가게 이름(영문)으로 병원 찾기 ----------

def normalize_roman(text: str) -> str:
    """카드 가게 이름(WON GWANG DAE HAK GYO)과 로마자 표기의 차이를 줄여 비교용으로 바꿉니다."""
    t = re.sub(r"[^a-z]", "", (text or "").lower())
    # 옛 표기(CHONBUK: ㅓ를 o로)도 맞도록 o·eo·u를 같게 봅니다
    for a, b in (("yeo", "yu"), ("eo", "u"), ("eu", "u"), ("ae", "e"), ("oe", "e"), ("oo", "u"), ("o", "u"),
                 ("ee", "i"), ("sh", "s"), ("ch", "j"), ("k", "g"), ("t", "d"), ("p", "b"), ("r", "l")):
        t = t.replace(a, b)
    return t


def hospital_from_merchant(text: str, hospitals: list) -> tuple:
    """가게 이름에 병원 찾을 글자의 로마자가 들어 있으면 (찾을 글자, 영문명). 6글자 이상 맞아야 인정."""
    t = normalize_roman(text)
    best, best_len = ("", ""), 0
    if not t:
        return best
    for pair in hospitals:
        if len(pair) < 2:
            continue
        key, en = str(pair[0]).strip(), str(pair[1]).strip()
        k = normalize_roman(romanize_korean(compact(key)))
        if len(k) >= 6 and k in t and len(k) > best_len:
            best, best_len = (key, en), len(k)
    return best


# ---------- 자동 입력표 (분류마다 찾을 글자·Comment·Business Purpose·Expense Type) ----------

CATEGORY_HEADER = ["분류", "영수증에서 찾을 글자(쉼표로 구분)", "Comment", "Business Purpose",
                   "Expense Type이 Undefined일 때", "설명", "City(비우면 영수증 주소)", "LCRT", "Receipt Status",
                   "Personal Expense(O면 체크)"]
CATEGORY_ROW_LEN = len(CATEGORY_HEADER)
# 한 줄(항목)의 칸 위치
C_NAME, C_KEYWORDS, C_COMMENT, C_PURPOSE, C_TYPE, C_NOTE, C_CITY, C_LCRT, C_RECEIPT, C_PERSONAL = range(CATEGORY_ROW_LEN)
NO_CATEGORY = "(분류 없음)"            # 예전 판이 만든 줄(이제 쓰지 않음, 읽을 때 건너뜀)
CITY_FROM_RECEIPT = "(영수증 주소에서 찾기)"
LCRT_CHOICES = ["NA"]
RECEIPT_CHOICES = ["Receipt", "No Receipt"]


def _category_row(values) -> list:
    return ([str(v or "").strip() for v in values] + [""] * CATEGORY_ROW_LEN)[:CATEGORY_ROW_LEN]


def _is_no_category(name: str) -> bool:
    return compact(name) in ("(분류없음)", "분류없음")


def category_rows_from_cfg(cfg: dict) -> list:
    """설정 파일의 키워드·입력 문구로 처음 자동입력표를 만듭니다. 위에 있는 분류가 먼저 맞춰집니다."""
    priority = list(cfg.get("분류_우선순위") or [])
    order = priority + [c for c in (cfg.get("입력_문구") or {}) if c not in priority]
    types = cfg.get("Undefined일_때_선택할_유형") or {}
    rows = []
    for cat in order:
        t = (cfg.get("입력_문구") or {}).get(cat, {})
        rows.append(_category_row([cat, ", ".join((cfg.get("키워드") or {}).get(cat, [])), t.get("Comment", ""),
                                   t.get("Business Purpose", ""), types.get(cat, ""), ""]))
    if rows:
        rows[0][C_NOTE] = "위에 있는 분류가 먼저 맞춰집니다. {hospital}(또는 {병원}) 자리에 병원 영문명이 들어가고, 병원을 못 찾으면 ' in {hospital}'은 빠집니다."
    return rows


def upgrade_category_table(table_path: Path, settings_path: Path, log=None) -> int:
    """예전 자동입력표에 새 기본 찾을 글자(식대·숙박)를 한 번만 더하고, 숙박 항목이 없으면 식대 아래에 넣습니다.
    사용자가 적은 글자와 문구는 그대로 두고 없는 글자만 뒤에 붙입니다. → 더한 글자 수"""
    if not table_path.exists():
        return 0                                          # 새로 만들 때는 기본값에 이미 들어 있음
    cfg, _ = load_settings(settings_path)
    if int(cfg.get("_기본_찾을_글자_판") or 0) >= DEFAULT_WORDS_VERSION:
        return 0
    rows = load_category_rows(table_path, cfg)
    added = 0
    for row in rows:
        words = DEFAULT_WORDS_TO_ADD.get(row[C_NAME])
        if not words:
            continue
        have = [k.strip() for k in re.split(r"[,，/]", row[C_KEYWORDS]) if k.strip()]
        known = {compact(k).lower() for k in have}
        new = [w for w in words if compact(w).lower() not in known]
        if new:
            row[C_KEYWORDS] = ", ".join(have + new)
            added += len(new)
    if not any(r[C_NAME] == "숙박" for r in rows):
        defaults = {r[C_NAME]: r for r in category_rows_from_cfg(default_settings())}
        at = next((i + 1 for i, r in enumerate(rows) if r[C_NAME] == "식대"), len(rows))
        rows.insert(at, _category_row(defaults["숙박"][:C_NOTE] + [""]))
        added += len(ACCOM_USER_WORDS + ACCOM_MORE_WORDS)
    if added:
        save_category_rows(table_path, rows)
        if log:
            log(f"자동 입력표에 기본 찾을 글자 {added}개를 더했습니다(식대·숙박). [자동 입력 편집]에서 고칠 수 있습니다.")
    cfg["_기본_찾을_글자_판"] = DEFAULT_WORDS_VERSION
    save_settings(settings_path, cfg)
    return added


def item_extras(row) -> dict:
    """항목의 City·LCRT·Receipt Status·Personal Expense (빈칸은 기본: 영수증 주소의 도시, NA, Receipt, 체크 안 함)"""
    city = row[C_CITY] if row[C_CITY] != CITY_FROM_RECEIPT else ""
    return {"도시": city, "LCRT": row[C_LCRT], "Receipt Status": row[C_RECEIPT],
            "Personal Expense": row[C_PERSONAL].strip().upper() in ("O", "Y", "YES", "V", "✓", "TRUE", "1")}


def apply_category_table(cfg: dict, path: Path, log=None) -> list:
    """자동입력표.csv(없으면 설정으로 만듦)를 읽어 분류·찾을 글자·입력 문구·Expense Type과
    City·LCRT·Receipt Status·Personal Expense를 설정에 반영합니다. → 분류 목록"""
    if not path.exists():
        write_csv(path, CATEGORY_HEADER, category_rows_from_cfg(cfg))
    order, keywords, texts, types, extras = [], {}, {}, {}, {}
    for r in read_csv_dicts(path):
        row = _category_row(r.values())
        cat = row[C_NAME]
        if not cat or _is_no_category(cat) or cat in order:
            continue
        order.append(cat)
        keywords[cat] = [compact(k) for k in re.split(r"[,，/]", row[C_KEYWORDS]) if compact(k)]
        texts[cat] = {name: value.replace("{병원}", "{hospital}")
                      for name, value in (("Comment", row[C_COMMENT]), ("Business Purpose", row[C_PURPOSE])) if value}
        if row[C_TYPE]:
            types[cat] = row[C_TYPE]
        extras[cat] = item_extras(row)
    if not order:
        if log:
            log(f"자동 입력표에 분류가 없어 설정 파일의 값을 씁니다: {path.name}")
        return list(cfg.get("분류_우선순위") or CATEGORIES)
    cfg["분류_우선순위"], cfg["키워드"], cfg["입력_문구"], cfg["Undefined일_때_선택할_유형"] = order, keywords, texts, types
    cfg["항목_추가_입력"] = extras
    if log:
        log(f"자동 입력표: 분류 {len(order)}개 ({', '.join(order)})")
    return order


def load_category_rows(path: Path, cfg: dict) -> list:
    """자동입력표 → [[분류, 찾을 글자, Comment, Business Purpose, Expense Type, 설명, City, LCRT, Receipt Status, Personal], ...]"""
    if not path.exists():
        write_csv(path, CATEGORY_HEADER, category_rows_from_cfg(cfg))
    rows = []
    for r in read_csv_dicts(path):
        row = _category_row(r.values())
        if row[C_NAME] and not _is_no_category(row[C_NAME]):
            rows.append(row)
    return rows


def save_category_rows(path: Path, rows: list):
    write_csv(path, CATEGORY_HEADER, [_category_row(r) for r in rows if str(r[C_NAME]).strip()])


def category_problems(rows: list) -> list:
    """저장 전에 알려 줄 것: 같은 이름, 찾을 글자는 있는데 문구가 빈 항목, 여러 항목에 같은 찾을 글자"""
    problems, names, owner = [], set(), {}
    for r in rows:
        name, kw = r[0].strip(), [compact(k) for k in re.split(r"[,，/]", r[1]) if compact(k)]
        if name in names:
            problems.append(f"'{name}' 항목이 두 번 있습니다(위쪽 것만 씁니다).")
        names.add(name)
        if kw and not (r[2].strip() or r[3].strip()):
            problems.append(f"'{name}': 찾을 글자는 있는데 Comment·Business Purpose가 비어 있습니다.")
        for k in kw:
            if k in owner and owner[k] != name:
                problems.append(f"'{k}' 글자가 '{owner[k]}'와 '{name}'에 다 있습니다(위쪽 '{owner[k]}'로 맞춰집니다).")
            owner.setdefault(k, name)
    return problems


HOSP_CHECK_NOTE = "영문 이름 확인"
HOSP_USER_NOTE = "직접 고침"


def hospital_problems(rows: list) -> list:
    """병원 목록 저장 전에 알려 줄 것: 같은 찾을 글자, 영문 이름이 빈 줄, 짧은 찾을 글자"""
    problems, seen = [], {}
    for key, en, *_ in rows:
        k = compact(key)
        if not k:
            continue
        if k in seen:
            problems.append(f"'{key}' 찾을 글자가 두 번 있습니다(위쪽 '{seen[k]}'만 씁니다).")
        seen.setdefault(k, en or key)
        if not en.strip():
            problems.append(f"'{key}': Concur에 쓸 영문 이름이 비어 있습니다.")
        if len(k) < 2:
            problems.append(f"'{key}': 찾을 글자가 너무 짧아 다른 글자에도 맞을 수 있습니다.")
    return problems


def remember_expense_types(settings_path: Path, types: list) -> int:
    """Concur에서 읽은 Expense Type을 설정 'Concur_경비_유형'에 더합니다(편집 창에서 고르기용). → 새로 더한 수"""
    cfg, _ = load_settings(settings_path)
    known = list(cfg.get("Concur_경비_유형") or [])
    new = [t for t in types if t and t not in known and t.lower() != "undefined"]
    if new:
        cfg["Concur_경비_유형"] = known + new
        save_settings(settings_path, cfg)
    return len(new)


def add_expense_types_to_table(path: Path, types: list, cfg: dict) -> int:
    """Concur에서 읽은 Expense Type 중 자동입력표에 없는 것을 줄로 더합니다('(분류 없음)' 위).
    찾을 글자를 비워 두므로 사용자가 적기 전에는 영수증에 맞춰지지 않습니다."""
    if not path.exists():
        write_csv(path, CATEGORY_HEADER, category_rows_from_cfg(cfg))
    rows = [_category_row(r.values()) for r in read_csv_dicts(path)]
    rows = [r for r in rows if r[C_NAME] and not _is_no_category(r[C_NAME])]
    known = {r[C_NAME] for r in rows} | {r[C_TYPE] for r in rows if r[C_TYPE]}
    new = [_category_row([t, "", "", "", t, "Concur 항목: '영수증에서 찾을 글자'를 적어야 영수증에 맞춰집니다"])
           for t in types if t and t not in known and t.lower() != "undefined"]
    if not new:
        return 0
    write_csv(path, CATEGORY_HEADER, rows + new)
    return len(new)


# ============================================================
# 처리할 달
# ============================================================

MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July",
               "August", "September", "October", "November", "December"]
OLD_REPORT_NAME_FORMATS = ("{month_name} Expenses",)


@dataclass
class TargetMonth:
    year: int
    month: int

    @property
    def prefix(self) -> str:
        return f"{self.year}-{self.month:02d}"

    @property
    def month_name(self) -> str:
        return MONTH_NAMES[self.month - 1]

    @property
    def month_abbr(self) -> str:
        return MONTH_NAMES[self.month - 1][:3]

    @property
    def first_day_us(self) -> str:
        return f"{self.month:02d}/01/{self.year}"

    @property
    def last_day_us(self) -> str:
        return f"{self.month:02d}/{calendar.monthrange(self.year, self.month)[1]:02d}/{self.year}"

    def report_name(self, fmt: str, **extra) -> str:
        return fmt.format(month_name=self.month_name, month_abbr=self.month_abbr, year=self.year, month=self.month,
                          **extra)


def resolve_target_month(text=None, today=None) -> TargetMonth:
    if text:
        m = re.fullmatch(r"\s*(\d{4})-(\d{1,2})\s*", text)
        if not m or not 1 <= int(m.group(2)) <= 12:
            raise RuntimeError("처리할 달은 2026-10 처럼 입력하세요.")
        return TargetMonth(int(m.group(1)), int(m.group(2)))
    today = today or date.today()
    return TargetMonth(today.year - 1, 12) if today.month == 1 else TargetMonth(today.year, today.month - 1)


def is_unfinished_month(tm: TargetMonth, today=None) -> bool:
    """이번 달처럼 아직 끝나지 않은 달인지"""
    today = today or date.today()
    return (tm.year, tm.month) >= (today.year, today.month)


def recent_months(count: int = 12, today=None, include_current: bool = False) -> list:
    """지난달부터 거꾸로 count개. include_current면 맨 앞에 이번 달을 넣습니다."""
    today = today or date.today()
    tm = resolve_target_month(None, today)
    out = [f"{today.year}-{today.month:02d}"] if include_current else []
    y, m = tm.year, tm.month
    for _ in range(count):
        out.append(f"{y}-{m:02d}")
        y, m = (y - 1, 12) if m == 1 else (y, m - 1)
    return out


# ============================================================
# 1. 영수증 보정 (receipt_preprocess.py 와 같은 처리)
# ============================================================

HEIC_EXTENSIONS = {".heic", ".heif"}          # iPhone 기본 사진 형식
SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png"} | HEIC_EXTENSIONS
PROCESSED_SUFFIX = "_보정됨"


def _cv2():
    try:
        import cv2
        import numpy as np
        return cv2, np
    except ImportError:
        raise RuntimeError("opencv-python, numpy 가 필요합니다:  pip install opencv-python numpy")


_HEIF_READY = None


def _register_heif() -> bool:
    """HEIC 사진을 Pillow로 열 수 있게 합니다(pillow-heif)."""
    global _HEIF_READY
    if _HEIF_READY is None:
        try:
            from pillow_heif import register_heif_opener
            register_heif_opener()
            _HEIF_READY = True
        except ImportError:
            _HEIF_READY = False
    return _HEIF_READY


def imread_unicode(path: Path):
    """한글 경로도 읽습니다. jpg는 사진 회전 정보를 반영하고, heic는 pillow-heif로 읽습니다."""
    cv2, np = _cv2()
    if path.suffix.lower() in HEIC_EXTENSIONS:
        if not _register_heif():
            raise RuntimeError("HEIC 사진을 읽으려면 pillow-heif 가 필요합니다:  pip install -r requirements.txt")
        try:
            from PIL import Image
            with Image.open(path) as im:      # 회전은 pillow-heif가 이미 반영함
                rgb = np.asarray(im.convert("RGB"))
            return cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR)
        except Exception:
            return None
    try:
        return cv2.imdecode(np.fromfile(str(path), dtype=np.uint8), cv2.IMREAD_COLOR)
    except Exception:
        return None


def imwrite_unicode(path: Path, image, ext: str = ".png", params=None) -> bool:
    cv2, _ = _cv2()
    ok, encoded = cv2.imencode(ext, image, params or [])
    if not ok:
        return False
    encoded.tofile(str(path))
    return path.exists() and path.stat().st_size > 0


def preprocess_receipt(image, scale: float):
    cv2, np = _cv2()
    enlarged = cv2.resize(image, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    denoised = cv2.bilateralFilter(enlarged, d=5, sigmaColor=20, sigmaSpace=20)
    l_ch, a_ch, b_ch = cv2.split(cv2.cvtColor(denoised, cv2.COLOR_BGR2LAB))
    clahe = cv2.createCLAHE(clipLimit=1.5, tileGridSize=(8, 8))
    enhanced = cv2.cvtColor(cv2.merge((clahe.apply(l_ch), a_ch, b_ch)), cv2.COLOR_LAB2BGR)
    blurred = cv2.GaussianBlur(enhanced, (0, 0), 1.0)
    sharpened = cv2.addWeighted(enhanced, 1.25, blurred, -0.25, 0)
    return np.clip(sharpened, 0, 255).astype(np.uint8)


NAME_DATETIME = re.compile(r"(20\d{2})-(\d{2})-(\d{2})(?:[_ T]?(\d{2})(\d{2})(\d{2}))?")
# 갤럭시·안드로이드·카카오톡: 20261103_123456.jpg, IMG_20261103_123456.jpg, KakaoTalk_20261103_123456789.jpg
NAME_DATETIME_COMPACT = re.compile(r"(?<!\d)(20\d{2})(\d{2})(\d{2})[_-](\d{2})(\d{2})(\d{2})")
EXIF_DATETIME = re.compile(r"\s*(\d{4}):(\d{2}):(\d{2})[ T](\d{2}):(\d{2}):(\d{2})")


def file_name_datetime(name: str):
    """파일 이름의 촬영 시각: 단축어(2026-10-22_123342.jpg), 갤럭시·안드로이드(20261022_123342.jpg),
    카카오톡(KakaoTalk_20261022_123342123.jpg, 보낸 시각)"""
    m = NAME_DATETIME.search(name) or NAME_DATETIME_COMPACT.search(name)
    if not m:
        return None
    try:
        return datetime(*(int(v) for v in m.groups() if v is not None))
    except ValueError:
        return None


def exif_taken_at(path: Path):
    """사진 속에 저장된 촬영 시각(EXIF). 읽을 수 없으면 None"""
    try:
        from PIL import Image
    except ImportError:
        return None
    if path.suffix.lower() in HEIC_EXTENSIONS and not _register_heif():
        return None
    try:
        with Image.open(path) as im:
            exif = im.getexif()
            sub = exif.get_ifd(0x8769)
            values = [sub.get(0x9003), sub.get(0x9004), exif.get(0x0132)]   # 촬영, 디지털화, 수정 시각
    except Exception:
        return None
    for v in values:
        if isinstance(v, bytes):
            v = v.decode("ascii", errors="ignore")
        m = EXIF_DATETIME.match(str(v or ""))
        if m:
            try:
                return datetime(*(int(x) for x in m.groups()))
            except ValueError:
                continue
    return None


def photo_taken_at(path: Path, st, cache: dict) -> datetime:
    """촬영 시각: 파일 이름 날짜 → 사진 속 촬영 시각 → 파일 수정 시각"""
    from_name = file_name_datetime(path.name)
    if from_name:
        return from_name
    c = cache.get(path.name)
    if c and c.get("size") == st.st_size and c.get("mtime") == int(st.st_mtime):
        try:
            return datetime.fromisoformat(c["taken"])
        except (KeyError, ValueError):
            pass
    taken = exif_taken_at(path) or datetime.fromtimestamp(st.st_mtime).replace(microsecond=0)
    cache[path.name] = {"size": st.st_size, "mtime": int(st.st_mtime), "taken": taken.isoformat()}
    return taken


def photo_window(cfg: dict, tm: TargetMonth) -> tuple:
    """(처리할 달 1일, 그 달이 끝나고 허용 일수가 지난 날) — 이 사이에 찍은 사진을 봅니다."""
    start = datetime(tm.year, tm.month, 1)
    month_end = datetime(tm.year + 1, 1, 1) if tm.month == 12 else datetime(tm.year, tm.month + 1, 1)
    return start, month_end + timedelta(days=int(cfg.get("영수증_촬영_허용_일수", 60)))


IMPORTED_FOLDER_NAME = "가져온 영수증"
SKIP_FOLDER_NAMES = {"concur_expenses_processed", "$recycle.bin", "system volume information"}


def _norm_folder(path) -> str:
    return os.path.normcase(os.path.normpath(str(path)))


def receipt_folders(cfg: dict) -> list:
    """영수증을 찾을 폴더들: 기본 폴더(영수증_원본_폴더) + 추가 폴더(영수증_추가_폴더). 빈 값·같은 폴더는 뺌"""
    extra = cfg.get("영수증_추가_폴더") or []
    if isinstance(extra, str):
        extra = [extra]
    out, seen = [], set()
    for f in [cfg.get("영수증_원본_폴더", "")] + list(extra):
        text = str(f or "").strip()
        if text and _norm_folder(text) not in seen:
            seen.add(_norm_folder(text))
            out.append(Path(text))
    return out


def set_receipt_folders(cfg: dict, folders: list):
    """폴더 목록을 설정에 저장(첫 폴더 = 기본 폴더)"""
    folders = [str(f) for f in folders if str(f).strip()]
    first = folders[0] if folders else ""
    if first != cfg.get("영수증_원본_폴더"):
        cfg["영수증_원본_폴더"] = first
        if first:
            cfg["보정본_폴더"] = str(Path(first) / "Concur_Expenses_Processed")
    cfg["영수증_추가_폴더"] = folders[1:]


def missing_folder_message(source: Path) -> str:
    try:
        similar = [d.name for d in source.parent.iterdir()
                   if d.is_dir() and d.name.lower().startswith(source.name.lower()[:6])]
    except OSError:
        similar = []
    hint = (f"\n같은 위치에 비슷한 폴더가 있습니다: {', '.join(similar[:5])}" if similar else "")
    return (f"영수증 폴더가 없습니다: {source}{hint}\n"
            f"프로그램 창의 [영수증 불러오기] 탭에서 영수증 폴더를 다시 고르세요.")


def iter_photo_files(folder: Path):
    """폴더와 하위 폴더의 사진 파일(숨김 폴더·예전 보정본 폴더는 뺌)"""
    for root, dirs, files in os.walk(folder, onerror=lambda e: None):
        dirs[:] = [d for d in dirs if not d.startswith(".") and d.lower() not in SKIP_FOLDER_NAMES]
        for name in files:
            p = Path(root) / name
            if name.startswith(".") or p.suffix.lower() not in SUPPORTED_EXTENSIONS or p.stem.endswith(PROCESSED_SUFFIX):
                continue
            yield p


def list_originals(cfg: dict, tm: TargetMonth, cache=None, log=None) -> list:
    """처리할 달 영수증일 수 있는 사진 [(경로, 촬영 시각), ...] — 촬영 순서대로.
    영수증을 나중에 찍을 수도 있으므로 처리할 달 1일부터, 그 달이 끝나고 허용 일수 안에 찍은 사진까지.
    영수증 폴더 여러 곳(하위 폴더 포함)에서 찾고, 같은 사진(이름·크기가 같음)이 여러 곳에 있으면 한 번만 씁니다."""
    log = log or (lambda s="": None)
    folders = receipt_folders(cfg)
    if not folders:
        raise RuntimeError("영수증 폴더가 정해지지 않았습니다. 프로그램 창의 [영수증 불러오기] 탭에서 폴더를 넣으세요.")
    existing = [d for d in folders if d.is_dir()]
    if not existing:
        raise RuntimeError(missing_folder_message(folders[0]))
    for d in folders:
        if d not in existing:
            log(f"   영수증 폴더가 없어 건너뜁니다: {d}")
    cache = {} if cache is None else cache
    start, end = photo_window(cfg, tm)
    too_old = start - timedelta(days=2)
    found, seen, names, same_name, dup = [], set(), set(), set(), 0
    for folder in existing:
        for p in iter_photo_files(folder):
            try:
                st = p.stat()
            except OSError:
                continue
            # 파일 수정 시각은 촬영 시각과 같거나 늦으므로, 처리할 달보다 앞이면 열어 볼 필요가 없음
            if not file_name_datetime(p.name) and datetime.fromtimestamp(st.st_mtime) < too_old:
                continue
            key = (p.name.lower(), st.st_size)
            if key in seen:                       # 같은 사진이 다른 폴더에도 있음
                dup += 1
                continue
            if key[0] in names:                   # 이름은 같은데 다른 사진
                same_name.add(p.name)
            seen.add(key)
            names.add(key[0])
            taken = photo_taken_at(p, st, cache)
            if start <= taken < end:
                found.append((p, taken))
    if dup:
        log(f"   여러 폴더에 같은 사진이 있어 한 번만 씁니다: {dup}장")
    for name in sorted(same_name):
        log(f"   이름이 같은 다른 사진이 있습니다: {name} (영수증 목록에서 헷갈릴 수 있으니 한쪽 이름을 바꾸세요)")
    return sorted(found, key=lambda x: (x[1], x[0].name.lower(), str(x[0]).lower()))


def count_photos_per_folder(cfg: dict, tm: TargetMonth, cache=None) -> list:
    """[(폴더, 처리할 달 영수증일 수 있는 사진 수 또는 None(폴더 없음)), ...] — [영수증 불러오기] 탭에서 확인용"""
    out = []
    for folder in receipt_folders(cfg):
        if not folder.is_dir():
            out.append((folder, None))
            continue
        one = dict(cfg, 영수증_원본_폴더=str(folder), 영수증_추가_폴더=[])
        out.append((folder, len(list_originals(one, tm, cache))))
    return out


def import_receipt_files(paths, dest: Path) -> dict:
    """아무 곳에 있는 사진 파일을 '가져온 영수증' 폴더로 복사합니다(원본은 그대로, 날짜 정보 유지).
    → {"copied": [...], "same": 이미 있던 수, "skipped": 사진이 아닌 파일 이름들}"""
    dest.mkdir(parents=True, exist_ok=True)
    result = {"copied": [], "same": 0, "skipped": []}
    for src in map(Path, paths):
        if not src.is_file() or src.suffix.lower() not in SUPPORTED_EXTENSIONS:
            result["skipped"].append(src.name)
            continue
        target = dest / src.name
        n = 2
        while target.exists():
            if target.stat().st_size == src.stat().st_size and _same_file_content(target, src):
                break
            target = dest / f"{src.stem}_{n}{src.suffix}"
            n += 1
        if target.exists():
            result["same"] += 1
            continue
        shutil.copy2(src, target)
        result["copied"].append(target.name)
    return result


def _same_file_content(a: Path, b: Path) -> bool:
    try:
        with open(a, "rb") as fa, open(b, "rb") as fb:
            while True:
                x, y = fa.read(1 << 20), fb.read(1 << 20)
                if x != y:
                    return False
                if not x:
                    return True
    except OSError:
        return False


def suggested_receipt_folders(home: Path = None, env=None) -> list:
    """이 PC에서 영수증 사진이 있을 만한 곳 [(이름, 폴더), ...] — 있는 폴더만"""
    home = Path(home or Path.home())
    env = os.environ if env is None else env
    onedrives = []
    for key in ("OneDriveConsumer", "OneDrive"):
        if env.get(key):
            onedrives.append(Path(env[key]))
    onedrives.append(home / "OneDrive")
    candidates = [
        ("iCloud Drive 영수증 폴더 (아이폰)", [default_receipt_folder(home)]),
        ("iCloud 사진 (아이폰)", [home / "Pictures" / "iCloud Photos" / "Photos", home / "Pictures" / "iCloud Photos"]),
        ("OneDrive 갤럭시 사진 (Samsung Gallery)",
         [od / pic / "Samsung Gallery" for od in onedrives for pic in ("Pictures", "사진")]),
        ("OneDrive 카메라 앨범",
         [od / pic / roll for od in onedrives for pic in ("Pictures", "사진") for roll in ("Camera Roll", "카메라 앨범")]),
        ("구글 드라이브", [Path("G:/내 드라이브"), Path("G:/My Drive"), home / "Google Drive" / "My Drive",
                      home / "Google Drive" / "내 드라이브", home / "Google Drive"]),
        ("카카오톡 받은 파일", [home / "Documents" / "카카오톡 받은 파일"]
         + [od / doc / "카카오톡 받은 파일" for od in onedrives for doc in ("Documents", "문서")]),
        ("다운로드", [home / "Downloads"]),
    ]
    out, seen = [], set()
    for label, paths in candidates:
        for path in paths:
            try:
                ok = path.is_dir()
            except OSError:
                ok = False
            if ok and _norm_folder(path) not in seen:
                seen.add(_norm_folder(path))
                out.append((label, path))
                break
    return out


def processed_path(cfg: dict, original: Path) -> Path:
    return Path(cfg["보정본_폴더"]) / f"{original.stem}{PROCESSED_SUFFIX}.png"


def make_enhanced_copy(cfg: dict, original: Path, work_dir: Path, log) -> Path:
    """금액·날짜를 못 읽은 영수증을 다시 읽기 위한 보정 사본(임시 jpg, 읽고 나면 지움). 실패하면 None.
    iPhone 사진은 이미 커서 확대는 효과가 거의 없고 느리기만 해서, 작은 사진(긴 변 2000픽셀 미만)만 확대합니다."""
    cv2, _ = _cv2()
    img = imread_unicode(original)
    if img is None:
        log(f"   읽기 실패: {original.name}")
        return None
    scale = float(cfg.get("보정_확대_배율", 2.0)) if max(img.shape[:2]) < 2000 else 1.0
    work_dir.mkdir(parents=True, exist_ok=True)
    out = work_dir / f"pre_{original.stem}.jpg"
    try:
        ok = imwrite_unicode(out, preprocess_receipt(img, scale), ".jpg", [cv2.IMWRITE_JPEG_QUALITY, 95])
    except Exception as e:
        log(f"   보정 실패: {original.name} ({e})")
        return None
    return out if ok else None


def info_score(rec) -> int:
    """읽어 낸 정보가 얼마나 되는지(날짜·금액·분류·병원)"""
    return (2 if rec.date_from_ocr else 0) + (2 if rec.total else 0) + (1 if rec.category else 0) + (1 if rec.hospital_en else 0)


def old_processed_note(cfg: dict) -> str:
    """예전 판이 만든 보정본(png) 폴더가 남아 있으면 지워도 된다는 안내"""
    folder = Path(str(cfg.get("보정본_폴더") or ""))
    try:
        files = [f for f in folder.glob(f"*{PROCESSED_SUFFIX}.png")] if folder.is_dir() else []
    except OSError:
        files = []
    if not files:
        return ""
    mb = sum(f.stat().st_size for f in files) / 1024 / 1024
    return (f"예전 판이 만든 보정 사진 {len(files)}개(약 {mb:,.0f}MB)가 {folder} 에 남아 있습니다. "
            f"이제 쓰지 않으니 폴더째 지워도 됩니다(iCloud 용량도 차지합니다).")


# ============================================================
# 2. Windows OCR (PAD의 'Windows OCR'과 같은 엔진)
# ============================================================

class WindowsOcr:
    def __init__(self, temp_dir: Path):
        try:
            from winrt.windows.globalization import Language
            from winrt.windows.media.ocr import OcrEngine
        except ImportError:
            raise RuntimeError("Windows OCR 패키지가 필요합니다:  pip install -r requirements.txt")
        self.engine = None
        for tag in ("ko-KR", "ko"):
            lang = Language(tag)
            if OcrEngine.is_language_supported(lang):
                self.engine = OcrEngine.try_create_from_language(lang)
                break
        if self.engine is None:
            raise RuntimeError("Windows 한국어 OCR이 없습니다. 설정 > 시간 및 언어 > 언어 > 한국어 > 언어 옵션에서 '광학 문자 인식'을 설치하세요.")
        self.max_dim = int(OcrEngine.max_image_dimension)
        self.temp_dir = temp_dir
        temp_dir.mkdir(parents=True, exist_ok=True)

    def _fit_size(self, path: Path) -> tuple:
        """OCR용 임시 사본(jpg): 사진 회전 정보를 반영하고, heic는 Windows가 못 읽을 수 있어서 바꿉니다.
        → (읽을 파일, 임시 파일인지). 임시 파일은 읽고 나면 지웁니다."""
        cv2, _ = _cv2()
        img = imread_unicode(path)
        if img is None:
            return path, False
        too_big = max(img.shape[:2]) > self.max_dim
        if path.suffix.lower() == ".png" and not too_big:
            return path, False
        if too_big:
            ratio = (self.max_dim - 1) / max(img.shape[:2])
            img = cv2.resize(img, None, fx=ratio, fy=ratio, interpolation=cv2.INTER_AREA)
        self.temp_dir.mkdir(parents=True, exist_ok=True)
        tmp = self.temp_dir / f"ocr_{path.stem}.jpg"
        imwrite_unicode(tmp, img, ".jpg", [cv2.IMWRITE_JPEG_QUALITY, 95])
        return tmp, True

    async def _read(self, path: Path) -> str:
        from winrt.windows.graphics.imaging import BitmapDecoder
        from winrt.windows.storage import FileAccessMode, StorageFile
        file = await StorageFile.get_file_from_path_async(str(path.resolve()))
        stream = await file.open_async(FileAccessMode.READ)
        try:
            decoder = await BitmapDecoder.create_async(stream)
            bitmap = await decoder.get_software_bitmap_async()
            result = await self.engine.recognize_async(bitmap)
            return "\n".join(line.text for line in result.lines)
        finally:
            close = getattr(stream, "close", None)
            if close:
                close()

    def read(self, path: Path) -> str:
        src, temp = self._fit_size(path)
        try:
            return asyncio.run(self._read(src))
        finally:
            if temp:
                try:
                    src.unlink()
                except OSError:
                    pass


# ============================================================
# 3. 영수증 분석
# ============================================================

@dataclass
class Receipt:
    original: Path
    ocr_file: Path
    ocr_text: str = ""
    dates: list = field(default_factory=list)       # 날짜 후보 ["2026-10-27", ...]
    date_from_ocr: bool = False
    amounts: list = field(default_factory=list)     # 금액 후보
    total: int = 0                                  # 대표 금액(가장 큰 금액)
    category: str = ""
    hospital_key: str = ""
    hospital_en: str = ""
    use: bool = True
    edited: list = field(default_factory=list)      # 사용자가 고친 항목
    used_for: str = ""
    city: str = ""                                  # 영수증에서 찾은 Concur 도시(없으면 기본 도시)
    hospital_new: str = ""                          # 병원 목록에 없어 영수증 글자에서 찾은 병원 이름
    taken: object = None                            # 촬영 시각(datetime)
    is_receipt: bool = True                         # 영수증으로 보이는지(자동 판단)
    ocr_dates: list = field(default_factory=list)   # 영수증 글자에서 읽은 날짜
    skip_reason: str = ""                           # 이 달 목록에서 자동으로 뺀 이유
    hospital_src: str = ""                          # 병원을 어디서 찾았는지

    @property
    def main_date(self) -> str:
        return self.dates[0] if self.dates else ""


DATE_IN_TEXT = re.compile(r"(?<!\d)(20\d{2})\s*[-./년]\s*(\d{1,2})\s*[-./월]\s*(\d{1,2})(?!\d)")
DATE_IN_NAME = re.compile(r"(20\d{2})-(\d{2})-(\d{2})")
AMOUNT_WITH_SEP = re.compile(r"(?<![\d.,])\d{1,3}(?:[,.]\d{3})+(?![\d,]|\.\d)")
AMOUNT_WON = re.compile(r"(?<![\d,.])(\d{3,7})\s*원")
# 쉼표 없이 찍힌 금액(예: 주차요금 6000)은 금액 글자 바로 뒤에 있을 때만
AMOUNT_NEAR_WORD = re.compile(r"(?:합계|총액|금액|요금|결제|받을|청구|판매|주차료|이용료)[^\d\n]{0,8}(\d{3,7})(?![\d,.\-/:년월])")
# 연도 두 자리(26.11.03), 날짜 글자 뒤 8자리(거래일시 20261103)
DATE_SHORT_YEAR = re.compile(r"(?<![\d.])(2\d)\s*[-./]\s*(\d{1,2})\s*[-./]\s*(\d{1,2})(?![\d.])")
DATE_COMPACT = re.compile(r"(?:일시|일자|날짜|거래일|승인일|판매일|결제일|입차|출차)\D{0,6}(20\d{2})(\d{2})(\d{2})(?!\d)")


def normalize_date(y, m, d):
    try:
        return date(int(y), int(m), int(d)).isoformat()
    except ValueError:
        return None


def parse_date_cell(text: str):
    """엑셀에서 고친 날짜 칸: 2026-10-27, 2026.10.27, 2026/10/27, 10/27/2026 모두 허용"""
    t = (text or "").strip()
    m = re.search(r"(20\d{2})\s*[-./년]\s*(\d{1,2})\s*[-./월]\s*(\d{1,2})", t)
    if m:
        return normalize_date(*m.groups())
    m = re.search(r"(\d{1,2})/(\d{1,2})/(20\d{2})", t)
    if m:
        return normalize_date(m.group(3), m.group(1), m.group(2))
    return None


def parse_amount_cell(text: str):
    digits = re.sub(r"\D", "", str(text or "").split(".")[0])
    return int(digits) if digits else None


def find_ocr_dates(ocr_text: str) -> list:
    found = []
    candidates = [m.groups() for m in DATE_IN_TEXT.finditer(ocr_text)]
    candidates += [m.groups() for m in DATE_COMPACT.finditer(ocr_text)]
    candidates += [("20" + y, mo, d) for y, mo, d in (m.groups() for m in DATE_SHORT_YEAR.finditer(ocr_text))]
    for y, mo, d in candidates:
        iso = normalize_date(y, mo, d)
        if iso and iso not in found:
            found.append(iso)
    return found


def file_name_date(name: str):
    m = DATE_IN_NAME.search(name)
    return normalize_date(*m.groups()) if m else None


def _fix_amount_text(text: str) -> str:
    """OCR이 0을 O로 읽은 것(6,0OO)과 쉼표 앞뒤 띄어쓰기(6, 000)를 바로잡습니다."""
    t = text or ""
    for _ in range(3):
        t = re.sub(r"(?<=[\d,])[Oo]|[Oo](?=[\d,])", "0", t)
    return re.sub(r"(\d)\s*([,.])\s*(\d{3})(?!\d)", r"\1\2\3", t)


def find_amounts(ocr_text: str) -> list:
    ocr_text = _fix_amount_text(ocr_text)
    values = [int(re.sub(r"\D", "", m.group(0))) for m in AMOUNT_WITH_SEP.finditer(ocr_text)]
    values += [int(m.group(1)) for m in AMOUNT_WON.finditer(ocr_text)]
    values += [int(m.group(1)) for m in AMOUNT_NEAR_WORD.finditer(ocr_text)]
    uniq = []
    for v in values:
        if v > 0 and v not in uniq:
            uniq.append(v)
    return uniq


def compact(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def detect_category(text_compact: str, keywords: dict, priority=None) -> str:
    text = (text_compact or "").lower()                   # 영어 글자는 대소문자 상관없이(HOTEL·Hotel·hotel)
    for cat in (priority or ["식대", "카페", "주차"]):
        if any(k and k.lower() in text for k in keywords.get(cat, [])):
            return cat
    return ""


_ORG_WORDS = ("사회복지법인", "의료법인", "의료재단", "재단법인", "학교법인", "법인", "재단")


def find_hospital_name_in_text(ocr_text: str) -> str:
    """병원 목록에 없어도 영수증에 'OO병원'·'OO의료원'이 찍혀 있으면 그 이름. 예: '의료법인 OO재단 목포한국병원' → 목포한국병원"""
    for line in (ocr_text or "").splitlines():
        tokens = [re.sub(r"\(주\)|㈜|\(재\)|\(의\)", "", t) for t in line.split()]
        for i, tok in enumerate(tokens):
            word = re.sub(r"[^가-힣]", "", tok)
            if word in ("병원", "의료원") and i > 0:                  # '원광대학교 병원'처럼 띄어 쓴 경우
                word = re.sub(r"[^가-힣]", "", tokens[i - 1]) + word
            m = re.match(r"^([가-힣]*?(?:병원|의료원))", word)
            if not m:
                continue
            name = m.group(1)
            for org in _ORG_WORDS:
                if org in name:
                    name = name.split(org)[-1]
            if len(name) >= 4 and not name.endswith("동물병원"):
                return name
    return ""


def city_name(full: str) -> str:
    """'Gwangju (ex Kwangju), KOREA, REPUBLIC OF' → 'Gwangju (ex Kwangju)'"""
    return (full or "").split(",")[0].strip()


def resolve_city(text: str, cities: dict) -> str:
    """사용자가 적은 도시(광주, Gwangju, 전체 이름)를 Concur 도시 이름으로"""
    t = (text or "").strip()
    if not t:
        return ""
    if t in (cities or {}):
        return cities[t]
    for full in (cities or {}).values():
        if city_name(full).lower() == t.lower() or city_name(full).split(" (")[0].lower() == t.lower():
            return full
    return t if "KOREA" in t.upper() else f"{t}, KOREA, REPUBLIC OF"


PROVINCES = ("서울특별시|부산광역시|대구광역시|인천광역시|광주광역시|대전광역시|울산광역시|세종특별자치시|경기도|"
             "강원특별자치도|강원도|충청북도|충청남도|전북특별자치도|전라북도|전라남도|경상북도|경상남도|제주특별자치도|"
             "서울|부산|대구|인천|광주|대전|울산|세종|경기|강원|충북|충남|전북|전남|경북|경남|제주")
METRO_CITIES = ("서울", "부산", "대구", "인천", "광주", "대전", "울산")


def detect_city(text_compact: str, cities: dict, hospital_key: str = "", low_priority=("서울",)) -> str:
    """영수증 주소에서 도시를 찾습니다. 못 찾으면 빈칸(→ 기본 도시).
    'OO시/OO광역시/OO군' → '전북 익산…'처럼 도 이름 뒤의 도시 → '광주 북구…'처럼 광역시 뒤의 구 → 병원 이름 앞의 도시.
    서울처럼 본사 주소로 자주 찍히는 도시(low_priority)는 다른 도시가 없을 때만 씁니다."""
    text = text_compact or ""
    found = []                                                        # (낮은 우선순위인지, 위치, 도시)
    names = sorted(cities or {}, key=len, reverse=True)

    def add(ko, pos):
        found.append((ko in (low_priority or ()), pos, cities[ko]))

    for ko in names:
        for m in re.finditer(re.escape(ko) + r"(?:광역시|특별시|특별자치시|시|군)", text):
            if ko == "광주" and "경기" in text[max(0, m.start() - 4):m.start()]:
                continue                                              # 경기도 광주시
            add(ko, m.start())
            break
        if ko in METRO_CITIES:
            m = re.search(re.escape(ko) + r"[가-힣]{1,3}구", text)
            if m and not (ko == "광주" and "경기" in text[max(0, m.start() - 4):m.start()]):
                add(ko, m.start())
    if names:
        for m in re.finditer(r"(" + PROVINCES + r")(" + "|".join(re.escape(n) for n in names) + r")", text):
            prov, ko = m.group(1), m.group(2)
            if prov.startswith(ko) or (ko == "광주" and prov.startswith("경기")):
                continue
            add(ko, m.start(2))
    if found:
        return min(found)[2]
    key = compact(hospital_key)
    for ko in sorted(cities or {}, key=len, reverse=True):
        if key.startswith(ko) and not key[len(ko):].startswith("향"):   # 순천향대는 순천이 아님
            return cities[ko]
    return ""


def detect_hospital(text_compact: str, mapping: list) -> tuple:
    best = ("", "")
    for pair in mapping:
        if len(pair) < 2:
            continue
        key, en = str(pair[0]).strip(), str(pair[1]).strip()
        if key and key in text_compact and len(key) > len(best[0]):
            best = (key, en)
    return best


def looks_like_receipt(text_compact: str, amounts: list, has_ocr_date: bool, category: str, words) -> bool:
    """금액이 있고 (영수증 글자 / OCR 날짜 / 분류 키워드) 중 하나가 있으면 영수증.
    금액을 못 읽었어도 영수증 글자가 두 개 이상이면 영수증으로 봅니다."""
    hits = sum(1 for w in (words or []) if w and w in text_compact)
    if amounts and (hits or has_ocr_date or category):
        return True
    return hits >= 2


HOSP_FROM_OCR, HOSP_FROM_DAY, HOSP_FROM_VENDOR, HOSP_FROM_USER = "영수증 글자", "같은 날 영수증", "Concur 가게 이름", "직접 고침"


def fill_same_day_hospitals(receipts: list) -> int:
    """병원을 못 찾은 영수증에 같은 날 다른 영수증(주차 등)의 병원을 채웁니다. 같은 날 병원이 여러 곳이면 주차 영수증 것만."""
    by_day = {}
    for r in receipts:
        if r.use and r.hospital_en and r.hospital_src in (HOSP_FROM_OCR, HOSP_FROM_USER):
            by_day.setdefault(r.main_date, []).append(r)
    filled = 0
    for r in receipts:
        if not r.use or r.hospital_en or not r.main_date or "병원" in r.edited:
            continue
        cands = by_day.get(r.main_date, [])
        names = {c.hospital_en for c in cands}
        if len(names) > 1:
            names = {c.hospital_en for c in cands if c.category == "주차"}
        if len(names) == 1:
            r.hospital_en, r.hospital_src = names.pop(), HOSP_FROM_DAY
            filled += 1
    return filled


def fill_same_day_cities(receipts: list) -> int:
    """도시를 못 찾은 영수증에 같은 날 다른 영수증의 도시를 채웁니다(한 곳일 때만)."""
    by_day = {}
    for r in receipts:
        if r.use and r.city and r.main_date:
            by_day.setdefault(r.main_date, set()).add(r.city)
    filled = 0
    for r in receipts:
        if r.use and not r.city and "도시" not in r.edited and len(by_day.get(r.main_date, ())) == 1:
            r.city = next(iter(by_day[r.main_date]))
            filled += 1
    return filled


def analyze_receipt(original: Path, ocr_file: Path, ocr_text: str, cfg: dict, hospitals: list,
                    taken=None) -> Receipt:
    rec = Receipt(original=original, ocr_file=ocr_file, ocr_text=ocr_text, taken=taken)
    rec.ocr_dates = find_ocr_dates(ocr_text)
    rec.dates = list(rec.ocr_dates)
    rec.date_from_ocr = bool(rec.dates)
    fdate = taken.date().isoformat() if taken else file_name_date(original.name)
    if fdate and fdate not in rec.dates:
        rec.dates.append(fdate)          # 촬영 날짜는 참고용 후보(영수증 날짜를 못 읽으면 대신 씀)
    rec.amounts = find_amounts(ocr_text)
    rec.total = max(rec.amounts) if rec.amounts else 0
    text_c = compact(ocr_text)
    rec.category = detect_category(text_c, cfg["키워드"], cfg.get("분류_우선순위"))
    rec.hospital_key, rec.hospital_en = detect_hospital(text_c, hospitals)
    if not rec.hospital_en and cfg.get("영수증에서_병원_찾기", True):
        name = find_hospital_name_in_text(ocr_text)
        if name:
            rec.hospital_key, rec.hospital_en, rec.hospital_new = hospital_key(name), hospital_english_name(name), name
    rec.hospital_src = HOSP_FROM_OCR if rec.hospital_en else ""
    rec.city = detect_city(text_c, cfg.get("도시_목록") or DEFAULT_CITIES, rec.hospital_key,
                           cfg.get("도시_낮은_우선순위", ["서울"]))
    rec.is_receipt = looks_like_receipt(text_c, rec.amounts, rec.date_from_ocr, rec.category,
                                        cfg.get("영수증_판별_단어"))
    rec.use = rec.is_receipt
    return rec


def receipt_issues(rec: Receipt, receipts: list) -> list:
    issues = []
    if not rec.use:
        return issues
    if not rec.date_from_ocr and "날짜" not in rec.edited:
        issues.append("날짜 못 읽음(촬영 날짜 사용)")
    if not rec.total:
        issues.append("금액 못 읽음")
    if not rec.category:
        issues.append("분류 없음")
    # 병원 이름이 없는 영수증이 많아 '병원 없음'은 확인할 것으로 치지 않음(Comment에서 ' in {hospital}'만 빠짐)
    if duplicate_count(rec, receipts) > 1:
        issues.append("같은 날짜·금액 영수증 있음")
    return issues


def duplicate_count(rec: Receipt, receipts: list) -> int:
    return sum(1 for r in receipts if r.use and r.main_date == rec.main_date and r.total == rec.total)


# ---------- 영수증 목록 (엑셀로 고치면 다음 실행에도 유지) ----------

LIST_HEADER = ["파일", "촬영 시각", "사용(O/X)", "날짜", "금액", "분류", "병원 영문", "병원 근거", "도시", "확인 필요",
               "OCR 앞부분", "[자동] 사용", "[자동] 날짜", "[자동] 금액", "[자동] 분류", "[자동] 병원 영문", "[자동] 도시",
               "날짜 후보", "금액 후보", "찾은 병원 글자", "사진 위치"]
NOT_RECEIPT_NOTE = "영수증 아닌 것 같음(자동 제외)"
UNDATED_NOTE = "날짜 못 읽음: 영수증 날짜를 적어 주세요(자동 제외)"
OTHER_MONTH_NOTE = "다른 달 영수증({date})(자동 제외)"


def place_in_month(rec: Receipt, tm: TargetMonth):
    """영수증에 적힌 날짜를 최우선으로 처리할 달 영수증인지 정합니다.
    촬영 날짜는 영수증 날짜를 못 읽었을 때만 참고합니다. → rec.skip_reason, rec.use"""
    prefix = tm.prefix + "-"
    date_edited = "날짜" in rec.edited
    rec.skip_reason = ""
    if not rec.is_receipt and not date_edited:      # 날짜를 직접 적었으면 영수증으로 봄
        rec.skip_reason = NOT_RECEIPT_NOTE
    else:
        receipt_dates = rec.dates if date_edited else rec.ocr_dates
        if receipt_dates:
            mine = [d for d in receipt_dates if d.startswith(prefix)]
            if mine:
                rec.dates = [mine[0]] + [d for d in rec.dates if d != mine[0]]
            else:
                rec.skip_reason = OTHER_MONTH_NOTE.format(date=receipt_dates[0])
        elif not rec.main_date.startswith(prefix):
            rec.skip_reason = UNDATED_NOTE
    rec.use = not rec.skip_reason


def list_csv_path(results_dir: Path, tm: TargetMonth) -> Path:
    return results_dir / f"영수증목록_{tm.prefix}.csv"


def parse_use_cell(text: str, default: bool) -> bool:
    t = (text or "").strip().upper()
    if t in ("X", "N", "NO", "아니오", "제외"):
        return False
    if t in ("O", "Y", "YES", "예", "사용"):
        return True
    return default


def apply_user_edits(rec: Receipt, row: dict, tm=None, cities=None, categories=None):
    """'[자동]' 칸과 다르게 바뀐 칸만 사용자가 고친 것으로 보고 반영합니다."""
    g = lambda k: str(row.get(k) or "").strip()
    user_use = parse_use_cell(g("사용(O/X)"), rec.use)
    auto_use = parse_use_cell(g("[자동] 사용"), True)     # 예전 목록에는 이 칸이 없음 → O로 봄
    user_date, auto_date = parse_date_cell(g("날짜")), parse_date_cell(g("[자동] 날짜"))
    if user_date and user_date != auto_date:
        rec.dates, rec.edited = [user_date], rec.edited + ["날짜"]
    user_amt, auto_amt = parse_amount_cell(g("금액")), parse_amount_cell(g("[자동] 금액"))
    if user_amt is not None and user_amt != auto_amt:
        rec.total, rec.amounts, rec.edited = user_amt, [user_amt], rec.edited + ["금액"]
    if g("분류") != g("[자동] 분류"):
        rec.category = g("분류") if g("분류") in (categories or CATEGORIES) else ""
        rec.edited.append("분류")
    if g("병원 영문") != g("[자동] 병원 영문"):
        rec.hospital_en = g("병원 영문")
        rec.hospital_src = HOSP_FROM_USER if rec.hospital_en else ""
        rec.edited.append("병원")
    if "도시" in row and g("도시") != g("[자동] 도시"):
        rec.city = resolve_city(g("도시"), cities or DEFAULT_CITIES)
        rec.edited.append("도시")
    if tm is not None and "날짜" in rec.edited:
        place_in_month(rec, tm)          # 고친 날짜로 이 달 영수증인지 다시 정함
    if user_use != auto_use:
        rec.use = user_use
        rec.edited.append("사용")


def is_auto_excluded(rec: Receipt) -> bool:
    return not rec.use and bool(rec.skip_reason)


def _list_order(r: Receipt):
    """이 달 영수증(영수증 날짜 순) → 날짜 적어야 할 것 → 다른 달 영수증 → 영수증 아님"""
    if not is_auto_excluded(r):
        rank = 0
    elif r.skip_reason == NOT_RECEIPT_NOTE:
        rank = 3
    elif r.skip_reason == UNDATED_NOTE:
        rank = 1
    else:
        rank = 2
    return (rank, r.main_date, r.taken or datetime.min, r.original.name.lower())


def save_receipt_list(receipts: list, auto: dict, path: Path):
    rows = []
    order = sorted(receipts, key=_list_order)
    for r in order:
        a = auto[r.original.name]
        issues = receipt_issues(r, receipts) or ([r.skip_reason] if is_auto_excluded(r) else [])
        rows.append([
            r.original.name, f"{r.taken:%Y-%m-%d %H:%M}" if r.taken else "",
            "O" if r.use else "X", r.main_date, r.total or "", r.category, r.hospital_en, r.hospital_src,
            city_name(r.city), ", ".join(issues), compact(r.ocr_text)[:100],
            "O" if a.use else "X", a.main_date, a.total or "", a.category, a.hospital_en, city_name(a.city),
            " ".join(a.dates), " ".join(str(x) for x in a.amounts), a.hospital_key, str(r.original),
        ])
    write_csv(path, LIST_HEADER, rows)


# ---------- 영수증 확인 창(엑셀 대신)에서 쓰는 도우미 ----------
REVIEW_USER_COLUMNS = ("사용(O/X)", "날짜", "금액", "분류", "병원 영문", "도시")


def review_row_edited(row: dict) -> bool:
    """영수증 목록 한 줄을 사용자가 고쳤는지('[자동]' 칸과 비교, apply_user_edits와 같은 기준)"""
    g = lambda k: str(row.get(k) or "").strip()
    if parse_use_cell(g("사용(O/X)"), True) != parse_use_cell(g("[자동] 사용"), True):
        return True
    if parse_date_cell(g("날짜")) and parse_date_cell(g("날짜")) != parse_date_cell(g("[자동] 날짜")):
        return True
    if parse_amount_cell(g("금액")) is not None and parse_amount_cell(g("금액")) != parse_amount_cell(g("[자동] 금액")):
        return True
    return any(g(k) != g(f"[자동] {k}") for k in ("분류", "병원 영문", "도시"))


def review_status(row: dict) -> str:
    """영수증 확인 창 목록의 상태: ✕ 뺌 / ✎ 고침 / ⚠ 확인 / ✓"""
    if not parse_use_cell(str(row.get("사용(O/X)") or ""), True):
        return "✕ 뺌"
    if review_row_edited(row):
        return "✎ 고침"
    if str(row.get("확인 필요") or "").strip():
        return "⚠ 확인"
    return "✓"


def save_review_rows(path: Path, rows: list):
    header = LIST_HEADER + [k for k in (rows[0].keys() if rows else []) if k not in LIST_HEADER]
    write_csv(path, header, [[row.get(h, "") for h in header] for row in rows])


def find_receipt_photo(row: dict, cfg: dict, index: dict = None):
    """영수증 목록 한 줄의 사진 파일: '사진 위치' 칸 → 영수증 폴더들에서 같은 이름"""
    where = str(row.get("사진 위치") or "").strip()
    if where and Path(where).is_file():
        return Path(where)
    if index is None:
        index = {}
    if not index:
        for folder in receipt_folders(cfg):
            if folder.is_dir():
                for f in iter_photo_files(folder):
                    index.setdefault(f.name, f)
        index.setdefault("", None)
    return index.get(str(row.get("파일") or ""))


def load_preview_image(path: Path, max_w: int, max_h: int):
    """사진을 미리보기 크기로 읽습니다(회전 정보 반영, HEIC 포함) → PIL 이미지"""
    from PIL import Image, ImageOps
    if path.suffix.lower() in HEIC_EXTENSIONS:
        _register_heif()
    with Image.open(path) as im:
        try:
            im.draft("RGB", (max_w * 2, max_h * 2))
        except Exception:
            pass
        im = ImageOps.exif_transpose(im)
        im = im.convert("RGB")
        im.thumbnail((max_w, max_h))
        return im


def _load_json(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save_json(path: Path, data: dict):
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")


def add_found_hospitals(receipts: list, csv_path: Path, log) -> int:
    """영수증 글자에서 새로 찾은 병원을 병원목록.csv에 더합니다(다음부터는 목록에서 찾음)."""
    rows = load_hospital_rows(csv_path)
    keys = {compact(r[0]) for r in rows}
    added = []
    for r in receipts:
        if r.use and r.hospital_new and r.hospital_src == HOSP_FROM_OCR and compact(r.hospital_key) not in keys:
            keys.add(compact(r.hospital_key))
            added.append([r.hospital_key, r.hospital_en, f"영수증에서 찾음: {r.hospital_new} (영문 이름 확인)"])
    if added:
        write_csv(csv_path, HOSPITAL_HEADER, rows + added)
        log(f"   영수증에서 새 병원 {len(added)}곳을 찾아 병원 목록에 더했습니다: "
            + ", ".join(f"{a[0]}→{a[1]}" for a in added) + " (영문 이름은 [병원 목록 편집]에서 확인)")
    return len(added)


def build_receipts(cfg: dict, tm: TargetMonth, hospitals: list, results_dir: Path, work_dir: Path, log: Log,
                   hospital_csv=None) -> list:
    date_cache_path = results_dir / "photo_date_cache.json"
    date_cache = _load_json(date_cache_path)
    originals = list_originals(cfg, tm, date_cache, log)
    _save_json(date_cache_path, date_cache)
    start, end = photo_window(cfg, tm)
    log(f"{tm.prefix} 영수증을 찾을 사진: {len(originals)}장 "
        f"({start:%Y-%m-%d} ~ {end - timedelta(days=1):%Y-%m-%d} 촬영. 어느 달인지는 영수증 날짜로 정함)")
    if not originals:
        return []
    use_pre = bool(cfg.get("영수증_보정_사용", True))

    cache_path = results_dir / "ocr_cache.json"
    cache = _load_json(cache_path)
    engine = []
    counts = {"보정": 0, "나아짐": 0}
    note = old_processed_note(cfg)
    if note:
        log("   " + note)

    def read_text(path: Path, st, key: str = None) -> str:
        key = key or path.name
        c = cache.get(key)
        if c and c.get("size") == st.st_size and c.get("mtime") == int(st.st_mtime) and "text" in c:
            return c["text"]
        if not engine:
            log("OCR 중... (사진 1장에 1~2초)")
            engine.append(WindowsOcr(work_dir))
        try:
            text = engine[0].read(path)
        except Exception as e:
            log(f"   OCR 실패: {path.name} ({e})")
            text = ""
        cache[key] = {"size": st.st_size, "mtime": int(st.st_mtime), "text": text}
        return text

    def enhanced_text(orig: Path, st) -> str:
        """보정 사본으로 다시 읽은 글자(결과는 기억해 두고, 보정 사본은 바로 지움)"""
        key = orig.name + "#보정"
        c = cache.get(key)
        if c and c.get("size") == st.st_size and c.get("mtime") == int(st.st_mtime) and "text" in c:
            return c["text"]
        copy = make_enhanced_copy(cfg, orig, work_dir, log)
        if copy is None:
            return ""
        try:
            return read_text(copy, st, key)
        finally:
            try:
                copy.unlink()
            except OSError:
                pass

    auto, receipts = {}, []
    try:
        for orig, taken in originals:
            st = orig.stat()
            ocr_file, text = orig, read_text(orig, st)
            first = analyze_receipt(orig, orig, text, cfg, hospitals, taken)
            first_is_receipt = first.is_receipt
            # 사진은 그대로 읽고, 영수증인데 금액이나 날짜를 못 읽었을 때만 보정해서 한 번 더 읽음
            if use_pre and first_is_receipt and (not first.total or not first.date_from_ocr):
                counts["보정"] += 1
                better = enhanced_text(orig, st)
                if better.strip() and info_score(analyze_receipt(orig, orig, better, cfg, hospitals, taken)) > info_score(first):
                    text = better
                    counts["나아짐"] += 1

            def analyze():
                r = analyze_receipt(orig, ocr_file, text, cfg, hospitals, taken)
                if first_is_receipt and not r.is_receipt:
                    r.is_receipt = True
                place_in_month(r, tm)
                return r

            auto[orig.name] = analyze()
            receipts.append(analyze())
    finally:
        _save_json(cache_path, cache)       # 중간에 멈춰도 읽은 결과는 남김
    if counts["보정"]:
        log(f"   금액·날짜를 못 읽어 보정해서 다시 읽은 영수증 {counts['보정']}장, 그중 더 잘 읽힌 것 {counts['나아짐']}장")

    for group in (list(auto.values()), receipts):
        fill_same_day_hospitals(group)
        fill_same_day_cities(group)
    list_path = list_csv_path(results_dir, tm)
    if list_path.exists():
        try:
            rows = {str(r.get("파일") or "").strip(): r for r in read_csv_dicts(list_path)}
        except PermissionError:
            raise RuntimeError(f"영수증 목록 파일이 열려 있어 읽을 수 없습니다. 엑셀을 닫고 다시 실행하세요.\n{list_path}")
        for rec in receipts:
            if rec.original.name in rows:
                apply_user_edits(rec, rows[rec.original.name], tm, cfg.get("도시_목록"), cfg.get("분류_우선순위"))
        fill_same_day_hospitals(receipts)
        fill_same_day_cities(receipts)
    if hospital_csv is not None:
        try:
            add_found_hospitals(receipts, Path(hospital_csv), log)
        except PermissionError:
            log("   병원 목록 파일이 열려 있어 새 병원을 더하지 못했습니다. 엑셀을 닫고 다시 실행하세요.")
    try:
        save_receipt_list(receipts, auto, list_path)
    except PermissionError:
        raise RuntimeError(f"영수증 목록 파일이 엑셀에 열려 있습니다. 엑셀을 닫고 다시 실행하세요.\n{list_path}")

    for rec in receipts:
        if is_auto_excluded(rec):
            continue
        flag = "" if rec.use else "  [사용 안 함]"
        edited = f"  (고침: {', '.join(rec.edited)})" if rec.edited else ""
        log(f"   {rec.original.name} | {rec.main_date or '-'} | {rec.total:,}원 | {rec.category or '분류없음'} | "
            f"{rec.hospital_en or '병원없음'} | {city_name(rec.city) or '기본 도시'}{edited}{flag}")
    excluded = [r for r in receipts if is_auto_excluded(r)]
    if excluded:
        other = sum(1 for r in excluded if r.skip_reason.startswith("다른 달"))
        undated = sum(1 for r in excluded if r.skip_reason == UNDATED_NOTE)
        log(f"   목록 맨 아래로 뺀 사진 {len(excluded)}장: 날짜 못 읽음 {undated}장 / 다른 달 영수증 {other}장 / "
            f"영수증 아님 {len(excluded) - other - undated}장")
        if undated:
            log("   날짜를 못 읽은 영수증은 영수증 목록에 날짜를 적어 주면 다음 실행부터 씁니다.")
    return receipts


# ============================================================
# 4. 경비 ↔ 영수증 매칭
# ============================================================

def concur_date_to_iso(text: str):
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", text or "")
    return normalize_date(m.group(3), m.group(1), m.group(2)) if m else None


def concur_amount_to_int(text: str):
    t = re.sub(r"[^\d.]", "", (text or "").replace(",", ""))
    try:
        return int(round(float(t))) if t else None
    except ValueError:
        return None


def find_receipt(receipts: list, iso_date: str, amount: int):
    """날짜·금액이 맞는 영수증. 영수증 날짜(대표 날짜)가 같은 것을 먼저 보고, 촬영 날짜 등 다른 후보는 나중에 봅니다."""
    pool = [r for r in receipts if r.use and not r.used_for and iso_date in r.dates]
    for main_first in (True, False):
        candidates = [r for r in pool if (r.main_date == iso_date) == main_first]
        for r in candidates:
            if r.total == amount:
                return r
        for r in candidates:
            if amount in r.amounts:
                return r
    return None


def find_receipt_by_date_only(receipts: list, iso_date: str):
    """금액을 못 읽은 영수증이 그날 한 장뿐이면 그 영수증(확인 필요로 표시)"""
    pool = [r for r in receipts if r.use and not r.used_for and not r.total and r.main_date == iso_date]
    return pool[0] if len(pool) == 1 else None


def city_to_choose(rec, current: str, default: str, fixed: str = ""):
    """Concur 도시 칸에 고를 도시(바꿀 필요가 없으면 None).
    영수증이 있으면 (자동 입력 항목에 정해 둔 도시 →) 영수증에서 찾은 도시, 못 찾았으면 기본 도시(광주)로 맞춥니다
    (카드 정보로 서울 등이 미리 들어가 있어도). 영수증 목록에서 도시를 직접 고쳤으면 그 도시가 먼저입니다.
    영수증이 없는 경비는 칸이 비어 있을 때만 기본 도시를 넣습니다."""
    if rec is None:
        return default if not (current or "").strip() else None
    want = (rec.city if "도시" in rec.edited else "") or fixed or rec.city or default
    return want if city_name(current).lower() != city_name(want).lower() else None


def texts_for(rec: Receipt, cfg: dict) -> dict:
    tmpl = cfg["입력_문구"].get(rec.category) if rec and rec.category else None
    if not tmpl:
        return {}
    out = {}
    for name, t in tmpl.items():
        t = t.replace("{병원}", "{hospital}")
        if not rec.hospital_en:
            t = re.sub(r"\s+(?:in|at|@)\s+\{hospital\}|\{hospital\}", "", t)
        out[name] = t.replace("{hospital}", rec.hospital_en).strip()
    return out


def prepare_upload_file(rec: Receipt, cfg: dict, work_dir: Path) -> Path:
    """원본(기본) 또는 보정본을 고르고, 5MB를 넘거나 HEIC면 JPEG 사본을 만듭니다."""
    cv2, _ = _cv2()
    src = rec.ocr_file if str(cfg.get("업로드_이미지", "원본")).strip() == "보정본" else rec.original
    limit = float(cfg.get("업로드_최대_MB", 5)) * 1024 * 1024
    must_convert = src.suffix.lower() in HEIC_EXTENSIONS     # Concur가 HEIC를 못 받을 수 있음
    if not must_convert and src.stat().st_size <= limit:
        return src
    img = imread_unicode(src)
    if img is None:
        if must_convert:
            raise RuntimeError(f"HEIC 사진을 JPEG로 바꾸지 못했습니다: {src.name}")
        return src
    work_dir.mkdir(parents=True, exist_ok=True)
    out = work_dir / f"{src.stem}_업로드용.jpg"
    scale, quality = 1.0, 90
    for _ in range(20):
        im = img if scale == 1.0 else cv2.resize(img, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
        ok, enc = cv2.imencode(".jpg", im, [cv2.IMWRITE_JPEG_QUALITY, quality])
        if ok and len(enc) <= limit * 0.95:
            enc.tofile(str(out))
            return out
        if quality > 75:
            quality -= 5
        else:
            scale *= 0.85
    raise RuntimeError(f"업로드 크기 제한 안으로 줄이지 못했습니다: {src.name}")


# ============================================================
# 5. Concur 자동화 (Playwright + PC에 설치된 Edge)
# ============================================================

JS_READ_ROWS = r"""
(sel) => {
  let root = document;
  try { root = document.querySelector(sel || '.entries-list') || document; } catch (e) {}
  const boxes = Array.from(root.querySelectorAll('input[type="checkbox"]'));
  const seen = new Set();
  const out = [];
  for (const cb of boxes) {
    let row = cb.closest('[role="row"], tr, .sapcnqr-data-grid-list__row');
    if (!row) {
      let c = cb.parentElement;
      for (let i = 0; i < 12 && c; i++, c = c.parentElement) {
        if (c.querySelectorAll('input[type="checkbox"]').length > 1) break;
        row = c;
      }
    }
    if (!row || seen.has(row)) continue;
    seen.add(row);
    out.push(Array.from(row.querySelectorAll('*')).filter(e => e.children.length === 0)
      .map(e => (e.textContent || '').trim()).filter(t => t));
  }
  return out;
}
"""


def _amount_in_cell(text: str):
    t = (text or "").strip()
    m = re.fullmatch(r"(?:([A-Z]{3}|₩)\s*)?(\d[\d,]*(?:\.\d+)?)", t)
    if not m or not (m.group(1) or "," in m.group(2) or len(re.sub(r"\D", "", m.group(2).split(".")[0])) >= 3):
        return None
    return concur_amount_to_int(m.group(2))


def parse_entry_rows(rows: list) -> list:
    """리포트 경비 목록의 줄 글자들 → [(날짜, 금액, 줄 글자)]"""
    out = []
    for texts in rows or []:
        iso = next((d for d in (concur_date_to_iso(t) for t in texts) if d), None)
        amount = next((a for a in (_amount_in_cell(t) for t in texts) if a is not None), None)
        if iso and amount is not None:
            out.append((iso, amount, " ".join(t for t in texts if t)[:200]))
    return out


def apply_vendor_hospital(rec, vendor: str, hospitals: list) -> bool:
    """영수증 글자나 직접 입력으로 병원을 못 정한 영수증에 Concur 가게 이름의 병원을 씁니다(같은 날 추정보다 우선)."""
    if rec is None or not vendor or rec.hospital_src in (HOSP_FROM_OCR, HOSP_FROM_USER) or "병원" in rec.edited:
        return False
    _, en = hospital_from_merchant(vendor, hospitals)
    if not en:
        return False
    rec.hospital_en, rec.hospital_src = en, HOSP_FROM_VENDOR
    return True


JS_INSPECT = r"""
() => {
  const vis = e => !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
  const labelOf = e => {
    if (e.labels && e.labels.length) return Array.from(e.labels).map(l => l.innerText.trim()).join(' / ');
    const ids = e.getAttribute('aria-labelledby');
    if (ids) return ids.split(/\s+/).map(x => (document.getElementById(x) || {}).innerText || '').join(' ').trim();
    return e.getAttribute('aria-label') || e.getAttribute('placeholder') || '';
  };
  const clean = t => (t || '').replace(/\s+/g, ' ').trim();
  const fields = Array.from(document.querySelectorAll(
      'input:not([type=hidden]), textarea, select, [role=combobox], [role=textbox], [contenteditable=true]'))
    .filter(vis).slice(0, 200)
    .map(e => ({tag: e.tagName.toLowerCase(), type: e.getAttribute('type') || e.getAttribute('role') || '',
                id: e.id || '', name: e.getAttribute('name') || '', label: clean(labelOf(e)).slice(0, 60),
                value: clean(e.value || e.innerText).slice(0, 40)}));
  const buttons = Array.from(document.querySelectorAll('button, [role=button], a[href]')).filter(vis)
    .map(e => clean(e.innerText || e.getAttribute('aria-label')).slice(0, 40)).filter(t => t);
  return {url: location.href.split('?')[0], title: document.title, fields, buttons: Array.from(new Set(buttons)).slice(0, 80)};
}
"""


JS_SELECT_MONTH = r"""
({month, year}) => {
  const re = new RegExp('^0?' + month + '/\\d{1,2}/' + year + '$');
  const boxes = Array.from(document.querySelectorAll('input[type="checkbox"].sapcnqr-checkbox__input'));
  const seen = new Set();
  let matched = 0, clicked = 0;
  const samples = [];
  for (const cb of boxes) {
    let row = cb.closest('[role="row"], tr, .sapcnqr-data-grid-list__row');
    if (!row) {
      let c = cb.parentElement;
      for (let i = 0; i < 12 && c; i++, c = c.parentElement) {
        if (c.querySelectorAll('input[type="checkbox"]').length > 1) break;
        row = c;
      }
    }
    if (!row || seen.has(row)) continue;
    seen.add(row);
    const leaves = Array.from(row.querySelectorAll('*')).filter(e => e.children.length === 0);
    const texts = leaves.map(e => (e.textContent || '').trim()).filter(t => t);
    if (samples.length < 3) samples.push(texts.join(' | ').slice(0, 150));
    if (!texts.some(t => re.test(t))) continue;
    matched++;
    if (!cb.checked) { cb.click(); clicked++; }
  }
  return {total: boxes.length, matched, clicked, samples};
}
"""


A_TAG = re.compile(r"<a\b([^>]*)>(.*?)</a>", re.I | re.S)
HREF_ATTR = re.compile(r"""href\s*=\s*(["'])(.*?)\1""", re.I | re.S)
SIGNIN_URL = re.compile(r"/nui/signin|/signin|/login|/sso|saml", re.I)


SIGNIN_LINK_TEXT = re.compile(r"(로그인|sign[\s-]?in|log[\s-]?in).{0,20}concur|concur.{0,20}(로그인|sign[\s-]?in|log[\s-]?in)", re.I)


def extract_login_link(html_body: str, button_text: str):
    """메일 본문(HTML)에서 글자가 button_text(‘|’로 여러 개)인 링크의 주소.
    글자가 똑같은 링크 → 글자를 포함하는 링크 → 'Concur'와 '로그인/Sign in'이 든 링크 순서로 고릅니다."""
    wants = [compact(w).lower() for w in str(button_text or "").split("|") if compact(w)]
    found = []
    for m in A_TAG.finditer(html_body or ""):
        h = HREF_ATTR.search(m.group(1))
        if not h:
            continue
        url = html_lib.unescape(h.group(2)).strip()
        if not url.lower().startswith(("https://", "http://")):
            continue
        raw = html_lib.unescape(re.sub(r"<[^>]+>", " ", m.group(2)))
        text = compact(raw).lower()
        if text in wants:
            found.append((0, url))
        elif any(w in text for w in wants):
            found.append((1, url))
        elif SIGNIN_LINK_TEXT.search(" ".join(raw.split())):
            found.append((2, url))
    return min(found)[1] if found else None


def _is_login_mail(item, subjects: list) -> bool:
    """제목이 로그인 메일 제목(한국어/영어 중 하나)이거나 보낸 사람이 concur"""
    subject = compact(getattr(item, "Subject", "") or "").lower()
    if any(s and s in subject for s in subjects):
        return True
    sender = str(getattr(item, "SenderEmailAddress", "") or "").lower()
    return "concur" in sender


def _newest_login_link(inbox, subject: str, button: str, since: datetime):
    subjects = [compact(x).lower() for x in str(subject or "").split("|") if compact(x)]
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)          # 최근 메일부터
    item = items.GetFirst()
    for _ in range(30):
        if item is None:
            break
        try:
            received = item.ReceivedTime.replace(tzinfo=None)   # Outlook 시각(PC 기준)
            if received < since - timedelta(minutes=2):
                break                                           # 더 오래된 메일만 남음
            if getattr(item, "Class", None) == 43 and _is_login_mail(item, subjects):
                link = extract_login_link(item.HTMLBody or "", button)
                if link:
                    return link
        except Exception:
            pass
        item = items.GetNext()
    return None


def find_login_link_in_outlook(cfg: dict, since: datetime, log, check_stop=None, poll_s: float = 3.0):
    """Outlook 받은 편지함에서 since 이후에 온 Concur 로그인 메일의 [SAP Concur 로그인] 링크를 찾습니다."""
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        log("   Outlook 메일을 읽으려면 pywin32 가 필요합니다:  pip install -r requirements.txt")
        return None
    subject = str(cfg.get("로그인_메일_제목") or "SAP Concur 로그인 요청|Sign-In Request to SAP Concur")
    button = str(cfg.get("로그인_메일_버튼") or "SAP Concur 로그인|Sign in to SAP Concur")
    timeout = float(cfg.get("로그인_메일_대기_초", 120))
    pythoncom.CoInitialize()
    try:
        try:
            inbox = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI").GetDefaultFolder(6)
        except Exception as e:
            log(f"   Outlook을 열지 못했습니다. (새 Outlook은 지원하지 않습니다) {e}")
            return None
        deadline = time.time() + timeout
        while True:
            link = _newest_login_link(inbox, subject, button, since)
            if link:
                log("   Outlook에서 로그인 메일을 찾았습니다.")
                return link
            if time.time() >= deadline:
                log(f"   {int(timeout)}초 동안 로그인 메일을 찾지 못했습니다.")
                return None
            if check_stop:
                check_stop()
            time.sleep(poll_s)
    finally:
        pythoncom.CoUninitialize()


PASTE_LINK_MSG = ("Outlook에서 로그인 메일을 찾지 못했습니다.\n\n"
                  "메일의 [SAP Concur 로그인](영어 메일은 [Sign in to SAP Concur]) 버튼을 마우스 오른쪽으로 눌러 [하이퍼링크 복사]를 고른 뒤 여기에 붙여 넣으세요.\n"
                  "버튼을 그냥 누르면 다른 Edge 창에서 열려서 이 프로그램 창은 로그인되지 않습니다.\n\n"
                  "[취소]를 누르면 열린 Edge 창에서 직접 로그인할 때까지 기다립니다.")


class ConcurBot:
    def __init__(self, cfg: dict, log: Log, ui=None, stop_flag=None):
        self.cfg = cfg
        self.log = log
        self.ui = ui
        self.stop_flag = stop_flag
        self.base = cfg["Concur_주소"].rstrip("/")
        self.t = cfg["화면_문구"]
        self.timeout_ms = int(cfg.get("대기시간_초", 60)) * 1000
        self.pw = self.ctx = self.page = None
        self.poll_s = 1.0
        self.login_timeout_s = 600
        self.diagnose_after_s = 15
        self.last_expense_types = []     # [화면 조사]에서 읽은 Expense Type 목록
        self.upload_wait_s = 15          # 영수증 칸(업로드 버튼)이 뜰 때까지 기다리는 시간
        self.upload_done_s = 30          # 올린 뒤 끝날 때까지 기다리는 시간
        self.find_mail_link = find_login_link_in_outlook
        self.hospitals = []          # 가게 이름으로 병원을 찾을 때 씀
        self.entries = []            # 리포트 경비 목록 [(날짜, 금액, 줄 글자)]

    def _check_stop(self):
        if self.stop_flag is not None and self.stop_flag.is_set():
            raise Cancelled()

    def start(self):
        try:
            from playwright.sync_api import sync_playwright
        except ImportError:
            raise RuntimeError("playwright 가 필요합니다:  pip install playwright")
        profile = Path(self.cfg["브라우저_프로필_폴더"])
        profile.mkdir(parents=True, exist_ok=True)
        self.pw = sync_playwright().start()
        self.ctx = self.pw.chromium.launch_persistent_context(str(profile), **self.launch_options())
        self.ctx.set_default_timeout(self.timeout_ms)
        self.page = self.ctx.pages[0] if self.ctx.pages else self.ctx.new_page()

    def launch_options(self) -> dict:
        return dict(
            channel="msedge", headless=False, no_viewport=True,
            locale="ko-KR",                  # Concur 화면이 한국어로 나오게(영어로 바뀌어 버튼을 못 찾던 문제)
            args=["--start-maximized", "--lang=ko-KR", "--disable-features=Translate"],   # 번역 창도 안 뜨게
            chromium_sandbox=True,           # '--no-sandbox' 경고줄이 뜨지 않게
        )

    def stop(self):
        for closer in (lambda: self.ctx and self.ctx.close(), lambda: self.pw and self.pw.stop()):
            try:
                closer()
            except Exception:
                pass

    # ---------- 도우미 ----------
    def _exact(self, text: str):
        return re.compile(r"^\s*" + re.escape(text) + r"\s*$")

    def sel(self, key: str) -> str:
        """설정 '화면_요소'의 위치(없으면 기본값)"""
        return str((self.cfg.get("화면_요소") or {}).get(key) or DEFAULT_ELEMENTS[key])

    def _visible_css(self, key: str) -> str:
        return ", ".join(part.strip() + ":visible" for part in self.sel(key).split(",") if part.strip())

    def field(self, key: str, timeout_s: float = None):
        """입력칸: 위치(화면_요소)로 찾고, 못 찾으면 칸 이름(화면_요소_라벨)으로 찾습니다. 끝내 못 찾으면 무엇을 고칠지 알려 줍니다."""
        page = self.page
        label = str((self.cfg.get("화면_요소_라벨") or {}).get(key) or DEFAULT_ELEMENT_LABELS.get(key, ""))
        locs = [page.locator(self.sel(key))]
        if label:
            body = "|".join(re.escape(a.strip()) for a in label.split("|") if a.strip())
            locs.append(page.get_by_label(re.compile(r"^\s*(?:" + body + r")\s*\*?\s*$", re.I)))
        deadline = time.time() + (self.timeout_ms / 1000 if timeout_s is None else timeout_s)
        while True:
            el = self._first_shown(locs)
            if el is not None:
                return el
            if time.time() >= deadline:
                raise RuntimeError(f"Concur 화면에서 '{key}' 칸을 찾지 못했습니다(위치 {self.sel(key)}). "
                                   f"Concur 화면이 바뀌었을 수 있습니다. 실행기록 폴더(기록·화면 캡처)를 보내 주시면 고칠 수 있습니다"
                                   f"(직접 고치려면 화면 조사(python concur_auto.py --inspect) 파일을 보고 설정 '화면_요소'를 고치세요).")
            time.sleep(0.3)

    def _alts(self, key: str) -> list:
        """화면 문구의 여러 글자(한국어|영어)"""
        return [a.strip() for a in str(self.t.get(key, key)).split("|") if a.strip()]

    def _re(self, key: str, exact: bool = True):
        body = "|".join(re.escape(a) for a in self._alts(key))
        return re.compile(r"^\s*(?:" + body + r")\s*$", re.I) if exact else re.compile(body, re.I)

    def _visible(self, locator, timeout_ms: int = 3000) -> bool:
        try:
            locator.wait_for(state="visible", timeout=timeout_ms)
            return True
        except Exception:
            return False

    def click_checkbox(self, locator):
        locator.wait_for(state="attached")
        try:
            locator.click(timeout=5000)
        except Exception:
            locator.evaluate("el => el.click()")

    def type_text(self, locator, text: str):
        """칸에 글을 넣습니다. 한꺼번에 넣고(fill), 칸이 받지 않으면 한 글자씩 칩니다.
        설정 '글자_입력_방식'이 '한 글자씩'이면 처음부터 한 글자씩(예전 방식, 한 글자에 0.01초)."""
        locator.wait_for(state="visible")
        locator.click()
        typed = False
        if str(self.cfg.get("글자_입력_방식", "한꺼번에")).strip() != "한 글자씩":
            try:
                locator.fill(text)
                typed = locator.input_value() == text
            except Exception:
                typed = False
        if not typed:
            locator.press("Control+A")
            locator.press("Delete")
            if text:
                locator.press_sequentially(text, delay=10)
        locator.press("Tab")

    def find_select(self, css: str):
        spans = self.page.locator(css)
        for i in range(spans.count()):
            span = spans.nth(i)
            try:
                value = (span.text_content() or "").strip()
                if span.is_visible():
                    return span, value
                control = span.locator("xpath=ancestor::*[contains(@class,'field__control')][1]")
                if control.count() and control.first.is_visible():
                    return control.first, value
            except Exception:
                continue
        return None, ""

    def _try_choose_city(self, trigger, city_text: str) -> bool:
        try:
            self.choose_city(trigger, city_text)
            return True
        except Exception as e:
            self.log(f"   도시 '{city_name(city_text)}'를 고르지 못했습니다: {str(e).splitlines()[0][:80]}")
            try:
                self.page.keyboard.press("Escape")
            except Exception:
                pass
            return False

    def choose_city(self, trigger, city_text: str):
        """도시 드롭다운: 정확한 이름이 보이면 고르고, 없으면 앞 단어(예: Mokpo)로 검색해 그 단어와 KOREA가 든 항목을 고릅니다."""
        page = self.page
        items = page.locator(self._visible_css("드롭다운 항목"))
        trigger.click()
        exact = items.filter(has_text=self._exact(city_text))
        if not self._visible(exact.first, 3000):
            head = city_name(city_text).split(" (")[0].strip()
            page.keyboard.press("Control+A")
            page.keyboard.type(head, delay=30)
            exact = items.filter(has_text=self._exact(city_text))
            if not self._visible(exact.first, 3000):
                exact = items.filter(has_text=re.compile(r"\b" + re.escape(head) + r"\b.*KOREA", re.I))
                if not self._visible(exact.first, 5000):
                    page.keyboard.press("Escape")
                    raise RuntimeError(f"도시 목록에서 '{head}'를 찾지 못했습니다.")
        exact.first.click()
        time.sleep(0.5)

    def choose_option(self, trigger, option_text: str):
        page = self.page
        options = page.locator(self._visible_css("드롭다운 항목")).filter(has_text=self._exact(option_text))
        trigger.click()
        if not self._visible(options.first, 4000):
            page.keyboard.type(option_text.split(",")[0][:20], delay=30)
            if not self._visible(options.first, 6000):
                page.keyboard.press("Escape")
                raise RuntimeError(f"드롭다운에서 '{option_text}' 항목을 찾지 못했습니다.")
        options.first.click()
        time.sleep(0.5)

    def wait_until_closed(self, stop_flag=None):
        """사용자가 Edge 창을 닫을 때까지(또는 [중지]를 누를 때까지) 기다립니다."""
        while not (stop_flag is not None and stop_flag.is_set()):
            try:
                pages = [pg for pg in self.ctx.pages if not pg.is_closed()]
            except Exception:
                return
            if not pages:
                return
            try:
                pages[-1].wait_for_timeout(1000)
            except Exception:
                pass

    def read_dropdown_options(self, key: str = "Expense Type", max_rounds: int = 80) -> list:
        """드롭다운(예: Expense Type)을 열어 항목 글자를 모두 읽고 닫습니다. 목록이 길면 아래로 내려 가며 읽습니다."""
        trigger, _ = self.find_select(self.sel(key))
        if trigger is None:
            return []
        page = self.page
        try:
            trigger.click()
        except Exception:
            return []
        items = page.locator(self._visible_css("드롭다운 항목"))
        try:
            items.first.wait_for(state="visible", timeout=5000)
        except Exception:
            page.keyboard.press("Escape")
            return []
        found = []
        for _ in range(max_rounds):
            for t in items.all_inner_texts():
                t = " ".join(t.split())
                if t and t not in found and t.lower() not in ("recently used", "최근에 사용됨", "all", "모두"):
                    found.append(t)
            try:
                moved = items.last.evaluate("""el => {
                    let c = el.parentElement;
                    while (c && c.scrollHeight <= c.clientHeight + 2) c = c.parentElement;
                    if (!c || c === document.body || c === document.documentElement) return false;
                    const before = c.scrollTop; c.scrollTop = before + Math.max(40, c.clientHeight * 0.9);
                    return c.scrollTop > before; }""")
            except Exception:
                moved = False
            if not moved:
                break
            time.sleep(0.3)
        page.keyboard.press("Escape")
        return found

    def inspect_screen(self, path: Path) -> Path:
        """지금 Concur 화면의 입력칸(id·name·칸 이름)·버튼 글자와, 설정 '화면_요소'를 이 화면에서 찾는지를 파일로 남깁니다."""
        info = self.page.evaluate(JS_INSPECT)
        lines = [f"화면 조사 {datetime.now():%Y-%m-%d %H:%M:%S} | 프로그램 {APP_VERSION}",
                 f"주소: {info.get('url', '')}", f"제목: {info.get('title', '')}", "",
                 "[설정 '화면_요소'를 이 화면에서 찾는지] (다른 화면의 칸은 '없음'이 정상입니다)"]
        for key in DEFAULT_ELEMENTS:
            try:
                n = self.page.locator(self.sel(key)).count()
            except Exception as e:
                n = f"선택자 오류: {str(e).splitlines()[0][:60]}"
            lines.append(f"  {key:<22} {self.sel(key):<55} {'찾음' if n else '없음'}{f'({n})' if isinstance(n, int) and n else ''}")
        lines += ["", "[이 화면의 입력칸]  종류 | id | name | 칸 이름 | 지금 값"]
        for f in info.get("fields", []):
            lines.append(f"  {f['tag']}/{f['type']} | #{f['id']} | {f['name']} | {f['label']} | {f['value']}")
        lines += ["", "[이 화면의 버튼·링크 글자]"] + [f"  {b}" for b in info.get("buttons", [])]
        try:
            self.last_expense_types = self.read_dropdown_options("Expense Type")
        except Exception:
            self.last_expense_types = []
        if self.last_expense_types:
            lines += ["", f"[Expense Type 목록] {len(self.last_expense_types)}개"] + [f"  {t}" for t in self.last_expense_types]
        path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        try:
            self.page.screenshot(path=str(path.with_suffix(".png")))
        except Exception:
            pass
        return path

    def screenshot(self, name: str):
        try:
            path = self.log.folder / f"{datetime.now():%H%M%S}_{name}.png"
            self.page.screenshot(path=str(path))
            self.log(f"   화면 저장: {path.name}")
        except Exception:
            pass

    # ---------- 로그인 ----------
    def _shown(self, locator) -> bool:
        try:
            return locator.count() > 0 and locator.first.is_visible()
        except Exception:
            return False

    def _first_shown(self, locators):
        """여러 찾는 방법 중 화면에 보이는 첫 요소"""
        for loc in locators:
            try:
                n = loc.count()
            except Exception:
                continue
            for i in range(min(n, 5)):
                el = loc.nth(i)
                try:
                    if el.is_visible():
                        return el
                except Exception:
                    pass
        return None

    def _button_like(self, texts, exact: bool = True, pattern=None):
        """글자(여러 개 가능)가 보이는 버튼. pattern을 주면 그 정규식으로 찾습니다."""
        page = self.page
        texts = [texts] if isinstance(texts, str) else list(texts)
        if pattern is None:
            body = "|".join(re.escape(t) for t in texts)
            pattern = re.compile(r"^\s*(?:" + body + r")\s*$", re.I) if exact else re.compile(body, re.I)
        values = ", ".join(f"input[type=submit][value='{v}'], input[type=button][value='{v}']"
                           for v in (t.replace("'", "\\'") for t in texts)) or "input[type=submit][value='']"
        return self._first_shown([
            page.get_by_role("button", name=pattern),
            page.get_by_role("link", name=pattern),
            page.locator("button, [role=button], a, div[tabindex], span[tabindex]").filter(has_text=pattern),
            page.locator(values),
            page.get_by_text(pattern),
        ])

    def _email_box(self):
        page = self.page
        return self._first_shown([
            page.get_by_label(re.compile("이메일|사용자 이름|SSO|user|email", re.I)),
            page.locator("input[type=email], input[autocomplete=username], input[name*=user i], input[id*=user i]"),
            page.locator("input[type=text], input:not([type])"),
            page.get_by_role("textbox"),
        ])

    def _next_button(self):
        return self._button_like(self._alts("다음"))

    def _email_link_button(self):
        found = self._button_like(self._alts("이메일 링크로 로그인"))
        if found is None:      # 글자가 조금 달라도 '이메일 … 링크' / 'email … link'가 든 버튼
            found = self._button_like([], pattern=re.compile(r"이메일.{0,6}링크|e-?mail.{0,20}link|link.{0,20}e-?mail", re.I))
        return found

    def on_signin_screen(self) -> bool:
        return (self._next_button() is not None and self._email_box() is not None) or self._email_link_button() is not None

    def _diagnose_login(self):
        """로그인 화면에서 누를 것을 못 찾을 때, 화면에 보이는 버튼·입력칸을 기록에 남깁니다(고칠 때 씀)."""
        try:
            buttons = self.page.evaluate("""() => Array.from(document.querySelectorAll(
                'button, [role=button], input[type=submit], input[type=button], a'))
                .filter(e => e.offsetParent !== null)
                .map(e => (e.innerText || e.value || e.getAttribute('aria-label') || '').trim().slice(0, 30))
                .filter(t => t).slice(0, 15)""")
            inputs = self.page.evaluate("""() => Array.from(document.querySelectorAll('input'))
                .filter(e => e.offsetParent !== null)
                .map(e => (e.type || '') + '/' + (e.name || e.id || '')).slice(0, 10)""")
        except Exception as e:
            buttons, inputs = [f"(읽지 못함: {e})"], []
        self.log(f"   로그인 화면에서 누를 버튼을 찾지 못했습니다. 주소: {self.page.url.split('?')[0]}")
        self.log(f"   화면의 버튼: {buttons} / 입력칸: {inputs}")
        self.screenshot("login_screen")

    def is_logged_in(self) -> bool:
        url = self.page.url or ""
        if "concursolutions.com" not in url or SIGNIN_URL.search(url) or self.on_signin_screen():
            return False
        for loc in (self.page.get_by_role("link", name=self._re("리포트 생성", exact=False)),
                    self.page.get_by_text(self._re("리포트 생성", exact=False)),
                    self.page.locator("#cnqr-app-content")):
            try:
                if loc.count() > 0:
                    return True
            except Exception:
                pass
        return False

    def _email_step(self) -> bool:
        """첫 로그인 화면: 이메일 칸 확인 → [다음]. 버튼을 못 누르면 이메일 칸에서 Enter."""
        next_btn = self._next_button()
        if next_btn is None:
            return False
        if time.time() - self._next_clicked_at < 8:           # 방금 눌렀으면 화면이 바뀔 때까지 기다림
            return True
        box = self._email_box()
        value = ""
        if box is not None:
            try:
                value = (box.input_value() or "").strip()
            except Exception:
                value = "?"                                     # 입력칸이 아닌 요소면 값이 있다고 봄
        if box is not None and not value:
            email = str(self.cfg.get("로그인_이메일") or "").strip()
            if not email:
                if not self._asked_email:
                    self.log("열린 Edge 창에서 회사 이메일을 입력하고 [다음]을 누르세요. "
                             "(프로그램 창의 'Concur 로그인 이메일'에 적어 두면 다음부터 자동으로 넣습니다)")
                    self._asked_email = True
                return True
            box.fill(email)
        self._next_attempts += 1
        self._next_clicked_at = time.time()
        if self._next_attempts % 2 == 1:
            self.log("로그인: 이메일 확인 → [다음]")
            try:
                next_btn.click(timeout=5000)
                return True
            except Exception as e:
                self.log(f"   [다음]을 누르지 못해 Enter로 넘어갑니다. ({str(e).splitlines()[0][:80]})")
        else:
            self.log("   화면이 넘어가지 않아 이메일 칸에서 Enter를 누릅니다.")
        if box is not None:
            box.press("Enter")
        else:
            next_btn.click(force=True, timeout=5000)
        if self._next_attempts >= 3 and not self._diagnosed:
            self._diagnosed = True
            self._diagnose_login()
        return True

    def _email_link_step(self) -> bool:
        """두 번째 로그인 화면: [이메일 링크로 로그인] → Outlook 메일의 링크를 이 창에서 엽니다."""
        button = self._email_link_button()
        if button is None:
            return False
        if self._link_requested_at is not None:              # 한 번만 요청
            return True
        self._link_requested_at = datetime.now()
        self.log("로그인: [이메일 링크로 로그인] → Outlook에서 로그인 메일을 기다립니다. (보통 10초 안에 옵니다)")
        button.click()
        link = self.find_mail_link(self.cfg, self._link_requested_at, self.log, self._check_stop)
        if not link and self.ui is not None and hasattr(self.ui, "ask_text"):
            link = self.ui.ask_text("로그인 링크 붙여넣기", PASTE_LINK_MSG)
        link = (link or "").strip()
        if link.lower().startswith(("https://", "http://")):
            self.log("로그인 메일의 링크를 이 Edge 창에서 엽니다.")
            self.page.goto(link, wait_until="domcontentloaded")
        else:
            self.log("열린 Edge 창에서 직접 로그인하세요. 로그인되면 자동으로 이어서 진행합니다.")
        return True

    def wait_login(self, need_home: bool = True):
        """이미 로그인돼 있으면 바로 진행합니다. 로그인 화면이면 '로그인_방법'에 따라 이메일 링크로 자동 로그인합니다."""
        page = self.page
        page.goto(self.base + "/home", wait_until="domcontentloaded")
        auto = str(self.cfg.get("로그인_방법", "이메일 링크")).strip() == "이메일 링크"
        self.log("Concur 로그인 확인 중... "
                 + ("로그인 화면이 나오면 이메일 링크로 자동 로그인합니다." if auto
                    else "로그인 화면이 뜨면 열린 Edge 창에서 직접 로그인하세요.")
                 + f" (최대 {self.login_timeout_s // 60}분)")
        self._link_requested_at, self._asked_email, self._next_clicked_at = None, False, 0.0
        self._next_attempts, self._diagnosed = 0, False
        last_error = ""
        started = time.time()
        deadline = started + self.login_timeout_s
        while time.time() < deadline:
            self._check_stop()
            try:
                if self.is_logged_in():
                    self.log("로그인 확인")
                    return
                if auto:
                    known = self._email_step() or self._email_link_step()
                    if (not known and not self._diagnosed and SIGNIN_URL.search(self.page.url or "")
                            and time.time() - started > self.diagnose_after_s):
                        self._diagnosed = True
                        self._diagnose_login()
            except Cancelled:
                raise
            except Exception as e:
                msg = str(e).splitlines()[0][:150] if str(e) else type(e).__name__
                if msg != last_error:
                    self.log(f"   로그인 화면 처리 중 문제(계속 기다립니다): {msg}")
                    last_error = msg
            time.sleep(self.poll_s)
        raise RuntimeError(f"로그인 대기 시간({self.login_timeout_s // 60}분)이 지났습니다.")

    # ---------- 리포트 생성 ----------
    def create_report(self, tm: TargetMonth):
        page = self.page
        name = tm.report_name(self.cfg["리포트_이름_형식"])
        purpose = tm.report_name(self.cfg.get("리포트_Business_Purpose", "service"), report_name=name)
        comment = tm.report_name(self.cfg.get("리포트_Comment_형식", "{report_name}"), report_name=name)
        self.log(f"리포트 생성: {name} ({tm.first_day_us} ~ {tm.last_day_us}), Business Purpose={purpose}, Comment={comment}")
        try:
            page.get_by_role("link", name=self._re("리포트 생성", exact=False)).first.click(timeout=15000)
        except Exception:
            page.goto(self.base + "/nui/expense?modal=header&context=new-report")
        self.type_text(self.field("리포트 이름"), name)
        self.type_text(self.field("리포트 시작일"), tm.first_day_us)
        self.type_text(self.field("리포트 종료일"), tm.last_day_us)
        self.type_text(self.field("리포트 Business Purpose"), purpose)
        self.type_text(self.field("리포트 Comment"), comment)
        page.get_by_role("button", name=self._re("리포트 생성")).last.click()
        page.wait_for_url(re.compile(r"/nui/expense/reports/"))
        time.sleep(3)

    # ---------- 지난달 경비 추가 ----------
    # ---------- 이미 만든 리포트 ----------
    def go_home(self):
        try:
            self.page.goto(self.base + "/home")
            self.page.wait_for_load_state()
        except Exception:
            pass

    def find_report(self, tm: TargetMonth):
        """Concur 리포트 목록에 같은 이름(예: Monthly Expense for Nov)의 리포트가 있으면 그 항목, 없으면 None(홈 화면으로 돌아감)"""
        name = tm.report_name(self.cfg["리포트_이름_형식"])
        page = self.page
        try:
            page.goto(self.base + str(self.cfg.get("리포트_목록_주소") or "/nui/expense"))
            item = page.get_by_text(name, exact=True).first
            if self._visible(item, int(float(self.cfg.get("리포트_찾기_대기_초", 8)) * 1000)):
                self.log(f"Concur에 같은 이름의 리포트가 있습니다: {name}")
                return item
        except Exception as e:
            self.log(f"   리포트 목록을 보지 못했습니다: {str(e).splitlines()[0][:80] if str(e) else type(e).__name__}")
        self.log(f"같은 이름의 리포트가 없어 새로 만듭니다: {name}")
        self.go_home()
        return None

    def open_report(self, item):
        item.click()
        self.page.wait_for_url(re.compile(r"/nui/expense/reports/"))
        time.sleep(2)

    def report_is_empty(self) -> bool:
        """연 리포트에 경비가 하나도 없는지(빈 리포트 카드가 보이면 True)"""
        page = self.page
        entries, empty = page.locator(self.sel("경비 목록")), page.locator(self.sel("사용 가능한 경비 카드"))
        deadline = time.time() + self.timeout_ms / 1000
        while time.time() < deadline:
            if self._shown(empty):
                return True
            if self._shown(entries):
                return False
            time.sleep(0.5)
        return False

    def add_available_expenses(self, tm: TargetMonth):
        page = self.page
        card = page.locator(self.sel("사용 가능한 경비 카드")).filter(has_text=self._re("사용 가능한 경비", exact=False))
        if not self._visible(card.first, 20000):
            card = page.get_by_text(self._re("사용 가능한 경비", exact=False))
        card.first.click()
        page.wait_for_url(re.compile(r"view=available_expenses"))
        # 경비 목록(체크박스)이 뜨거나 '사용 가능한 경비 없음'이 뜰 때까지
        boxes, empty = page.locator(self.sel("경비 체크박스")), page.get_by_text(self._re("사용 가능한 경비 없음", exact=False))
        deadline = time.time() + self.timeout_ms / 1000
        while True:
            if self._shown(empty):
                self.log("사용 가능한 경비가 없습니다(0건).")
                raise NoExpenses("Concur에 사용 가능한 경비가 없습니다(0건). 카드 결제가 아직 들어오지 않았을 수 있습니다.")
            try:
                if boxes.count() > 0:
                    break
            except Exception:
                pass
            if time.time() >= deadline:
                raise RuntimeError("사용 가능한 경비 목록이 뜨지 않았습니다.")
            time.sleep(0.5)
        time.sleep(2)
        r = page.evaluate(JS_SELECT_MONTH, {"month": tm.month, "year": tm.year})
        self.log(f"사용 가능한 경비 중 {tm.prefix} 경비 {r['matched']}건 선택")
        if r["matched"] == 0:
            self.log("   그 달 경비를 찾지 못했습니다. 목록 예시: " + " / ".join(r["samples"]))
            self.screenshot("available_expenses")
            raise NoExpenses(f"사용 가능한 경비 중 {tm.prefix} 경비가 없습니다. (처리할 달이 맞는지 확인하세요)")
        page.get_by_role("button", name=self._re("리포트에 추가")).click()
        page.locator(self.sel("경비 목록")).first.wait_for(state="visible")
        time.sleep(3)

    # ---------- 경비 편집 화면 열기 ----------
    def read_vendor(self) -> str:
        """경비 편집 화면의 가게(공급업체) 이름"""
        page = self.page
        locs = [page.locator(f"{sel}:visible") for sel in (self.cfg.get("가게_이름_칸") or [])]
        label = str(self.cfg.get("가게_이름_라벨") or "").strip()
        if label:
            locs.append(page.get_by_label(re.compile(label, re.I)))
        for loc in locs:
            try:
                if loc.count() == 0:
                    continue
                el = loc.first
                try:
                    value = el.input_value(timeout=1000)
                except Exception:
                    value = el.text_content(timeout=1000) or ""
                value = " ".join(value.split())
                if value:
                    return value
            except Exception:
                continue
        return ""

    def vendor_text(self, iso, amount) -> str:
        """가게 이름. 편집 화면에서 못 읽으면 리포트 경비 목록에서 같은 날짜·금액 줄의 글자"""
        value = self.read_vendor()
        if value:
            return value
        for e_iso, e_amount, text in self.entries:
            if e_iso == iso and e_amount == amount:
                return text
        return ""

    def open_first_expense(self):
        page = self.page
        # 경비 목록이 뜨거나, 경비가 하나도 없는 빈 리포트 화면이 뜰 때까지
        entries, empty = page.locator(self.sel("경비 목록")), page.locator(self.sel("사용 가능한 경비 카드"))
        deadline = time.time() + self.timeout_ms / 1000
        while not self._shown(entries):
            if self._shown(empty):
                raise NoExpenses("이 리포트에는 경비가 없습니다. 먼저 Concur에서 경비를 리포트에 넣어 주세요.")
            if time.time() >= deadline:
                break
            time.sleep(0.5)
        entries.first.wait_for(state="visible")
        try:
            self.entries = parse_entry_rows(page.evaluate(JS_READ_ROWS, self.sel("경비 목록")))
        except Exception:
            self.entries = []
        boxes = page.locator(self.sel("경비 목록 체크박스"))
        try:
            boxes.first.wait_for(state="attached", timeout=10000)
            self.click_checkbox(boxes.first)
        except Exception:
            self.click_checkbox(page.locator(self.sel("경비 목록")).locator(self.sel("경비 체크박스")).nth(1))
        toolbar_edit = page.locator(self.sel("도구 모음")).get_by_role("button", name=self._re("편집"))
        if self._visible(toolbar_edit.first, 5000):
            toolbar_edit.first.click()
        else:
            page.get_by_role("button", name=self._re("편집")).first.click()
        self.field("경비 금액")

    def progress(self) -> tuple:
        el = self.page.get_by_text(re.compile(r"^\s*\d+\s*(?:/|of)\s*\d+\s*(?:개)?\s*$")).first
        el.wait_for(state="visible")
        a, b = re.findall(r"\d+", el.inner_text())[:2]
        return int(a), int(b)

    def wait_progress(self, index: int):
        deadline = time.time() + self.timeout_ms / 1000
        while time.time() < deadline:
            try:
                if self.progress()[0] == index:
                    time.sleep(1.0)
                    return
            except Exception:
                pass
            time.sleep(0.5)
        raise RuntimeError(f"{index}번째 경비 화면으로 넘어가지 않았습니다.")

    # ---------- 경비 하나 입력 ----------
    def fill_expense(self, rec, upload, checks: list) -> list:
        """Concur 경비 화면에 위에서 아래 순서로 넣습니다: Expense Type → Business Purpose → City → LCRT →
        Receipt Status → Personal Expense → Comment → 영수증 첨부. 값은 자동 입력 편집의 항목에서 가져옵니다.
        (Transaction Date·Amount·Currency는 카드 정보로 들어와 있어 건드리지 않음)"""
        page, t, done = self.page, self.t, []
        try:
            page.locator(f'{self.sel("Receipt Status")}, {self.sel("Comment")}').first.wait_for(state="attached", timeout=8000)
        except Exception:
            pass
        category = rec.category if rec else ""
        extra = (self.cfg.get("항목_추가_입력") or {}).get(category) or {}
        texts = texts_for(rec, self.cfg) if rec else {}

        def safely(step: str, action) -> bool:
            """칸 하나를 못 채워도 다음 칸과 영수증 첨부는 계속합니다(결과 파일 '확인 필요'에 남김)."""
            try:
                action()
                return True
            except Exception as e:
                msg = str(e).splitlines()[0][:80] if str(e) else type(e).__name__
                self.log(f"   {step}을(를) 고르지 못했습니다: {msg}")
                checks.append(f"{step} 확인")
                try:
                    page.keyboard.press("Escape")
                except Exception:
                    pass
                return False

        def type_field(field_name):
            text = texts.get(field_name)
            if text:
                self.type_text(self.field(field_name), text)
                done.append(f"{field_name} 입력")

        # 1. Expense Type (Concur에 Undefined일 때만, 고르면 화면 칸이 바뀔 수 있어 맨 먼저)
        trigger, value = self.find_select(self.sel("Expense Type"))
        if trigger is not None and value == t["Undefined"]:
            new_type = (self.cfg.get("Undefined일_때_선택할_유형") or {}).get(category) if category else ""
            if not new_type:
                checks.append("Expense Type 직접 고르기(맞는 자동 입력 항목 없음)")
            elif safely(f"Expense Type({new_type})", lambda: self.choose_option(trigger, new_type)):
                done.append(f"Expense Type={new_type}")

        # 2. Business Purpose
        type_field("Business Purpose")

        # 3. City
        trigger, value = self.find_select(self.sel("도시"))
        if trigger is not None:
            default = self.cfg["기본_도시"]
            city = city_to_choose(rec, value, default, extra.get("도시", ""))
            if city:
                if self._try_choose_city(trigger, city):
                    done.append(f"도시={city_name(city)}")
                elif city != default and self._try_choose_city(trigger, default):
                    done.append(f"도시={city_name(default)}(기본)")
                    checks.append(f"도시 '{city_name(city)}'를 Concur 목록에서 못 찾아 기본 도시로")
                else:
                    checks.append("도시 확인")

        # 4. LCRT
        trigger, value = self.find_select(self.sel("LCRT"))
        want = extra.get("LCRT") or t["NA"]
        if trigger is not None and value != want and safely(f"LCRT({want})", lambda: self.choose_option(trigger, want)):
            done.append(f"LCRT={want}")

        # 5. Receipt Status (영수증을 붙일 때)
        if rec:
            trigger, value = self.find_select(self.sel("Receipt Status"))
            want_receipt = extra.get("Receipt Status") or t["Receipt"]
            if trigger is not None and value != want_receipt and safely(
                    f"Receipt Status({want_receipt})", lambda: self.choose_option(trigger, want_receipt)):
                done.append(f"Receipt Status={want_receipt}")

        # 6. Personal Expense (항목에서 체크한 것만)
        if extra.get("Personal Expense"):
            box = page.get_by_label(self._re("Personal Expense", exact=False)).first
            if safely("Personal Expense", lambda: box.check(timeout=3000)):
                done.append("Personal Expense 체크")

        # 7. Comment
        type_field("Comment")

        # 8. 영수증 첨부(오른쪽 영수증 칸)
        if rec and upload:
            self.attach_receipt(upload, done, checks)
        return done

    def _upload_control(self):
        page = self.page
        return self._first_shown([
            page.locator(self.sel("영수증 업로드 버튼")),
            page.get_by_role("button", name=self._re("영수증 업로드", exact=False)),
            page.locator("button, [role=button], a").filter(has_text=self._re("영수증 업로드", exact=False)),
        ])

    def _attached_mark(self):
        page = self.page
        return self._first_shown([
            page.get_by_role("button", name=self._re("첨부된 영수증", exact=False)),
            page.locator("button, [role=button], a").filter(has_text=self._re("첨부된 영수증", exact=False)),
        ])

    def attach_receipt(self, upload: Path, done: list, checks: list):
        """영수증 칸이 뜰 때까지 기다렸다가(최대 upload_wait_s초) [영수증 업로드]로 올립니다.
        예전에는 버튼을 1.5초 안에 못 찾으면 '이미 첨부됨'으로 보고 건너뛰어, 영수증 칸이 늦게 뜨면 첨부가 빠졌습니다."""
        page = self.page
        control = attached = None
        deadline = time.time() + self.upload_wait_s
        while True:
            control = self._upload_control()
            attached = None if control is not None else self._attached_mark()
            if control is not None or attached is not None or time.time() >= deadline:
                break
            time.sleep(0.5)
        if control is None and attached is not None:
            done.append("이미 첨부됨(건너뜀)")
            return
        files = page.locator("input[type=file]")
        if control is not None:
            try:
                with page.expect_file_chooser(timeout=15000) as chooser:
                    control.click()
                chooser.value.set_files(str(upload))
            except Exception:
                if files.count() == 0:
                    raise
                files.last.set_input_files(str(upload))
        elif files.count() > 0:
            files.last.set_input_files(str(upload))       # 버튼은 못 찾았지만 숨은 파일 칸이 있음
        else:
            checks.append("영수증 첨부 버튼을 못 찾음(첨부됐는지 확인)")
            self.log("   영수증 첨부 버튼을 찾지 못했습니다. 화면을 저장합니다.")
            self.screenshot("upload_button")
            return
        # 올라가면 [영수증 업로드] 버튼이 사라짐
        finished = False
        end = time.time() + self.upload_done_s
        while time.time() < end:
            if self._upload_control() is None:
                finished = True
                break
            time.sleep(0.5)
        if not finished:
            checks.append("영수증이 올라갔는지 확인")
            self.log("   업로드가 끝났는지 확인하지 못했습니다. Concur 화면에서 영수증이 붙었는지 확인하세요.")
        time.sleep(1)
        done.append(f"영수증 첨부({upload.name})")

    # ---------- 전체 경비 반복 ----------
    def process_expenses(self, receipts: list, work_dir: Path, stop_flag=None) -> list:
        page = self.page
        current, total = self.progress()
        self.log(f"경비 입력 시작: 총 {total}건")
        results = []
        for i in range(current, total + 1):
            if stop_flag and stop_flag.is_set():
                self.log(f"사용자가 중지했습니다. {i}번째 경비 화면에서 저장하고 멈춥니다.")
                try:
                    page.get_by_role("button", name=self._re("경비 저장")).first.click()
                    time.sleep(3)
                except Exception:
                    self.log("   저장 버튼을 누르지 못했습니다. Concur 화면에서 직접 저장하세요.")
                break
            row = {"순번": i, "Concur 날짜": "", "Concur 금액": "", "가게 이름": "", "영수증": "", "분류": "",
                   "병원": "", "병원 근거": "", "도시": "", "처리": "", "확인 필요": "", "오류": ""}
            checks = []
            try:
                self.wait_progress(i)
                date_text = self.field("경비 날짜").input_value()
                amount_text = self.field("경비 금액").input_value()
                row["Concur 날짜"], row["Concur 금액"] = date_text, amount_text
                iso, amount = concur_date_to_iso(date_text), concur_amount_to_int(amount_text)
                rec = find_receipt(receipts, iso, amount) if iso and amount is not None else None
                by_date_only = False
                if rec is None and iso:
                    rec = find_receipt_by_date_only(receipts, iso)
                    by_date_only = rec is not None
                try:
                    row["가게 이름"] = self.vendor_text(iso, amount)
                except Exception:
                    pass
                upload = None
                if rec is None:
                    checks.append("영수증 없음")
                else:
                    rec.used_for = f"{i}번 {date_text} {amount_text}"
                    apply_vendor_hospital(rec, row["가게 이름"], self.hospitals)
                    row["영수증"], row["분류"] = rec.original.name, rec.category
                    row["병원"], row["병원 근거"] = rec.hospital_en, rec.hospital_src
                    row["도시"] = city_name(rec.city) or f"기본({city_name(self.cfg['기본_도시'])})"
                    if by_date_only:
                        checks.append("금액 못 읽은 영수증을 날짜로 첨부")
                    if not rec.category:
                        checks.append("분류 없음(영수증만 첨부)")
                    if duplicate_count(rec, receipts) > 1:
                        checks.append("같은 날짜·금액 영수증 여러 장")
                    upload = prepare_upload_file(rec, self.cfg, work_dir)
                done = self.fill_expense(rec, upload, checks)
                row["처리"] = ", ".join(done) or "변경 없음"
                self.log(f"   ({i}/{total}) {date_text} {amount_text}원 → "
                         f"{rec.original.name if rec else '영수증 없음'} | {row['처리']}")
            except Exception as e:
                row["오류"] = str(e).splitlines()[0][:200]
                checks.append("오류")
                self.log(f"   ({i}/{total}) 오류: {row['오류']}")
                self.screenshot(f"expense_{i}")
            row["확인 필요"] = ", ".join(checks)
            results.append(row)
            try:
                if i == total:
                    page.get_by_role("button", name=self._re("경비 저장")).first.click()
                    time.sleep(3)
                    self.log("마지막 경비 저장 완료")
                else:
                    page.get_by_role("button", name=self._re("다음 경비", exact=False)).first.click()
            except Exception as e:
                self.log(f"   다음 경비로 넘어가지 못해 중단합니다: {e}")
                self.screenshot("next_failed")
                break
        return results


def save_results_csv(results: list, path: Path):
    if results:
        write_csv(path, list(results[0].keys()), [list(r.values()) for r in results])


# ============================================================
# 6. 전체 작업 (창/명령 프롬프트 공통)
# ============================================================

class ConsoleUI:
    def log(self, line):
        print(line, flush=True)

    def confirm(self, title, msg) -> bool:
        return input(f"\n[{title}] {msg}\n(Enter=확인 / n=취소): ").strip().lower() not in ("n", "no", "아니오")

    def yes_no_cancel(self, title, msg):
        a = input(f"\n[{title}] {msg}\n(y=예 / n=아니오 / c=취소): ").strip().lower()
        return None if a in ("c", "취소") else a in ("y", "yes", "예", "")

    def ask_text(self, title, msg):
        a = input(f"\n[{title}] {msg}\n(붙여 넣고 Enter / 그냥 Enter=건너뛰기): ").strip()
        return a or None

    def info(self, title, msg):
        print(f"\n[{title}] {msg}")

    def open(self, path):
        open_path(path)


def unfinished_month_note(tm: TargetMonth) -> str:
    return (f"{tm.prefix}: 아직 끝나지 않은 달입니다. "
            f"Concur에 아직 들어오지 않은 경비와 앞으로 받을 영수증은 빠집니다. "
            f"달이 끝난 뒤 남은 경비를 Concur에서 리포트에 추가하고 ③으로 이어서 넣으면 됩니다.")


def run_job(mode: str, month_text, ui, base: Path, stop_flag=None):
    """mode: ocr(영수증 분석만) / full(전체) / edit(만든 리포트에 입력만)"""
    cfg, _ = load_settings(base / SETTINGS_FILE_NAME)
    tm = resolve_target_month(month_text)
    results_dir = base / "결과"
    results_dir.mkdir(parents=True, exist_ok=True)
    log = Log(base / "실행기록", ui.log)
    work_dir = log.folder / "임시"
    bot = None
    keep_open = False
    title = {"ocr": "영수증 분석", "full": "전체 실행", "edit": "리포트에 입력", "inspect": "화면 조사"}[mode]
    if mode == "inspect":
        try:
            log(f"===== 화면 조사 | 프로그램 {APP_VERSION} =====")
            bot = ConcurBot(cfg, log, ui, stop_flag)
            bot.start()
            bot.wait_login(need_home=False)
            made, types_added = [], 0
            while ui.confirm("화면 조사", "열린 Edge 창에서 조사할 Concur 화면(예: 경비 편집 화면, 리포트 만들기 화면)을 연 뒤 [확인]을 누르세요.\n"
                                         "다 끝났으면 [취소]를 누르세요."):
                path = bot.inspect_screen(log.folder / f"화면조사_{datetime.now():%Y%m%d_%H%M%S}.txt")
                made.append(path)
                log(f"화면 조사 파일: {path.name} (같은 이름의 .png 화면 캡처도 있습니다)")
                if bot.last_expense_types:
                    added = remember_expense_types(base / SETTINGS_FILE_NAME, bot.last_expense_types)
                    types_added += len(bot.last_expense_types)
                    log(f"Concur Expense Type {len(bot.last_expense_types)}개를 읽었습니다(새로 {added}개). "
                        "자동 입력 편집 창의 항목 이름 목록에서 고를 수 있습니다.")
            if made:
                ui.open(log.folder)
                ui.info("화면 조사", f"화면 조사 파일 {len(made)}개를 '실행기록' 폴더에 만들었습니다.\n"
                                   "'없음'으로 나온 칸은 이 파일의 id·name을 보고 설정 파일의 '화면_요소'를 고치거나, "
                                   "파일을 Claude에게 보내 고쳐 달라고 하세요."
                                   + (f"\n\nConcur Expense Type {types_added}개를 읽었습니다. "
                                      "[자동 입력 편집]의 [새 항목]이나 항목 이름 목록에서 고르면 됩니다." if types_added else ""))
        except Cancelled:
            log("중단했습니다.")
        except Exception as e:
            log("오류: " + str(e))
            ui.info("오류", f"{e}\n\n자세한 내용: {log.path.name}")
        finally:
            if bot:
                bot.stop()
            log.close()
        return
    try:
        log(f"===== {title} | {tm.prefix} ({tm.month_name}) | 프로그램 {APP_VERSION} =====")
        shutil.rmtree(work_dir, ignore_errors=True)          # 지난 실행의 임시 파일
        if is_unfinished_month(tm):
            log(unfinished_month_note(tm))
        try:
            upgrade_category_table(base / CATEGORY_FILE_NAME, base / SETTINGS_FILE_NAME, log)
            apply_category_table(cfg, base / CATEGORY_FILE_NAME, log)
        except PermissionError:
            log("자동 입력표가 엑셀에 열려 있어 읽지 못했습니다. 엑셀을 닫고 다시 실행하세요.")
            raise Cancelled()
        except Exception as e:
            log(f"자동 입력표를 읽지 못해 설정 파일의 값으로 진행합니다: {e}")
        try:
            import_hospitals_from_excel(cfg, base / HOSPITAL_FILE_NAME, log)
        except Exception as e:
            log(f"병원 엑셀을 읽지 못해 지금 병원 목록으로 진행합니다: {e}")
        hospitals = load_hospitals(base / HOSPITAL_FILE_NAME)
        receipts = build_receipts(cfg, tm, hospitals, results_dir, work_dir, log, base / HOSPITAL_FILE_NAME)
        list_path = list_csv_path(results_dir, tm)
        need = [r for r in receipts if receipt_issues(r, receipts) or r.skip_reason == UNDATED_NOTE]
        used = [r for r in receipts if r.use]
        skipped = [r for r in receipts if is_auto_excluded(r)]
        log(f"{tm.prefix} 영수증 {len(used)}장 분석 완료(목록 맨 아래로 뺀 사진 {len(skipped)}장), "
            f"확인 필요 {len(need)}장 → {list_path.name}")

        if mode == "ocr":
            if receipts and hasattr(ui, "review"):
                ui.review(list_path)
                log("영수증 확인 창에서 사용 여부·날짜·금액·분류·병원·도시를 고르거나 고친 뒤 [저장]하세요. 다음 실행부터 고친 내용을 씁니다.")
                return
            if receipts and cfg.get("결과_자동_열기", True):
                ui.open(list_path)
            ui.info("영수증 분석 완료",
                    f"영수증 {len(used)}장 중 확인 필요 {len(need)}장입니다.\n"
                    + (f"이 달 영수증이 아니라서 뺀 사진 {len(skipped)}장(날짜 못 읽음·다른 달·영수증 아님)은 목록 맨 아래에 있습니다.\n"
                       if skipped else "")
                    + "\n"
                    + "영수증 목록 파일에서 날짜·금액·분류·병원을 고치거나 '사용'을 O/X로 바꿀 수 있습니다.\n"
                    "저장하면 다음 실행 때 고친 내용이 그대로 쓰입니다.")
            return
        if not used and not ui.confirm("영수증 없음", f"{tm.prefix} 영수증이 없습니다. 영수증 없이 경비만 처리할까요?"):
            raise Cancelled()

        if need and cfg.get("입력_전_확인", True):
            ans = ui.yes_no_cancel(
                "입력 전 확인",
                f"확인이 필요한 영수증이 {len(need)}장 있습니다.\n"
                + "\n".join(f"· {r.original.name}: {', '.join(receipt_issues(r, receipts) or [r.skip_reason])}"
                            for r in need[:8])
                + ("\n..." if len(need) > 8 else "")
                + ("\n\n[예] 영수증 확인 창에서 고치기" if hasattr(ui, "review") else "\n\n[예] 영수증 목록을 열어 고치기")
                + "\n[아니오] 그대로 진행\n[취소] 중단")
            if ans is None:
                raise Cancelled()
            if ans:
                if hasattr(ui, "review"):
                    ui.review(list_path, wait=True)           # 창을 닫으면 이어서
                else:
                    ui.open(list_path)
                    if not ui.confirm("영수증 목록 수정", "엑셀에서 고친 뒤 저장하고 엑셀을 닫았으면 [확인]을 누르세요.\n([취소]는 중단)"):
                        raise Cancelled()
                receipts = build_receipts(cfg, tm, hospitals, results_dir, work_dir, log, base / HOSPITAL_FILE_NAME)

        bot = ConcurBot(cfg, log, ui, stop_flag)
        bot.hospitals = hospitals
        bot.start()
        keep_open = bool(cfg.get("작업_후_Edge_열어두기", True))
        if mode == "edit":
            bot.wait_login(need_home=False)
            if not ui.confirm("리포트 열기", "열린 Edge 창에서 입력할 리포트를 열어 경비 목록이 보이게 한 뒤 [확인]을 누르세요."):
                raise Cancelled()
        else:
            bot.wait_login()
            name = tm.report_name(cfg["리포트_이름_형식"])
            found = bot.find_report(tm)
            again = False
            if found is not None:
                again = ui.yes_no_cancel("이미 만든 리포트", (
                    f"Concur에 '{name}' 리포트가 이미 있습니다.\n\n"
                    "[예] 그 리포트에 이어서 넣기\n[아니오] 새 리포트 만들기\n[취소] 중단"))
                if again is None:
                    raise Cancelled()
            if again:
                bot.open_report(found)
                if bot.report_is_empty():
                    bot.add_available_expenses(tm)
                else:
                    log("이미 들어 있는 경비에 이어서 입력합니다. 새로 들어온 카드 경비는 Concur에서 이 리포트에 넣은 뒤 다시 ②를 누르세요.")
            else:
                if found is not None:
                    bot.go_home()
                bot.create_report(tm)
                bot.add_available_expenses(tm)
        bot.open_first_expense()
        results = bot.process_expenses(receipts, work_dir, stop_flag)

        results_path = results_dir / f"경비처리_{tm.prefix}.csv"
        save_results_csv(results, results_path)
        matched = sum(1 for r in results if r["영수증"])
        checks = [r for r in results if r["확인 필요"]]
        unused = [r.original.name for r in receipts if r.use and not r.used_for]
        log(f"완료: 경비 {len(results)}건, 영수증 연결 {matched}건, 확인 필요 {len(checks)}건")
        if unused:
            log(f"어느 경비와도 맞지 않은 영수증 {len(unused)}장: {', '.join(unused)}")
        if cfg.get("결과_자동_열기", True):
            ui.open(results_path)
        ui.info("완료",
                f"경비 {len(results)}건 중 영수증 연결 {matched}건, 확인 필요 {len(checks)}건\n"
                + (f"맞는 경비가 없던 영수증 {len(unused)}장\n" if unused else "")
                + "\n경비처리 결과 파일의 '확인 필요' 칸을 보고 Concur에서 확인한 뒤 직접 제출하세요.\n"
                + ("Edge 창은 열어 둡니다. 확인·수정이 끝나면 Edge 창을 닫으세요(닫아야 다음 작업을 할 수 있습니다)."
                   if keep_open else "[확인]을 누르면 Edge 창이 닫힙니다."))
    except Cancelled:
        log("중단했습니다.")
        keep_open = False
    except NoExpenses as e:
        log(f"넣을 경비가 없어 끝냅니다: {e}")
        ui.info("넣을 경비 없음", f"{e}\n\n프로그램 작업은 여기서 끝냅니다."
                + ("\nConcur 창은 그대로 두니, 필요하면 직접 경비를 만들거나 나중에 다시 실행하세요. 다 보면 Edge 창을 닫으세요."
                   if keep_open else ""))
    except Exception as e:
        log("오류: " + str(e))
        log(traceback.format_exc())
        if bot and bot.page:
            bot.screenshot("unexpected_error")
        ui.info("오류", f"{e}\n\n자세한 내용: {log.path.name}"
                + ("\n\nEdge 창은 열어 둡니다. 확인이 끝나면 Edge 창을 닫으세요." if keep_open else ""))
    finally:
        if bot:
            if keep_open and not (stop_flag is not None and stop_flag.is_set()):
                log("Edge 창은 그대로 둡니다. Concur에서 확인·수정이 끝나면 Edge 창을 닫으세요. (닫아야 다음 작업을 할 수 있습니다)")
                bot.wait_until_closed(stop_flag)
                log("Edge 창이 닫혔습니다.")
            bot.stop()
        shutil.rmtree(work_dir, ignore_errors=True)          # OCR·업로드용 임시 사본
        log(f"실행 기록: {log.path}")
        log.close()


# ============================================================
# 7. 창(GUI)
# ============================================================

# ---------- 창 모양(색·글꼴) ----------
THEME = {
    "bg": "#f2f4f7", "card": "#ffffff", "border": "#dfe4ea", "text": "#1d2d3e", "muted": "#6b7787",
    "accent": "#0a6ed1", "accent_hover": "#0854a0", "accent_soft": "#e7f1fc", "accent_off": "#9fc4ea",
    "header": "#0e3a63", "header_sub": "#b6cfe8", "warn": "#b05000", "hover": "#edf1f6", "pressed": "#e0e7ef",
    "field": "#ffffff", "heading": "#f4f6f9",
}


def enable_dpi_awareness():
    """Windows 화면 배율(125%·150%)에서 글자가 흐리지 않게 합니다(창을 만들기 전에)."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def ui_font_family(tkfont) -> str:
    try:
        families = set(tkfont.families())
    except Exception:
        families = set()
    for name in ("Malgun Gothic", "맑은 고딕", "NanumGothic", "NanumBarunGothic", "Noto Sans CJK KR", "Apple SD Gothic Neo"):
        if name in families:
            return name
    return "TkDefaultFont"


def make_check_images(tk, size: int, master=None, gap: int = 0) -> dict:
    """체크 상자 그림(둥근 네모 + ✓, 오른쪽에 글자와 띄울 빈칸 gap): 꺼짐·마우스 올림·켜짐"""
    c = THEME
    r = max(2, size // 5)
    pts = [(0.24, 0.52), (0.43, 0.70), (0.77, 0.32)]
    thick = max(1.6, size * 0.13)

    def seg_dist(x, y, a, b):
        (ax, ay), (bx, by) = a, b
        dx, dy = bx - ax, by - ay
        t = max(0.0, min(1.0, ((x - ax) * dx + (y - ay) * dy) / (dx * dx + dy * dy)))
        return ((x - ax - t * dx) ** 2 + (y - ay - t * dy) ** 2) ** 0.5

    def corner_dist(x, y):
        cx, cy = min(x, size - 1 - x), min(y, size - 1 - y)
        if cx < r and cy < r:
            return ((r - cx - 0.5) ** 2 + (r - cy - 0.5) ** 2) ** 0.5
        return 0.0

    def draw(fill, border, check=None):
        img = tk.PhotoImage(master=master, width=size + gap, height=size)
        scaled = [(ax * size, ay * size) for ax, ay in pts]
        for y in range(size):
            for x in range(size):
                d = corner_dist(x, y)
                if d > r:
                    continue
                edge = x in (0, size - 1) or y in (0, size - 1) or d > r - 1
                color = border if edge else fill
                if check and not edge and min(seg_dist(x + 0.5, y + 0.5, scaled[0], scaled[1]),
                                              seg_dist(x + 0.5, y + 0.5, scaled[1], scaled[2])) <= thick / 2:
                    color = check
                img.put(color, to=(x, y))
        return img

    return {"off": draw(c["field"], "#b3bdc8"), "hover": draw(c["accent_soft"], c["accent"]),
            "on": draw(c["accent"], c["accent"], "#ffffff"), "on_hover": draw(c["accent_hover"], c["accent_hover"], "#ffffff")}


def make_dot_images(tk, size: int, master=None, gap: int = 0) -> dict:
    """목록 앞 상태 점: 설정됨(파랑)·확인 필요(주황)·안 함(회색 테두리)"""
    c = THEME

    def draw(fill, ring):
        img = tk.PhotoImage(master=master, width=size + gap, height=size)
        r = size / 2
        for y in range(size):
            for x in range(size):
                d = ((x + 0.5 - r) ** 2 + (y + 0.5 - r) ** 2) ** 0.5
                if d <= r - 0.3:
                    img.put(ring if (fill is None and d > r - 1.6) else (fill or c["card"]), to=(x, y))
        return img

    return {"dot_on": draw(c["accent"], None), "dot_warn": draw("#e08a1e", None), "dot_off": draw(None, "#a9b3be")}


def apply_theme(root, ttk, tkfont) -> dict:
    """평면적이고 깔끔한 디자인(clam 테마에 색·글꼴·여백). → {'px': 배율 맞춘 픽셀, 'font': 글꼴, 'scale': 배율}"""
    c = THEME
    try:
        scale = max(1.0, float(root.winfo_fpixels("1i")) / 96.0)
    except Exception:
        scale = 1.0

    def px(n):
        return int(round(n * scale))

    family = ui_font_family(tkfont)

    def font(size, *style):
        return (family, size, *style)

    root.configure(bg=c["bg"])
    st = ttk.Style(root)
    st.theme_use("clam")
    st.configure(".", background=c["bg"], foreground=c["text"], font=font(10), bordercolor=c["border"],
                 lightcolor=c["bg"], darkcolor=c["bg"], troughcolor=c["bg"], focuscolor=c["accent"],
                 selectbackground=c["accent_soft"], selectforeground=c["text"], insertcolor=c["text"])
    # 바탕·카드
    st.configure("TFrame", background=c["bg"])
    st.configure("Card.TFrame", background=c["card"], bordercolor=c["border"], lightcolor=c["border"],
                 darkcolor=c["border"], relief="solid", borderwidth=1)
    st.configure("Pane.TFrame", background=c["card"])
    st.configure("TLabel", background=c["bg"], foreground=c["text"])
    st.configure("Card.TLabel", background=c["card"])
    st.configure("CardMuted.TLabel", background=c["card"], foreground=c["muted"], font=font(9))
    st.configure("Field.TLabel", background=c["card"], foreground=c["muted"])
    st.configure("Value.TLabel", background=c["card"], foreground=c["accent"], font=font(10, "bold"))
    st.configure("Section.TLabel", background=c["card"], foreground=c["text"], font=font(11, "bold"))
    st.configure("AccentText.TLabel", background=c["card"], foreground=c["accent"], font=font(10, "bold"))
    st.configure("Warn.TLabel", background=c["card"], foreground=c["warn"], font=font(9))
    st.configure("Status.TLabel", background=c["card"], foreground=c["muted"], font=font(9, "bold"))
    st.configure("TSeparator", background=c["border"])
    # 버튼
    flat = {"relief": "raised", "borderwidth": 1, "focusthickness": 0}
    btn = "#f6f8fa"
    st.configure("TButton", background=btn, foreground=c["text"], bordercolor="#cbd3dc", lightcolor=btn,
                 darkcolor=btn, padding=(px(14), px(6)), font=font(10), width=-6, **flat)
    st.map("TButton", background=[("disabled", c["bg"]), ("pressed", c["pressed"]), ("active", c["hover"])],
           lightcolor=[("disabled", c["bg"]), ("pressed", c["pressed"]), ("active", c["hover"])],
           darkcolor=[("disabled", c["bg"]), ("pressed", c["pressed"]), ("active", c["hover"])],
           foreground=[("disabled", "#a3abb5")], bordercolor=[("active", "#aeb9c5"), ("disabled", c["border"])])
    for name, size, pad in (("Accent.TButton", 10, (px(16), px(6))), ("BigAccent.TButton", 12, (px(18), px(14)))):
        st.configure(name, background=c["accent"], foreground="#ffffff", bordercolor=c["accent"], lightcolor=c["accent"],
                     darkcolor=c["accent"], padding=pad, font=font(size, "bold"), **flat)
        st.map(name, background=[("disabled", c["accent_off"]), ("pressed", c["accent_hover"]), ("active", c["accent_hover"])],
               lightcolor=[("disabled", c["accent_off"]), ("pressed", c["accent_hover"]), ("active", c["accent_hover"])],
               darkcolor=[("disabled", c["accent_off"]), ("pressed", c["accent_hover"]), ("active", c["accent_hover"])],
               bordercolor=[("disabled", c["accent_off"]), ("active", c["accent_hover"])],
               foreground=[("disabled", "#ffffff")])
    st.configure("Big.TButton", padding=(px(18), px(14)), font=font(12, "bold"))
    st.configure("Tool.TButton", padding=(px(12), px(5)), font=font(9), width=0)
    # 입력칸
    st.configure("TEntry", fieldbackground=c["field"], bordercolor=c["border"], lightcolor=c["field"],
                 darkcolor=c["field"], padding=(px(8), px(5)))
    st.map("TEntry", bordercolor=[("focus", c["accent"])], lightcolor=[("focus", c["accent"])])
    st.configure("TCombobox", fieldbackground=c["field"], background=c["field"], bordercolor=c["border"],
                 lightcolor=c["field"], darkcolor=c["field"], arrowcolor=c["muted"], padding=(px(8), px(4)),
                 arrowsize=px(13))
    st.map("TCombobox", fieldbackground=[("readonly", c["field"]), ("disabled", c["bg"])],
           background=[("active", c["hover"]), ("readonly", c["field"])],
           selectbackground=[("readonly", c["field"])], selectforeground=[("readonly", c["text"])],
           bordercolor=[("focus", c["accent"])], lightcolor=[("focus", c["accent"])], arrowcolor=[("active", c["accent"])])
    root.option_add("*TCombobox*Listbox.font", font(10))
    root.option_add("*TCombobox*Listbox.background", c["field"])
    root.option_add("*TCombobox*Listbox.selectBackground", c["accent_soft"])
    root.option_add("*TCombobox*Listbox.selectForeground", c["text"])
    root.option_add("*Menu.font", font(10))
    # 체크(둥근 네모 + ✓ 그림)
    import tkinter
    checks = make_check_images(tkinter, px(16), root, gap=px(8))
    st.element_create("Nice.indicator", "image", checks["off"],
                      ("selected", "active", checks["on_hover"]), ("selected", checks["on"]),
                      ("active", checks["hover"]), sticky="")
    st.layout("TCheckbutton", [("Checkbutton.padding", {"sticky": "nswe", "children": [
        ("Nice.indicator", {"side": "left", "sticky": ""}),
        ("Checkbutton.label", {"side": "left", "sticky": "nswe"})]})])
    for name, bg in (("TCheckbutton", c["bg"]), ("Card.TCheckbutton", c["card"])):
        st.configure(name, background=bg, foreground=c["text"], padding=(0, px(3)))
        st.map(name, background=[("active", bg)])
    # 탭
    st.configure("TNotebook", background=c["bg"], borderwidth=0, tabmargins=(0, 0, 0, 0), bordercolor=c["border"],
                 lightcolor=c["card"], darkcolor=c["border"])
    st.configure("TNotebook.Tab", background=c["bg"], foreground=c["muted"], padding=(px(20), px(9)), borderwidth=0,
                 font=font(10, "bold"), bordercolor=c["bg"], lightcolor=c["bg"], focuscolor=c["card"])
    st.map("TNotebook.Tab", background=[("selected", c["card"]), ("active", c["hover"])],
           foreground=[("selected", c["accent"]), ("active", c["text"])],
           lightcolor=[("selected", c["card"])], bordercolor=[("selected", c["border"])],
           expand=[("selected", (0, 0, 0, 0))])
    # 목록(표)
    st.configure("Treeview", background=c["field"], fieldbackground=c["field"], foreground=c["text"],
                 rowheight=px(28), bordercolor=c["border"], lightcolor=c["field"], darkcolor=c["field"], font=font(10))
    st.map("Treeview", background=[("selected", c["accent_soft"])], foreground=[("selected", c["text"])])
    st.configure("Treeview.Heading", background=c["heading"], foreground=c["muted"], font=font(9, "bold"),
                 relief="flat", padding=(px(8), px(6)), bordercolor=c["border"], lightcolor=c["heading"],
                 darkcolor=c["heading"])
    st.map("Treeview.Heading", background=[("active", c["hover"])])
    # 왼쪽 목록(항목·병원): 줄 높이를 넉넉하게, 펼침 표시 없이
    st.configure("Items.Treeview", rowheight=px(32), font=font(10), indent=0)
    st.layout("Items.Treeview.Item", [("Treeitem.padding", {"sticky": "nswe", "children": [
        ("Treeitem.image", {"side": "left", "sticky": ""}),
        ("Treeitem.focus", {"side": "left", "sticky": "", "children": [("Treeitem.text", {"side": "left", "sticky": ""})]})]})])
    # 스크롤바
    for name in ("Vertical.TScrollbar", "Horizontal.TScrollbar"):
        st.configure(name, background="#dce2e9", troughcolor=c["card"], bordercolor=c["card"], arrowcolor="#9aa5b1",
                     lightcolor="#dce2e9", darkcolor="#dce2e9", gripcount=0, arrowsize=px(11), relief="flat")
        st.map(name, background=[("active", "#c3ccd6")], lightcolor=[("active", "#c3ccd6")], darkcolor=[("active", "#c3ccd6")])
    images = dict(checks, **make_dot_images(tkinter, px(10), root, gap=px(10)))
    return {"px": px, "font": font, "scale": scale, "family": family, "images": images}


def auto_wrap(label, margin: int = 0):
    """글자를 지금 칸 폭에 맞춰 줄바꿈(창이 좁아도 오른쪽이 잘리지 않게). 라벨은 sticky="ew"로 놓아야 합니다."""
    def on_resize(event):
        width = max(60, event.width - margin)
        if str(label.cget("wraplength")) != str(width):
            label.configure(wraplength=width)
    label.bind("<Configure>", on_resize, add="+")
    return label


def hline(parent, tk):
    """연한 구분선"""
    return tk.Frame(parent, height=1, bg=THEME["border"])


def item_tree(parent, ttk, app, first, second):
    """왼쪽 목록 표(점 + 이름, 두 번째 칸): first/second = (제목, 폭)"""
    c = THEME
    tree = ttk.Treeview(parent, columns=("second",), show="tree headings", selectmode="browse", style="Items.Treeview")
    tree.heading("#0", text=first[0], anchor="w")
    tree.heading("second", text=second[0], anchor="w")
    tree.column("#0", width=first[1], minwidth=app.px(90), stretch=True)
    tree.column("second", width=second[1], minwidth=app.px(90), stretch=True)
    tree.tag_configure("off", foreground=c["muted"])
    tree.tag_configure("on", foreground=c["text"])
    tree.tag_configure("warn", foreground=c["text"])
    sb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
    sb.pack(side="right", fill="y")
    tree.pack(side="left", fill="both", expand=True)
    tree.configure(yscrollcommand=sb.set)
    return tree


def text_box(parent, tk, ttk, theme, height=8, wrap="word"):
    """평면 글 상자(실행 기록 등) + 얇은 스크롤바 → (바깥 프레임, Text)"""
    c, px, font = THEME, theme["px"], theme["font"]
    frame = tk.Frame(parent, bg=c["card"], highlightthickness=1, highlightbackground=c["border"], highlightcolor=c["border"])
    text = tk.Text(frame, height=height, wrap=wrap, relief="flat", borderwidth=0, highlightthickness=0, bg=c["card"],
                   fg=c["text"], font=font(9), padx=px(10), pady=px(8), insertbackground=c["text"],
                   selectbackground=c["accent_soft"], selectforeground=c["text"], spacing1=px(1), spacing3=px(1))
    sb = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
    text.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    text.pack(side="left", fill="both", expand=True)
    return frame, text


class GuiApp:
    def __init__(self, base: Path):
        import tkinter as tk
        from tkinter import filedialog, messagebox, simpledialog, ttk
        from tkinter import font as tkfont
        self.sd, self.ttk = simpledialog, ttk
        self.editor = None
        self.tk, self.mb, self.fd = tk, messagebox, filedialog
        self.base = base
        enable_dpi_awareness()
        root = self.root = tk.Tk()
        root.title(f"Concur 경비 자동 입력 ({APP_VERSION})")
        self.theme = apply_theme(root, ttk, tkfont)
        self.px, self.font = self.theme["px"], self.theme["font"]
        px, font, c = self.px, self.font, THEME
        self.restored = self.offer_restore_on_first_run() if not (base / SETTINGS_FILE_NAME).exists() else None
        self.cfg, created = load_settings(base / SETTINGS_FILE_NAME)
        load_hospitals(base / HOSPITAL_FILE_NAME)
        self.q = queue.Queue()
        self.worker = None
        self.stop_flag = threading.Event()

        width = min(px(920), root.winfo_screenwidth() - px(40))
        height = min(px(830), root.winfo_screenheight() - px(80))
        root.geometry(f"{width}x{height}")
        root.minsize(min(px(780), width), min(px(620), height))

        # 위쪽 머리띠
        header = tk.Frame(root, bg=c["header"])
        header.pack(fill="x")
        title_box = tk.Frame(header, bg=c["header"])
        title_box.pack(side="left", padx=px(22), pady=(px(10), px(9)))
        tk.Label(title_box, text="Concur 경비 자동 입력", bg=c["header"], fg="#ffffff", font=font(16, "bold")).pack(anchor="w")
        tk.Label(title_box, text="영수증 사진을 읽어 Concur 경비 리포트를 채웁니다 · 제출은 직접 하세요", bg=c["header"],
                 fg=c["header_sub"], font=font(9)).pack(anchor="w", pady=(px(2), 0))
        tk.Label(header, text=APP_VERSION, bg=c["header"], fg=c["header_sub"], font=font(9)).pack(
            side="right", anchor="s", padx=px(22), pady=(0, px(12)))

        body = ttk.Frame(root, padding=(px(16), px(14), px(16), px(16)))
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(1, weight=1)
        nb = self.notebook = ttk.Notebook(body)
        nb.grid(row=0, column=0, sticky="nsew")
        f = ttk.Frame(nb, style="Pane.TFrame", padding=(px(22), px(14)))
        f.columnconfigure(1, weight=1)
        nb.add(f, text="실행")
        self.sources_frame = ttk.Frame(nb, style="Pane.TFrame", padding=(px(22), px(14)))
        nb.add(self.sources_frame, text="영수증 불러오기")
        self.backup_frame = ttk.Frame(nb, style="Pane.TFrame", padding=(px(22), px(14)))
        nb.add(self.backup_frame, text="백업·되살리기")

        # [실행] 탭: 기본 정보
        r = 0
        ttk.Label(f, text="기본 정보", style="Section.TLabel").grid(row=r, column=0, columnspan=3, sticky="w", pady=(0, px(8)))
        r += 1

        def field_label(text):
            ttk.Label(f, text=text, style="Field.TLabel").grid(row=r, column=0, sticky="w", pady=px(5), padx=(0, px(16)))

        field_label("처리할 달")
        months = recent_months(include_current=True)                     # 맨 위가 이번 달
        self.month = tk.StringVar(value=resolve_target_month(None).prefix)  # 처음 선택은 지난달
        mf = ttk.Frame(f, style="Pane.TFrame")
        mf.grid(row=r, column=1, columnspan=2, sticky="w")
        self.month_box = ttk.Combobox(mf, textvariable=self.month, values=months, width=10, state="readonly")
        self.month_box.pack(side="left")
        self.month_box.bind("<<ComboboxSelected>>", lambda e: self.on_month_changed())
        ttk.Label(mf, text="처음에는 지난달 · 맨 위는 이번 달", style="CardMuted.TLabel").pack(side="left", padx=(px(10), 0))
        r += 1
        field_label("영수증 폴더")
        self.folder_summary = ttk.Label(f, text="", style="Value.TLabel")
        self.folder_summary.grid(row=r, column=1, columnspan=2, sticky="w")
        r += 1
        field_label("기본 도시")
        self.city = tk.StringVar(value=self.cfg["기본_도시"])
        ttk.Entry(f, textvariable=self.city).grid(row=r, column=1, columnspan=2, sticky="ew", pady=px(4))
        r += 1
        field_label("Concur 로그인 이메일")
        self.login_email = tk.StringVar(value=str(self.cfg.get("로그인_이메일") or ""))
        ttk.Entry(f, textvariable=self.login_email).grid(row=r, column=1, sticky="ew", pady=px(4))
        self.auto_login = tk.BooleanVar(value=str(self.cfg.get("로그인_방법", "이메일 링크")).strip() == "이메일 링크")
        ttk.Checkbutton(f, text="이메일 링크로 자동 로그인", variable=self.auto_login, style="Card.TCheckbutton").grid(
            row=r, column=2, sticky="w", padx=(px(12), 0))
        r += 1
        self.pre = tk.BooleanVar(value=bool(self.cfg.get("영수증_보정_사용", True)))
        self.review = tk.BooleanVar(value=bool(self.cfg.get("입력_전_확인", True)))
        opts = ttk.Frame(f, style="Pane.TFrame")
        opts.grid(row=r, column=1, columnspan=2, sticky="w", pady=(px(4), 0))
        ttk.Checkbutton(opts, text="금액·날짜를 못 읽은 영수증은 보정해서 한 번 더 읽기", variable=self.pre,
                        style="Card.TCheckbutton").pack(anchor="w")
        ttk.Checkbutton(opts, text="Concur 입력 전에 확인이 필요한 영수증 알려 주기", variable=self.review,
                        style="Card.TCheckbutton").pack(anchor="w")
        r += 1
        hline(f, tk).grid(row=r, column=0, columnspan=3, sticky="ew", pady=(px(14), px(14)))
        r += 1

        # 실행 버튼
        act = ttk.Frame(f, style="Pane.TFrame")
        act.grid(row=r, column=0, columnspan=3, sticky="ew")
        act.columnconfigure(0, weight=1, uniform="act")
        act.columnconfigure(1, weight=1, uniform="act")
        self.run_buttons = []
        for col, (text, mode, style, caption) in enumerate((
                ("① 영수증 분석만", "ocr", "Big.TButton", "사진의 글자만 읽고 영수증 확인 창을 엽니다. Concur는 열지 않습니다."),
                ("② 전체 실행", "full", "BigAccent.TButton", "리포트 만들기 → 경비 넣기 → 영수증 첨부·입력 → 저장 (같은 이름의 리포트가 있으면 이어서)"))):
            b = ttk.Button(act, text=text, style=style, command=lambda m=mode: self.start(m))
            b.grid(row=0, column=col, sticky="ew", padx=(0 if col == 0 else px(6), px(6) if col == 0 else 0))
            self.run_buttons.append(b)
            auto_wrap(ttk.Label(act, text=caption, style="CardMuted.TLabel", wraplength=px(390), justify="left")).grid(
                row=1, column=col, sticky="new", pady=(px(6), 0), padx=(0 if col == 0 else px(6), px(6) if col == 0 else 0))
        r += 1

        tools = ttk.Frame(f, style="Pane.TFrame")
        tools.grid(row=r, column=0, columnspan=3, sticky="ew", pady=(px(18), 0))
        ttk.Label(tools, text="바로 가기", style="CardMuted.TLabel").pack(side="left", padx=(0, px(10)))
        for text, cmd in (("영수증 확인", self.open_list), ("병원 목록 편집", self.open_hospitals),
                          ("자동 입력 편집", self.open_category_table)):
            ttk.Button(tools, text=text, style="Tool.TButton", command=cmd).pack(side="left", padx=(0, px(6)))
        r += 1
        auto_wrap(ttk.Label(f, style="CardMuted.TLabel", wraplength=px(800), justify="left", text=(
            "실행 중에는 열린 Edge 창을 클릭하지 마세요(다른 프로그램은 써도 됩니다). "
            "[영수증 확인]에서 고친 내용은 다음 실행에도 그대로 쓰입니다."))).grid(
            row=r, column=0, columnspan=3, sticky="ew", pady=(px(12), 0))

        # 아래: 실행 기록
        log_card = ttk.Frame(body, style="Card.TFrame", padding=(px(14), px(10), px(14), px(14)))
        log_card.grid(row=1, column=0, sticky="nsew", pady=(px(14), 0))
        log_card.columnconfigure(0, weight=1)
        log_card.rowconfigure(1, weight=1)
        ttk.Label(log_card, text="실행 기록", style="Section.TLabel").grid(row=0, column=0, sticky="w", pady=(0, px(8)))
        self.status = ttk.Label(log_card, text="● 준비됨", style="Status.TLabel")
        self.status.grid(row=0, column=1, sticky="e", pady=(0, px(8)))
        box, self.text = text_box(log_card, tk, ttk, self.theme, height=5)
        box.grid(row=1, column=0, columnspan=2, sticky="nsew")
        self.text.configure(state="disabled")

        self.sources = ReceiptSourcesTab(self, self.sources_frame)
        self.backups = BackupTab(self, self.backup_frame)
        nb.bind("<<NotebookTabChanged>>", lambda e: self.on_tab_changed())

        root.protocol("WM_DELETE_WINDOW", self.on_close)
        root.after(100, self.poll)
        if self.restored:
            self.append(f"백업에서 되살렸습니다: {self.restored} (자동 입력 규칙·병원 목록·설정)")
        self.auto_backup("켤 때")
        if created:
            self.append("처음 실행입니다. 영수증 폴더·기본 도시를 확인하고, [병원 목록 편집]에서 담당 병원을 넣어 주세요.")
            self.append("영수증 사진이 있는 곳은 [영수증 불러오기] 탭에서 넣습니다. 아이폰은 iCloud 사진 폴더, "
                        "갤럭시는 OneDrive 사진 폴더, 그 밖에는 카카오톡 받은 파일이나 [사진 가져오기]로 넣으면 됩니다.")
            self.append("병원 목록에는 영수증에서 확인한 병원(원광대병원, 목포한국병원)만 들어 있습니다. "
                        "영수증에서 새 병원이 보이면 자동으로 더합니다.")

    # ---------- 화면 동작 ----------
    def append(self, line):
        self.text.configure(state="normal")
        self.text.insert("end", line + "\n")
        self.text.see("end")
        self.text.configure(state="disabled")

    def on_tab_changed(self):
        try:
            current = self.notebook.nametowidget(self.notebook.select())
        except Exception:
            return
        if current is self.sources_frame:
            self.sources.auto_count()
        elif current is self.backup_frame:
            self.backups.refresh()

    def on_month_changed(self):
        self.sources.counts = {}
        self.sources.refresh()
        self.on_tab_changed()

    # ---------- 백업 ----------
    def offer_restore_on_first_run(self):
        """새 폴더에 받은 프로그램을 처음 켤 때: 문서 폴더의 최근 백업을 되살릴지 묻습니다. → 되살린 백업 이름 또는 None"""
        folder = backup_dir()
        backups = list_backups(folder)
        if not backups:
            return None
        path, info = backups[0]
        when, kind, version, _ = backup_label(path, info)
        self.root.withdraw()
        try:
            if not self.mb.askyesno("예전 설정 되살리기", (
                    f"예전에 쓰던 설정 백업이 있습니다.\n\n만든 시각: {when} ({kind})\n프로그램: {version}\n\n"
                    "자동 입력 규칙·병원 목록·설정·영수증 목록을 되살릴까요?\n"
                    "(아니요를 눌러도 나중에 [백업·되살리기] 탭에서 되살릴 수 있습니다)")):
                return None
            restore_backup(path, self.base, folder)
            return f"{when} ({path.name})"
        except Exception as e:
            self.mb.showerror("예전 설정 되살리기", str(e))
            return None
        finally:
            self.root.deiconify()

    def auto_backup(self, note="자동 백업"):
        try:
            made = auto_backup(self.base, backup_dir(load_settings(self.base / SETTINGS_FILE_NAME)[0]), note)
        except Exception as e:
            self.append(f"자동 백업 실패: {e}")
            return None
        if made is not None and getattr(self, "backups", None) is not None:
            self.backups.refresh()
        return made

    def reload_from_settings(self):
        """되살린 설정을 창에 다시 보여 줌"""
        self.cfg, _ = load_settings(self.base / SETTINGS_FILE_NAME)
        load_hospitals(self.base / HOSPITAL_FILE_NAME)
        self.city.set(self.cfg["기본_도시"])
        self.login_email.set(str(self.cfg.get("로그인_이메일") or ""))
        self.auto_login.set(str(self.cfg.get("로그인_방법", "이메일 링크")).strip() == "이메일 링크")
        self.pre.set(bool(self.cfg.get("영수증_보정_사용", True)))
        self.review.set(bool(self.cfg.get("입력_전_확인", True)))
        self.sources.folders = [str(f) for f in receipt_folders(self.cfg)]
        self.sources.counts = {}
        self.sources.refresh()

    def update_folder_summary(self, folders: list):
        if not folders:
            text = "없음 — [영수증 불러오기] 탭에서 넣으세요"
        else:
            names = [Path(f).name or str(f) for f in folders]
            text = f"{len(folders)}곳: " + ", ".join(names[:3]) + (" …" if len(names) > 3 else "")
        self.folder_summary.configure(text=text)

    def open_list(self):
        """[영수증 확인]: 처리할 달의 영수증 확인 창"""
        path = list_csv_path(self.base / "결과", resolve_target_month(self.month.get()))
        if path.exists():
            self.open_review(path)
        else:
            self.mb.showinfo("영수증 확인", "아직 이 달의 영수증 목록이 없습니다. ①을 먼저 실행하세요.")

    def open_review(self, path, on_close=None):
        """영수증 확인 창(이미 열려 있으면 닫고 새로)"""
        old = getattr(self, "review_win", None)
        if old is not None and old.win.winfo_exists():
            if Path(path) == old.path and on_close is None:
                old.win.lift()
                return old
            old.close(force=True)
        try:
            self.review_win = ReceiptReview(self, Path(path), on_close=on_close)
        except Exception as e:
            self.mb.showerror("영수증 확인", f"영수증 목록을 열지 못했습니다: {e}")
            if on_close:
                on_close()
            return None
        return self.review_win

    def open_category_table(self):
        """자동 입력 편집 창(이미 열려 있으면 앞으로)"""
        if self.editor is not None and self.editor.win.winfo_exists():
            self.editor.win.lift()
            return
        try:
            self.editor = CategoryEditor(self)
        except PermissionError:
            self.mb.showerror("자동 입력 편집", "자동입력표.csv가 엑셀에 열려 있습니다. 엑셀을 닫고 다시 누르세요.")

    def open_hospitals(self):
        """병원 목록 편집 창(엑셀 대신, 이미 열려 있으면 앞으로)"""
        old = getattr(self, "hosp_editor", None)
        if old is not None and old.win.winfo_exists():
            old.win.lift()
            return old
        try:
            self.hosp_editor = HospitalEditor(self)
        except PermissionError:
            self.mb.showerror("병원 목록 편집", "병원목록.csv가 엑셀에 열려 있습니다. 엑셀을 닫고 다시 누르세요.")
            return None
        return self.hosp_editor

    def save_form(self):
        cfg, _ = load_settings(self.base / SETTINGS_FILE_NAME)
        set_receipt_folders(cfg, self.sources.folders)
        cfg["기본_도시"] = self.city.get().strip()
        cfg["영수증_보정_사용"] = bool(self.pre.get())
        cfg["입력_전_확인"] = bool(self.review.get())
        cfg["로그인_이메일"] = self.login_email.get().strip()
        cfg["로그인_방법"] = "이메일 링크" if self.auto_login.get() else "직접"
        save_settings(self.base / SETTINGS_FILE_NAME, cfg)

    def set_running(self, running: bool):
        for b in self.run_buttons:
            b.configure(state="disabled" if running else "normal")
        self.status.configure(text="● 실행 중… (끝나면 여기에 알려 드립니다)" if running else "● 준비됨",
                              foreground=THEME["accent"] if running else THEME["muted"])

    def start(self, mode):
        if self.worker and self.worker.is_alive():
            return
        tm = resolve_target_month(self.month.get())
        if mode == "full" and is_unfinished_month(tm) and not self.mb.askokcancel(
                "전체 실행", f"{unfinished_month_note(tm)}\n\n계속할까요?"):
            return
        try:
            self.save_form()
        except Exception as e:
            self.mb.showerror("설정 저장 실패", str(e))
            return
        self.stop_flag.clear()
        self.set_running(True)
        ui = GuiBridge(self.q)
        self.worker = threading.Thread(target=self._work, args=(mode, self.month.get(), ui), daemon=True)
        self.worker.start()

    def _work(self, mode, month, ui):
        try:
            run_job(mode, month, ui, self.base, self.stop_flag)
        except Exception as e:
            ui.log(f"오류: {e}")
        finally:
            self.q.put(("done",))

    def poll(self):
        try:
            while True:
                item = self.q.get_nowait()
                kind = item[0]
                if kind == "log":
                    self.append(item[1])
                elif kind == "open":
                    open_path(item[1])
                elif kind == "ask":
                    _, how, title, msg, holder, ev = item
                    if how == "okcancel":
                        holder[0] = self.mb.askokcancel(title, msg, parent=self.root)
                    elif how == "yesnocancel":
                        holder[0] = self.mb.askyesnocancel(title, msg, parent=self.root)
                    elif how == "text":
                        holder[0] = self.sd.askstring(title, msg, parent=self.root)
                    else:
                        self.mb.showinfo(title, msg, parent=self.root)
                    ev.set()
                elif kind == "call":
                    item[1]()
                elif kind == "review":
                    _, path, ev = item
                    self.open_review(path, on_close=ev.set if ev is not None else None)
                elif kind == "done":
                    self.set_running(False)
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def on_close(self):
        if self.worker and self.worker.is_alive():
            if not self.mb.askyesno("종료", "작업 중입니다. 종료하면 Edge 창도 닫힙니다. 종료할까요?"):
                return
            self.auto_backup("끌 때")
            self.root.destroy()
            os._exit(0)
        self.auto_backup("끌 때")
        self.root.destroy()

    def run(self):
        self.root.mainloop()


PHONE_GUIDE = (
    ("아이폰", "[폴더 추가 ▾]에서 iCloud 사진이나 iCloud Drive 영수증 폴더를 고르면, 찍기만 해도 들어옵니다."),
    ("갤럭시", "갤러리 앱 → 설정 → 'OneDrive와 동기화'를 켠 뒤, [폴더 추가 ▾]에서 OneDrive 갤럭시 사진을 고르세요."),
    ("그 밖에", "카카오톡 '나와의 채팅'으로 원본 사진을 보내 PC에서 저장(카카오톡 받은 파일)하거나, "
              "USB·메일로 받은 사진은 PC에 옮긴 뒤 [사진 가져오기…]로 넣으세요."),
)


class ReceiptSourcesTab:
    """[영수증 불러오기] 탭: 영수증을 찾을 폴더 여러 곳([폴더 추가 ▾]에 이 PC에서 찾은 곳), 사진 가져오기, 이 달 사진 수(자동)"""

    def __init__(self, app, parent):
        tk, ttk = app.tk, app.ttk
        px = app.px
        self.app, self.tk = app, tk
        cfg, _ = load_settings(app.base / SETTINGS_FILE_NAME)
        self.folders = [str(f) for f in receipt_folders(cfg)]
        self.counts = {}
        self.counting, self.recount = False, False
        parent.columnconfigure(0, weight=1)

        ttk.Label(parent, text="영수증을 찾을 폴더", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        auto_wrap(ttk.Label(parent, style="CardMuted.TLabel", wraplength=px(820), justify="left", text=(
            "영수증 사진이 있는 폴더를 모두 넣어 두세요. 하위 폴더까지 찾고, 영수증이 아닌 사진은 자동으로 뺍니다. "
            "오른쪽 숫자는 처리할 달 영수증일 수 있는 사진 수입니다. 같은 영수증을 두 가지 방법으로 넣지는 마세요(두 장으로 보임)."))
                  ).grid(row=1, column=0, sticky="ew", pady=(px(2), px(8)))

        box = ttk.Frame(parent, style="Pane.TFrame")
        box.grid(row=2, column=0, sticky="ew")
        box.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(box, columns=("folder", "count"), show="headings", height=4, selectmode="browse")
        self.tree.heading("folder", text="폴더")
        self.tree.heading("count", text="이 달 사진")
        self.tree.column("folder", width=px(300), minwidth=px(150), anchor="w", stretch=True)
        self.tree.column("count", width=px(100), minwidth=px(80), anchor="center", stretch=False)
        self.tree.grid(row=0, column=0, sticky="ew")
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=sb.set)

        bf = ttk.Frame(parent, style="Pane.TFrame")
        bf.grid(row=3, column=0, sticky="ew", pady=(px(10), 0))
        self.add_button = ttk.Button(bf, text="폴더 추가 ▾", style="Accent.TButton", command=self.show_add_menu)
        self.add_button.pack(side="left")
        ttk.Button(bf, text="사진 가져오기…", command=self.import_dialog).pack(side="left", padx=(px(6), 0))
        ttk.Button(bf, text="폴더 열기", command=self.open_selected).pack(side="left", padx=(px(6), 0))
        ttk.Button(bf, text="빼기", command=self.remove_selected).pack(side="left", padx=(px(6), 0))
        self.import_status = ttk.Label(bf, text="", style="AccentText.TLabel")
        self.import_status.pack(side="left", padx=(px(12), 0))

        hline(parent, tk).grid(row=4, column=0, sticky="ew", pady=(px(12), px(10)))
        ttk.Label(parent, text="휴대폰별 사진 넣는 방법", style="Section.TLabel").grid(row=5, column=0, sticky="w", pady=(0, px(4)))
        guide = ttk.Frame(parent, style="Pane.TFrame")
        guide.grid(row=6, column=0, sticky="ew")
        guide.columnconfigure(1, weight=1)
        for i, (phone, how) in enumerate(PHONE_GUIDE):
            ttk.Label(guide, text=phone, style="Field.TLabel").grid(row=i, column=0, sticky="nw", padx=(0, px(14)), pady=px(2))
            auto_wrap(ttk.Label(guide, text=how, style="Card.TLabel", wraplength=px(700), justify="left")).grid(
                row=i, column=1, sticky="ew", pady=px(2))
        self.refresh()

    # ---------- 목록 ----------
    def imported_folder(self) -> Path:
        return self.app.base / IMPORTED_FOLDER_NAME

    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        for i, folder in enumerate(self.folders):
            if not Path(folder).is_dir():
                count = "폴더 없음"
            elif _norm_folder(folder) in self.counts:
                count = f"{self.counts[_norm_folder(folder)]}장"
            else:
                count = "세는 중…" if self.counting else "-"
            self.tree.insert("", "end", iid=str(i), values=(folder, count))
        self.app.update_folder_summary(self.folders)

    def save(self):
        path = self.app.base / SETTINGS_FILE_NAME
        cfg, _ = load_settings(path)
        set_receipt_folders(cfg, self.folders)
        save_settings(path, cfg)

    def add_folder(self, folder) -> bool:
        folder = str(Path(folder))
        if any(_norm_folder(f) == _norm_folder(folder) for f in self.folders):
            return False
        self.folders.append(folder)
        self.save()
        self.refresh()
        self.auto_count()
        return True

    def add_place(self, label, path):
        if self.add_folder(path):
            self.app.append(f"영수증 폴더 추가: {label} ({path})")

    def show_add_menu(self):
        """[폴더 추가 ▾]: 이 PC에서 찾은 곳(iCloud·OneDrive 갤럭시 사진·구글 드라이브·카카오톡·다운로드) + 직접 고르기"""
        tk, c = self.tk, THEME
        menu = tk.Menu(self.add_button, tearoff=0, bg=c["card"], fg=c["text"], activebackground=c["accent_soft"],
                       activeforeground=c["text"], disabledforeground="#a3abb5", relief="flat", borderwidth=1,
                       font=self.app.font(10))
        added = {_norm_folder(f) for f in self.folders}
        places = suggested_receipt_folders()
        for label, path in places:
            if _norm_folder(path) in added:
                menu.add_command(label=f"✓  {label}  (이미 있음)", state="disabled")
            else:
                shown = str(path) if len(str(path)) <= 48 else "…" + str(path)[-47:]
                menu.add_command(label=f"{label}   —   {shown}", command=lambda l=label, p=path: self.add_place(l, p))
        if places:
            menu.add_separator()
        menu.add_command(label="다른 폴더 직접 고르기…", command=self.add_folder_dialog)
        self.add_menu = menu
        x, y = self.add_button.winfo_rootx(), self.add_button.winfo_rooty() + self.add_button.winfo_height()
        try:
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()
        return menu

    def add_folder_dialog(self):
        start = self.folders[0] if self.folders and Path(self.folders[0]).is_dir() else str(Path.home())
        d = self.app.fd.askdirectory(title="영수증 사진이 있는 폴더 고르기", initialdir=start)
        if d:
            if self.add_folder(d):
                self.app.append(f"영수증 폴더 추가: {Path(d)}")
            else:
                self.app.mb.showinfo("폴더 추가", "이미 목록에 있는 폴더입니다.")

    def selected(self):
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    def remove_selected(self, ask=True):
        i = self.selected()
        if i is None:
            self.app.mb.showinfo("빼기", "뺄 폴더를 목록에서 고르세요.")
            return
        if ask and not self.app.mb.askyesno(
                "빼기", f"이 폴더를 목록에서 뺄까요? (폴더와 사진은 지우지 않습니다)\n\n{self.folders[i]}"):
            return
        del self.folders[i]
        self.save()
        self.refresh()

    def open_selected(self):
        i = self.selected()
        if i is None:
            self.app.mb.showinfo("폴더 열기", "열 폴더를 목록에서 고르세요.")
            return
        open_path(Path(self.folders[i]))

    # ---------- 이 달 사진 세기(탭을 열 때·폴더를 바꿀 때 자동) ----------
    def visible(self) -> bool:
        try:
            return self.app.notebook.nametowidget(self.app.notebook.select()) is self.app.sources_frame
        except Exception:
            return False

    def auto_count(self):
        if not self.visible():
            return
        if self.counting:
            self.recount = True
            return
        self.count_photos(quiet=True)

    def count_photos(self, wait=False, quiet=False):
        tm = resolve_target_month(self.app.month.get())
        cfg, _ = load_settings(self.app.base / SETTINGS_FILE_NAME)
        set_receipt_folders(cfg, self.folders)
        self.counting = True
        self.refresh()
        if not quiet:
            self.app.append(f"{tm.prefix} 영수증일 수 있는 사진을 세는 중… (사진이 많으면 조금 걸립니다)")

        def work():
            try:
                cache_path = self.app.base / "결과" / "photo_date_cache.json"
                cache = _load_json(cache_path)
                counts = count_photos_per_folder(cfg, tm, cache)
                try:
                    cache_path.parent.mkdir(parents=True, exist_ok=True)
                    _save_json(cache_path, cache)
                except OSError:
                    pass
                self.app.q.put(("call", lambda: self.show_counts(tm, counts, quiet)))
            except Exception as e:
                self.app.q.put(("call", lambda: self.count_failed(e)))

        t = threading.Thread(target=work, daemon=True)
        t.start()
        if wait:
            t.join()
        return t

    def count_failed(self, e):
        self.counting = False
        self.refresh()
        self.app.append(f"사진 세기 실패: {e}")

    def show_counts(self, tm, counts, quiet=False):
        self.counting = False
        if not quiet:
            start, end = photo_window(load_settings(self.app.base / SETTINGS_FILE_NAME)[0], tm)
            self.app.append(f"{tm.prefix} 영수증일 수 있는 사진({start:%m/%d} ~ {end - timedelta(days=1):%m/%d} 촬영):")
        for folder, n in counts:
            if n is not None:
                self.counts[_norm_folder(folder)] = n
            if not quiet:
                self.app.append(f"   {folder}: " + ("폴더 없음" if n is None else f"{n}장"))
        self.refresh()
        if self.recount:
            self.recount = False
            self.auto_count()

    # ---------- 사진 가져오기 ----------
    def import_dialog(self):
        files = self.app.fd.askopenfilenames(
            title="가져올 영수증 사진 고르기 (여러 장 고를 수 있음)",
            filetypes=[("사진", "*.jpg *.jpeg *.png *.heic *.heif"), ("모든 파일", "*.*")])
        if files:
            self.import_files(files)

    def import_files(self, files) -> dict:
        dest = self.imported_folder()
        result = import_receipt_files(files, dest)
        self.add_folder(dest)
        parts = [f"{len(result['copied'])}장 가져옴"]
        if result["same"]:
            parts.append(f"이미 있던 사진 {result['same']}장")
        if result["skipped"]:
            parts.append(f"사진이 아닌 파일 {len(result['skipped'])}개는 뺌")
        msg = ", ".join(parts)
        self.import_status.configure(text=msg)
        self.app.append(f"사진 가져오기: {msg} → {dest}")
        self.counts.pop(_norm_folder(dest), None)
        self.refresh()
        self.auto_count()
        return result


class BackupTab:
    """[백업·되살리기] 탭: 자동 입력 규칙·병원 목록·설정·영수증 목록을 zip 하나로 백업하고 되살리기"""

    def __init__(self, app, parent):
        tk, ttk = app.tk, app.ttk
        px = app.px
        self.app = app
        self.items = []
        parent.columnconfigure(0, weight=1)
        ttk.Label(parent, text="백업·되살리기", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        auto_wrap(ttk.Label(parent, style="CardMuted.TLabel", wraplength=px(820), justify="left", text=(
            "자동 입력 규칙·병원 목록·설정·영수증 목록을 파일 하나(zip)로 묶어 둡니다. 켤 때·끌 때·자동 입력을 저장할 때 "
            "바뀐 것이 있으면 자동으로 백업합니다(최근 20개). 새 판을 같은 폴더에 덮어쓰면 그대로 쓰이고, 새 폴더에 두면 "
            "처음 켤 때 되살릴지 물어봅니다. 되살리기 전 상태도 자동으로 백업하므로 언제든 되돌릴 수 있습니다."))
        ).grid(row=1, column=0, sticky="ew", pady=(px(2), px(10)))

        folder_row = ttk.Frame(parent, style="Pane.TFrame")
        folder_row.grid(row=2, column=0, sticky="ew", pady=(0, px(8)))
        ttk.Label(folder_row, text="백업 폴더", style="Field.TLabel").pack(side="left")
        self.folder_label = ttk.Label(folder_row, text="", style="Card.TLabel")
        self.folder_label.pack(side="left", padx=(px(10), 0))
        ttk.Button(folder_row, text="폴더 열기", style="Tool.TButton", command=self.open_folder).pack(side="right")

        box = ttk.Frame(parent, style="Pane.TFrame")
        box.grid(row=3, column=0, sticky="ew")
        box.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(box, columns=("when", "kind", "version", "files"), show="headings", height=6,
                                 selectmode="browse")
        for col, text, width, anchor in (("when", "만든 시각", 140, "w"), ("kind", "종류", 170, "w"),
                                         ("version", "프로그램 판", 220, "w"), ("files", "파일 수", 70, "center")):
            self.tree.heading(col, text=text)
            self.tree.column(col, width=px(width), minwidth=px(60), anchor=anchor)
        self.tree.grid(row=0, column=0, sticky="ew")
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=sb.set)

        bf = ttk.Frame(parent, style="Pane.TFrame")
        bf.grid(row=4, column=0, sticky="ew", pady=(px(10), 0))
        ttk.Button(bf, text="지금 백업하기", style="Accent.TButton", command=self.backup_now).pack(side="left")
        ttk.Button(bf, text="고른 백업 되살리기", command=self.restore_selected).pack(side="left", padx=(px(6), 0))
        ttk.Button(bf, text="다른 백업 파일에서 되살리기…", command=self.restore_from_file).pack(side="left", padx=(px(6), 0))
        self.status = ttk.Label(parent, text="", style="AccentText.TLabel", wraplength=px(820), justify="left")
        self.status.grid(row=5, column=0, sticky="w", pady=(px(10), 0))
        self.refresh()

    def folder(self) -> Path:
        return backup_dir(load_settings(self.app.base / SETTINGS_FILE_NAME)[0])

    def refresh(self):
        folder = self.folder()
        self.folder_label.configure(text=str(folder))
        self.items = list_backups(folder)
        self.tree.delete(*self.tree.get_children())
        for i, (path, info) in enumerate(self.items):
            self.tree.insert("", "end", iid=str(i), values=backup_label(path, info))

    def open_folder(self):
        self.folder().mkdir(parents=True, exist_ok=True)
        open_path(self.folder())

    def backup_now(self):
        path = _unique_path(self.folder(), f"직접백업_{datetime.now():%Y%m%d_%H%M%S}")
        try:
            make_backup(self.app.base, path, "직접 백업")
        except Exception as e:
            self.app.mb.showerror("백업", str(e))
            return None
        self.status.configure(text=f"백업했습니다: {path.name}")
        self.app.append(f"백업했습니다: {path}")
        self.refresh()
        return path

    def _can_restore(self) -> bool:
        if self.app.worker and self.app.worker.is_alive():
            self.app.mb.showinfo("되살리기", "작업 중에는 되살릴 수 없습니다. 작업이 끝난 뒤 다시 누르세요.")
            return False
        if self.app.editor is not None and self.app.editor.win.winfo_exists():
            self.app.mb.showinfo("되살리기", "자동 입력 편집 창을 닫고 다시 누르세요.")
            return False
        hosp = getattr(self.app, "hosp_editor", None)
        if hosp is not None and hosp.win.winfo_exists():
            self.app.mb.showinfo("되살리기", "병원 목록 편집 창을 닫고 다시 누르세요.")
            return False
        return True

    def restore(self, path: Path, ask=True) -> bool:
        if not self._can_restore():
            return False
        info = read_backup_info(path)
        when, kind, version, files = backup_label(path, info) if info else ("", "", "", "0")
        if ask and not self.app.mb.askyesno("되살리기", (
                f"이 백업으로 되살릴까요?\n\n만든 시각: {when} ({kind})\n프로그램: {version}\n파일 {files}개\n\n"
                "지금 상태는 되살리기 전에 자동으로 백업해 둡니다.")):
            return False
        try:
            names = restore_backup(path, self.app.base, self.folder())
        except Exception as e:
            self.app.mb.showerror("되살리기", str(e))
            return False
        self.app.reload_from_settings()
        msg = f"되살렸습니다: {when} 백업, 파일 {len(names)}개 (자동 입력 규칙·병원 목록·설정·영수증 목록)"
        self.status.configure(text=msg)
        self.app.append(msg)
        self.refresh()
        return True

    def restore_selected(self):
        sel = self.tree.selection()
        if not sel:
            self.app.mb.showinfo("되살리기", "되살릴 백업을 목록에서 고르세요.")
            return False
        return self.restore(self.items[int(sel[0])][0])

    def restore_from_file(self):
        picked = self.app.fd.askopenfilename(title="되살릴 백업 파일 고르기", initialdir=str(self.folder()),
                                             filetypes=[("백업 파일", "*.zip"), ("모든 파일", "*.*")])
        if picked:
            return self.restore(Path(picked))
        return False


class ReceiptReview:
    """영수증 확인 창: ①에서 읽은 영수증을 사진과 함께 보며 사용 여부·날짜·금액·분류·병원·도시를 고르거나 고칩니다.
    고친 내용은 영수증 목록 파일(결과\\영수증목록_YYYY-MM.csv)에 저장되어 다음 실행에 그대로 쓰입니다(엑셀 대신)."""
    NO_CATEGORY_TEXT = "(분류 없음)"
    DEFAULT_CITY_TEXT = "(기본 도시)"

    def __init__(self, app, path: Path, on_close=None):
        tk, ttk = app.tk, app.ttk
        px, font, c = app.px, app.font, THEME
        self.app, self.tk, self.path, self.on_close = app, tk, Path(path), on_close
        self.rows = read_csv_dicts(self.path)
        self.cfg, _ = load_settings(app.base / SETTINGS_FILE_NAME)
        self.categories = [r[C_NAME] for r in load_category_rows(app.base / CATEGORY_FILE_NAME, self.cfg)]
        hospitals = {r[1] for r in load_hospital_rows(app.base / HOSPITAL_FILE_NAME) if len(r) > 1 and r[1]}
        hospitals |= {str(r.get("병원 영문") or "") for r in self.rows} - {""}
        self.hospitals = sorted(hospitals, key=str.lower)
        cities = {city_name(v) for v in (self.cfg.get("도시_목록") or {}).values()} | {city_name(self.cfg.get("기본_도시", ""))}
        cities |= {str(r.get("도시") or "") for r in self.rows}
        self.cities = [self.DEFAULT_CITY_TEXT] + sorted((x for x in cities if x), key=str.lower)
        self.current, self.loading, self.dirty = None, False, False
        self.images, self.photo_index = {}, {}
        self.preview_size = (px(330), px(440))

        win = self.win = tk.Toplevel(app.root)
        month = re.search(r"(\d{4}-\d{2})", self.path.name)
        self.month_text = month.group(1) if month else ""
        win.title(f"영수증 확인 {self.month_text}")
        win.configure(bg=c["bg"])
        width = min(px(1240), app.root.winfo_screenwidth() - px(40))
        height = min(px(720), app.root.winfo_screenheight() - px(90))
        win.geometry(f"{width}x{height}")
        win.minsize(min(px(1000), width), min(px(560), height))
        win.transient(app.root)
        win.protocol("WM_DELETE_WINDOW", self.close)
        outer = ttk.Frame(win, padding=(px(16), px(12)))
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(2, weight=1)
        outer.rowconfigure(2, weight=1)

        top = ttk.Frame(outer)
        top.grid(row=0, column=0, columnspan=3, sticky="ew")
        ttk.Label(top, text=f"영수증 확인  {self.month_text}", font=font(14, "bold")).pack(side="left")
        self.summary = ttk.Label(top, text="", foreground=c["muted"])
        self.summary.pack(side="left", padx=(px(14), 0))
        auto_wrap(ttk.Label(outer, foreground=c["muted"], font=font(9), justify="left", wraplength=px(1100), text=(
            "왼쪽에서 영수증을 고르면 사진과 읽은 내용이 나옵니다. 틀린 곳만 고르거나 고친 뒤 [저장]하세요. "
            "고친 내용은 다음 실행부터 그대로 쓰입니다. ✓ 괜찮음 · ⚠ 확인 필요 · ✎ 고침 · ✕ 뺌"))).grid(
            row=1, column=0, columnspan=3, sticky="ew", pady=(px(2), px(10)))

        # 왼쪽: 영수증 목록
        left = ttk.Frame(outer, style="Card.TFrame", padding=px(10))
        left.grid(row=2, column=0, sticky="nsw", padx=(0, px(12)))
        left.rowconfigure(1, weight=1)
        self.only_check = tk.BooleanVar(value=False)
        ttk.Checkbutton(left, text="확인 필요만 보기", variable=self.only_check, style="Card.TCheckbutton",
                        command=self.refresh_list).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, px(6)))
        self.tree = ttk.Treeview(left, columns=("state", "date", "amount", "cat"), show="headings", selectmode="browse")
        for col, text, w, anchor in (("state", "상태", 70, "w"), ("date", "날짜", 92, "center"),
                                     ("amount", "금액", 80, "e"), ("cat", "분류", 84, "w")):
            self.tree.heading(col, text=text)
            self.tree.column(col, width=px(w), minwidth=px(50), anchor=anchor, stretch=False)
        self.tree.grid(row=1, column=0, sticky="ns")
        sb = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        sb.grid(row=1, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.on_select())

        # 가운데: 사진
        mid = ttk.Frame(outer, style="Card.TFrame", padding=px(10))
        mid.grid(row=2, column=1, sticky="ns", padx=(0, px(12)))
        photo_box = tk.Frame(mid, bg="#eef1f5", width=self.preview_size[0], height=self.preview_size[1])
        photo_box.pack()
        photo_box.pack_propagate(False)                    # 사진 크기와 상관없이 칸 크기 고정
        self.photo = tk.Label(photo_box, bg="#eef1f5", fg=c["muted"], font=font(9), text="", justify="center")
        self.photo.pack(fill="both", expand=True)
        name_row = ttk.Frame(mid, style="Pane.TFrame")
        name_row.pack(fill="x", pady=(px(8), 0))
        self.file_label = ttk.Label(name_row, text="", style="CardMuted.TLabel")
        self.file_label.pack(side="left")
        ttk.Button(name_row, text="사진 크게 보기", style="Tool.TButton", command=self.open_photo).pack(side="right")

        # 오른쪽: 고칠 칸
        right = ttk.Frame(outer, style="Card.TFrame", padding=(px(18), px(12)))
        right.grid(row=2, column=2, sticky="nsew")
        right.columnconfigure(1, weight=1)
        self.vars = {}
        r = 0

        def label(text):
            ttk.Label(right, text=text, style="Field.TLabel").grid(row=r, column=0, sticky="nw", pady=(px(9), 0),
                                                                  padx=(0, px(14)))

        def hint(text):
            nonlocal r
            auto_wrap(ttk.Label(right, text=text, style="CardMuted.TLabel", justify="left", wraplength=px(380))).grid(
                row=r, column=1, sticky="ew", pady=(px(2), 0))
            r += 1

        def combo(key, text, values, readonly=False, hint_text=""):
            nonlocal r
            label(text)
            var = self.vars[key] = tk.StringVar()
            w = ttk.Combobox(right, textvariable=var, values=values, state="readonly" if readonly else "normal")
            w.grid(row=r, column=1, sticky="ew", pady=(px(7), 0))
            w.bind("<<ComboboxSelected>>", lambda e: self.on_edit())
            w.bind("<KeyRelease>", lambda e: self.on_edit())
            r += 1
            if hint_text:
                hint(hint_text)
            return w

        label("사용")
        self.use = tk.BooleanVar(value=True)
        ttk.Checkbutton(right, text="이 영수증을 Concur에 넣기", variable=self.use, style="Card.TCheckbutton",
                        command=self.on_edit).grid(row=r, column=1, sticky="w", pady=(px(7), 0))
        r += 1
        self.date_box = combo("date", "날짜", [], hint_text="영수증에 적힌 날짜 (예: 2026-10-27). ▼에 사진에서 읽은 날짜가 있습니다.")
        self.amount_box = combo("amount", "금액", [], hint_text="▼에 사진에서 읽은 금액이 있습니다.")
        self.cat_box = combo("cat", "분류", [self.NO_CATEGORY_TEXT] + self.categories, readonly=True,
                             hint_text="자동 입력 편집에서 만든 항목. 이 항목의 Comment·Expense Type이 들어갑니다.")
        self.hosp_box = combo("hosp", "병원 (영문)", self.hospitals, hint_text="비우면 Comment에서 ' in 병원'이 빠집니다.")
        self.city_box = combo("city", "도시", self.cities, hint_text="'(기본 도시)'면 기본 도시를 고릅니다.")
        label("확인할 것")
        self.issue_label = ttk.Label(right, text="", style="Warn.TLabel", justify="left", wraplength=px(380))
        auto_wrap(self.issue_label).grid(row=r, column=1, sticky="ew", pady=(px(9), 0))
        r += 1
        label("읽은 글자")
        self.ocr_label = ttk.Label(right, text="", style="CardMuted.TLabel", justify="left", wraplength=px(380))
        auto_wrap(self.ocr_label).grid(row=r, column=1, sticky="ew", pady=(px(9), 0))
        r += 1
        self.status = ttk.Label(right, text="", style="AccentText.TLabel", justify="left", wraplength=px(380))
        auto_wrap(self.status).grid(row=r, column=1, sticky="ew", pady=(px(12), 0))

        bottom = ttk.Frame(outer)
        bottom.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(px(12), 0))
        ttk.Button(bottom, text="닫기", command=self.close).pack(side="right", padx=(px(6), 0))
        ttk.Button(bottom, text="엑셀로 열기", command=self.open_excel).pack(side="right", padx=(px(6), 0))
        ttk.Button(bottom, text="저장", style="Accent.TButton", command=self.save).pack(side="right")

        self.refresh_list()
        first = self.tree.get_children()
        if first:
            self.tree.selection_set(first[0])
            self.on_select()
        win.lift()
        win.focus_force()

    # ---------- 목록 ----------
    def row_values(self, row) -> tuple:
        amount = parse_amount_cell(row.get("금액"))
        return (review_status(row), str(row.get("날짜") or ""), f"{amount:,}" if amount else "",
                str(row.get("분류") or ""))

    def refresh_list(self):
        keep = self.current
        self.tree.delete(*self.tree.get_children())
        for i, row in enumerate(self.rows):
            if self.only_check.get() and review_status(row) not in ("⚠ 확인",):
                continue
            self.tree.insert("", "end", iid=str(i), values=self.row_values(row))
        if keep is not None and self.tree.exists(str(keep)):
            self.tree.selection_set(str(keep))
        used = [r for r in self.rows if parse_use_cell(str(r.get("사용(O/X)") or ""), True)]
        need = [r for r in self.rows if review_status(r) == "⚠ 확인"]
        self.summary.configure(text=f"영수증 {len(used)}장 사용 · 확인 필요 {len(need)}장 · 뺀 사진 {len(self.rows) - len(used)}장")

    def on_select(self):
        sel = self.tree.selection()
        if not sel:
            return
        self.current = int(sel[0])
        row = self.rows[self.current]
        g = lambda k: str(row.get(k) or "").strip()
        self.loading = True
        self.use.set(parse_use_cell(g("사용(O/X)"), True))
        dates = list(dict.fromkeys([d for d in [g("날짜"), g("[자동] 날짜")] + g("날짜 후보").split() if d]))
        self.date_box.configure(values=dates)
        self.vars["date"].set(g("날짜"))
        amounts = [int(a) for a in g("금액 후보").split() if a.isdigit()]
        for a in (parse_amount_cell(g("금액")), parse_amount_cell(g("[자동] 금액"))):
            if a and a not in amounts:
                amounts.insert(0, a)
        self.amount_box.configure(values=[f"{a:,}" for a in amounts])
        amount = parse_amount_cell(g("금액"))
        self.vars["amount"].set(f"{amount:,}" if amount else "")
        self.vars["cat"].set(g("분류") or self.NO_CATEGORY_TEXT)
        self.vars["hosp"].set(g("병원 영문"))
        self.vars["city"].set(g("도시") or self.DEFAULT_CITY_TEXT)
        self.loading = False
        self.issue_label.configure(text=g("확인 필요") or "없음")
        self.ocr_label.configure(text=g("OCR 앞부분") or "(읽은 글자 없음)")
        self.status.configure(text="")
        self.show_photo(row)

    # ---------- 사진 ----------
    def photo_path(self, row):
        return find_receipt_photo(row, self.cfg, self.photo_index)

    def show_photo(self, row):
        path = self.photo_path(row)
        self.file_label.configure(text=str(row.get("파일") or ""))
        if path is None:
            self.photo.configure(image="", text="사진 파일을 찾지 못했습니다.\n(영수증 폴더에서 옮기거나 지웠을 수 있습니다)")
            return
        key = str(path)
        if key not in self.images:
            try:
                from PIL import ImageTk
                self.images[key] = ImageTk.PhotoImage(load_preview_image(path, *self.preview_size), master=self.win)
            except ImportError:
                self.photo.configure(image="", text="사진 미리보기에는 Pillow가 필요합니다.\n[사진 크게 보기]로 여세요.")
                return
            except Exception as e:
                self.photo.configure(image="", text=f"사진을 열지 못했습니다.\n{str(e)[:60]}")
                return
        self.photo.configure(image=self.images[key], text="")

    def open_photo(self):
        if self.current is None:
            return
        path = self.photo_path(self.rows[self.current])
        if path is not None:
            open_path(path)

    # ---------- 고치기 ----------
    def on_edit(self):
        if self.loading or self.current is None:
            return
        row = self.rows[self.current]
        row["사용(O/X)"] = "O" if self.use.get() else "X"
        row["날짜"] = self.vars["date"].get().strip()
        row["금액"] = self.vars["amount"].get().strip()
        cat = self.vars["cat"].get().strip()
        row["분류"] = "" if cat == self.NO_CATEGORY_TEXT else cat
        row["병원 영문"] = self.vars["hosp"].get().strip()
        city = self.vars["city"].get().strip()
        row["도시"] = "" if city == self.DEFAULT_CITY_TEXT else city
        warn = []
        if row["날짜"] and not parse_date_cell(row["날짜"]):
            warn.append("날짜는 2026-10-27처럼 적어 주세요.")
        if row["금액"] and parse_amount_cell(row["금액"]) is None:
            warn.append("금액은 숫자로 적어 주세요.")
        self.status.configure(text=" ".join(warn), style="Warn.TLabel" if warn else "AccentText.TLabel")
        self.dirty = True
        self.win.title(f"영수증 확인 {self.month_text} *")
        if self.tree.exists(str(self.current)):
            self.tree.item(str(self.current), values=self.row_values(row))
        used = [r for r in self.rows if parse_use_cell(str(r.get("사용(O/X)") or ""), True)]
        need = [r for r in self.rows if review_status(r) == "⚠ 확인"]
        self.summary.configure(text=f"영수증 {len(used)}장 사용 · 확인 필요 {len(need)}장 · 뺀 사진 {len(self.rows) - len(used)}장")

    def save(self) -> bool:
        try:
            save_review_rows(self.path, self.rows)
        except PermissionError:
            self.app.mb.showerror("저장", "영수증 목록 파일이 엑셀에 열려 있습니다. 엑셀을 닫고 다시 저장하세요.", parent=self.win)
            return False
        self.dirty = False
        self.win.title(f"영수증 확인 {self.month_text}")
        edited = sum(1 for r in self.rows if review_row_edited(r))
        self.status.configure(text=f"저장했습니다. 고친 영수증 {edited}장 (다음 실행부터 씁니다)", style="AccentText.TLabel")
        self.app.append(f"영수증 확인: 저장했습니다(고친 영수증 {edited}장). 다음 실행부터 고친 내용을 씁니다.")
        return True

    def open_excel(self):
        if self.dirty and not self.save():
            return
        open_path(self.path)

    def close(self, force=False):
        if self.dirty and not force:
            ans = self.app.mb.askyesnocancel("영수증 확인", "고친 내용을 저장할까요?", parent=self.win)
            if ans is None or (ans and not self.save()):
                return
        self.win.destroy()
        if self.on_close:
            callback, self.on_close = self.on_close, None
            callback()


class HospitalEditor:
    """병원 목록 편집 창(엑셀 대신): 영수증에서 찾을 글자 → Concur에 쓸 영문 이름"""

    def __init__(self, app):
        tk, ttk = app.tk, app.ttk
        px, font, c = app.px, app.font, THEME
        self.app, self.tk = app, tk
        self.path = app.base / HOSPITAL_FILE_NAME
        load_hospital_rows(self.path)                  # 없으면 만들고, 예전 기본 목록은 정리
        # 영문 이름이 아직 빈 줄도 지우지 않게 파일의 줄을 그대로 읽음
        self.rows = [([str(v or "").strip() for v in r.values()] + ["", "", ""])[:3] for r in read_csv_dicts(self.path)]
        self.rows = [r for r in self.rows if r[0] or r[1]]
        self.images = app.theme["images"]
        self.current, self.visible, self.dirty, self.loading = None, [], False, False

        win = self.win = tk.Toplevel(app.root)
        win.title("병원 목록 편집")
        win.configure(bg=c["bg"])
        width = min(px(960), app.root.winfo_screenwidth() - px(40))
        height = min(px(600), app.root.winfo_screenheight() - px(90))
        win.geometry(f"{width}x{height}")
        win.minsize(min(px(820), width), min(px(480), height))
        win.transient(app.root)
        win.protocol("WM_DELETE_WINDOW", self.close)
        outer = ttk.Frame(win, padding=(px(16), px(14)))
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(1, weight=1)
        auto_wrap(ttk.Label(outer, foreground=c["muted"], font=font(9), justify="left", wraplength=px(900), text=(
            "영수증에 '찾을 글자'가 보이면 그 병원으로 보고, Comment의 {hospital} 자리에 영문 이름을 넣습니다. "
            "주황 점은 자동으로 만든 영문 이름이라 공식 이름과 같은지 확인이 필요한 병원입니다. "
            "영수증에서 새 병원이 보이면 자동으로 여기에 더해집니다."))).grid(
            row=0, column=0, columnspan=2, sticky="ew", pady=(0, px(10)))

        left = ttk.Frame(outer, style="Card.TFrame", padding=px(12))
        left.grid(row=1, column=0, sticky="nsw", padx=(0, px(12)))
        left.rowconfigure(1, weight=1)
        sf = ttk.Frame(left, style="Pane.TFrame")
        sf.grid(row=0, column=0, sticky="ew", pady=(0, px(8)))
        ttk.Label(sf, text="찾기", style="Field.TLabel").pack(side="left")
        self.filter = tk.StringVar()
        fe = ttk.Entry(sf, textvariable=self.filter, width=22)
        fe.pack(side="left", padx=(px(8), 0), fill="x", expand=True)
        fe.bind("<KeyRelease>", lambda e: self.refresh_list())
        box = tk.Frame(left, bg=c["card"], highlightthickness=1, highlightbackground=c["border"], highlightcolor=c["border"])
        box.grid(row=1, column=0, sticky="nsew")
        self.tree = item_tree(box, ttk, app, ("찾을 글자", px(130)), ("영문 이름", px(220)))
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.on_select())
        bf = ttk.Frame(left, style="Pane.TFrame")
        bf.grid(row=2, column=0, sticky="ew", pady=(px(8), 0))
        ttk.Button(bf, text="새 병원", style="Tool.TButton", command=self.add).pack(side="left", padx=(0, px(4)))
        ttk.Button(bf, text="삭제", style="Tool.TButton", command=self.remove).pack(side="left")

        right = ttk.Frame(outer, style="Card.TFrame", padding=(px(18), px(14)))
        right.grid(row=1, column=1, sticky="nsew")
        right.columnconfigure(1, weight=1)
        self.vars = {}
        r = 0
        for key, label, hint in (
                ("key", "영수증에서 찾을 글자", "영수증에 이 글자가 있으면 이 병원으로 봅니다 (예: 원광대). 너무 짧으면 다른 병원에도 맞을 수 있습니다."),
                ("en", "Concur에 쓸 영문 이름", "Comment의 {hospital} 자리에 들어갑니다 (예: Wonkwang University Hospital).")):
            ttk.Label(right, text=label, style="Field.TLabel").grid(row=r, column=0, sticky="w", pady=(px(7), 0), padx=(0, px(14)))
            var = self.vars[key] = tk.StringVar()
            w = ttk.Entry(right, textvariable=var)
            w.grid(row=r, column=1, sticky="ew", pady=(px(7), 0))
            w.bind("<KeyRelease>", lambda e: self.on_edit())
            if key == "key":
                self.key_entry = w
            r += 1
            auto_wrap(ttk.Label(right, text=hint, style="CardMuted.TLabel", justify="left", wraplength=px(420))).grid(
                row=r, column=1, sticky="ew", pady=(px(2), 0))
            r += 1
        ttk.Label(right, text="메모", style="Field.TLabel").grid(row=r, column=0, sticky="nw", pady=(px(12), 0), padx=(0, px(14)))
        self.note_label = ttk.Label(right, text="", style="Card.TLabel", justify="left", wraplength=px(420))
        auto_wrap(self.note_label).grid(row=r, column=1, sticky="ew", pady=(px(12), 0))
        r += 1
        self.check_label = ttk.Label(right, text="", style="Warn.TLabel", justify="left", wraplength=px(420))
        auto_wrap(self.check_label).grid(row=r, column=1, sticky="ew", pady=(px(12), 0))
        r += 1
        self.check_button = ttk.Button(right, text="영문 이름 확인함", style="Tool.TButton", command=self.confirm_name)
        self.check_button.grid(row=r, column=1, sticky="w", pady=(px(6), 0))
        self.check_button.grid_remove()
        r += 1

        bottom = ttk.Frame(outer)
        bottom.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(px(12), 0))
        ttk.Button(bottom, text="닫기", command=self.close).pack(side="right", padx=(px(6), 0))
        ttk.Button(bottom, text="엑셀로 열기", command=self.open_excel).pack(side="right", padx=(px(6), 0))
        ttk.Button(bottom, text="저장", style="Accent.TButton", command=self.save).pack(side="right")

        self.refresh_list()
        if self.visible:
            self.select_visible(0)

    # ---------- 목록 ----------
    def item_view(self, row) -> dict:
        state = "off" if not row[1].strip() else ("warn" if HOSP_CHECK_NOTE in row[2] else "on")
        dot = {"on": "dot_on", "warn": "dot_warn", "off": "dot_off"}[state]
        return {"text": row[0], "image": self.images[dot], "values": (row[1] or "—",), "tags": (state,)}

    def refresh_list(self):
        q = self.filter.get().strip().lower()
        self.visible = [i for i, row in enumerate(self.rows) if not q or q in " ".join(row).lower()]
        self.tree.delete(*self.tree.get_children())
        for i in self.visible:
            self.tree.insert("", "end", iid=str(i), **self.item_view(self.rows[i]))
        if self.current in self.visible:
            self.tree.selection_set(str(self.current))
            self.tree.see(str(self.current))

    def select_visible(self, pos: int):
        iid = str(self.visible[pos])
        self.tree.selection_set(iid)
        self.tree.see(iid)
        self.on_select()

    def on_select(self):
        sel = self.tree.selection()
        if not sel:
            return
        self.current = int(sel[0])
        row = self.rows[self.current]
        self.loading = True
        self.vars["key"].set(row[0])
        self.vars["en"].set(row[1])
        self.loading = False
        self.show_note()

    def show_note(self):
        row = self.rows[self.current] if self.current is not None else ["", "", ""]
        self.note_label.configure(text=row[2] or "—")
        if HOSP_CHECK_NOTE in row[2]:
            self.check_label.configure(text="자동으로 만든 영문 이름입니다. 공식 이름과 같은지 확인하고, "
                                            "맞으면 [영문 이름 확인함]을 누르세요(고치면 저절로 확인됨).")
            self.check_button.grid()
        else:
            self.check_label.configure(text="")
            self.check_button.grid_remove()

    # ---------- 고치기 ----------
    def on_edit(self):
        if self.loading or self.current is None:
            return
        row = self.rows[self.current]
        new_en = self.vars["en"].get().strip()
        if new_en != row[1] and HOSP_CHECK_NOTE in row[2]:
            row[2] = HOSP_USER_NOTE                     # 영문 이름을 직접 고쳤으면 확인한 것으로(엑셀 가져오기가 덮지 않음)
        row[0], row[1] = self.vars["key"].get().strip(), new_en
        self.mark_dirty()
        if self.tree.exists(str(self.current)):
            self.tree.item(str(self.current), **self.item_view(row))
        self.show_note()

    def confirm_name(self):
        if self.current is None:
            return
        self.rows[self.current][2] = "영문 이름 확인함"
        self.mark_dirty()
        self.tree.item(str(self.current), **self.item_view(self.rows[self.current]))
        self.show_note()

    def mark_dirty(self):
        self.dirty = True
        self.win.title("병원 목록 편집 *")

    def add(self):
        at = self.current + 1 if self.current is not None else len(self.rows)
        self.rows.insert(at, ["", "", HOSP_USER_NOTE])
        self.current = at
        self.mark_dirty()
        self.filter.set("")
        self.refresh_list()
        self.select_visible(self.visible.index(at))
        self.key_entry.focus_set()
        return at

    def remove(self):
        if self.current is None:
            return
        row = self.rows[self.current]
        if row[0] and not self.app.mb.askyesno("삭제", f"'{row[0]}' 병원을 지울까요?", parent=self.win):
            return
        at = self.current
        self.rows.pop(at)
        self.current = None
        self.mark_dirty()
        self.refresh_list()
        if self.visible:
            near = min(range(len(self.visible)), key=lambda k: (abs(self.visible[k] - at), self.visible[k] >= at))
            self.select_visible(near)
        else:
            for v in self.vars.values():
                v.set("")
            self.show_note()

    def save(self) -> bool:
        rows = [r for r in self.rows if r[0].strip()]
        problems = hospital_problems(rows)
        if problems and not self.app.mb.askyesno(
                "저장", "확인할 것이 있습니다.\n\n" + "\n".join(f"· {p}" for p in problems[:10]) + "\n\n그래도 저장할까요?",
                parent=self.win):
            return False
        try:
            write_csv(self.path, HOSPITAL_HEADER, rows)
        except PermissionError:
            self.app.mb.showerror("저장", "병원목록.csv가 엑셀에 열려 있습니다. 엑셀을 닫고 다시 저장하세요.", parent=self.win)
            return False
        self.dirty = False
        self.win.title("병원 목록 편집")
        self.app.append(f"병원 목록을 저장했습니다: {len(rows)}곳 (다음 실행부터 씁니다)")
        return True

    def open_excel(self):
        if self.dirty and not self.save():
            return
        open_path(self.path)

    def close(self):
        if self.dirty:
            ans = self.app.mb.askyesnocancel("병원 목록 편집", "바꾼 내용을 저장할까요?", parent=self.win)
            if ans is None or (ans and not self.save()):
                return
        self.win.destroy()


class CategoryEditor:
    """자동 입력 편집 창: 왼쪽에서 항목을 고르고, 오른쪽에서 이름(한글 마음대로)·찾을 글자와
    Concur 경비 화면의 칸(Expense Type·Business Purpose·City·LCRT·Receipt Status·Personal Expense·Comment)을
    Concur 화면 위에서 아래 순서대로 정합니다."""
    # (키, 칸 위치)
    FIELDS = (("name", C_NAME), ("keywords", C_KEYWORDS), ("etype", C_TYPE), ("purpose", C_PURPOSE), ("city", C_CITY),
              ("lcrt", C_LCRT), ("receipt", C_RECEIPT), ("comment", C_COMMENT))

    def __init__(self, app):
        tk, ttk = app.tk, app.ttk
        self.app, self.tk = app, tk
        self.path = app.base / CATEGORY_FILE_NAME
        upgrade_category_table(self.path, app.base / SETTINGS_FILE_NAME, app.append)
        cfg, _ = load_settings(app.base / SETTINGS_FILE_NAME)
        self.rows = load_category_rows(self.path, cfg)
        self.type_choices = sorted(set(cfg.get("Concur_경비_유형") or KNOWN_EXPENSE_TYPES)
                                   | {r[C_TYPE] for r in self.rows if r[C_TYPE]}, key=str.lower)
        self.default_city = cfg.get("기본_도시", "")
        cities = list(dict.fromkeys([self.default_city] + list((cfg.get("도시_목록") or {}).values())
                                    + [r[C_CITY] for r in self.rows if r[C_CITY] and r[C_CITY] != CITY_FROM_RECEIPT]))
        self.city_choices = [CITY_FROM_RECEIPT] + sorted((c for c in cities if c), key=str.lower)
        self.current, self.visible, self.dirty, self.loading = None, [], False, False

        px, font, c = app.px, app.font, THEME
        win = self.win = tk.Toplevel(app.root)
        win.title("자동 입력 편집")
        win.configure(bg=c["bg"])
        width = min(px(1080), app.root.winfo_screenwidth() - px(40))
        height = min(px(680), app.root.winfo_screenheight() - px(90))
        win.geometry(f"{width}x{height}")
        win.minsize(min(px(900), width), min(px(560), height))
        win.transient(app.root)
        win.protocol("WM_DELETE_WINDOW", self.close)
        outer = ttk.Frame(win, padding=(px(16), px(14)))
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(1, weight=1)

        ttk.Label(outer, foreground=c["muted"], font=font(9), wraplength=px(1020), justify="left", text=(
            "왼쪽에서 항목을 고르고 오른쪽에 적으면 됩니다. 항목 이름은 한글로 마음대로 적고, 그 아래 칸들은 Concur 경비 화면에 "
            "위에서 아래 순서대로 들어갑니다. 파란 점 = 설정됨(영수증에 맞춰짐), 회색 점 = 아직 안 함. "
            "'영수증에서 찾을 글자'가 비어 있는 항목은 영수증에 맞춰지지 않으니, 실제로 쓰는 항목만 채우세요. "
            "위에 있는 항목이 먼저 맞춰집니다(▲▼로 순서 바꿈).")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, px(10)))

        # 왼쪽: 목록 카드
        left = ttk.Frame(outer, style="Card.TFrame", padding=px(12))
        left.grid(row=1, column=0, sticky="nsw", padx=(0, px(12)))
        left.rowconfigure(2, weight=1)
        sf = ttk.Frame(left, style="Pane.TFrame")
        sf.grid(row=0, column=0, columnspan=2, sticky="ew")
        ttk.Label(sf, text="찾기", style="Field.TLabel").pack(side="left")
        self.filter = tk.StringVar()
        fe = ttk.Entry(sf, textvariable=self.filter, width=22)
        fe.pack(side="left", padx=(px(8), 0), fill="x", expand=True)
        fe.bind("<KeyRelease>", lambda e: self.refresh_list())
        self.only_on = tk.BooleanVar(value=False)
        ttk.Checkbutton(left, text="설정한 항목만 보기", variable=self.only_on, style="Card.TCheckbutton",
                        command=self.refresh_list).grid(row=1, column=0, columnspan=2, sticky="w", pady=px(6))
        lb_frame = tk.Frame(left, bg=c["card"], highlightthickness=1, highlightbackground=c["border"],
                            highlightcolor=c["border"])
        lb_frame.grid(row=2, column=0, columnspan=2, sticky="nsew")
        self.images = app.theme["images"]
        self.tree = item_tree(lb_frame, ttk, app, ("항목", px(140)), ("Expense Type", px(170)))
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.on_select())
        bf = ttk.Frame(left, style="Pane.TFrame")
        bf.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(px(8), 0))
        for text, cmd in (("새 항목", self.add), ("삭제", self.remove), ("▲ 위로", lambda: self.move(-1)),
                          ("▼ 아래로", lambda: self.move(1))):
            ttk.Button(bf, text=text, style="Tool.TButton", command=cmd).pack(side="left", padx=(0, px(4)))

        # 오른쪽: 입력 칸 카드 (항목 이름·찾을 글자, 그다음 Concur 경비 화면 위에서 아래 순서)
        right = ttk.Frame(outer, style="Card.TFrame", padding=(px(18), px(12)))
        right.grid(row=1, column=1, sticky="nsew")
        right.columnconfigure(1, weight=1)
        self.vars = {}
        hint_style = {"style": "CardMuted.TLabel", "wraplength": px(430), "justify": "left"}
        r = 0

        def label(text, row, top=7):
            ttk.Label(right, text=text, style="Field.TLabel").grid(row=row, column=0, sticky="w", pady=(px(top), 0),
                                                                  padx=(0, px(14)))

        def hint(text, row):
            ttk.Label(right, text=text, **hint_style).grid(row=row, column=1, sticky="w", pady=(px(2), 0))

        def entry(key, text, hint_text="", combo=None, readonly=False):
            nonlocal r
            label(text, r)
            var = self.vars[key] = tk.StringVar()
            if combo is None:
                w = ttk.Entry(right, textvariable=var)
            else:
                w = ttk.Combobox(right, textvariable=var, values=combo, state="readonly" if readonly else "normal")
                w.bind("<<ComboboxSelected>>", lambda e: self.on_edit())
            w.grid(row=r, column=1, sticky="ew", pady=(px(7), 0))
            w.bind("<KeyRelease>", lambda e: self.on_edit())
            r += 1
            if hint_text:
                hint(hint_text, r)
                r += 1
            return w

        self.name_entry = entry("name", "항목 이름", "한글로 알아보기 쉬운 이름 (예: 톨게이트, 택시, 주유)")
        entry("keywords", "영수증에서 찾을 글자", "영수증에 이 글자가 있으면 이 항목 (쉼표로 여러 개)")
        hline(right, tk).grid(row=r, column=0, columnspan=2, sticky="ew", pady=(px(14), px(8)))
        r += 1
        ttk.Label(right, text="Concur 경비 화면에 넣을 값 (화면 위에서 아래 순서대로 입력)", style="AccentText.TLabel").grid(
            row=r, column=0, columnspan=2, sticky="w")
        r += 1
        self.etype_box = entry("etype", "Expense Type", "Concur에 Undefined일 때 고름 (비우면 그대로 두고 '확인 필요')",
                               combo=self.type_choices)
        entry("purpose", "Business Purpose")
        self.city_box = entry("city", "City", "보통은 영수증 주소에서 찾기 (못 찾으면 기본 도시)", combo=self.city_choices)
        entry("lcrt", "LCRT", combo=LCRT_CHOICES)
        entry("receipt", "Receipt Status", combo=RECEIPT_CHOICES)
        label("Personal Expense", r)
        self.personal = tk.BooleanVar(value=False)
        ttk.Checkbutton(right, text="체크하기 (개인 경비로 처리할 항목만)", style="Card.TCheckbutton",
                        variable=self.personal, command=self.on_edit).grid(row=r, column=1, sticky="w", pady=(px(7), 0))
        r += 1
        entry("comment", "Comment")

        self.status = ttk.Label(right, text="", style="Warn.TLabel", wraplength=px(430), justify="left")
        self.status.grid(row=r, column=1, sticky="w", pady=(px(12), 0))

        # 아래: 저장
        bottom = ttk.Frame(outer)
        bottom.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(px(12), 0))
        ttk.Button(bottom, text="닫기", command=self.close).pack(side="right", padx=(px(6), 0))
        ttk.Button(bottom, text="엑셀로 열기", command=self.open_excel).pack(side="right", padx=(px(6), 0))
        ttk.Button(bottom, text="저장", style="Accent.TButton", command=self.save).pack(side="right")

        self.refresh_list()
        first = next((i for i, row in enumerate(self.visible) if self.rows[row][1].strip()), 0 if self.visible else None)
        if first is not None:
            self.select_visible(first)

    # ---------- 목록 ----------
    def item_view(self, row: list) -> dict:
        """목록 한 줄: 점(설정됨 파랑/안 함 회색) + 항목 이름, Expense Type"""
        on = bool(row[C_KEYWORDS].strip())
        return {"text": row[C_NAME], "image": self.images["dot_on" if on else "dot_off"],
                "values": (row[C_TYPE] or "—",), "tags": ("on" if on else "off",)}

    def refresh_list(self):
        q = self.filter.get().strip().lower()
        self.visible = [i for i, row in enumerate(self.rows)
                        if (not q or q in " ".join(row[:5]).lower()) and (not self.only_on.get() or row[1].strip())]
        self.tree.delete(*self.tree.get_children())
        for i in self.visible:
            self.tree.insert("", "end", iid=str(i), **self.item_view(self.rows[i]))
        if self.current in self.visible:
            self.tree.selection_set(str(self.current))
            self.tree.see(str(self.current))

    def select_visible(self, pos: int):
        iid = str(self.visible[pos])
        self.tree.selection_set(iid)
        self.tree.see(iid)
        self.on_select()

    def on_select(self):
        sel = self.tree.selection()
        if not sel:
            return
        self.current = int(sel[0])
        row = self.rows[self.current]
        self.loading = True
        for key, col in self.FIELDS:
            self.vars[key].set(row[col])
        if not row[C_CITY]:
            self.vars["city"].set(CITY_FROM_RECEIPT)
        if not row[C_LCRT]:
            self.vars["lcrt"].set(LCRT_CHOICES[0])
        if not row[C_RECEIPT]:
            self.vars["receipt"].set(RECEIPT_CHOICES[0])
        self.personal.set(item_extras(row)["Personal Expense"])
        self.loading = False
        self.update_status()

    # ---------- 입력 ----------
    def on_edit(self):
        if self.loading or self.current is None:
            return
        row = self.rows[self.current]
        for key, col in self.FIELDS:
            row[col] = self.vars[key].get().strip()
        if row[C_CITY] == CITY_FROM_RECEIPT:
            row[C_CITY] = ""
        row[C_PERSONAL] = "O" if self.personal.get() else ""
        self.mark_dirty()
        if self.tree.exists(str(self.current)):
            self.tree.item(str(self.current), **self.item_view(row))
        self.update_status()

    def update_status(self):
        if self.current is None:
            self.status.configure(text="")
            return
        row = self.rows[self.current]
        self.status.configure(text="" if row[1].strip() else
                              "아직 '영수증에서 찾을 글자'가 비어 있어 이 항목은 영수증에 맞춰지지 않습니다.")

    def mark_dirty(self):
        self.dirty = True
        self.win.title("자동 입력 편집 *")

    def add(self):
        """[새 항목]: '새 항목' 줄을 더하고 이름 칸을 골라 둠 → 바로 한글 이름을 적으면 됨"""
        names = {r[0] for r in self.rows}
        name, n = "새 항목", 2
        while name in names:
            name, n = f"새 항목 {n}", n + 1
        self.add_named(name)
        self.name_entry.focus_set()
        self.name_entry.select_range(0, "end")
        return name

    def add_named(self, name: str):
        """이름으로 항목을 더하고 고릅니다(이미 있으면 그 항목을 고름)."""
        existing = next((i for i, r in enumerate(self.rows) if r[0] == name), None)
        if existing is None:
            at = self.current + 1 if self.current is not None else 0
            self.rows.insert(at, _category_row([name]))
            existing = at
            self.mark_dirty()
        self.current = existing
        self.filter.set("")
        self.only_on.set(False)
        self.refresh_list()
        self.select_visible(self.visible.index(existing))

    def remove(self):
        if self.current is None:
            return
        name = self.rows[self.current][0]
        if not self.app.mb.askyesno("삭제", f"'{name}' 항목을 지울까요?", parent=self.win):
            return
        at = self.current
        self.rows.pop(at)
        self.current = None
        self.mark_dirty()
        self.refresh_list()
        if self.visible:                       # 지운 자리 옆 항목을 고른 채로
            near = min(range(len(self.visible)), key=lambda k: (abs(self.visible[k] - at), self.visible[k] >= at))
            self.select_visible(near)
        else:
            for v in self.vars.values():
                v.set("")
            self.status.configure(text="")

    def move(self, delta: int):
        i = self.current
        if i is None or not 0 <= i + delta < len(self.rows):
            return
        self.rows[i], self.rows[i + delta] = self.rows[i + delta], self.rows[i]
        self.current = i + delta
        self.refresh_list()
        self.mark_dirty()

    # ---------- 저장 ----------
    def save(self) -> bool:
        problems = category_problems(self.rows)
        if problems and not self.app.mb.askyesno(
                "저장", "확인할 것이 있습니다.\n\n" + "\n".join(f"· {p}" for p in problems[:10]) + "\n\n그래도 저장할까요?",
                parent=self.win):
            return False
        try:
            save_category_rows(self.path, self.rows)
        except PermissionError:
            self.app.mb.showerror("저장", "자동입력표.csv가 엑셀에 열려 있습니다. 엑셀을 닫고 다시 저장하세요.", parent=self.win)
            return False
        self.dirty = False
        self.win.title("자동 입력 편집")
        on = sum(1 for r in self.rows if r[1].strip())
        self.app.append(f"자동 입력을 저장했습니다: 설정된 항목 {on}개 (다음 실행부터 씁니다)")
        self.status.configure(text=f"저장했습니다. 설정된 항목 {on}개.")
        self.app.auto_backup("자동 입력 저장 때")
        return True

    def open_excel(self):
        if self.dirty and not self.save():
            return
        open_path(self.path)

    def close(self):
        if self.dirty:
            ans = self.app.mb.askyesnocancel("자동 입력 편집", "바꾼 내용을 저장할까요?", parent=self.win)
            if ans is None or (ans and not self.save()):
                return
        self.win.destroy()


class GuiBridge:
    """작업 스레드 → 창으로 메시지/질문 전달"""
    def __init__(self, q):
        self.q = q

    def _ask(self, how, title, msg):
        holder, ev = [None], threading.Event()
        self.q.put(("ask", how, title, msg, holder, ev))
        ev.wait()
        return holder[0]

    def log(self, line):
        self.q.put(("log", line))

    def confirm(self, title, msg):
        return bool(self._ask("okcancel", title, msg))

    def yes_no_cancel(self, title, msg):
        return self._ask("yesnocancel", title, msg)

    def ask_text(self, title, msg):
        return self._ask("text", title, msg)

    def info(self, title, msg):
        self._ask("info", title, msg)

    def open(self, path):
        self.q.put(("open", path))

    def review(self, path, wait=False):
        """영수증 확인 창을 엽니다(wait=True면 창을 닫을 때까지 기다림)."""
        ev = threading.Event() if wait else None
        self.q.put(("review", path, ev))
        if ev is not None:
            ev.wait()
        return True


# ============================================================
# 메인
# ============================================================

def main():
    parser = argparse.ArgumentParser(description="Concur 경비 자동 입력")
    parser.add_argument("--cli", action="store_true", help="창 대신 명령 프롬프트로 실행")
    parser.add_argument("--month", help="처리할 달 (예: 2026-10, 이번 달도 됨). 없으면 지난달")
    parser.add_argument("--ocr-only", action="store_true", help="영수증 분석만")
    parser.add_argument("--edit-only", action="store_true", help="이미 만든 리포트에 입력만")
    parser.add_argument("--full", action="store_true", help="전체 실행")
    parser.add_argument("--inspect", action="store_true", help="Concur 화면 조사(입력칸 위치를 파일로 남김)")
    args = parser.parse_args()
    base = app_dir()

    cli = args.cli or args.ocr_only or args.edit_only or args.full or args.inspect
    if not cli:
        try:
            app = GuiApp(base)
        except ImportError:
            print("tkinter가 없어 명령 프롬프트 모드로 실행합니다.")
        except Exception as e:
            try:
                import tkinter
                from tkinter import messagebox
                tkinter.Tk().withdraw()
                messagebox.showerror("Concur 경비 자동 입력", str(e))
            except Exception:
                print(e)
            return
        else:
            app.run()
            return

    mode = ("ocr" if args.ocr_only else "edit" if args.edit_only else "full" if args.full
            else "inspect" if args.inspect else "")
    if not mode:
        print("1) 영수증 분석만   2) 전체 실행   3) 만든 리포트에 입력만   4) 화면 조사")
        mode = {"1": "ocr", "2": "full", "3": "edit", "4": "inspect"}.get(input("번호: ").strip(), "")
        if not mode:
            return
        if not args.month and mode != "inspect":
            last, this = resolve_target_month(None).prefix, recent_months(1, include_current=True)[0]
            args.month = input(f"처리할 달 (Enter={last} 지난달, 이번 달은 {this}): ").strip() or None
    run_job(mode, args.month, ConsoleUI(), base)


if __name__ == "__main__":
    main()
