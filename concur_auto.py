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
import html as html_lib
import json
import os
import queue
import re
import subprocess
import sys
import threading
import time
import traceback
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

APP_VERSION = "2026-11-13 화면 조사"     # 고칠 때마다 바꿉니다. 실행 기록 첫 줄과 창 제목에 나옵니다.
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
# Concur가 바뀌면 설정 파일의 '화면_요소'만 고치면 됩니다([화면 조사] 파일 참고). 쉼표로 여러 개를 적으면 그중 하나를 찾습니다.
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
    "리포트 시작일": "Start Date|시작일|시작 날짜",
    "리포트 종료일": "End Date|종료일|종료 날짜",
    "리포트 Business Purpose": "Business Purpose|비즈니스 목적|업무 목적",
    "리포트 Comment": "Comment|코멘트|의견",
    "경비 날짜": "Transaction Date|거래 날짜|거래일",
    "경비 금액": "Amount|금액",
    "Comment": "Comment|코멘트|의견",
    "Business Purpose": "Business Purpose|비즈니스 목적|업무 목적",
}

OLD_DEFAULT_CITY = "Daejeon (ex Taejon), KOREA, REPUBLIC OF"
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
        "영수증_보정_사용": True,
        "보정_확대_배율": 2.0,
        "입력_전_확인": True,
        "결과_자동_열기": True,
        "_업로드_설명": "원본 또는 보정본. Concur는 1장당 5MB 제한이 있어 넘으면 자동으로 줄여서 올립니다.",
        "업로드_이미지": "원본",
        "업로드_최대_MB": 5,
        "기본_도시": "Gwangju (ex Kwangju), KOREA, REPUBLIC OF",
        "_도시_설명": "영수증 주소(예: 익산시, 광주광역시)나 병원 이름(예: 목포한국병원)에서 도시를 찾아 Concur 도시를 고릅니다. 못 찾으면 기본_도시. 왼쪽은 영수증에서 찾을 한글, 오른쪽은 Concur 도시 목록에 보이는 이름 그대로.",
        "도시_목록": dict(DEFAULT_CITIES),
        "_도시_낮은_우선순위_설명": "주차 회사·카드사 본사 주소처럼 영수증에 자주 찍히지만 실제 장소가 아닌 도시. 다른 도시가 없을 때만 씁니다.",
        "도시_낮은_우선순위": ["서울"],
        "_영수증_병원_설명": "병원 목록에 없어도 영수증에 'OO병원'이 보이면 그 이름을 쓰고 병원 목록에 더합니다(영문 이름은 자동 변환).",
        "영수증에서_병원_찾기": True,
        "리포트_이름_형식": "{month_name} Expenses",
        "_키워드_설명": "아래 키워드·분류_우선순위·입력_문구·Undefined일_때_선택할_유형은 자동입력표.csv를 처음 만들 때만 씁니다. 그 뒤로는 자동입력표.csv(창의 [자동 입력표 편집])에서 고치세요.",
        "키워드": {
            "주차": ["주차", "입차", "출차", "주자"],
            "카페": ["메리카", "리카노", "테이크", "카페"],
            "식대": ["버터"],
        },
        "분류_우선순위": ["식대", "카페", "주차"],
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
        },
        "_Undefined_설명": "자동입력표.csv를 처음 만들 때만 씁니다. Expense Type이 Undefined인 경비에 고를 유형. 결과 파일에 '확인 필요'로 표시됩니다.",
        "Undefined일_때_선택할_유형": {
            "주차": "Parking & Toll",
            "카페": "Other Staff Expenses",
            "식대": "Other Staff Expenses",
            "분류없음": "Other Staff Expenses",
        },
        "_화면_문구_설명": "Concur 화면에 보이는 글자. '|'로 나눠 여러 개를 적으면 그중 하나가 보이면 됩니다(예: 다음|Next). 화면이 영어로 바뀌어도 되도록 영어도 같이 적어 두었습니다.",
        "화면_문구": {
            "리포트 생성": "리포트 생성|Create Report",
            "다음": "다음|Next|Continue",
            "이메일 링크로 로그인": "이메일 링크로 로그인|Sign in with email link|Email me a sign in link|Email sign in link|Email link",
            "사용 가능한 경비": "사용 가능한 경비|Available Expenses",
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
        },
        "_화면_요소_설명": "Concur 화면의 입력칸 위치(CSS 선택자). Concur가 바뀌어 칸을 못 찾으면 창의 [화면 조사]로 만든 파일의 id·name을 보고 여기를 고칩니다. 예: 입력칸 id가 txtComment면 \"#txtComment\". 쉼표로 여러 개를 적으면 그중 하나를 찾습니다.",
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
                   "Expense Type이 Undefined일 때", "설명"]
NO_CATEGORY = "(분류 없음)"


def category_rows_from_cfg(cfg: dict) -> list:
    """설정 파일의 키워드·입력 문구로 처음 자동입력표를 만듭니다. 위에 있는 분류가 먼저 맞춰집니다."""
    priority = list(cfg.get("분류_우선순위") or [])
    order = priority + [c for c in (cfg.get("입력_문구") or {}) if c not in priority]
    types = cfg.get("Undefined일_때_선택할_유형") or {}
    rows = []
    for cat in order:
        t = (cfg.get("입력_문구") or {}).get(cat, {})
        rows.append([cat, ", ".join((cfg.get("키워드") or {}).get(cat, [])), t.get("Comment", ""),
                     t.get("Business Purpose", ""), types.get(cat, ""), ""])
    if rows:
        rows[0][5] = "위에 있는 분류가 먼저 맞춰집니다. {hospital}(또는 {병원}) 자리에 병원 영문명이 들어가고, 병원을 못 찾으면 ' in {hospital}'은 빠집니다."
    rows.append([NO_CATEGORY, "", "", "", types.get("분류없음", "Other Staff Expenses"),
                 "어느 분류에도 맞지 않는 영수증과 영수증 없는 경비의 Expense Type"])
    return rows


def apply_category_table(cfg: dict, path: Path, log=None) -> list:
    """자동입력표.csv(없으면 설정으로 만듦)를 읽어 분류·찾을 글자·입력 문구·Expense Type을 설정에 반영합니다. → 분류 목록"""
    if not path.exists():
        write_csv(path, CATEGORY_HEADER, category_rows_from_cfg(cfg))
    order, keywords, texts = [], {}, {}
    types = {"분류없음": (cfg.get("Undefined일_때_선택할_유형") or {}).get("분류없음", "Other Staff Expenses")}
    for row in read_csv_dicts(path):
        vals = [str(v or "").strip() for v in row.values()] + [""] * 6
        cat, kw, comment, purpose, etype = vals[:5]
        if not cat:
            continue
        if compact(cat) in ("(분류없음)", "분류없음"):
            if etype:
                types["분류없음"] = etype
            continue
        if cat in order:
            continue
        order.append(cat)
        keywords[cat] = [compact(k) for k in re.split(r"[,，/]", kw) if compact(k)]
        texts[cat] = {name: value.replace("{병원}", "{hospital}")
                      for name, value in (("Comment", comment), ("Business Purpose", purpose)) if value}
        if etype:
            types[cat] = etype
    if not order:
        if log:
            log(f"자동 입력표에 분류가 없어 설정 파일의 값을 씁니다: {path.name}")
        return list(cfg.get("분류_우선순위") or CATEGORIES)
    cfg["분류_우선순위"], cfg["키워드"], cfg["입력_문구"], cfg["Undefined일_때_선택할_유형"] = order, keywords, texts, types
    if log:
        log(f"자동 입력표: 분류 {len(order)}개 ({', '.join(order)})")
    return order


# ============================================================
# 처리할 달
# ============================================================

MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July",
               "August", "September", "October", "November", "December"]


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
    def first_day_us(self) -> str:
        return f"{self.month:02d}/01/{self.year}"

    @property
    def last_day_us(self) -> str:
        return f"{self.month:02d}/{calendar.monthrange(self.year, self.month)[1]:02d}/{self.year}"

    def report_name(self, fmt: str) -> str:
        return fmt.format(month_name=self.month_name, year=self.year, month=self.month)


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
EXIF_DATETIME = re.compile(r"\s*(\d{4}):(\d{2}):(\d{2})[ T](\d{2}):(\d{2}):(\d{2})")


def file_name_datetime(name: str):
    """단축어로 만든 파일 이름(2026-10-22_123342.jpg)의 촬영 시각"""
    m = NAME_DATETIME.search(name)
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


def list_originals(cfg: dict, tm: TargetMonth, cache=None) -> list:
    """처리할 달 영수증일 수 있는 사진 [(경로, 촬영 시각), ...] — 촬영 순서대로.
    영수증을 나중에 찍을 수도 있으므로 처리할 달 1일부터, 그 달이 끝나고 허용 일수 안에 찍은 사진까지."""
    source = Path(cfg["영수증_원본_폴더"])
    if not source.is_dir():
        try:
            similar = [d.name for d in source.parent.iterdir()
                       if d.is_dir() and d.name.lower().startswith(source.name.lower()[:6])]
        except OSError:
            similar = []
        hint = (f"\n같은 위치에 비슷한 폴더가 있습니다: {', '.join(similar[:5])}\n"
                f"프로그램 창의 [찾기]로 영수증 폴더를 다시 고르세요.") if similar else ""
        raise RuntimeError(f"영수증 폴더가 없습니다: {source}{hint}")
    cache = {} if cache is None else cache
    start, end = photo_window(cfg, tm)
    too_old = start - timedelta(days=2)
    found = []
    for p in source.iterdir():
        if not (p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS) or p.name.startswith("."):
            continue
        if p.stem.endswith(PROCESSED_SUFFIX):
            continue
        st = p.stat()
        # 파일 수정 시각은 촬영 시각과 같거나 늦으므로, 처리할 달보다 앞이면 열어 볼 필요가 없음
        if not file_name_datetime(p.name) and datetime.fromtimestamp(st.st_mtime) < too_old:
            continue
        taken = photo_taken_at(p, st, cache)
        if start <= taken < end:
            found.append((p, taken))
    return sorted(found, key=lambda x: (x[1], x[0].name.lower()))


def processed_path(cfg: dict, original: Path) -> Path:
    return Path(cfg["보정본_폴더"]) / f"{original.stem}{PROCESSED_SUFFIX}.png"


def preprocess_one(cfg: dict, original: Path, out: Path, log: Log) -> str:
    """영수증 한 장 보정 → 'made' 또는 'failed'"""
    cv2, _ = _cv2()
    out.parent.mkdir(parents=True, exist_ok=True)
    img = imread_unicode(original)
    if img is None:
        log(f"   읽기 실패: {original.name}")
        return "failed"
    try:
        ok = imwrite_unicode(out, preprocess_receipt(img, float(cfg["보정_확대_배율"])), ".png",
                             [cv2.IMWRITE_PNG_COMPRESSION, 3])
    except Exception as e:
        log(f"   보정 실패: {original.name} ({e})")
        return "failed"
    return "made" if ok else "failed"


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

    def _fit_size(self, path: Path) -> Path:
        """jpg는 사진 회전 정보를 반영하고 heic는 Windows가 못 읽을 수 있어서 png로 바꿔 OCR합니다."""
        cv2, _ = _cv2()
        img = imread_unicode(path)
        if img is None:
            return path
        too_big = max(img.shape[:2]) > self.max_dim
        if path.suffix.lower() == ".png" and not too_big:
            return path
        if too_big:
            ratio = (self.max_dim - 1) / max(img.shape[:2])
            img = cv2.resize(img, None, fx=ratio, fy=ratio, interpolation=cv2.INTER_AREA)
        tmp = self.temp_dir / f"ocr_{path.stem}.png"
        imwrite_unicode(tmp, img)
        return tmp

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
        return asyncio.run(self._read(self._fit_size(path)))


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
    for cat in (priority or ["식대", "카페", "주차"]):
        if any(k and k in text_compact for k in keywords.get(cat, [])):
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


def detect_city(text_compact: str, cities: dict, hospital_key: str = "", low_priority=("서울",)) -> str:
    """영수증 주소의 'OO시/OO광역시/OO군' → 병원 이름 앞의 도시 이름 순서로 찾습니다. 못 찾으면 빈칸.
    서울처럼 본사 주소로 자주 찍히는 도시(low_priority)는 다른 도시가 없을 때만 씁니다."""
    found = []                                                        # (낮은 우선순위인지, 위치, 도시)
    for ko, full in (cities or {}).items():
        for m in re.finditer(re.escape(ko) + r"(?:광역시|특별시|특별자치시|시|군)", text_compact or ""):
            if ko == "광주" and "경기" in text_compact[max(0, m.start() - 4):m.start()]:
                continue                                              # 경기도 광주시
            found.append((ko in (low_priority or ()), m.start(), full))
            break
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
    if not rec.hospital_en:
        issues.append("병원 없음")
    if duplicate_count(rec, receipts) > 1:
        issues.append("같은 날짜·금액 영수증 있음")
    return issues


def duplicate_count(rec: Receipt, receipts: list) -> int:
    return sum(1 for r in receipts if r.use and r.main_date == rec.main_date and r.total == rec.total)


# ---------- 영수증 목록 (엑셀로 고치면 다음 실행에도 유지) ----------

LIST_HEADER = ["파일", "촬영 시각", "사용(O/X)", "날짜", "금액", "분류", "병원 영문", "병원 근거", "도시", "확인 필요",
               "OCR 앞부분", "[자동] 사용", "[자동] 날짜", "[자동] 금액", "[자동] 분류", "[자동] 병원 영문", "[자동] 도시",
               "날짜 후보", "금액 후보", "찾은 병원 글자"]
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
            " ".join(a.dates), " ".join(str(x) for x in a.amounts), a.hospital_key,
        ])
    write_csv(path, LIST_HEADER, rows)


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
    originals = list_originals(cfg, tm, date_cache)
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
    counts = {"made": 0, "failed": 0}

    def read_text(path: Path, st) -> str:
        c = cache.get(path.name)
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
        cache[path.name] = {"size": st.st_size, "mtime": int(st.st_mtime), "text": text}
        return text

    auto, receipts = {}, []
    try:
        for orig, taken in originals:
            st = orig.stat()
            ocr_file, text = orig, read_text(orig, st)
            first_is_receipt = analyze_receipt(orig, orig, text, cfg, hospitals, taken).is_receipt
            # 보정은 영수증으로 보이는 사진만 (일반 사진까지 보정하면 느리고 용량이 큼)
            if use_pre and first_is_receipt:
                out = processed_path(cfg, orig)
                if not out.exists():
                    counts[preprocess_one(cfg, orig, out, log)] += 1
                if out.exists():
                    better = read_text(out, st)
                    if better.strip():
                        ocr_file, text = out, better

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
    if counts["made"] or counts["failed"]:
        log(f"   보정: 새로 {counts['made']}장 / 실패 {counts['failed']}장")

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


def city_to_choose(rec, current: str, default: str):
    """Concur 도시 칸에 고를 도시(바꿀 필요가 없으면 None).
    영수증이 있으면 영수증에서 찾은 도시, 못 찾았으면 기본 도시(광주)로 맞춥니다(카드 정보로 서울 등이 미리 들어가 있어도).
    영수증이 없는 경비는 칸이 비어 있을 때만 기본 도시를 넣습니다."""
    if rec is None:
        return default if not (current or "").strip() else None
    want = rec.city or default
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
                                   f"Concur 화면이 바뀌었으면 [화면 조사]로 확인한 뒤 설정 '화면_요소'를 고치세요.")
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
        locator.wait_for(state="visible")
        locator.click()
        locator.press("Control+A")
        locator.press("Delete")
        if text:
            locator.press_sequentially(text, delay=15)
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
        self.log(f"리포트 생성: {name} ({tm.first_day_us} ~ {tm.last_day_us})")
        try:
            page.get_by_role("link", name=self._re("리포트 생성", exact=False)).first.click(timeout=15000)
        except Exception:
            page.goto(self.base + "/nui/expense?modal=header&context=new-report")
        self.type_text(self.field("리포트 이름"), name)
        self.type_text(self.field("리포트 시작일"), tm.first_day_us)
        self.type_text(self.field("리포트 종료일"), tm.last_day_us)
        self.type_text(self.field("리포트 Business Purpose"), name)
        self.type_text(self.field("리포트 Comment"), name)
        page.get_by_role("button", name=self._re("리포트 생성")).last.click()
        page.wait_for_url(re.compile(r"/nui/expense/reports/"))
        time.sleep(3)

    # ---------- 지난달 경비 추가 ----------
    def add_available_expenses(self, tm: TargetMonth):
        page = self.page
        card = page.locator(self.sel("사용 가능한 경비 카드")).filter(has_text=self._re("사용 가능한 경비", exact=False))
        if not self._visible(card.first, 20000):
            card = page.get_by_text(self._re("사용 가능한 경비", exact=False))
        card.first.click()
        page.wait_for_url(re.compile(r"view=available_expenses"))
        page.locator(self.sel("경비 체크박스")).first.wait_for(state="attached")
        time.sleep(2)
        r = page.evaluate(JS_SELECT_MONTH, {"month": tm.month, "year": tm.year})
        self.log(f"사용 가능한 경비 중 {tm.prefix} 경비 {r['matched']}건 선택")
        if r["matched"] == 0:
            self.log("   지난달 경비를 찾지 못했습니다. 목록 예시: " + " / ".join(r["samples"]))
            self.screenshot("available_expenses")
            raise RuntimeError("추가할 경비가 없습니다.")
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
        page.locator(self.sel("경비 목록")).first.wait_for(state="visible")
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
        """rec가 있으면 문구·Receipt·첨부까지, 없어도 LCRT·Expense Type·도시는 처리"""
        page, t, done = self.page, self.t, []
        try:
            page.locator(f'{self.sel("Receipt Status")}, {self.sel("Comment")}').first.wait_for(state="attached", timeout=8000)
        except Exception:
            pass

        for field_name in ("Comment", "Business Purpose"):
            text = texts_for(rec, self.cfg).get(field_name) if rec else None
            if text:
                self.type_text(self.field(field_name), text)
                done.append(f"{field_name} 입력")

        if rec:
            trigger, value = self.find_select(self.sel("Receipt Status"))
            if trigger is not None and value == t["No Receipt"]:
                self.choose_option(trigger, t["Receipt"])
                done.append("Receipt 선택")

        trigger, _ = self.find_select(self.sel("LCRT"))
        if trigger is not None:
            self.choose_option(trigger, t["NA"])
            done.append("LCRT=NA")

        trigger, value = self.find_select(self.sel("Expense Type"))
        if trigger is not None and value == t["Undefined"]:
            types = self.cfg["Undefined일_때_선택할_유형"]
            new_type = types.get((rec.category if rec else "") or "분류없음") or types.get("분류없음")
            if new_type:
                self.choose_option(trigger, new_type)
                done.append(f"Expense Type={new_type}")
                checks.append(f"Expense Type 자동 선택({new_type})")

        trigger, value = self.find_select(self.sel("도시"))
        if trigger is not None:
            city = city_to_choose(rec, value, self.cfg["기본_도시"])
            if city:
                self.choose_city(trigger, city)
                done.append(f"도시={city_name(city)}")

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
                    elif not rec.hospital_en:
                        checks.append("병원명 없음")
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
    title = {"ocr": "영수증 분석", "full": "전체 실행", "edit": "리포트에 입력", "inspect": "화면 조사"}[mode]
    if mode == "inspect":
        try:
            log(f"===== 화면 조사 | 프로그램 {APP_VERSION} =====")
            bot = ConcurBot(cfg, log, ui, stop_flag)
            bot.start()
            bot.wait_login(need_home=False)
            made = []
            while ui.confirm("화면 조사", "열린 Edge 창에서 조사할 Concur 화면(예: 경비 편집 화면, 리포트 만들기 화면)을 연 뒤 [확인]을 누르세요.\n"
                                         "다 끝났으면 [취소]를 누르세요."):
                path = bot.inspect_screen(log.folder / f"화면조사_{datetime.now():%Y%m%d_%H%M%S}.txt")
                made.append(path)
                log(f"화면 조사 파일: {path.name} (같은 이름의 .png 화면 캡처도 있습니다)")
            if made:
                ui.open(log.folder)
                ui.info("화면 조사", f"화면 조사 파일 {len(made)}개를 '실행기록' 폴더에 만들었습니다.\n"
                                   "'없음'으로 나온 칸은 이 파일의 id·name을 보고 설정 파일의 '화면_요소'를 고치거나, "
                                   "파일을 Claude에게 보내 고쳐 달라고 하세요.")
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
        if is_unfinished_month(tm):
            log(unfinished_month_note(tm))
        try:
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
                + "\n\n[예] 영수증 목록을 열어 고치기\n[아니오] 그대로 진행\n[취소] 중단")
            if ans is None:
                raise Cancelled()
            if ans:
                ui.open(list_path)
                if not ui.confirm("영수증 목록 수정", "엑셀에서 고친 뒤 저장하고 엑셀을 닫았으면 [확인]을 누르세요.\n([취소]는 중단)"):
                    raise Cancelled()
                receipts = build_receipts(cfg, tm, hospitals, results_dir, work_dir, log, base / HOSPITAL_FILE_NAME)

        bot = ConcurBot(cfg, log, ui, stop_flag)
        bot.hospitals = hospitals
        bot.start()
        if mode == "edit":
            bot.wait_login(need_home=False)
            if not ui.confirm("리포트 열기", "열린 Edge 창에서 입력할 리포트를 열어 경비 목록이 보이게 한 뒤 [확인]을 누르세요."):
                raise Cancelled()
        else:
            bot.wait_login()
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
                  "[확인]을 누르면 Edge 창이 닫힙니다.")
    except Cancelled:
        log("중단했습니다.")
    except Exception as e:
        log("오류: " + str(e))
        log(traceback.format_exc())
        if bot and bot.page:
            bot.screenshot("unexpected_error")
        ui.info("오류", f"{e}\n\n자세한 내용: {log.path.name}")
    finally:
        if bot:
            bot.stop()
        log(f"실행 기록: {log.path}")
        log.close()


# ============================================================
# 7. 창(GUI)
# ============================================================

class GuiApp:
    def __init__(self, base: Path):
        import tkinter as tk
        from tkinter import filedialog, messagebox, scrolledtext, simpledialog, ttk
        self.sd = simpledialog
        self.tk, self.mb, self.fd = tk, messagebox, filedialog
        self.base = base
        self.cfg, created = load_settings(base / SETTINGS_FILE_NAME)
        load_hospitals(base / HOSPITAL_FILE_NAME)
        self.q = queue.Queue()
        self.worker = None
        self.stop_flag = threading.Event()

        root = self.root = tk.Tk()
        root.title(f"Concur 경비 자동 입력 ({APP_VERSION})")
        root.geometry("820x720")
        root.minsize(720, 560)
        style = ttk.Style()
        for name in ("TLabel", "TButton", "TCheckbutton"):
            style.configure(name, font=("Malgun Gothic", 10))
        style.configure("Big.TButton", font=("Malgun Gothic", 11, "bold"))

        f = ttk.Frame(root, padding=14)
        f.pack(fill="both", expand=True)
        f.columnconfigure(1, weight=1)

        ttk.Label(f, text="처리할 달").grid(row=0, column=0, sticky="w", pady=3)
        months = recent_months(include_current=True)                     # 맨 위가 이번 달
        self.month = tk.StringVar(value=resolve_target_month(None).prefix)  # 처음 선택은 지난달
        mf = ttk.Frame(f)
        mf.grid(row=0, column=1, columnspan=2, sticky="w")
        ttk.Combobox(mf, textvariable=self.month, values=months, width=10, state="readonly").pack(side="left")
        ttk.Label(mf, foreground="#555", text="  처음에는 지난달이 골라져 있습니다. 맨 위의 이번 달도 고를 수 있습니다.").pack(side="left")

        ttk.Label(f, text="영수증 폴더").grid(row=1, column=0, sticky="w", pady=3)
        self.folder = tk.StringVar(value=self.cfg["영수증_원본_폴더"])
        ttk.Entry(f, textvariable=self.folder).grid(row=1, column=1, sticky="ew")
        ttk.Button(f, text="찾기", command=self.pick_folder).grid(row=1, column=2, padx=(6, 0))

        ttk.Label(f, text="기본 도시").grid(row=2, column=0, sticky="w", pady=3)
        self.city = tk.StringVar(value=self.cfg["기본_도시"])
        ttk.Entry(f, textvariable=self.city).grid(row=2, column=1, columnspan=2, sticky="ew")

        ttk.Label(f, text="Concur 로그인 이메일").grid(row=3, column=0, sticky="w", pady=3)
        lf = ttk.Frame(f)
        lf.grid(row=3, column=1, columnspan=2, sticky="ew")
        lf.columnconfigure(0, weight=1)
        self.login_email = tk.StringVar(value=str(self.cfg.get("로그인_이메일") or ""))
        ttk.Entry(lf, textvariable=self.login_email).grid(row=0, column=0, sticky="ew")
        self.auto_login = tk.BooleanVar(value=str(self.cfg.get("로그인_방법", "이메일 링크")).strip() == "이메일 링크")
        ttk.Checkbutton(lf, text="이메일 링크로 자동 로그인 (Outlook 메일 사용)", variable=self.auto_login).grid(row=0, column=1, padx=(8, 0))

        self.pre = tk.BooleanVar(value=bool(self.cfg.get("영수증_보정_사용", True)))
        ttk.Checkbutton(f, text="영수증 보정 사용 (사진이 선명하면 꺼도 됩니다)", variable=self.pre).grid(row=4, column=1, sticky="w", pady=(6, 0))
        self.review = tk.BooleanVar(value=bool(self.cfg.get("입력_전_확인", True)))
        ttk.Checkbutton(f, text="Concur 입력 전에 확인이 필요한 영수증 알려 주기", variable=self.review).grid(row=5, column=1, sticky="w")

        bf = ttk.Frame(f)
        bf.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(12, 4))
        self.run_buttons = []
        for text, mode in (("① 영수증 분석만", "ocr"), ("② 전체 실행", "full"), ("③ 만든 리포트에 입력만", "edit")):
            b = ttk.Button(bf, text=text, style="Big.TButton", command=lambda m=mode: self.start(m))
            b.pack(side="left", expand=True, fill="x", padx=3, ipady=8)
            self.run_buttons.append(b)
        self.stop_button = ttk.Button(bf, text="중지", command=self.stop, state="disabled")
        self.stop_button.pack(side="left", padx=3, ipady=8)

        sf = ttk.Frame(f)
        sf.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(2, 6))
        for text, cmd in (("영수증 목록 열기", self.open_list), ("병원 목록 편집", self.open_hospitals),
                          ("병원 엑셀 가져오기", self.import_hospital_excel),
                          ("자동 입력표 편집", self.open_category_table),
                          ("화면 조사", lambda: self.start("inspect")),
                          ("결과 폴더 열기", lambda: open_path(self.base / "결과")),
                          ("설정 파일 열기", lambda: open_path(self.base / SETTINGS_FILE_NAME, notepad=True))):
            ttk.Button(sf, text=text, command=cmd).pack(side="left", padx=3)

        ttk.Label(f, foreground="#555", wraplength=760, justify="left", text=(
            "영수증 폴더는 iCloud 사진 폴더나 iCloud Drive 폴더 모두 됩니다. 파일 이름·형식은 상관없고, 영수증이 아닌 사진은 자동으로 뺍니다.\n"
            "어느 달 영수증인지는 영수증에 적힌 날짜로 정하므로 늦게 찍어도 됩니다.\n"
            "처음에는 ①로 OCR 결과를 확인하세요. 영수증 목록(엑셀)에서 고친 내용은 다음 실행에도 그대로 쓰입니다.\n"
            "②③ 실행 중에는 열린 Edge 창을 클릭하지 마세요. 다른 프로그램은 써도 됩니다. 제출은 직접 하세요."
        )).grid(row=8, column=0, columnspan=3, sticky="w", pady=(0, 6))

        self.text = scrolledtext.ScrolledText(f, height=16, state="disabled", font=("Malgun Gothic", 9))
        self.text.grid(row=9, column=0, columnspan=3, sticky="nsew")
        f.rowconfigure(9, weight=1)

        root.protocol("WM_DELETE_WINDOW", self.on_close)
        root.after(100, self.poll)
        if created:
            self.append("처음 실행입니다. 영수증 폴더·기본 도시를 확인하고, [병원 목록 편집]에서 담당 병원을 넣어 주세요.")
            self.append("iPhone으로 찍기만 한다면 영수증 폴더로 iCloud 사진 폴더(보통 사진\\iCloud Photos\\Photos)를 고르세요.")
            self.append("병원 목록에는 영수증에서 확인한 병원(원광대병원, 목포한국병원)만 들어 있습니다. "
                        "영수증에서 새 병원이 보이면 자동으로 더하고, [병원 엑셀 가져오기]로 회사 엑셀(MR·CT·XR 탭)에서 가져올 수도 있습니다.")

    # ---------- 화면 동작 ----------
    def append(self, line):
        self.text.configure(state="normal")
        self.text.insert("end", line + "\n")
        self.text.see("end")
        self.text.configure(state="disabled")

    def pick_folder(self):
        d = self.fd.askdirectory(initialdir=self.folder.get() or str(Path.home()))
        if d:
            self.folder.set(str(Path(d)))

    def open_list(self):
        path = list_csv_path(self.base / "결과", resolve_target_month(self.month.get()))
        if path.exists():
            open_path(path)
        else:
            self.mb.showinfo("영수증 목록", "아직 이 달의 영수증 목록이 없습니다. ①을 먼저 실행하세요.")

    def open_category_table(self):
        path = self.base / CATEGORY_FILE_NAME
        if not path.exists():
            apply_category_table(load_settings(self.base / SETTINGS_FILE_NAME)[0], path)
        open_path(path)

    def open_hospitals(self):
        load_hospitals(self.base / HOSPITAL_FILE_NAME)
        open_path(self.base / HOSPITAL_FILE_NAME)

    def import_hospital_excel(self):
        cfg, _ = load_settings(self.base / SETTINGS_FILE_NAME)
        current = str(cfg.get("병원_엑셀_파일") or "")
        picked = self.fd.askopenfilename(
            title="병원 정보 엑셀 고르기 (MR·CT·XR 탭이 있는 파일)",
            initialdir=str(Path(current).parent) if current else str(Path.home()),
            filetypes=[("Excel 파일", "*.xlsx *.xlsm"), ("모든 파일", "*.*")])
        if not picked:
            return
        cfg["병원_엑셀_파일"] = str(Path(picked))
        save_settings(self.base / SETTINGS_FILE_NAME, cfg)
        try:
            read, added = import_hospitals_from_excel(cfg, self.base / HOSPITAL_FILE_NAME, self.append)
        except Exception as e:
            self.mb.showerror("병원 엑셀 가져오기", str(e))
            return
        self.mb.showinfo("병원 엑셀 가져오기",
                         f"병원 {read}곳을 읽었고, 병원 목록에 새로 {added}곳을 넣었습니다. (MR·CT·XR 탭)\n\n"
                         "엑셀에 영문 이름이 없는 병원은 한글을 자동으로 바꾼 것이라 공식 영문명과 다를 수 있습니다. "
                         "이어서 열리는 병원 목록에서 확인하고 고쳐 주세요.\n"
                         "앞으로는 실행할 때마다 이 엑셀에서 새 병원을 자동으로 더합니다.")
        if added:
            open_path(self.base / HOSPITAL_FILE_NAME)

    def save_form(self):
        cfg, _ = load_settings(self.base / SETTINGS_FILE_NAME)
        folder = self.folder.get().strip()
        if folder != cfg["영수증_원본_폴더"]:
            cfg["영수증_원본_폴더"] = folder
            cfg["보정본_폴더"] = str(Path(folder) / "Concur_Expenses_Processed")
        cfg["기본_도시"] = self.city.get().strip()
        cfg["영수증_보정_사용"] = bool(self.pre.get())
        cfg["입력_전_확인"] = bool(self.review.get())
        cfg["로그인_이메일"] = self.login_email.get().strip()
        cfg["로그인_방법"] = "이메일 링크" if self.auto_login.get() else "직접"
        save_settings(self.base / SETTINGS_FILE_NAME, cfg)

    def set_running(self, running: bool):
        for b in self.run_buttons:
            b.configure(state="disabled" if running else "normal")
        self.stop_button.configure(state="normal" if running else "disabled")

    def start(self, mode):
        if self.worker and self.worker.is_alive():
            return
        tm = resolve_target_month(self.month.get())
        note = unfinished_month_note(tm) if is_unfinished_month(tm) else ""
        if mode == "full" and not self.mb.askokcancel(
                "전체 실행", f"{self.month.get()} 리포트를 새로 만듭니다.\n이미 만들었다면 ③을 쓰세요."
                + (f"\n\n{note}" if note else "") + "\n\n계속할까요?"):
            return
        if mode == "edit" and note and not self.mb.askokcancel("리포트에 입력", f"{note}\n\n계속할까요?"):
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

    def stop(self):
        self.stop_flag.set()
        self.append("중지 요청: 지금 처리 중인 경비가 끝나면 멈춥니다.")

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
                elif kind == "done":
                    self.set_running(False)
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def on_close(self):
        if self.worker and self.worker.is_alive():
            if not self.mb.askyesno("종료", "작업 중입니다. 종료하면 Edge 창도 닫힙니다. 종료할까요?"):
                return
            self.root.destroy()
            os._exit(0)
        self.root.destroy()

    def run(self):
        self.root.mainloop()


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
