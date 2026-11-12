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

SETTINGS_FILE_NAME = "concur_settings.json"
HOSPITAL_FILE_NAME = "병원목록.csv"
CATEGORIES = ("주차", "카페", "식대")

DEFAULT_HOSPITALS = [
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
        "영수증_보정_사용": True,
        "보정_확대_배율": 2.0,
        "입력_전_확인": True,
        "결과_자동_열기": True,
        "_업로드_설명": "원본 또는 보정본. Concur는 1장당 5MB 제한이 있어 넘으면 자동으로 줄여서 올립니다.",
        "업로드_이미지": "원본",
        "업로드_최대_MB": 5,
        "기본_도시": "Daejeon (ex Taejon), KOREA, REPUBLIC OF",
        "리포트_이름_형식": "{month_name} Expenses",
        "_키워드_설명": "OCR 결과(띄어쓰기 무시)에 이 글자가 있으면 해당 분류. 여러 분류가 맞으면 분류_우선순위 앞쪽이 이깁니다.",
        "키워드": {
            "주차": ["주차", "입차", "출차", "주자"],
            "카페": ["메리카", "리카노", "테이크", "카페"],
            "식대": ["버터"],
        },
        "분류_우선순위": ["식대", "카페", "주차"],
        "_영수증_판별_설명": "폴더에 일반 사진이 섞여 있어도 됩니다. 금액이 보이고 아래 글자·OCR 날짜·분류 키워드 중 하나가 있으면 영수증으로 봅니다. 아니면 영수증 목록 맨 아래에 '사용 X'로 넣습니다.",
        "영수증_판별_단어": ["합계", "결제", "승인", "카드", "영수증", "금액", "부가세", "사업자", "매출", "가맹점", "거래", "주차"],
        "_입력_문구_설명": "{hospital} 자리에 병원 영문명. 병원을 못 찾으면 ' in {hospital}' 부분은 빠집니다.",
        "입력_문구": {
            "주차": {"Comment": "Parking fee for service in {hospital}", "Business Purpose": "For service"},
            "카페": {"Comment": "Cafe for wifi in {hospital}", "Business Purpose": "For wifi"},
            "식대": {"Comment": "Overtime meal fee for service in {hospital}", "Business Purpose": "For Overtime meal"},
        },
        "_Undefined_설명": "Expense Type이 Undefined인 경비에 고를 유형. 결과 파일에 '확인 필요'로 표시됩니다.",
        "Undefined일_때_선택할_유형": {
            "주차": "Parking & Toll",
            "카페": "Other Staff Expenses",
            "식대": "Other Staff Expenses",
            "분류없음": "Other Staff Expenses",
        },
        "_화면_문구_설명": "Concur 화면에 보이는 글자. Concur 언어 설정이 바뀌면 여기만 고칩니다.",
        "화면_문구": {
            "리포트 생성": "리포트 생성",
            "사용 가능한 경비": "사용 가능한 경비",
            "리포트에 추가": "리포트에 추가",
            "편집": "편집",
            "다음 경비": "다음 경비",
            "경비 저장": "경비 저장",
            "No Receipt": "No Receipt",
            "Receipt": "Receipt",
            "NA": "NA",
            "Undefined": "Undefined",
        },
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
    return merged, False


def save_settings(path: Path, cfg: dict):
    path.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


def load_hospitals(path: Path) -> list:
    if not path.exists():
        write_csv(path, ["OCR에서 찾을 글자", "Concur에 쓸 영문명"], DEFAULT_HOSPITALS)
    pairs = []
    for row in read_csv_dicts(path):
        vals = [str(v or "").strip() for v in row.values()]
        if len(vals) >= 2 and vals[0] and vals[1]:
            pairs.append([vals[0], vals[1]])
    return pairs


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
    taken: object = None                            # 촬영 시각(datetime)
    is_receipt: bool = True                         # 영수증으로 보이는지(자동 판단)
    ocr_dates: list = field(default_factory=list)   # 영수증 글자에서 읽은 날짜
    skip_reason: str = ""                           # 이 달 목록에서 자동으로 뺀 이유

    @property
    def main_date(self) -> str:
        return self.dates[0] if self.dates else ""


DATE_IN_TEXT = re.compile(r"(?<!\d)(20\d{2})\s*[-./년]\s*(\d{1,2})\s*[-./월]\s*(\d{1,2})(?!\d)")
DATE_IN_NAME = re.compile(r"(20\d{2})-(\d{2})-(\d{2})")
AMOUNT_WITH_SEP = re.compile(r"(?<![\d.,])\d{1,3}(?:[,.]\d{3})+(?![\d,]|\.\d)")
AMOUNT_WON = re.compile(r"(?<![\d,.])(\d{3,7})\s*원")


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
    for m in DATE_IN_TEXT.finditer(ocr_text):
        iso = normalize_date(*m.groups())
        if iso and iso not in found:
            found.append(iso)
    return found


def file_name_date(name: str):
    m = DATE_IN_NAME.search(name)
    return normalize_date(*m.groups()) if m else None


def find_amounts(ocr_text: str) -> list:
    values = [int(re.sub(r"\D", "", m.group(0))) for m in AMOUNT_WITH_SEP.finditer(ocr_text)]
    values += [int(m.group(1)) for m in AMOUNT_WON.finditer(ocr_text)]
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

LIST_HEADER = ["파일", "촬영 시각", "사용(O/X)", "날짜", "금액", "분류", "병원 영문", "확인 필요", "OCR 앞부분",
               "[자동] 사용", "[자동] 날짜", "[자동] 금액", "[자동] 분류", "[자동] 병원 영문",
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


def apply_user_edits(rec: Receipt, row: dict, tm=None):
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
        rec.category = g("분류") if g("분류") in CATEGORIES else ""
        rec.edited.append("분류")
    if g("병원 영문") != g("[자동] 병원 영문"):
        rec.hospital_en = g("병원 영문")
        rec.edited.append("병원")
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
            "O" if r.use else "X", r.main_date, r.total or "", r.category, r.hospital_en,
            ", ".join(issues), compact(r.ocr_text)[:100],
            "O" if a.use else "X", a.main_date, a.total or "", a.category, a.hospital_en,
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


def build_receipts(cfg: dict, tm: TargetMonth, hospitals: list, results_dir: Path, work_dir: Path, log: Log) -> list:
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

    list_path = list_csv_path(results_dir, tm)
    if list_path.exists():
        try:
            rows = {str(r.get("파일") or "").strip(): r for r in read_csv_dicts(list_path)}
        except PermissionError:
            raise RuntimeError(f"영수증 목록 파일이 열려 있어 읽을 수 없습니다. 엑셀을 닫고 다시 실행하세요.\n{list_path}")
        for rec in receipts:
            if rec.original.name in rows:
                apply_user_edits(rec, rows[rec.original.name], tm)
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
            f"{rec.hospital_en or '병원없음'}{edited}{flag}")
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


def texts_for(rec: Receipt, cfg: dict) -> dict:
    tmpl = cfg["입력_문구"].get(rec.category) if rec and rec.category else None
    if not tmpl:
        return {}
    out = {}
    for name, t in tmpl.items():
        if not rec.hospital_en:
            t = t.replace(" in {hospital}", "")
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


class ConcurBot:
    def __init__(self, cfg: dict, log: Log):
        self.cfg = cfg
        self.log = log
        self.base = cfg["Concur_주소"].rstrip("/")
        self.t = cfg["화면_문구"]
        self.timeout_ms = int(cfg.get("대기시간_초", 60)) * 1000
        self.pw = self.ctx = self.page = None

    def start(self):
        try:
            from playwright.sync_api import sync_playwright
        except ImportError:
            raise RuntimeError("playwright 가 필요합니다:  pip install playwright")
        profile = Path(self.cfg["브라우저_프로필_폴더"])
        profile.mkdir(parents=True, exist_ok=True)
        self.pw = sync_playwright().start()
        self.ctx = self.pw.chromium.launch_persistent_context(
            str(profile), channel="msedge", headless=False, no_viewport=True, args=["--start-maximized"])
        self.ctx.set_default_timeout(self.timeout_ms)
        self.page = self.ctx.pages[0] if self.ctx.pages else self.ctx.new_page()

    def stop(self):
        for closer in (lambda: self.ctx and self.ctx.close(), lambda: self.pw and self.pw.stop()):
            try:
                closer()
            except Exception:
                pass

    # ---------- 도우미 ----------
    def _exact(self, text: str):
        return re.compile(r"^\s*" + re.escape(text) + r"\s*$")

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

    def choose_option(self, trigger, option_text: str):
        page = self.page
        options = page.locator(".sapcnqr-listbox-item__wrapper:visible").filter(has_text=self._exact(option_text))
        trigger.click()
        if not self._visible(options.first, 4000):
            page.keyboard.type(option_text.split(",")[0][:20], delay=30)
            if not self._visible(options.first, 6000):
                page.keyboard.press("Escape")
                raise RuntimeError(f"드롭다운에서 '{option_text}' 항목을 찾지 못했습니다.")
        options.first.click()
        time.sleep(0.5)

    def screenshot(self, name: str):
        try:
            path = self.log.folder / f"{datetime.now():%H%M%S}_{name}.png"
            self.page.screenshot(path=str(path))
            self.log(f"   화면 저장: {path.name}")
        except Exception:
            pass

    # ---------- 로그인 ----------
    def wait_login(self, need_home: bool = True):
        page = self.page
        page.goto(self.base + "/home", wait_until="domcontentloaded")
        self.log("Concur 로그인 확인 중... 로그인 화면이 뜨면 열린 Edge 창에서 직접 로그인하세요. (최대 10분)")
        create = page.get_by_role("link", name=self.t["리포트 생성"])
        deadline = time.time() + 600
        while time.time() < deadline:
            try:
                if "concursolutions.com" in page.url and create.count() > 0:
                    self.log("로그인 확인")
                    return
                if not need_home and "concursolutions.com" in page.url and page.locator("#cnqr-app-content").count() > 0:
                    self.log("로그인 확인")
                    return
            except Exception:
                pass
            time.sleep(2)
        raise RuntimeError("로그인 대기 시간(10분)이 지났습니다.")

    # ---------- 리포트 생성 ----------
    def create_report(self, tm: TargetMonth):
        page = self.page
        name = tm.report_name(self.cfg["리포트_이름_형식"])
        self.log(f"리포트 생성: {name} ({tm.first_day_us} ~ {tm.last_day_us})")
        try:
            page.get_by_role("link", name=self.t["리포트 생성"]).first.click(timeout=15000)
        except Exception:
            page.goto(self.base + "/nui/expense?modal=header&context=new-report")
        page.locator("#name").wait_for(state="visible")
        self.type_text(page.locator("#name"), name)
        self.type_text(page.locator("#custom11-date-input-field-input"), tm.first_day_us)
        self.type_text(page.locator("#custom12-date-input-field-input"), tm.last_day_us)
        self.type_text(page.locator("#businessPurpose:visible").first, name)
        self.type_text(page.locator("#comment:visible").first, name)
        page.get_by_role("button", name=self.t["리포트 생성"], exact=True).last.click()
        page.wait_for_url(re.compile(r"/nui/expense/reports/"))
        time.sleep(3)

    # ---------- 지난달 경비 추가 ----------
    def add_available_expenses(self, tm: TargetMonth):
        page = self.page
        card = page.locator(".report-empty-state-actions__card-body").filter(has_text=self.t["사용 가능한 경비"])
        if not self._visible(card.first, 20000):
            card = page.get_by_text(self.t["사용 가능한 경비"])
        card.first.click()
        page.wait_for_url(re.compile(r"view=available_expenses"))
        page.locator("input[type=checkbox].sapcnqr-checkbox__input").first.wait_for(state="attached")
        time.sleep(2)
        r = page.evaluate(JS_SELECT_MONTH, {"month": tm.month, "year": tm.year})
        self.log(f"사용 가능한 경비 중 {tm.prefix} 경비 {r['matched']}건 선택")
        if r["matched"] == 0:
            self.log("   지난달 경비를 찾지 못했습니다. 목록 예시: " + " / ".join(r["samples"]))
            self.screenshot("available_expenses")
            raise RuntimeError("추가할 경비가 없습니다.")
        page.get_by_role("button", name=self.t["리포트에 추가"], exact=True).click()
        page.locator(".entries-list").first.wait_for(state="visible")
        time.sleep(3)

    # ---------- 경비 편집 화면 열기 ----------
    def open_first_expense(self):
        page = self.page
        page.locator(".entries-list").first.wait_for(state="visible")
        boxes = page.locator(".entries-list .sapcnqr-dropzone input.sapcnqr-checkbox__input")
        try:
            boxes.first.wait_for(state="attached", timeout=10000)
            self.click_checkbox(boxes.first)
        except Exception:
            self.click_checkbox(page.locator(".entries-list input.sapcnqr-checkbox__input").nth(1))
        toolbar_edit = page.locator(".sapcnqr-toolbar__toolbar").get_by_role("button", name=self.t["편집"], exact=True)
        if self._visible(toolbar_edit.first, 5000):
            toolbar_edit.first.click()
        else:
            page.get_by_role("button", name=self.t["편집"], exact=True).first.click()
        page.locator("#transactionAmount").first.wait_for(state="attached")

    def progress(self) -> tuple:
        el = self.page.get_by_text(re.compile(r"^\s*\d+\s*/\s*\d+\s*개\s*$")).first
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
            page.locator('[name="receiptType"], #comment').first.wait_for(state="attached", timeout=8000)
        except Exception:
            pass

        for field_id, field_name in (("comment", "Comment"), ("businessPurpose", "Business Purpose")):
            text = texts_for(rec, self.cfg).get(field_name) if rec else None
            if text:
                self.type_text(page.locator(f"#{field_id}:visible").first, text)
                done.append(f"{field_name} 입력")

        if rec:
            trigger, value = self.find_select('[name="receiptType"]')
            if trigger is not None and value == t["No Receipt"]:
                self.choose_option(trigger, t["Receipt"])
                done.append("Receipt 선택")

        trigger, _ = self.find_select("#custom32 .sapcnqr-select-field__input")
        if trigger is not None:
            self.choose_option(trigger, t["NA"])
            done.append("LCRT=NA")

        trigger, value = self.find_select('[name="expenseType"]')
        if trigger is not None and value == t["Undefined"]:
            types = self.cfg["Undefined일_때_선택할_유형"]
            new_type = types.get((rec.category if rec else "") or "분류없음") or types.get("분류없음")
            if new_type:
                self.choose_option(trigger, new_type)
                done.append(f"Expense Type={new_type}")
                checks.append(f"Expense Type 자동 선택({new_type})")

        trigger, value = self.find_select("#location")
        if trigger is not None and not value:
            self.choose_option(trigger, self.cfg["기본_도시"])
            done.append("도시 선택")

        upload_btn = page.locator("#upload-receipt-button:visible").first
        if rec and upload and self._visible(upload_btn, 1500):
            try:
                with page.expect_file_chooser(timeout=15000) as chooser:
                    upload_btn.click()
                chooser.value.set_files(str(upload))
            except Exception:
                file_input = page.locator("input[type=file]")
                if file_input.count() == 0:
                    raise
                file_input.last.set_input_files(str(upload))
            try:
                upload_btn.wait_for(state="hidden", timeout=30000)
            except Exception:
                self.log("   업로드 완료 표시를 확인하지 못해 10초 더 기다립니다.")
                time.sleep(10)
            time.sleep(2)
            done.append(f"영수증 첨부({upload.name})")
        elif rec:
            done.append("이미 첨부됨(건너뜀)")
        return done

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
                    page.get_by_role("button", name=self.t["경비 저장"], exact=True).first.click()
                    time.sleep(3)
                except Exception:
                    self.log("   저장 버튼을 누르지 못했습니다. Concur 화면에서 직접 저장하세요.")
                break
            row = {"순번": i, "Concur 날짜": "", "Concur 금액": "", "영수증": "", "분류": "", "처리": "", "확인 필요": "", "오류": ""}
            checks = []
            try:
                self.wait_progress(i)
                date_text = page.locator("#transactionDate-date-input-field-input:visible").first.input_value()
                amount_text = page.locator("#transactionAmount:visible").first.input_value()
                row["Concur 날짜"], row["Concur 금액"] = date_text, amount_text
                iso, amount = concur_date_to_iso(date_text), concur_amount_to_int(amount_text)
                rec = find_receipt(receipts, iso, amount) if iso and amount is not None else None
                upload = None
                if rec is None:
                    checks.append("영수증 없음")
                else:
                    rec.used_for = f"{i}번 {date_text} {amount_text}"
                    row["영수증"], row["분류"] = rec.original.name, rec.category
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
                    page.get_by_role("button", name=self.t["경비 저장"], exact=True).first.click()
                    time.sleep(3)
                    self.log("마지막 경비 저장 완료")
                else:
                    page.get_by_role("button", name=self.t["다음 경비"]).first.click()
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
    hospitals = load_hospitals(base / HOSPITAL_FILE_NAME)
    tm = resolve_target_month(month_text)
    results_dir = base / "결과"
    results_dir.mkdir(parents=True, exist_ok=True)
    log = Log(base / "실행기록", ui.log)
    work_dir = log.folder / "임시"
    bot = None
    title = {"ocr": "영수증 분석", "full": "전체 실행", "edit": "리포트에 입력"}[mode]
    try:
        log(f"===== {title} | {tm.prefix} ({tm.month_name}) =====")
        if is_unfinished_month(tm):
            log(unfinished_month_note(tm))
        receipts = build_receipts(cfg, tm, hospitals, results_dir, work_dir, log)
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
                receipts = build_receipts(cfg, tm, hospitals, results_dir, work_dir, log)

        bot = ConcurBot(cfg, log)
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
        from tkinter import filedialog, messagebox, scrolledtext, ttk
        self.tk, self.mb, self.fd = tk, messagebox, filedialog
        self.base = base
        self.cfg, created = load_settings(base / SETTINGS_FILE_NAME)
        load_hospitals(base / HOSPITAL_FILE_NAME)
        self.q = queue.Queue()
        self.worker = None
        self.stop_flag = threading.Event()

        root = self.root = tk.Tk()
        root.title("Concur 경비 자동 입력")
        root.geometry("820x660")
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

        self.pre = tk.BooleanVar(value=bool(self.cfg.get("영수증_보정_사용", True)))
        ttk.Checkbutton(f, text="영수증 보정 사용 (사진이 선명하면 꺼도 됩니다)", variable=self.pre).grid(row=3, column=1, sticky="w", pady=(6, 0))
        self.review = tk.BooleanVar(value=bool(self.cfg.get("입력_전_확인", True)))
        ttk.Checkbutton(f, text="Concur 입력 전에 확인이 필요한 영수증 알려 주기", variable=self.review).grid(row=4, column=1, sticky="w")

        bf = ttk.Frame(f)
        bf.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(12, 4))
        self.run_buttons = []
        for text, mode in (("① 영수증 분석만", "ocr"), ("② 전체 실행", "full"), ("③ 만든 리포트에 입력만", "edit")):
            b = ttk.Button(bf, text=text, style="Big.TButton", command=lambda m=mode: self.start(m))
            b.pack(side="left", expand=True, fill="x", padx=3, ipady=8)
            self.run_buttons.append(b)
        self.stop_button = ttk.Button(bf, text="중지", command=self.stop, state="disabled")
        self.stop_button.pack(side="left", padx=3, ipady=8)

        sf = ttk.Frame(f)
        sf.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(2, 6))
        for text, cmd in (("영수증 목록 열기", self.open_list), ("병원 목록 편집", self.open_hospitals),
                          ("결과 폴더 열기", lambda: open_path(self.base / "결과")),
                          ("설정 파일 열기", lambda: open_path(self.base / SETTINGS_FILE_NAME, notepad=True))):
            ttk.Button(sf, text=text, command=cmd).pack(side="left", padx=3)

        ttk.Label(f, foreground="#555", wraplength=760, justify="left", text=(
            "영수증 폴더는 iCloud 사진 폴더나 iCloud Drive 폴더 모두 됩니다. 파일 이름·형식은 상관없고, 영수증이 아닌 사진은 자동으로 뺍니다.\n"
            "어느 달 영수증인지는 영수증에 적힌 날짜로 정하므로 늦게 찍어도 됩니다.\n"
            "처음에는 ①로 OCR 결과를 확인하세요. 영수증 목록(엑셀)에서 고친 내용은 다음 실행에도 그대로 쓰입니다.\n"
            "②③ 실행 중에는 열린 Edge 창을 클릭하지 마세요. 다른 프로그램은 써도 됩니다. 제출은 직접 하세요."
        )).grid(row=7, column=0, columnspan=3, sticky="w", pady=(0, 6))

        self.text = scrolledtext.ScrolledText(f, height=16, state="disabled", font=("Malgun Gothic", 9))
        self.text.grid(row=8, column=0, columnspan=3, sticky="nsew")
        f.rowconfigure(8, weight=1)

        root.protocol("WM_DELETE_WINDOW", self.on_close)
        root.after(100, self.poll)
        if created:
            self.append("처음 실행입니다. 영수증 폴더·기본 도시를 확인하고, [병원 목록 편집]에서 담당 병원을 넣어 주세요.")
            self.append("iPhone으로 찍기만 한다면 영수증 폴더로 iCloud 사진 폴더(보통 사진\\iCloud Photos\\Photos)를 고르세요.")
            self.append("기본 병원 목록은 대전·충청권입니다. 짧은 글자는 다른 병원에도 걸립니다(예: '남대'는 '전남대'에도 맞음).")

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

    def open_hospitals(self):
        load_hospitals(self.base / HOSPITAL_FILE_NAME)
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
    args = parser.parse_args()
    base = app_dir()

    cli = args.cli or args.ocr_only or args.edit_only or args.full
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

    mode = "ocr" if args.ocr_only else "edit" if args.edit_only else "full" if args.full else ""
    if not mode:
        print("1) 영수증 분석만   2) 전체 실행   3) 만든 리포트에 입력만")
        mode = {"1": "ocr", "2": "full", "3": "edit"}.get(input("번호: ").strip(), "")
        if not mode:
            return
        if not args.month:
            last, this = resolve_target_month(None).prefix, recent_months(1, include_current=True)[0]
            args.month = input(f"처리할 달 (Enter={last} 지난달, 이번 달은 {this}): ").strip() or None
    run_job(mode, args.month, ConsoleUI(), base)


if __name__ == "__main__":
    main()
