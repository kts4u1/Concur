# -*- coding: utf-8 -*-
"""병원 찾기: 병원 정보 엑셀 가져오기(한글→영문 자동 변환), 같은 날 영수증의 병원, Concur 가게 이름"""

import os
from datetime import datetime
from pathlib import Path

import pytest

import concur_auto as ca


# ---------- 한글 → 영문 ----------

@pytest.mark.parametrize("ko, en", [
    ("원광", "wongwang"), ("전북", "jeonbuk"), ("한림", "hallim"), ("단국", "danguk"),
    ("순천향", "suncheonhyang"), ("백제", "baekje"), ("을지", "eulji"), ("삼육", "samyuk"), ("신라", "silla"),
])
def test_romanize(ko, en):
    assert ca.romanize_korean(ko) == en


@pytest.mark.parametrize("raw, key, en", [
    ("원광대학교병원", "원광대", "Wongwang University Hospital"),
    ("원광대학교 산본병원", "원광대학교산본", "Wongwang University Sanbon Hospital"),
    ("국립소방병원", "국립소방", "National Sobang Hospital"),
    ("대전성모병원", "대전성모", "Daejeon St. Mary's Hospital"),
    ("군산의료원", "군산의료원", "Gunsan Medical Center"),       # '군산'만으로는 군산 시내 영수증에 다 걸려서 전체 이름
    ("전북대학교병원 3.0T (전주)", "전북대", "Jeonbuk University Hospital"),
])
def test_hospital_key_and_english(raw, key, en):
    assert ca.hospital_key(raw) == key
    assert ca.hospital_english_name(raw) == en


# ---------- 병원 정보 엑셀 ----------

def make_excel(path, header_row=3, sheet="MR System Information"):
    openpyxl = pytest.importorskip("openpyxl")
    wb = openpyxl.Workbook()
    wb.active.title = "Summary"
    ws = wb.create_sheet(sheet)
    ws.cell(1, 1, "MR 장비 현황")
    header = ["No", "지역", "병원명", "모델", "S/N"]
    data = [
        [1, "익산", "원광대학교병원", "Ingenia 3.0T", "A1"],
        [2, "익산", "원광대학교병원 1.5T", "Achieva", "A2"],      # 같은 병원 두 번째 장비
        [3, "전주", "전북대학교병원", "Elition", "A3"],
        [4, "대전", "건양대학교병원", "Ingenia", "A4"],           # 기본 목록에 이미 있는 병원(건양대)
        [5, "군산", "군산의료원", "Prodiva", "A5"],
        [6, "", None, "", ""],
    ]
    for ci, v in enumerate(header, 1):
        ws.cell(header_row, ci, v)
    for ri, row in enumerate(data, header_row + 1):
        for ci, v in enumerate(row, 1):
            ws.cell(ri, ci, v)
    wb.save(path)
    return path


def test_read_hospital_names_auto_column(tmp_path):
    p = make_excel(tmp_path / "mr.xlsx")
    assert ca.read_hospital_names_from_excel(p, "MR System Information") == \
        ["원광대학교병원", "전북대학교병원", "건양대학교병원", "군산의료원"]


@pytest.mark.parametrize("column", ["병원명", "C", "c"])
def test_read_hospital_names_given_column(tmp_path, column):
    p = make_excel(tmp_path / "mr.xlsx")
    assert ca.read_hospital_names_from_excel(p, "mr system information", column)[0] == "원광대학교병원"


def test_read_hospital_names_errors(tmp_path):
    p = make_excel(tmp_path / "mr.xlsx", sheet="Other")
    with pytest.raises(RuntimeError, match="탭이 없습니다"):
        ca.read_hospital_names_from_excel(p, "MR System Information")
    p = make_excel(tmp_path / "mr2.xlsx")
    with pytest.raises(RuntimeError, match="열 제목"):
        ca.read_hospital_names_from_excel(p, "MR System Information", "고객 이름")


def test_import_adds_only_new_hospitals_and_keeps_user_edits(tmp_path):
    xlsx = make_excel(tmp_path / "mr.xlsx")
    csv_path = tmp_path / ca.HOSPITAL_FILE_NAME
    cfg = {"병원_엑셀_파일": str(xlsx), "병원_엑셀_탭": "MR System Information", "병원_엑셀_열": ""}
    lines = []
    assert ca.import_hospitals_from_excel(cfg, csv_path, lines.append) == (4, 3)   # 건양대는 이미 있음
    rows = ca.load_hospital_rows(csv_path)
    added = {r[0]: r for r in rows[len(ca.DEFAULT_HOSPITALS):]}
    assert added["원광대"][1] == "Wongwang University Hospital" and "자동 변환" in added["원광대"][2]
    assert [r[1] for r in rows if r[0] == "건양대"] == ["KYUH"]
    assert "병원 엑셀: 4곳 읽음, 병원 목록에 새로 3곳 추가" in lines[0]

    # 사용자가 영문 이름을 고친 뒤 다시 가져와도 그대로
    rows = [r if r[0] != "원광대" else ["원광대", "WKUH", r[2]] for r in rows]
    ca.write_csv(csv_path, ca.HOSPITAL_HEADER, rows)
    assert ca.import_hospitals_from_excel(cfg, csv_path, lines.append) == (4, 0)
    assert ["원광대", "WKUH"] in ca.load_hospitals(csv_path)


def test_import_without_file_setting_does_nothing(tmp_path):
    assert ca.import_hospitals_from_excel({}, tmp_path / "h.csv", print) == (0, 0)
    lines = []
    assert ca.import_hospitals_from_excel({"병원_엑셀_파일": str(tmp_path / "없음.xlsx")}, tmp_path / "h.csv", lines.append) == (0, 0)
    assert "병원 엑셀 파일이 없습니다" in lines[0]


def test_imported_hospital_is_found_on_receipt(tmp_path):
    xlsx = make_excel(tmp_path / "mr.xlsx")
    csv_path = tmp_path / ca.HOSPITAL_FILE_NAME
    ca.import_hospitals_from_excel({"병원_엑셀_파일": str(xlsx)}, csv_path, lambda s: None)
    p = Path("IMG_0001.JPG")
    rec = ca.analyze_receipt(p, p, "원광대학교병원 주차장 입차 출차 3,000원", ca.default_settings(),
                             ca.load_hospitals(csv_path), datetime(2026, 11, 3, 9, 0))
    assert rec.hospital_en == "Wongwang University Hospital" and rec.hospital_src == "영수증 글자"


def test_run_job_imports_excel(tmp_path):
    xlsx = make_excel(tmp_path / "mr.xlsx")
    src = tmp_path / "src"
    src.mkdir()
    base = tmp_path / "app"
    base.mkdir()
    ca.save_settings(base / ca.SETTINGS_FILE_NAME,
                     {"영수증_원본_폴더": str(src), "결과_자동_열기": False, "병원_엑셀_파일": str(xlsx)})

    class UI:
        lines = []

        def log(self, line):
            self.lines.append(line)

        def info(self, title, msg):
            pass

        def open(self, path):
            pass

    ui = UI()
    ca.run_job("ocr", "2026-10", ui, base)
    assert any("병원 엑셀: 4곳 읽음, 병원 목록에 새로 3곳 추가" in line for line in ui.lines)


# ---------- 같은 날 영수증의 병원 ----------

def rec(name, date, category, hospital="", src=""):
    r = ca.Receipt(original=Path(name), ocr_file=Path(name), dates=[date], category=category,
                   hospital_en=hospital, hospital_src=src)
    return r


def test_same_day_parking_hospital_fills_cafe_and_meal():
    parking = rec("p.jpg", "2026-11-03", "주차", "Wongwang University Hospital", "영수증 글자")
    cafe = rec("c.jpg", "2026-11-03", "카페")
    meal = rec("m.jpg", "2026-11-03", "식대")
    other_day = rec("o.jpg", "2026-11-04", "카페")
    assert ca.fill_same_day_hospitals([parking, cafe, meal, other_day]) == 2
    assert (cafe.hospital_en, cafe.hospital_src) == ("Wongwang University Hospital", "같은 날 영수증")
    assert meal.hospital_en == "Wongwang University Hospital"
    assert other_day.hospital_en == ""


def test_same_day_with_two_hospitals_prefers_parking_or_gives_up():
    a = rec("a.jpg", "2026-11-03", "주차", "Hospital A", "영수증 글자")
    b = rec("b.jpg", "2026-11-03", "카페", "Hospital B", "영수증 글자")
    meal = rec("m.jpg", "2026-11-03", "식대")
    ca.fill_same_day_hospitals([a, b, meal])
    assert meal.hospital_en == "Hospital A"

    a2 = rec("a2.jpg", "2026-11-05", "주차", "Hospital A", "영수증 글자")
    c2 = rec("c2.jpg", "2026-11-05", "주차", "Hospital C", "영수증 글자")
    meal2 = rec("m2.jpg", "2026-11-05", "식대")
    ca.fill_same_day_hospitals([a2, c2, meal2])
    assert meal2.hospital_en == ""                    # 어느 병원인지 모름 → '병원 없음'으로 남김


def test_same_day_does_not_fill_when_user_cleared_or_excluded():
    parking = rec("p.jpg", "2026-11-03", "주차", "Hospital A", "영수증 글자")
    cleared = rec("c.jpg", "2026-11-03", "카페")
    cleared.edited = ["병원"]
    excluded = rec("x.jpg", "2026-11-03", "카페")
    excluded.use = False
    ca.fill_same_day_hospitals([parking, cleared, excluded])
    assert cleared.hospital_en == "" and excluded.hospital_en == ""


def test_receipt_list_shows_hospital_source(tmp_path, monkeypatch):
    src = tmp_path / "src"
    src.mkdir()
    texts = {"2026-10-05_090000.jpg": "2026.10.05 건양대학교병원 주차 입차 3,000원",
             "2026-10-05_120000.jpg": "2026.10.05 아메리카노 4,500"}
    for name in texts:
        (src / name).write_bytes(b"x")

    class FakeOcr:
        def __init__(self, temp_dir):
            pass

        def read(self, path):
            return texts[path.name]

    monkeypatch.setattr(ca, "WindowsOcr", FakeOcr)
    cfg = ca.default_settings()
    cfg.update({"영수증_원본_폴더": str(src), "보정본_폴더": str(src / "p"), "영수증_보정_사용": False})
    results = tmp_path / "결과"
    results.mkdir()
    log = ca.Log(tmp_path / "log", sink=lambda s: None)
    try:
        recs = ca.build_receipts(cfg, ca.TargetMonth(2026, 10), ca.DEFAULT_HOSPITALS, results, tmp_path / "w", log)
    finally:
        log.close()
    assert [(r.hospital_en, r.hospital_src) for r in recs] == [("KYUH", "영수증 글자"), ("KYUH", "같은 날 영수증")]
    rows = ca.read_csv_dicts(ca.list_csv_path(results, ca.TargetMonth(2026, 10)))
    assert [(r["병원 영문"], r["병원 근거"], r["확인 필요"]) for r in rows] == \
        [("KYUH", "영수증 글자", ""), ("KYUH", "같은 날 영수증", "")]
    assert ca.texts_for(recs[1], cfg)["Comment"] == "Cafe for wifi in KYUH"


# ---------- Concur 가게 이름 ----------

HOSPITALS = [["원광대", "Wonkwang University Hospital"], ["대전성모", "Daejeon St. Mary's Hospital"],
             ["남대", "CNUH"], ["전북대", "Jeonbuk University Hospital"]]


@pytest.mark.parametrize("vendor, en", [
    ("WON GWANG DAE HAK GYO BYUNG WON", "Wonkwang University Hospital"),
    ("WONGWANGDAEHAKGYO BYEONGWON PARKING", "Wonkwang University Hospital"),
    ("DAEJEON SUNGMO HOSPITAL", "Daejeon St. Mary's Hospital"),
    ("CHONBUK DAE HAK GYO HOSPITAL", "Jeonbuk University Hospital"),
    ("STARBUCKS KOREA", ""),
    ("JEON NAM DAE", ""),                 # '남대'처럼 짧은 글자로는 찾지 않음
    ("", ""),
])
def test_hospital_from_merchant(vendor, en):
    assert ca.hospital_from_merchant(vendor, HOSPITALS)[1] == en


def test_vendor_hospital_priority():
    r = rec("a.jpg", "2026-11-03", "카페")
    assert ca.apply_vendor_hospital(r, "WON GWANG DAE HAK GYO BYUNG WON", HOSPITALS)
    assert (r.hospital_en, r.hospital_src) == ("Wonkwang University Hospital", "Concur 가게 이름")

    same_day = rec("b.jpg", "2026-11-03", "카페", "Hospital A", "같은 날 영수증")
    assert ca.apply_vendor_hospital(same_day, "WON GWANG DAE HAK GYO BYUNG WON", HOSPITALS)   # 가게 이름이 더 정확
    assert same_day.hospital_en == "Wonkwang University Hospital"

    from_ocr = rec("c.jpg", "2026-11-03", "주차", "Hospital A", "영수증 글자")
    assert not ca.apply_vendor_hospital(from_ocr, "WON GWANG DAE HAK GYO BYUNG WON", HOSPITALS)
    assert from_ocr.hospital_en == "Hospital A"


def test_parse_entry_rows():
    rows = [["", "Undefined", "WON GWANG DAE HAK GYO BYUNG WON", "11/03/2026", "KRW 3,100"],
            ["Parking & Toll", "STARBUCKS", "11/04/2026", "4,500"],
            ["모두 선택"]]
    assert ca.parse_entry_rows(rows) == [
        ("2026-11-03", 3100, "Undefined WON GWANG DAE HAK GYO BYUNG WON 11/03/2026 KRW 3,100"),
        ("2026-11-04", 4500, "Parking & Toll STARBUCKS 11/04/2026 4,500"),
    ]


sync_api = pytest.importorskip("playwright.sync_api")


@pytest.fixture
def page():
    exe = os.environ.get("PW_CHROMIUM") or "/opt/pw-browsers/chromium"
    with sync_api.sync_playwright() as p:
        b = p.chromium.launch(executable_path=exe if Path(exe).exists() else None)
        yield b.new_page()
        b.close()


def make_bot(page):
    bot = ca.ConcurBot(ca.default_settings(), lambda s="": None)
    bot.page = page
    return bot


def test_read_vendor_from_field_or_label(page):
    bot = make_bot(page)
    page.set_content('<input id="vendorDescription" value="  WON GWANG  DAE HAK GYO ">')
    assert bot.read_vendor() == "WON GWANG DAE HAK GYO"
    page.set_content('<label for="v">공급업체 이름</label><input id="v" value="DAEJEON SUNGMO">')
    assert bot.read_vendor() == "DAEJEON SUNGMO"
    page.set_content('<p>가게 칸 없음</p>')
    assert bot.read_vendor() == ""


def test_vendor_from_report_list_when_field_missing(page):
    bot = make_bot(page)
    page.set_content("""<div class="entries-list">
      <div role="row"><input type="checkbox"><span>모두 선택</span></div>
      <div role="row"><input type="checkbox"><span>Undefined</span><span>WON GWANG DAE HAK GYO BYUNG WON</span>
        <span>11/03/2026</span><span>KRW 3,100</span></div>
      <div role="row"><input type="checkbox"><span>STARBUCKS</span><span>11/04/2026</span><span>KRW 4,500</span></div>
    </div>""")
    bot.entries = ca.parse_entry_rows(page.evaluate(ca.JS_READ_ROWS))
    assert len(bot.entries) == 2
    text = bot.vendor_text("2026-11-03", 3100)
    assert "WON GWANG" in text
    assert ca.hospital_from_merchant(text, HOSPITALS)[1] == "Wonkwang University Hospital"
