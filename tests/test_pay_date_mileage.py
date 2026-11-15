# -*- coding: utf-8 -*-
"""승인일시·마일리지(2026-11-15 사용자 화면·설명)
- 주차 정기권 영수증: 사용시작·사용종료(다음 달)보다 [승인일시]가 영수증 날짜. 도시는 본사 주소(서울)가 아니라 주차장 이름(광주…).
- 마일리지: 차량 유지비 보조금 신청서(한 달 운행 거리 → 금액)가 경비 하나, 이 달 1일·다음 달 1일 계기판 사진을 인쇄해
  찍은 사진을 함께 올림. 카드 경비가 아니라 새 경비로 만듦. (글자는 모두 지어낸 값)"""

import os
from datetime import datetime
from pathlib import Path

import pytest

import concur_auto as ca

OCT, NOV, DEC = ca.TargetMonth(2026, 10), ca.TargetMonth(2026, 11), ca.TargetMonth(2026, 12)

PASS = """영수증
* 가맹점정보
상 호: (주)가나파킹
사업자번호: 123-45-67890
주 소: 서울특별시 중구 세종대로 1, 10층
* 상품정보
주차장명: 광주가나주차타워점
상 품 명: 유료정기권(99,000)
사용시작: 2026-12-01 00:00
사용종료: 2026-12-31 23:59
* 결제정보
[승인일시] 2026-11-02 10:51:01
[공급가액] 90,000 원
[부가세액] 9,000 원
[승인금액] 99,000 원
출력일시:2026-11-02 10:51"""


def make(name, text, taken):
    return ca.analyze_receipt(Path(name), Path(name), text, ca.default_settings(), [], taken)


def test_pay_date_words():
    assert ca.find_pay_dates(PASS) == ["2026-11-02"]
    assert ca.find_pay_dates("[숭인일시]2026-11-0210:51:01") == ["2026-11-02"]       # OCR이 '승'을 '숭'으로, 시각이 붙어 읽힘
    assert ca.find_pay_dates("거래일시: 26.11.03 12:00") == ["2026-11-03"]
    assert ca.find_pay_dates("결제일 20261104") == ["2026-11-04"]
    assert ca.find_pay_dates("스타벅스 2026.10.07 합계 4,500") == []
    assert ca.find_ocr_dates("입차 2026-11-0210:51") == ["2026-11-02"]                 # 날짜 뒤에 시각이 붙어도 날짜


def test_parking_pass_dated_by_approval():
    taken = datetime(2026, 11, 15, 9, 0)                                              # 영수증을 나중에 찍음
    rec = make("IMG_0617.HEIC", PASS, taken)
    assert rec.ocr_dates[0] == "2026-11-02" and rec.main_date == "2026-11-02"
    assert rec.total == 99000 and rec.category == "주차"
    assert rec.city.startswith("Gwangju")                                              # 본사 주소 서울이 아니라 주차장 이름
    ca.place_in_month(rec, NOV)
    assert rec.use and rec.main_date == "2026-11-02"
    assert ca.find_receipt([rec], "2026-11-02", 99000) is rec
    for tm in (OCT, DEC):                                                              # 12월(사용 기간) 영수증이 아님
        r = make("IMG_0617.HEIC", PASS, taken)
        ca.place_in_month(r, tm)
        assert not r.use and r.skip_reason == ca.OTHER_MONTH_NOTE.format(date="2026-11-02")


def test_receipt_without_pay_words_unchanged():
    rec = make("IMG_0001.JPG", "스타벅스 2026.10.07 합계 4,500", datetime(2026, 10, 7, 9, 0))
    assert rec.pay_dates == [] and rec.ocr_dates == ["2026-10-07"]


# ---------- 마일리지 ----------
FORM = """Car Record Entry [차량 유지비 보조금 신청서]
User Information Home [자택] Office [사무실]
Travel Information
Date [날짜] Day of the week [요일] Purpose [목적] Travel Info [행선지] From To Total km Amount
2026-10-01 Thu service office 가나 Hospital 250 108,250
2026-10-02 Fri service office 다라 Hospital 20 8,660
2026-10-03 Sat 0
2026-10-30 Fri service office 마바 Hospital 15 6,495
2026-10-31 Sat 0
Total Driven Distance (km) 1,250
Total Km Amount (Won) 541,250
Remark"""
ODO = "2026.10.01\n0 km/h 0.8 60647 km\n2026.11.01\n0 km/h 0.9 62440 km"
WORDS = {"마일리지": ca.MILEAGE_DEFAULT_WORDS, "카페": ["스타벅스"]}
ORDER = ["마일리지", "카페"]


def test_recognize_mileage_form():
    assert ca.mileage_form_info(FORM) == {"total": 541250, "km": 1250}
    assert ca.mileage_form_info(ODO) is None and ca.mileage_form_info(PASS) is None
    assert ca.detect_category(ca.compact(FORM), {"마일리지": ["Car Record"]}, ["마일리지"]) == "마일리지"   # 띄어 적은 찾을 글자도


def mileage_recs(odo_text=ODO, odo_taken=datetime(2026, 11, 3, 18, 1)):
    t = datetime(2026, 11, 3, 18, 0)
    rs = [make("IMG_5650.HEIC", FORM, t), make("IMG_5649.HEIC", odo_text, odo_taken),
          make("IMG_0002.JPG", "스타벅스 2026.11.01 합계 4,500", datetime(2026, 11, 1, 9, 0)),
          make("IMG_0003.JPG", "스타벅스 2026.10.07 합계 4,500", datetime(2026, 10, 7, 9, 0))]
    for r in rs:
        ca.place_in_month(r, OCT)
    return rs


def test_bundle_form_with_odometer_photo():
    rs = mileage_recs()
    assert ca.group_mileage_bundles(rs, OCT, WORDS, ORDER) == 1
    form, odo, nov_cafe, cafe = rs
    assert form.use and form.new_kind == "마일리지" and form.category == "마일리지"
    assert form.total == 541250 and form.main_date == "2026-10-31" and form.last_use == "2026-10-31"   # 그 달 말일
    assert [p.name for p in form.bundle] == ["IMG_5649.HEIC"]
    assert form.bundle_note == "마일리지 신청서 1,250km 541,250원, 계기판 사진 1장 함께 첨부"
    assert ca.receipt_issues(form, rs) == []
    assert not odo.use and odo.skip_reason == "마일리지 신청서(IMG_5650.HEIC)에 함께 첨부"
    assert not nov_cafe.bundled_into and cafe.use and not cafe.bundle                  # 11월 1일 영수증은 계기판 사진이 아님
    assert ca.find_receipt(rs, "2026-10-31", 541250) is form                         # 다시 돌릴 때 만든 경비와 맞춤


def test_odometer_without_dates_uses_photo_time():
    rs = mileage_recs("0 km/h 0.8 60647 km 0.9 62440 km")                             # 날짜 글자를 못 읽음
    ca.group_mileage_bundles(rs, OCT, WORDS, ORDER)
    assert [p.name for p in rs[0].bundle] == ["IMG_5649.HEIC"]


def test_missing_odometer_photo_is_flagged():
    rs = mileage_recs("0 km/h 0.8 60647 km", datetime(2026, 11, 20, 9, 0))              # 날짜도 없고 한참 뒤에 찍음
    ca.group_mileage_bundles(rs, OCT, WORDS, ORDER)
    form = rs[0]
    assert form.bundle == [] and form.new_kind == "마일리지" and form.use
    assert ca.receipt_issues(form, rs) == ["함께 올릴 계기판(주행거리) 사진을 못 찾음"]


def test_mileage_item_gets_default_words(tmp_path):
    settings, table = tmp_path / ca.SETTINGS_FILE_NAME, tmp_path / ca.CATEGORY_FILE_NAME
    cfg = ca.default_settings()
    cfg["_기본_찾을_글자_판"] = 5
    ca.save_settings(settings, cfg)
    ca.save_category_rows(table, [["마일리지", "마일리지", "Car mileage", "For service", "Car mileage", ""]])
    ca.upgrade_category_table(table, settings)
    words = ca.load_category_rows(table, ca.default_settings())[0][ca.C_KEYWORDS]
    assert words.startswith("마일리지, ") and "Car Record" in words and words.count("마일리지") == 1


# ---------- 영수증 목록·② 전체 실행 ----------
TEXTS = {"IMG_5650.JPG": FORM, "IMG_5649.JPG": ODO}
MILEAGE = ["마일리지", ", ".join(ca.MILEAGE_DEFAULT_WORDS), "Car mileage for service", "For service", "Car mileage",
           "", "", "", "", ""]


class FakeOcr:
    texts = TEXTS

    def __init__(self, work_dir):
        pass

    def read(self, path):
        return self.texts[Path(path).name]

    def read_zoomed(self, path):
        return ""


def photo_folder(tmp_path, names):
    src = tmp_path / "src"
    src.mkdir()
    for name in names:
        p = src / name
        p.write_bytes(b"\xff\xd8\xff\xd9" + name.encode())
        t = datetime(2026, 11, 3, 18, 0).timestamp()
        os.utime(p, (t, t))
    return src


def test_receipt_list_shows_mileage_bundle(tmp_path, monkeypatch):
    src = photo_folder(tmp_path, TEXTS)
    monkeypatch.setattr(ca, "WindowsOcr", FakeOcr)
    cfg = ca.default_settings()
    table = tmp_path / ca.CATEGORY_FILE_NAME
    ca.save_category_rows(table, ca.category_rows_from_cfg(cfg) + [MILEAGE])
    ca.apply_category_table(cfg, table)
    cfg.update({"영수증_원본_폴더": str(src), "영수증_보정_사용": False})
    results = tmp_path / "결과"
    results.mkdir()
    lines = []
    got = {r.original.name: r for r in ca.build_receipts(cfg, OCT, [], results, tmp_path / "w", lines.append)}
    assert got["IMG_5650.JPG"].use and got["IMG_5650.JPG"].category == "마일리지"
    assert not got["IMG_5649.JPG"].use
    rows = {r["파일"]: r for r in ca.read_csv_dicts(ca.list_csv_path(results, OCT))}
    assert rows["IMG_5650.JPG"]["함께 첨부"].startswith("IMG_5649.JPG | 마일리지 신청서 1,250km 541,250원")
    assert rows["IMG_5650.JPG"]["날짜"] == "2026-10-31" and rows["IMG_5650.JPG"]["금액"] == "541250"
    assert "함께 첨부" in rows["IMG_5649.JPG"]["확인 필요"]
    assert any("마일리지: IMG_5650.JPG 마일리지 신청서" in l for l in lines)


def test_full_run_creates_mileage_expense(tmp_path, monkeypatch):
    from test_toll_manual import UI, FakeBot
    src = photo_folder(tmp_path, ["IMG_5650.JPG"])                                   # 계기판 사진이 없어도 새 경비는 만듦
    base = tmp_path / "app"
    base.mkdir()
    cfg = ca.default_settings()
    cfg.update({"영수증_원본_폴더": str(src), "결과_자동_열기": False, "영수증_보정_사용": False})
    ca.save_settings(base / ca.SETTINGS_FILE_NAME, cfg)
    ca.save_category_rows(base / ca.CATEGORY_FILE_NAME, [MILEAGE])
    monkeypatch.setattr(ca, "WindowsOcr", FakeOcr)
    monkeypatch.setattr(ca, "ConcurBot", FakeBot)
    FakeBot.made = []
    ui = UI()
    ca.run_job("full", "2026-10", ui, base)
    assert FakeBot.made == [("IMG_5650.JPG", [], 541250, "2026-10-31", 1)]


def test_manual_mileage_expense(tmp_path):
    pytest.importorskip("cv2")
    sync_api = pytest.importorskip("playwright.sync_api")
    from PIL import Image
    from test_toll_manual import report_html
    html = report_html().replace("'Taxi'];", "'Taxi', 'Car mileage'];")
    form, odo = tmp_path / "IMG_5650.jpg", tmp_path / "IMG_5649.jpg"
    for f in (form, odo):
        Image.new("RGB", (300, 400), (240, 240, 240)).save(f, quality=90)
    cfg = ca.default_settings()
    table = tmp_path / ca.CATEGORY_FILE_NAME
    ca.save_category_rows(table, [MILEAGE])
    ca.apply_category_table(cfg, table)
    rec = ca.Receipt(original=form, ocr_file=form, bundle=[odo], total=541250, dates=["2026-10-31"],
                     last_use="2026-10-31", category="마일리지", new_kind="마일리지")
    exe = os.environ.get("PW_CHROMIUM") or "/opt/pw-browsers/chromium"
    with sync_api.sync_playwright() as p:
        b = p.chromium.launch(executable_path=exe if Path(exe).exists() else None)
        pg = b.new_page()
        pg.set_default_timeout(3000)
        pg.route("**/report", lambda route: route.fulfill(body=html, content_type="text/html; charset=utf-8"))
        bot = ca.ConcurBot(cfg, lambda s="": None)
        bot.page, bot.timeout_ms, bot.upload_wait_s, bot.upload_done_s = pg, 3000, 3, 3
        bot.report_url = "http://concur.test/report"
        bot.screenshot = lambda name: None
        rows = bot.create_manual_expenses([rec], tmp_path / "w")
        saved = pg.evaluate("saved")
        b.close()
    assert saved["type"] == "Car mileage" and saved["date"] == "10/31/2026"
    assert saved["amount"].replace(",", "") == "541250" and saved["comment"] == "Car mileage for service"
    assert saved["uploads"] == ["IMG_5650.jpg", "IMG_5649.jpg"]                      # 신청서 → 계기판 사진
    assert rows[0]["오류"] == "" and rows[0]["확인 필요"] == ""
