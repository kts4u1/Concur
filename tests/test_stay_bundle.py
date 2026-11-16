# -*- coding: utf-8 -*-
"""숙박 묶음(2026-11-15 사용자 설명): 숙박 경비에는 숙박 영수증 → 인보이스 → 숙박을 승인받은 메일 사진을 차례로 올림
(없는 것은 빼고). 인보이스는 영어일 수도 있고, 승인 메일은 메일에 적힌 숙박 날짜로 그날 숙박 영수증에 붙임.
톨게이트 사용내역의 작은 '총액'을 못 읽으면(표의 한 줄 금액 8,000원으로 읽음) 크게 다시 읽고, 그래도 못 읽으면 '확인 필요'.
(글자는 모두 지어낸 값)"""

import os
from datetime import datetime
from pathlib import Path

import concur_auto as ca

OCT = ca.TargetMonth(2026, 10)
T = datetime(2026, 11, 15, 10, 0)

STAY_SLIP = "가나모텔 카드 영수증\n사업자번호 123-45-67891\n카드 결제 감사합니다"          # 날짜·금액을 못 읽은 숙박 영수증
INVOICE = """가나 1호점
인보이스
상호 가나 1호점 견적일자 2026년 10월 20일
주소 대전광역시 유성구 가나로 1
업태 숙박업 종목 호텔업
순번 입실 일자 퇴실 일자 객실 타입 수량 단가 결제 금액
1 2026년 10월 20일 2026년 10월 21일 스탠다드 1 ₩75,000 ₩75,000
총계 ₩75,000"""
INVOICE_EN = """GRAND HOTEL JEONJU
INVOICE
Guest Name: Staff Room No: 1203
Arrival: 06-Oct-26 Departure: 07-Oct-26
Room Charge KRW 90,000
Total KRW 90,000"""
MAIL = """Re: 가나다병원 장비 설치 관련 숙박 승인요청
Manager
받는 사람 Engineer 참조 Staff
2026-09-30 (수) 오후 3:15
Dear all, Approved Thank you!
보낸 사람: Engineer <engineer@example.com>
보낸 날짜: 수요일, 9 월 30, 2026 15:06
받는 사람: Manager <manager@example.com>
제목: 가나다병원 장비 설치 관련 숙박 승인요청
안녕하십니까. 가나다병원 장비 설치 일정으로 10 월 6 일 7 일 8 일 홍길동님과 제가 전주에서 숙박 이후 설치 일정을
이어 나가려고 합니다. 이에 숙박 일정의 승인을 요청 드립니다. 2명"""


def make(name, text, taken=T):
    return ca.analyze_receipt(Path(name), Path(name), text, ca.default_settings(), [], taken)


def placed(*recs):
    for r in recs:
        ca.place_in_month(r, OCT)
    return list(recs)


def test_recognize_invoice_and_mail():
    assert ca.is_stay_invoice(INVOICE) and ca.is_stay_invoice(INVOICE_EN)
    assert not ca.is_stay_invoice(STAY_SLIP) and not ca.is_stay_invoice("스타벅스 2026.10.07 합계 4,500")
    assert ca.stay_approval_dates(MAIL) == ["2026-10-06", "2026-10-07", "2026-10-08"]     # 보낸 날짜 9-30은 뺌
    assert ca.stay_approval_dates("From: Engineer\nSubject: Hotel approval\nApproved. Stay Oct 6-8.\nSent: Sep 30, 2026") == [
        "2026-10-06", "2026-10-07", "2026-10-08"]
    assert ca.stay_approval_dates(STAY_SLIP + " 승인번호 1234") is None and ca.stay_approval_dates(INVOICE) is None
    assert make("IMG_1.JPG", INVOICE_EN).ocr_dates == ["2026-10-06", "2026-10-07"]           # 영어 날짜(06-Oct-26)


def test_invoice_fills_unread_stay_receipt():
    slip, inv = placed(make("IMG_5646.HEIC", STAY_SLIP), make("IMG_5647.HEIC", INVOICE, datetime(2026, 11, 15, 10, 1)))
    assert not slip.use and slip.category == "숙박"                                         # 날짜·금액을 못 읽어 빠졌던 영수증
    assert ca.group_stay_bundles([slip, inv], OCT, ca.default_settings()) == 1
    assert slip.use and slip.main_date == "2026-10-20" and slip.total == 75000              # 인보이스에서 읽은 날짜·금액
    assert slip.city.startswith("Daejeon") and slip.stay_in == "2026-10-20"
    assert slip.vendor == "가나 1호점" and inv.category == "숙박"                          # Vendor Name은 인보이스의 호텔 이름
    assert [p.name for p in slip.bundle] == ["IMG_5647.HEIC"] and slip.bundle_note == "인보이스 IMG_5647.HEIC 함께 첨부"
    assert not inv.use and inv.skip_reason == "숙박 영수증(IMG_5646.HEIC)에 함께 첨부"
    start, end, note = ca.stay_dates(slip, "2026-10-21", 6)
    assert (start.isoformat(), end.isoformat(), note) == ("2026-10-20", "2026-10-21", "")    # Date Range = 인보이스 입실일부터
    assert ca.find_receipt([slip, inv], "2026-10-21", 75000) is slip                        # 카드 거래일이 다음 날이어도


def test_mail_attached_to_each_night():
    night6 = make("IMG_0006.JPG", "가나호텔 2026-10-06 21:10 합계 80,000")
    night7 = make("IMG_0007.JPG", "가나호텔 2026-10-08 01:30 합계 80,000")             # 야근 뒤 새벽 영수증 = 7일 숙박
    other = make("IMG_0020.JPG", "다라모텔 2026-10-20 22:00 합계 70,000")
    inv = make("IMG_0008.JPG", INVOICE_EN.replace("90,000", "80,000"))
    mail = make("IMG_0009.PNG", MAIL, datetime(2026, 9, 30, 15, 20))
    rs = placed(night6, night7, other, inv, mail)
    assert ca.group_stay_bundles(rs, OCT, ca.default_settings()) == 2
    assert [p.name for p in night6.bundle] == ["IMG_0008.JPG", "IMG_0009.PNG"]               # 영수증 → 인보이스 → 승인 메일
    assert night6.bundle_note == "인보이스 IMG_0008.JPG, 승인 메일 IMG_0009.PNG 함께 첨부"
    assert [p.name for p in night7.bundle] == ["IMG_0009.PNG"] and other.bundle == []
    assert not mail.use and mail.skip_reason == "숙박 승인 메일 → 숙박 영수증(IMG_0006.JPG, IMG_0007.JPG)에 함께 첨부"


def test_invoice_alone_is_the_stay_receipt():
    inv, mail = placed(make("IMG_0008.JPG", INVOICE_EN), make("IMG_0009.PNG", MAIL))
    ca.group_stay_bundles([inv, mail], OCT, ca.default_settings())
    assert inv.use and inv.category == "숙박" and inv.total == 90000 and inv.stay_in == "2026-10-06"
    assert [p.name for p in inv.bundle] == ["IMG_0009.PNG"]


def test_other_receipts_untouched():
    cafe, slip = placed(make("IMG_1.JPG", "스타벅스 2026.10.07 합계 4,500"), make("IMG_2.JPG", "가나호텔 2026-10-07 22:00 합계 80,000"))
    assert ca.group_stay_bundles([cafe, slip], OCT, ca.default_settings()) == 0
    assert cafe.use and slip.use and not cafe.bundle and not slip.bundle


# ---------- 영수증 목록: 그 전 달에 찍은 승인 메일도 ----------
class FakeOcr:
    texts = {}

    def __init__(self, work_dir):
        pass

    def read(self, path):
        return self.texts[Path(path).name]

    def read_zoomed(self, path):
        return self.texts.get(Path(path).name + "#크게", "")


def photos(tmp_path, texts):
    src = tmp_path / "src"
    src.mkdir()
    for name in texts:
        if "#" not in name:
            (src / name).write_bytes(b"\xff\xd8\xff\xd9" + name.encode())
    return src


def build(tmp_path, monkeypatch, texts, tm=OCT):
    FakeOcr.texts = texts
    monkeypatch.setattr(ca, "WindowsOcr", FakeOcr)
    cfg = ca.default_settings()
    cfg.update({"영수증_원본_폴더": str(photos(tmp_path, texts)), "영수증_보정_사용": False})
    results = tmp_path / "결과"
    results.mkdir()
    lines = []
    recs = ca.build_receipts(cfg, tm, [], results, tmp_path / "w", lines.append)
    return {r.original.name: r for r in recs}, results, lines


def test_receipt_list_with_early_approval_mail(tmp_path, monkeypatch):
    texts = {"2026-10-06_211000.jpg": "가나호텔 2026-10-06 21:10 합계 80,000",
             "2026-09-30_152000.png": MAIL,                                        # 10월 전에 받은 승인 메일
             "2026-09-28_120000.jpg": "가족 사진"}                                  # 10월 전 다른 사진은 목록에 없음
    got, results, lines = build(tmp_path, monkeypatch, texts)
    assert set(got) == {"2026-10-06_211000.jpg", "2026-09-30_152000.png"}
    assert [p.name for p in got["2026-10-06_211000.jpg"].bundle] == ["2026-09-30_152000.png"]
    rows = {r["파일"]: r for r in ca.read_csv_dicts(ca.list_csv_path(results, OCT))}
    assert rows["2026-10-06_211000.jpg"]["함께 첨부"] == "2026-09-30_152000.png | 승인 메일 2026-09-30_152000.png 함께 첨부"
    assert "함께 첨부" in rows["2026-09-30_152000.png"]["확인 필요"]
    assert ca.review_status(rows["2026-09-30_152000.png"]) == "✓"
    assert [r["파일"] for r in ca.read_csv_dicts(ca.list_csv_path(results, OCT))] == [
        "2026-09-30_152000.png", "2026-10-06_211000.jpg"]                                # 함께 붙는 사진도 이 달 줄 사이에(날짜 순)
    assert any("숙박 승인 메일만 찾을 그 전 20일 사진 2장" in l for l in lines)


# ---------- 톨게이트 사용내역 총액 ----------
from test_toll_bundle import SHEET1, SHEET2, STATEMENT       # noqa: E402

SMALL = STATEMENT.replace("(총 13건 / 총액 61,500원)\n", "")       # 작은 글자 '총액'을 못 읽음 → 가장 큰 금액이 표의 한 줄


def test_statement_total_unread_is_flagged():
    stmt = make("IMG_5648.HEIC", SMALL)
    placed(stmt)
    ca.group_toll_bundles([stmt], OCT, {"톨게이트": ca.TOLL_DEFAULT_WORDS}, ["톨게이트"])
    assert stmt.total == 4700 and stmt.bundle_issue == ca.TOLL_TOTAL_UNREAD
    assert ca.TOLL_TOTAL_UNREAD in ca.receipt_issues(stmt, [stmt])
    stmt.edited.append("금액")                                                     # 영수증 확인 창에서 금액을 적으면
    assert ca.TOLL_TOTAL_UNREAD not in ca.receipt_issues(stmt, [stmt])


def test_statement_total_from_sheet_footer():
    stmt, s1, s2 = placed(make("IMG_5648.HEIC", SMALL), make("IMG_5651.HEIC", SHEET1), make("IMG_5652.HEIC", SHEET2))
    ca.group_toll_bundles([stmt, s1, s2], OCT, {"톨게이트": ca.TOLL_DEFAULT_WORDS}, ["톨게이트"])
    assert stmt.total == 61500 and stmt.bundle_issue == ""                          # 인쇄물 맨 아래 '총 13 건/61500 원'
    assert s1.total == 0 and s2.total == 0


def test_statement_total_read_with_zoom(tmp_path, monkeypatch):
    texts = {"IMG_5648.JPG": SMALL, "IMG_5648.JPG#크게": STATEMENT}                # 크게 다시 읽으면 '총액'이 읽힘
    got, _, lines = build(tmp_path, monkeypatch, texts)
    assert got["IMG_5648.JPG"].total == 61500 and got["IMG_5648.JPG"].bundle_issue == ""
    assert any("크게 다시 읽은 사진 1장, 그중 더 잘 읽힌 것 1장" in l for l in lines)


# ---------- 인터넷으로 예약한 숙박의 카드전표(가맹점 = 예약 회사) ----------
NOL_SLIP = """카드승인전표 인터넷 재발급용
카드번호 ****-123456-*1234 거래일시 2026.10.20 16:39:32
승인번호 00000001
가맹점명 주_놀유니버스 사업자등록번호 123-45-67891
가맹점주소 서울 구로구 디지털로 1
금액(AMOUNT) 68,181
부가세(VAT) 6,819
합계(TOTAL) 75,000"""
INVOICE_UNREAD = "가나 1호점\n주소 대전광역시 유성구 가나로 1\n업태 숙박업\n입실 일자 퇴실 일자 객실 타입"   # 제목·날짜·금액을 못 읽음


def test_booking_company_slip_is_stay():
    slip = make("IMG_5646.HEIC", NOL_SLIP)
    assert slip.category == "숙박" and slip.main_date == "2026-10-20" and slip.total == 75000
    assert ca.category_from_vendor("NOL UNIVERSE CO., LTD", ca.default_settings()) == "숙박"   # Concur 가맹점 이름으로도
    inv = make("IMG_5647.HEIC", INVOICE_UNREAD, datetime(2026, 11, 15, 10, 1))
    assert ca.is_stay_invoice(INVOICE_UNREAD)                                               # 입실·퇴실 칸이 있음
    placed(slip, inv)
    assert ca.group_stay_bundles([slip, inv], OCT, ca.default_settings()) == 1
    assert [p.name for p in slip.bundle] == ["IMG_5647.HEIC"] and not inv.use
    assert inv.main_date == "2026-10-20"                                                    # 날짜를 못 읽은 인보이스는 영수증 날짜로(11-15 아님)
    assert slip.vendor == "놀유니버스"                                                       # 인보이스에서 상호를 못 읽으면 전표의 가맹점명
    assert slip.city.startswith("Daejeon")                                                  # 예약 회사 주소(서울)가 아니라 호텔
    start, end, _ = ca.stay_dates(slip, "2026-10-20", 6)
    assert (start.isoformat(), end.isoformat()) == ("2026-10-20", "2026-10-21")             # 거래 시각 16:39 → 그날부터 1박


def test_stay_item_gets_booking_words(tmp_path):
    settings, table = tmp_path / ca.SETTINGS_FILE_NAME, tmp_path / ca.CATEGORY_FILE_NAME
    cfg = ca.default_settings()
    cfg["_기본_찾을_글자_판"] = 6
    ca.save_settings(settings, cfg)
    ca.save_category_rows(table, [["숙박", "모텔, 호텔, 야놀자", "Accommodation fee", "For service", "Accommodation", ""]])
    ca.upgrade_category_table(table, settings)
    words = ca.load_category_rows(table, ca.default_settings())[0][ca.C_KEYWORDS]
    assert words.startswith("모텔, 호텔, 야놀자, 놀유니버스") and "agoda" in words


def test_unreadable_stay_photo_goes_with_receipt():
    """제목·날짜·금액을 하나도 못 읽은 숙박 사진(인보이스 등)은 한 시간 안에 찍은 숙박 영수증에 붙임"""
    slip = make("IMG_5646.HEIC", NOL_SLIP)
    blur = make("IMG_5647.HEIC", "가나 1호점\n업태 숙박업", datetime(2026, 11, 15, 10, 2))
    far = make("IMG_5700.HEIC", "다라 호텔", datetime(2026, 11, 20, 10, 0))
    placed(slip, blur, far)
    assert ca.group_stay_bundles([slip, blur, far], OCT, ca.default_settings()) == 1
    assert [p.name for p in slip.bundle] == ["IMG_5647.HEIC"] and slip.bundle_note == "숙박 사진 IMG_5647.HEIC 함께 첨부"
    assert not blur.use and blur.main_date == "2026-10-20" and "IMG_5646.HEIC" in blur.skip_reason
    assert not far.bundled_into


def test_vendor_name_in_english():
    """Vendor Name은 영어로, 인보이스 맨 위 큰 제목(호텔 이름)을 먼저 씀 — 작은 '상 호' 칸은 잘못 읽히기 쉬움(실제 화면)"""
    garbled = "가나 쉼터 1호점\n인보이스\n상 호 죠q까최하녹 규섬 간력일자 2026년 10% 20일\n업태 숙박업 입실 퇴실"
    assert ca.invoice_title(garbled) == "가나 쉼터 1호점"
    assert ca.find_vendor_name("상 호 죠q까최하녹 규섬 간력일자 2026년 10% 20일") == ""              # 잘못 읽은 글자는 안 씀
    assert ca.find_vendor_name("상 호 가나 쉼터 1호점 견적일자 2026년 10월 20일") == "가나 쉼터 1호점"
    assert ca.vendor_english("가나 쉼터 1호점") == "Gana Swimteo No.1"
    assert ca.vendor_english("놀유니버스") == "NOL Universe" and ca.vendor_english("다라모텔") == "Dara Motel"
    assert ca.vendor_english("GRAND HOTEL JEONJU") == "GRAND HOTEL JEONJU"
    slip = make("IMG_5646.HEIC", NOL_SLIP)
    inv = make("IMG_5647.HEIC", garbled, datetime(2026, 11, 15, 10, 1))
    placed(slip, inv)
    ca.group_stay_bundles([slip, inv], OCT, ca.default_settings())
    assert slip.vendor == "가나 쉼터 1호점"


# ---------- 도시: 크게 읽은 글자에만 주소가 있는 인보이스(2026-11-16 숙박 도시가 Seoul) ----------
def test_city_from_zoomed_read_even_when_not_chosen(tmp_path, monkeypatch):
    inv_small = "가나 쉼터 1호점\n인보이스\n입실 일자 퇴실 일자 객실 타입"
    texts = {"2026-10-20_164000.jpg": NOL_SLIP, "2026-10-20_170000.jpg": inv_small,
             "2026-10-20_170000.jpg#크게": inv_small + "\n주소 대전광역시 유성구 가나로 1"}   # 날짜·금액은 똑같이 못 읽음
    got, _, _ = build(tmp_path, monkeypatch, texts)
    slip = got["2026-10-20_164000.jpg"]
    assert [p.name for p in slip.bundle] == ["2026-10-20_170000.jpg"]
    assert slip.city.startswith("Daejeon")                       # 예약 회사 주소(서울)가 아니라 크게 읽은 인보이스 주소


def test_booking_slip_ignores_remembered_city():
    """예약 회사 전표는 사업자번호가 늘 같아서, 한 번 고친 도시를 기억해 쓰면 다른 숙소도 그 도시가 됨"""
    slip = make("IMG_5646.HEIC", NOL_SLIP)
    memory = {slip.biz_no: {"분류": "숙박", "도시": "Daejeon (ex Taejon), KOREA, REPUBLIC OF", "병원 영문": "Gana Hospital"}}
    assert slip.biz_no and ca.apply_store_memory(slip, memory, ca.default_settings())
    assert slip.category == "숙박" and slip.city.startswith("Seoul") and slip.hospital_en == ""


def test_city_from_district_or_stay_name():
    """주소에 광역시 이름 없이 구 이름만 있거나(유성구 → 대전), 주소를 못 읽어도 숙소 이름의 동네(… 유성 1호점 → 대전)"""
    c = ca.DEFAULT_CITIES
    assert ca.detect_city(ca.compact("주소 유성구 봉명동 1"), c).startswith("Daejeon")
    assert ca.detect_city(ca.compact("서울 구로구 디지털로 1"), c).startswith("Seoul")      # 중구·구로구 같은 구는 그대로
    slip = make("IMG_5646.HEIC", NOL_SLIP)
    inv = make("IMG_5647.HEIC", "가나 쉼터 유성 1호점\n인보이스\n입실 일자 퇴실 일자 객실 타입", datetime(2026, 11, 15, 10, 1))
    placed(slip, inv)
    assert ca.group_stay_bundles([slip, inv], OCT, ca.default_settings()) == 1
    assert slip.vendor == "가나 쉼터 유성 1호점" and slip.city.startswith("Daejeon")
