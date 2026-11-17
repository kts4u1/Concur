# -*- coding: utf-8 -*-
"""2026-11-17 사용자 화면(영수증 확인 창 네 장)에서 고친 것:
1. 병원 이름이 없는 주차 정기권 영수증에 같은 날 원광대병원 영수증의 병원이 들어감 → 같은 날 병원을 빌려 오지 않음
2. '주차영수증'(주차장명·주차권·주차시간·주차요금)이 식대로 분류됨 → 찾을 글자가 가장 많이 나온 항목
3. 'SK쉴더스 전주 수병원 주차장'의 수병원을 못 찾음(4글자 미만이라 버림) → 짧은 이름도, 영문은 Jeonju Soo Hospital
4. 1,400원 주차 영수증이 411,400원 → 점선(......)을 4·1로 읽어 붙은 금액은 뺌"""

from pathlib import Path

import concur_auto as ca


def analyze(text, hospitals=None, taken=None):
    return ca.analyze_receipt(Path("IMG_1.HEIC"), Path("IMG_1.HEIC"), text, ca.default_settings(),
                              hospitals if hospitals is not None else ca.DEFAULT_HOSPITALS, taken)


# ---------- 2. 분류 ----------

PARKING_AT_HOSPITAL = """주차영수증
상 호 : 원광대병원
주 차 장 명 : 원광대학교병원
주 소 : 전북 익산시 무왕로 985
사 업 자 : 123-45-67891
주차권 번호 : 1234567
입 차 시 간 : 2026-10-01 16:36:31
출 차 시 간 : 2026-10-01 19:21:52
주 차 시 간 : 2시간 45분 21초
주 차 요 금 : 3,100
결 제 수 단 : 현대카드
식사 후 출차하세요
발 행 일 자 : 2026-10-01"""


def test_parking_receipt_with_one_meal_word_is_parking():
    rec = analyze(PARKING_AT_HOSPITAL)
    assert rec.category == "주차"                       # 예전: 식대(위쪽 항목의 글자 '식사'가 하나라도 있으면 식대)
    assert rec.hospital_en == "Wonkwang University Hospital" and rec.total == 3100


def test_tie_keeps_table_order():
    cfg = ca.default_settings()
    text = ca.compact("한식 뷔페 주차 2시간 무료 합계 9,000")      # 식대 1번, 주차 1번 → 위쪽 항목(식대)
    assert ca.detect_category(text, cfg["키워드"], cfg["분류_우선순위"]) == "식대"
    assert ca.detect_category(ca.compact("김치찌개 한식 식당 주차권 제공"), cfg["키워드"], cfg["분류_우선순위"]) == "식대"


def test_store_memory_does_not_turn_clear_parking_into_meal():
    """병원은 주차장·식당이 한 사업자번호라, 기억한 분류(식대)보다 영수증 글자가 뚜렷하면(주차 3번 이상) 글자를 따름"""
    cfg = ca.default_settings()
    for text in (PARKING_AT_HOSPITAL, PARKING_AT_HOSPITAL.replace("식사 후 출차하세요\n", "")):
        rec = analyze(text)
        assert rec.biz_no
        ca.apply_store_memory(rec, {rec.biz_no: {"분류": "식대"}}, cfg)
        assert rec.category == "주차"

    cafe = analyze("카페 한빛\n사업자 : 123-45-67891\n아이스티 4,500\n2026-10-03")
    ca.apply_store_memory(cafe, {cafe.biz_no: {"분류": "식대"}}, cfg)
    assert cafe.category == "식대"                      # 글자가 뚜렷하지 않으면 사용자가 고른 분류를 씀


# ---------- 1. 같은 날 병원 ----------

PASS_RECEIPT = """영수증
* 가맹점정보
상 호 : (주)하이파킹
주 소 : 서울특별시 서대문구 충정로 60
* 상품정보
주차장명: 광주영효주차타워점
상 품 명: 유료정기권(99,000)
사용시작: 2026-11-01 00:00
사용종료: 2026-11-30 23:59
* 결제정보
[승인일시] 2026-10-01 09:58:48
[승인금액] 99,000 원"""


def test_pass_receipt_does_not_borrow_same_day_hospital(tmp_path, monkeypatch):
    src = tmp_path / "src"
    src.mkdir()
    texts = {"2026-10-01_163631.jpg": PARKING_AT_HOSPITAL, "2026-10-01_095848.jpg": PASS_RECEIPT}
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
    recs = ca.build_receipts(cfg, ca.TargetMonth(2026, 10), ca.DEFAULT_HOSPITALS, results, tmp_path / "w", log)
    by = {r.original.name: r for r in recs}
    assert by["2026-10-01_095848.jpg"].hospital_en == ""          # 예전: Wonkwang University Hospital(같은 날 영수증)
    assert by["2026-10-01_095848.jpg"].total == 99000 and by["2026-10-01_095848.jpg"].category == "주차"
    assert ca.city_name(by["2026-10-01_095848.jpg"].city).startswith("Gwangju")
    assert by["2026-10-01_163631.jpg"].hospital_en == "Wonkwang University Hospital"


# ---------- 3. 짧은 병원 이름 ----------

SOO = """매출전표
No.0000000001
SK쉴더스 전주 수병원 주차장
전북특별자치도 전주시 완산구 서원로 63
2026-10-07 15:03:11
출구무인정산기
신용카드 9,400 원"""


def test_short_hospital_name_after_city():
    assert ca.find_hospital_name_in_text(SOO) == "전주수병원"
    assert ca.find_hospital_name_in_text("SK쉴더스전주수병원주차장") == "전주수병원"   # 띄어쓰기 없이 읽혀도
    assert ca.find_hospital_name_in_text("의료법인 OO재단 목포한국병원") == "목포한국병원"
    rec = analyze(SOO)
    assert rec.hospital_en == "Jeonju Soo Hospital" and rec.category == "주차" and rec.total == 9400
    auto = analyze(SOO, hospitals=[])                   # 병원 목록에 없으면 자동 변환
    assert (auto.hospital_new, auto.hospital_key, auto.hospital_en) == ("전주수병원", "전주수병원", "Jeonju Su Hospital")


def test_hospital_from_other_read_when_main_read_is_broken():
    """원래 읽기에서는 병원 줄이 깨졌는데(날짜·금액은 더 잘 읽음) 크게 읽은 글자에는 병원 이름이 있음"""
    cfg = ca.default_settings()
    rec = analyze("?과之자 우길홈로 주차 2026-10-07 15:03:11 9,400 원")
    assert rec.hospital_en == ""
    ca.fill_from_other_reads(rec, [SOO], cfg, ca.DEFAULT_HOSPITALS)
    assert (rec.hospital_en, rec.hospital_src) == ("Jeonju Soo Hospital", "영수증 글자")
    assert ca.city_name(rec.city) == "Jeonju"


def test_new_default_hospital_added_once_to_old_list(tmp_path):
    csv_path = tmp_path / ca.HOSPITAL_FILE_NAME
    settings = tmp_path / ca.SETTINGS_FILE_NAME
    ca.write_csv(csv_path, ca.HOSPITAL_HEADER, [["원광대", "Wonkwang University Hospital", "영수증에서 확인"],
                                                ["목포한국", "Mokpo Hankook Hospital", "영수증에서 확인"],
                                                ["전북대", "Jeonbuk National University Hospital", "직접 고침"]])
    lines = []
    assert ca.add_new_default_hospitals(csv_path, settings, lines.append) == 1
    rows = ca.load_hospital_rows(csv_path)
    assert ["전주수병원", "Jeonju Soo Hospital", "영수증에서 확인"] in rows and len(rows) == 4
    assert any("전주수병원" in line for line in lines)
    ca.write_csv(csv_path, ca.HOSPITAL_HEADER, rows[:3])            # 사용자가 지우면 다시 넣지 않음
    assert ca.add_new_default_hospitals(csv_path, settings) == 0
    assert len(ca.load_hospital_rows(csv_path)) == 3


# ---------- 4. 점선 뒤 금액 ----------

def test_dotted_leader_misread_amount_dropped():
    text = """[ 영 수 증 ]
상 호 : 일곡공영주차장
주 소 : 광주 북구 설죽로 495-6
사업자번호 : 234-56-78901
입차 일시 : ...............2026-10-15 12:07:35
출차 일시 : ...............2026-10-15 13:58:25
주차 요금 : ................411,400원
할인 요금 : ....................0원
카드 결제 : .................1,400원"""
    rec = analyze(text)
    assert rec.total == 1400 and 411400 not in rec.amounts and rec.category == "주차"
    assert ca.find_amounts("주차요금 ......1,400원") == [1400]          # 점선 바로 뒤 금액도 읽음
    assert ca.find_amounts("합계 11,400\n물 1,400") == [11400, 1400]   # 진짜 합계는 그대로
    assert ca.find_amounts("합계 41,400\n승인금액 41,400\n주차 1,400") == [41400, 1400]   # 두 번 나온 합계도 그대로


# ---------- 하이패스 사용내역·인쇄물 금액 (2026-11-17 두 번째 화면) ----------

STATEMENT_SMALL = """기간별 사용내역
카드구분 전체 카드사 전체 카드번호 0000-00**-****-0000
조회 구분 사용일 기준 조회 기간 2026년 10월 01일 ~ 2026년 10월 31일
13 2026/10/01 16:17:55 선불 0000-00**-****-0000 1종 광주 북김제 선불Hi-pass 4,700
12 2026/10/01 20:43:23 선불 0000-00**-****-0000 1종 북김제 광주 선불Hi-pass 4,700
4 2026/10/20 15:46:33 선불 0000-00**-****-0000 1종 남광산 유성 선불Hi-pass 8,000"""


def test_statement_without_readable_total_has_no_amount():
    """총액(작은 글자)을 못 읽은 사용내역은 표의 한 줄 금액(8,000)을 쓰지 않고 금액을 비워 둠"""
    nov = __import__("datetime").datetime(2026, 11, 15)
    rec = analyze(STATEMENT_SMALL, taken=nov)
    assert rec.total == 0 and rec.amounts == []
    read = analyze(STATEMENT_SMALL.replace("기간별 사용내역", "기간별 사용내역 (총 13건 / 총액 61,500원)"), taken=nov)
    assert read.total == 61500 and read.amounts == [61500]


def test_hipass_receipt_sheet_has_no_amount_but_single_toll_receipt_does():
    sheet = ("하이패스 빠르고 편리합니다 영수증 한국도로공사 북김제영업소 2026년10월01일 16시17분 통행료 4,700원(카드)\n"
             "하이패스 빠르고 편리합니다 영수증 광주영업소 2026년10월02일 08시05분 통행료 4,300원(카드) 합계 118,100")
    rec = analyze(sheet, taken=__import__("datetime").datetime(2026, 11, 15))
    assert rec.total == 0 and rec.amounts == [] and rec.main_date == "2026-10-01"
    single = analyze("하이패스 영수증 한국도로공사 광주영업소 2026년10월02일 08시05분 통행료 4,300원(카드)")
    assert single.total == 4300                          # 톨게이트에서 받은 영수증 한 장은 금액을 씀


# ---------- 주차 금액이 너무 큼 (2026-11-17 사용자: 주차비 411,400원은 나올 리 없음) ----------

PARK_411 = """[ 영 수 증 ]
상 호 : 일곡공영주차장
주 소 : 광주 북구 설죽로 495-6
입차 일시 : 2026-10-15 12:07:35
출차 일시 : 2026-10-15 13:58:25
주차 시간 : 1시간 50분
주차 요금 : 411,400원
할인 요금 : 0원
카드 결제 : 411,400원"""


def build_with_zoom(tmp_path, monkeypatch, main, zoomed):
    import os
    from datetime import datetime
    src = tmp_path / "src"
    src.mkdir()
    p = src / "IMG_5657.JPG"
    p.write_bytes(b"\xff\xd8\xff\xd9")
    t = datetime(2026, 11, 15, 10, 0).timestamp()
    os.utime(p, (t, t))
    calls = []

    class FakeOcr:
        def __init__(self, work_dir):
            pass

        def read(self, path):
            return main

        def read_zoomed(self, path):
            calls.append(path)
            return zoomed

    monkeypatch.setattr(ca, "WindowsOcr", FakeOcr)
    cfg = ca.default_settings()
    cfg.update({"영수증_원본_폴더": str(src), "영수증_보정_사용": False})
    results = tmp_path / "결과"
    results.mkdir()
    lines = []
    recs = ca.build_receipts(cfg, ca.TargetMonth(2026, 10), [], results, tmp_path / "w", lines.append)
    return recs[0], calls, lines


def test_too_big_parking_amount_fixed_by_zoomed_read(tmp_path, monkeypatch):
    zoomed = PARK_411.replace("411,400원", "1,400원")
    rec, calls, lines = build_with_zoom(tmp_path, monkeypatch, PARK_411, zoomed)
    assert len(calls) == 1 and rec.category == "주차"
    assert rec.total == 1400 and rec.amount_issue == "" and ca.receipt_issues(rec, [rec]) == []
    assert any("바로잡은 사진 1장" in line for line in lines)


def test_too_big_parking_amount_flagged_when_still_unclear(tmp_path, monkeypatch):
    rec, calls, _ = build_with_zoom(tmp_path, monkeypatch, PARK_411, PARK_411)    # 크게 읽어도 411,400
    assert len(calls) == 1 and rec.total == 411400
    assert any("주차 금액이 너무 큼(411,400원)" in i for i in ca.receipt_issues(rec, [rec]))
    rec.edited.append("금액")                                                     # 영수증 확인 창에서 고치면 사라짐
    assert not any("너무 큼" in i for i in ca.receipt_issues(rec, [rec]))


def test_normal_amounts_are_not_read_again(tmp_path, monkeypatch):
    rec, calls, _ = build_with_zoom(tmp_path, monkeypatch, PARK_411.replace("411,400원", "1,400원"), "")
    assert calls == [] and rec.total == 1400
    meal = analyze("한식 식당 2026-10-15 합계 412,000원")                        # 식대 큰 금액은 의심하지 않음
    assert not ca.amount_suspect(meal, ca.default_settings())


# ---------- 점선 지우고 다시 읽기 (2026-11-17: 크게 읽어도 11,400원) ----------

def test_remove_dot_leaders_keeps_commas_and_digits():
    import pytest
    cv2 = pytest.importorskip("cv2")
    np = pytest.importorskip("numpy")
    from PIL import Image, ImageDraw, ImageFont
    font = ImageFont.load_default(size=40)
    lines = ["FEE : ...................... 1,400", "CARD : ................. 1,400", "TIME 12:07:35  No. 1.5",
             "- - - - - - - - - - - - - - - -"]
    for paper, tilt in ((255, 0), (205, 1.2)):           # 흰 종이 / 회색 종이에 조금 기울어진 사진(실제 사진처럼)
        img = Image.new("RGB", (1400, 360), (paper,) * 3)
        d = ImageDraw.Draw(img)
        for k, t in enumerate(lines):
            d.text((30, 20 + k * 80), t, fill="black", font=font)
        img = img.rotate(tilt, fillcolor=(paper,) * 3, resample=Image.BICUBIC)
        arr = cv2.cvtColor(np.asarray(img), cv2.COLOR_RGB2BGR)
        out, n = ca.remove_dot_leaders(arr)
        assert n >= 35
        dark = lambda a: int((a < 110).sum())
        for k in (0, 1):
            y0, y1 = 10 + k * 80, 90 + k * 80
            x_amount = int(30 + d.textlength(lines[k][:-5], font=font))
            x_dots0, x_dots1 = int(30 + d.textlength(lines[k][:7], font=font)), int(30 + d.textlength(lines[k][:-6], font=font))
            assert dark(out[y0:y1, x_amount:]) >= dark(arr[y0:y1, x_amount:]) * 0.95   # 1,400(쉼표 포함)은 남김
            assert dark(out[y0:y1, x_dots0:x_dots1]) <= dark(arr[y0:y1, x_dots0:x_dots1]) * 0.05   # 점선은 지움
        assert dark(out[170:350]) >= dark(arr[170:350]) * 0.95                       # 시각·소수점·줄표는 그대로
    assert ca.join_spaced_amount("주차 요금 1 400원") == "주차 요금 1,400원"


def test_parking_amount_fixed_by_dotless_read(tmp_path, monkeypatch):
    """크게 읽기는 11,400(점 하나를 1로), 점선을 지운 사진은 1,400 두 번 → 1,400"""
    zoomed = PARK_411.replace("411,400원", "11,400원")
    dotless = PARK_411.replace("411,400원", "1,400원")
    fake = tmp_path / "dotless_IMG_5657.jpg"
    fake.write_bytes(b"x")
    monkeypatch.setattr(ca, "make_dotless_copy", lambda original, work_dir, log=None: fake)
    import os
    from datetime import datetime
    src = tmp_path / "src"
    src.mkdir()
    p = src / "IMG_5657.JPG"
    p.write_bytes(b"\xff\xd8\xff\xd9")
    t = datetime(2026, 11, 15, 10, 0).timestamp()
    os.utime(p, (t, t))

    class FakeOcr:
        def __init__(self, work_dir):
            pass

        def read(self, path):
            return dotless if "dotless" in Path(path).name else PARK_411

        def read_zoomed(self, path):
            return zoomed

    monkeypatch.setattr(ca, "WindowsOcr", FakeOcr)
    cfg = ca.default_settings()
    cfg.update({"영수증_원본_폴더": str(src), "영수증_보정_사용": False})
    results = tmp_path / "결과"
    results.mkdir()
    rec = ca.build_receipts(cfg, ca.TargetMonth(2026, 10), [], results, tmp_path / "w", lambda s: None)[0]
    assert rec.total == 1400 and rec.amount_issue == ""
    assert not fake.exists()                                                       # 임시 사본은 지움


# ---------- 계기판 사진은 위쪽 날짜의 달 (2026-11-17 사용자: "윗쪽의 날짜를 인식해서 그 달로") ----------

def mileage_pair(tmp_path, monkeypatch, month):
    import os
    from datetime import datetime
    src = tmp_path / "src"
    src.mkdir()
    texts = {"IMG_5649.JPG": "2026.10.01 0 km 0.8 2026.11.01 0 km 0.9",                      # 10월 계기판 사진
             "IMG_5660.JPG": "Car Record Entry [차량 유지비 보조금 신청서]\n2026-09-01 Tue service 250 108,250\n"
                             "2026-09-30 Wed service 20 8,660\nTotal Driven Distance (km) 2,494\n"
                             "Total Km Amount (Won) 1,079,730"}                              # 9월 신청서
    for name in texts:
        p = src / name
        p.write_bytes(b"\xff\xd8\xff\xd9" + name.encode())
        t = datetime(2026, 11, 15, 10, 0).timestamp()                              # 같은 날 찍음
        os.utime(p, (t, t))

    class FakeOcr:
        def __init__(self, work_dir):
            pass

        def read(self, path):
            return texts[Path(path).name]

        def read_zoomed(self, path):
            return ""

    monkeypatch.setattr(ca, "WindowsOcr", FakeOcr)
    cfg = ca.default_settings()
    cfg.update({"영수증_원본_폴더": str(src), "영수증_보정_사용": False})
    results = tmp_path / "결과"
    results.mkdir()
    recs = ca.build_receipts(cfg, ca.TargetMonth(2026, month), [], results, tmp_path / "w", lambda s: None)
    return {r.original.name: r for r in recs}


def test_october_dashboard_is_not_attached_to_september_form(tmp_path, monkeypatch):
    got = mileage_pair(tmp_path, monkeypatch, 9)
    dash, form = got["IMG_5649.JPG"], got["IMG_5660.JPG"]
    assert dash.bundled_into == "" and not dash.use and dash.main_date == "2026-10-01"      # 예전: 9월 신청서에 붙어 ✓
    assert dash.skip_reason == ca.OTHER_MONTH_NOTE.format(date="2026-10-01")
    assert ca.review_status({"사용(O/X)": "X", "확인 필요": dash.skip_reason, "붙는 영수증": ""}) .startswith("✕")
    assert form.use and form.bundle == []
    assert dash.category == "마일리지"                                             # 분류가 비지 않게(사용자 요청)


def test_dashboard_main_date_is_upper_date(tmp_path, monkeypatch):
    rec = analyze("2026.10.01 0 km 0.8 2026.11.01 0 km 0.9")
    assert ca.dashboard_month(rec) == "2026-10"
    ca.place_dashboards([rec], ca.TargetMonth(2026, 11))
    assert rec.main_date == "2026-10-01" and not rec.use                           # 11월에는 10월 사진
