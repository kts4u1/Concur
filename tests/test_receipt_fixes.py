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
