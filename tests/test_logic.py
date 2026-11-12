# -*- coding: utf-8 -*-
"""Windows 없이 돌릴 수 있는 부분(날짜·금액·분류·병원 추출, 영수증 목록, 매칭) 테스트"""

import json
from datetime import date
from pathlib import Path

import pytest

import concur_auto as ca


# ---------- 처리할 달 ----------

def test_default_month_is_last_month():
    assert ca.resolve_target_month(None, date(2026, 11, 12)) == ca.TargetMonth(2026, 10)
    assert ca.resolve_target_month(None, date(2027, 1, 5)) == ca.TargetMonth(2026, 12)


def test_month_text():
    assert ca.resolve_target_month("2026-3") == ca.TargetMonth(2026, 3)
    for bad in ("2026-13", "10월", "2026/10"):
        with pytest.raises(RuntimeError):
            ca.resolve_target_month(bad)


def test_target_month_properties():
    tm = ca.TargetMonth(2028, 2)
    assert tm.prefix == "2028-02"
    assert tm.month_name == "February"
    assert tm.first_day_us == "02/01/2028"
    assert tm.last_day_us == "02/29/2028"
    assert tm.report_name("{month_name} Expenses") == "February Expenses"


def test_recent_months_cross_year():
    assert ca.recent_months(3, date(2027, 2, 1)) == ["2027-01", "2026-12", "2026-11"]


# ---------- OCR 텍스트 분석 ----------

def test_find_ocr_dates():
    text = "승인일시 2026.10.27 14:02\n2026-10-27\n2026년 10월 3일\n2026.13.01"
    assert ca.find_ocr_dates(text) == ["2026-10-27", "2026-10-03"]


def test_find_amounts_ignores_dates_and_card_numbers():
    text = "아메리카노 4,500\n합계 9,000원\n주차요금 3000원\n카드번호 1234-56**-****\n2026.10.27"
    assert ca.find_amounts(text) == [4500, 9000, 3000]


def test_detect_category_uses_priority():
    cfg = ca.default_settings()
    kw, pr = cfg["키워드"], cfg["분류_우선순위"]
    assert ca.detect_category(ca.compact("주 차 요 금"), kw, pr) == "주차"
    assert ca.detect_category("아메리카노테이크아웃", kw, pr) == "카페"
    assert ca.detect_category("버터주차", kw, pr) == "식대"
    assert ca.detect_category("영수증", kw, pr) == ""


def test_detect_hospital_prefers_longest_key():
    hs = ca.OLD_DEFAULT_HOSPITALS
    assert ca.detect_hospital("세종충남대학교병원주차장", hs) == ("세종충남대", "SCNUH")
    assert ca.detect_hospital("충남대학교병원", hs) == ("남대", "CNUH")
    assert ca.detect_hospital("스타벅스", hs) == ("", "")


def test_analyze_receipt_falls_back_to_file_name_date():
    p = Path("2026-10-22_123342.jpg")
    rec = ca.analyze_receipt(p, p, "건양대병원 주차 입차 3,000원", ca.default_settings(), ca.OLD_DEFAULT_HOSPITALS)
    assert rec.dates == ["2026-10-22"] and not rec.date_from_ocr
    assert (rec.total, rec.category, rec.hospital_en) == (3000, "주차", "KYUH")
    assert ca.receipt_issues(rec, [rec]) == ["날짜 못 읽음(촬영 날짜 사용)"]


def test_analyze_receipt_ocr_date_comes_first():
    p = Path("2026-10-22_090000.jpg")
    rec = ca.analyze_receipt(p, p, "2026/10/21 카페 4,500", ca.default_settings(), ca.OLD_DEFAULT_HOSPITALS)
    assert rec.dates == ["2026-10-21", "2026-10-22"]
    assert rec.main_date == "2026-10-21"


# ---------- 영수증 목록에서 고친 내용 ----------

def _auto_receipt():
    p = Path("2026-10-22_123342.jpg")
    return ca.analyze_receipt(p, p, "건양대병원 주차 3,000원", ca.default_settings(), ca.OLD_DEFAULT_HOSPITALS)


def _row(**changes):
    row = {"파일": "2026-10-22_123342.jpg", "사용(O/X)": "O",
           "날짜": "2026-10-22", "[자동] 날짜": "2026-10-22",
           "금액": "3000", "[자동] 금액": "3000",
           "분류": "주차", "[자동] 분류": "주차",
           "병원 영문": "KYUH", "[자동] 병원 영문": "KYUH"}
    row.update(changes)
    return row


def test_edits_compare_against_auto_columns():
    # 엑셀이 날짜·금액 모양을 바꿔도([자동] 칸도 같이 바뀜) 고친 것으로 보지 않음
    rec = _auto_receipt()
    ca.apply_user_edits(rec, _row(**{"날짜": "10/22/2026", "[자동] 날짜": "10/22/2026", "금액": "3,000"}))
    assert rec.edited == [] and rec.total == 3000 and rec.use


def test_edits_change_date_and_amount():
    rec = _auto_receipt()
    ca.apply_user_edits(rec, _row(**{"날짜": "10/23/2026", "금액": "3,500"}))
    assert rec.dates == ["2026-10-23"] and rec.total == 3500 and rec.amounts == [3500]
    assert rec.edited == ["날짜", "금액"]


def test_edits_exclude_category_hospital():
    rec = _auto_receipt()
    ca.apply_user_edits(rec, _row(**{"사용(O/X)": "x", "분류": "식대", "병원 영문": ""}))
    assert rec.use is False
    assert rec.category == "식대" and rec.hospital_en == ""
    assert rec.edited == ["분류", "병원", "사용"]


def test_edits_unknown_category_becomes_empty():
    rec = _auto_receipt()
    ca.apply_user_edits(rec, _row(**{"분류": "기타"}))
    assert rec.category == ""


def test_receipt_list_edits_survive_rerun(tmp_path, monkeypatch):
    src = tmp_path / "src"
    src.mkdir()
    texts = {"2026-10-05_120000.jpg": "2026.10.05 건양대 주차 3,000원",
             "2026-10-06_120000.jpg": "아메리카노 4,500",
             "2026-11-01_120000.jpg": "2026.11.01 합계 2,000"}     # 11월 영수증은 목록 맨 아래로
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
    tm = ca.TargetMonth(2026, 10)
    try:
        recs = ca.build_receipts(cfg, tm, ca.OLD_DEFAULT_HOSPITALS, results, tmp_path / "w", log)
        assert [r.original.name for r in recs] == list(texts)
        assert recs[1].category == "카페" and recs[1].hospital_en == ""
        assert not recs[2].use and recs[2].skip_reason == "다른 달 영수증(2026-11-01)(자동 제외)"

        list_path = ca.list_csv_path(results, tm)
        rows = ca.read_csv_dicts(list_path)
        rows[0]["사용(O/X)"] = "X"
        rows[1]["병원 영문"] = "KYUH"
        ca.write_csv(list_path, ca.LIST_HEADER, [[r[h] for h in ca.LIST_HEADER] for r in rows])

        class NoOcr:
            def __init__(self, *a):
                raise AssertionError("두 번째 실행은 OCR 캐시를 써야 합니다")

        monkeypatch.setattr(ca, "WindowsOcr", NoOcr)
        recs = ca.build_receipts(cfg, tm, ca.OLD_DEFAULT_HOSPITALS, results, tmp_path / "w", log)
        assert recs[0].use is False
        assert recs[1].hospital_en == "KYUH" and recs[1].edited == ["병원"]

        rows = ca.read_csv_dicts(list_path)
        assert rows[1]["병원 영문"] == "KYUH" and rows[1]["[자동] 병원 영문"] == ""
    finally:
        log.close()


# ---------- 경비 ↔ 영수증 매칭 ----------

def _rec(name, dates, amounts, total):
    return ca.Receipt(original=Path(name), ocr_file=Path(name), dates=dates, amounts=amounts, total=total)


def test_find_receipt_prefers_total_then_any_amount():
    r1 = _rec("a.jpg", ["2026-10-05"], [4500, 9000], 9000)
    r2 = _rec("b.jpg", ["2026-10-05"], [4500], 4500)
    assert ca.find_receipt([r1, r2], "2026-10-05", 4500) is r2
    r2.used_for = "1번"
    assert ca.find_receipt([r1, r2], "2026-10-05", 4500) is r1
    assert ca.find_receipt([r1, r2], "2026-10-06", 9000) is None
    r1.use = False
    assert ca.find_receipt([r1, r2], "2026-10-05", 9000) is None


def test_texts_for_drops_hospital_phrase_when_unknown():
    cfg = ca.default_settings()
    rec = ca.Receipt(Path("a"), Path("a"), category="주차", hospital_en="KYUH")
    assert ca.texts_for(rec, cfg) == {"Comment": "Parking fee for service in KYUH", "Business Purpose": "For service"}
    rec.hospital_en = ""
    assert ca.texts_for(rec, cfg)["Comment"] == "Parking fee for service"
    rec.category = ""
    assert ca.texts_for(rec, cfg) == {}


def test_concur_values():
    assert ca.concur_date_to_iso("10/05/2026") == "2026-10-05"
    assert ca.concur_date_to_iso("") is None
    assert ca.concur_amount_to_int("12,000") == 12000
    assert ca.concur_amount_to_int("₩4,500.00") == 4500
    assert ca.concur_amount_to_int("") is None


# ---------- 설정 파일 / 병원 목록 ----------

def test_settings_created_then_merged(tmp_path):
    p = tmp_path / ca.SETTINGS_FILE_NAME
    _, created = ca.load_settings(p)
    assert created and p.exists()

    p.write_text(json.dumps({"키워드": {"식대": ["버터", "식당"]}, "기본_도시": "Gwangju"}, ensure_ascii=False),
                 encoding="utf-8")
    cfg, created = ca.load_settings(p)
    assert not created
    assert cfg["키워드"]["식대"] == ["버터", "식당"]
    assert cfg["키워드"]["주차"] == ca.default_settings()["키워드"]["주차"]
    assert cfg["기본_도시"] == "Gwangju"

    p.write_text("{bad", encoding="utf-8")
    with pytest.raises(RuntimeError):
        ca.load_settings(p)


def test_hospitals_read_cp949_from_excel(tmp_path):
    p = tmp_path / ca.HOSPITAL_FILE_NAME
    p.write_bytes("OCR에서 찾을 글자,Concur에 쓸 영문명\n전남대,CNUH Gwangju\n,빈칸\n".encode("cp949"))
    assert ca.load_hospitals(p) == [["전남대", "CNUH Gwangju"]]


def test_hospitals_default_file_created(tmp_path):
    p = tmp_path / ca.HOSPITAL_FILE_NAME
    assert ca.load_hospitals(p) == ca.DEFAULT_HOSPITALS
