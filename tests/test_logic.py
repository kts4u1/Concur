# -*- coding: utf-8 -*-
"""Windows 없이 돌릴 수 있는 부분(날짜·금액·분류·병원 추출, 영수증 목록, 매칭) 테스트"""

import json
from datetime import date
from pathlib import Path

import pytest

import concur_auto as ca


# ---------- 처리할 달 ----------

def test_기본값은_지난달():
    assert ca.resolve_target_month(None, date(2026, 11, 12)) == ca.TargetMonth(2026, 10)
    assert ca.resolve_target_month(None, date(2027, 1, 5)) == ca.TargetMonth(2026, 12)


def test_처리할_달은_2026_10_형식만_받음():
    assert ca.resolve_target_month("2026-3") == ca.TargetMonth(2026, 3)
    for bad in ("2026-13", "10월", "2026/10"):
        with pytest.raises(RuntimeError):
            ca.resolve_target_month(bad)


def test_리포트_이름과_기간_날짜():
    tm = ca.TargetMonth(2028, 2)
    assert tm.prefix == "2028-02"
    assert tm.month_name == "February"
    assert tm.first_day_us == "02/01/2028"
    assert tm.last_day_us == "02/29/2028"
    assert tm.report_name("{month_name} Expenses") == "February Expenses"


def test_달_목록은_해를_넘어가도_이어짐():
    assert ca.recent_months(3, date(2027, 2, 1)) == ["2027-01", "2026-12", "2026-11"]


# ---------- OCR 텍스트 분석 ----------

def test_OCR에서_날짜_찾기():
    text = "승인일시 2026.10.27 14:02\n2026-10-27\n2026년 10월 3일\n2026.13.01"
    assert ca.find_ocr_dates(text) == ["2026-10-27", "2026-10-03"]


def test_금액은_날짜와_카드번호를_빼고_찾음():
    text = "아메리카노 4,500\n합계 9,000원\n주차요금 3000원\n카드번호 1234-56**-****\n2026.10.27"
    assert ca.find_amounts(text) == [4500, 9000, 3000]


def test_분류는_우선순위대로_정함():
    cfg = ca.default_settings()
    kw, pr = cfg["키워드"], cfg["분류_우선순위"]
    assert ca.detect_category(ca.compact("주 차 요 금"), kw, pr) == "주차"
    assert ca.detect_category("아메리카노테이크아웃", kw, pr) == "카페"
    assert ca.detect_category("버터주차", kw, pr) == "식대"
    assert ca.detect_category("영수증", kw, pr) == ""


def test_병원은_더_긴_글자가_우선():
    hs = ca.DEFAULT_HOSPITALS
    assert ca.detect_hospital("세종충남대학교병원주차장", hs) == ("세종충남대", "SCNUH")
    assert ca.detect_hospital("충남대학교병원", hs) == ("남대", "CNUH")
    assert ca.detect_hospital("스타벅스", hs) == ("", "")


def test_OCR에_날짜가_없으면_파일명_날짜를_씀():
    p = Path("2026-10-22_123342.jpg")
    rec = ca.analyze_receipt(p, p, "건양대병원 주차 입차 3,000원", ca.default_settings(), ca.DEFAULT_HOSPITALS)
    assert rec.dates == ["2026-10-22"] and not rec.date_from_ocr
    assert (rec.total, rec.category, rec.hospital_en) == (3000, "주차", "KYUH")
    assert ca.receipt_issues(rec, [rec]) == ["날짜 못 읽음(파일명 날짜 사용)"]


def test_OCR_날짜가_파일명_날짜보다_먼저():
    p = Path("2026-10-22_090000.jpg")
    rec = ca.analyze_receipt(p, p, "2026/10/21 카페 4,500", ca.default_settings(), ca.DEFAULT_HOSPITALS)
    assert rec.dates == ["2026-10-21", "2026-10-22"]
    assert rec.main_date == "2026-10-21"


# ---------- 영수증 목록에서 고친 내용 ----------

def _auto_receipt():
    p = Path("2026-10-22_123342.jpg")
    return ca.analyze_receipt(p, p, "건양대병원 주차 3,000원", ca.default_settings(), ca.DEFAULT_HOSPITALS)


def _row(**changes):
    row = {"파일": "2026-10-22_123342.jpg", "사용(O/X)": "O",
           "날짜": "2026-10-22", "[자동] 날짜": "2026-10-22",
           "금액": "3000", "[자동] 금액": "3000",
           "분류": "주차", "[자동] 분류": "주차",
           "병원 영문": "KYUH", "[자동] 병원 영문": "KYUH"}
    row.update(changes)
    return row


def test_엑셀이_모양만_바꾼_칸은_고친_것으로_보지_않음():
    # 엑셀이 날짜·금액 모양을 바꿔도([자동] 칸도 같이 바뀜) 고친 것으로 보지 않음
    rec = _auto_receipt()
    ca.apply_user_edits(rec, _row(**{"날짜": "10/22/2026", "[자동] 날짜": "10/22/2026", "금액": "3,000"}))
    assert rec.edited == [] and rec.total == 3000 and rec.use


def test_날짜와_금액을_고치면_반영():
    rec = _auto_receipt()
    ca.apply_user_edits(rec, _row(**{"날짜": "10/23/2026", "금액": "3,500"}))
    assert rec.dates == ["2026-10-23"] and rec.total == 3500 and rec.amounts == [3500]
    assert rec.edited == ["날짜", "금액"]


def test_사용_X와_분류_병원_고치면_반영():
    rec = _auto_receipt()
    ca.apply_user_edits(rec, _row(**{"사용(O/X)": "x", "분류": "식대", "병원 영문": ""}))
    assert rec.use is False
    assert rec.category == "식대" and rec.hospital_en == ""
    assert rec.edited == ["분류", "병원"]


def test_목록에_없는_분류는_빈칸():
    rec = _auto_receipt()
    ca.apply_user_edits(rec, _row(**{"분류": "기타"}))
    assert rec.category == ""


def test_고친_내용은_다시_실행해도_유지되고_OCR은_다시_안_함(tmp_path, monkeypatch):
    src = tmp_path / "src"
    src.mkdir()
    texts = {"2026-10-05_120000.jpg": "2026.10.05 건양대 주차 3,000원",
             "2026-10-06_120000.jpg": "아메리카노 4,500"}
    for name in texts:
        (src / name).write_bytes(b"x")
    (src / "2026-11-01_120000.jpg").write_bytes(b"x")   # 다른 달은 제외

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
        recs = ca.build_receipts(cfg, tm, ca.DEFAULT_HOSPITALS, results, tmp_path / "w", log)
        assert [r.original.name for r in recs] == list(texts)
        assert recs[1].category == "카페" and recs[1].hospital_en == ""

        list_path = ca.list_csv_path(results, tm)
        rows = ca.read_csv_dicts(list_path)
        rows[0]["사용(O/X)"] = "X"
        rows[1]["병원 영문"] = "KYUH"
        ca.write_csv(list_path, ca.LIST_HEADER, [[r[h] for h in ca.LIST_HEADER] for r in rows])

        class NoOcr:
            def __init__(self, *a):
                raise AssertionError("두 번째 실행은 OCR 캐시를 써야 합니다")

        monkeypatch.setattr(ca, "WindowsOcr", NoOcr)
        recs = ca.build_receipts(cfg, tm, ca.DEFAULT_HOSPITALS, results, tmp_path / "w", log)
        assert recs[0].use is False
        assert recs[1].hospital_en == "KYUH" and recs[1].edited == ["병원"]

        rows = ca.read_csv_dicts(list_path)
        assert rows[1]["병원 영문"] == "KYUH" and rows[1]["[자동] 병원 영문"] == ""
    finally:
        log.close()


# ---------- 경비 ↔ 영수증 매칭 ----------

def _rec(name, dates, amounts, total):
    return ca.Receipt(original=Path(name), ocr_file=Path(name), dates=dates, amounts=amounts, total=total)


def test_경비_매칭은_합계_금액_먼저_그다음_다른_금액():
    r1 = _rec("a.jpg", ["2026-10-05"], [4500, 9000], 9000)
    r2 = _rec("b.jpg", ["2026-10-05"], [4500], 4500)
    assert ca.find_receipt([r1, r2], "2026-10-05", 4500) is r2
    r2.used_for = "1번"
    assert ca.find_receipt([r1, r2], "2026-10-05", 4500) is r1
    assert ca.find_receipt([r1, r2], "2026-10-06", 9000) is None
    r1.use = False
    assert ca.find_receipt([r1, r2], "2026-10-05", 9000) is None


def test_병원을_모르면_in_병원_문구를_뺌():
    cfg = ca.default_settings()
    rec = ca.Receipt(Path("a"), Path("a"), category="주차", hospital_en="KYUH")
    assert ca.texts_for(rec, cfg) == {"Comment": "Parking fee for service in KYUH", "Business Purpose": "For service"}
    rec.hospital_en = ""
    assert ca.texts_for(rec, cfg)["Comment"] == "Parking fee for service"
    rec.category = ""
    assert ca.texts_for(rec, cfg) == {}


def test_Concur_화면의_날짜와_금액_읽기():
    assert ca.concur_date_to_iso("10/05/2026") == "2026-10-05"
    assert ca.concur_date_to_iso("") is None
    assert ca.concur_amount_to_int("12,000") == 12000
    assert ca.concur_amount_to_int("₩4,500.00") == 4500
    assert ca.concur_amount_to_int("") is None


# ---------- 설정 파일 / 병원 목록 ----------

def test_설정_파일은_처음에_만들고_고친_값과_합침(tmp_path):
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


def test_엑셀이_ANSI로_저장한_병원_목록도_읽음(tmp_path):
    p = tmp_path / ca.HOSPITAL_FILE_NAME
    p.write_bytes("OCR에서 찾을 글자,Concur에 쓸 영문명\n전남대,CNUH Gwangju\n,빈칸\n".encode("cp949"))
    assert ca.load_hospitals(p) == [["전남대", "CNUH Gwangju"]]


def test_병원_목록이_없으면_기본값으로_만듦(tmp_path):
    p = tmp_path / ca.HOSPITAL_FILE_NAME
    assert ca.load_hospitals(p) == ca.DEFAULT_HOSPITALS
