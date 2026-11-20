# -*- coding: utf-8 -*-
"""2026-11-20 사용자 화면: 돼지찌개 식당·매운탕 식당 영수증이 '분류 없음'.
날짜·금액은 읽었지만 처음 읽은 글자에서 가게 이름·메뉴가 깨져('돼지찌개' → '도!|지씨개') 찾을 글자가 안 맞았음.
→ 분류를 못 찾은 이 달 영수증은 크게(그래도 없으면 보정해서) 다시 읽어 분류만 찾고, 날짜·금액은 처음 읽은 그대로.
(가게 이름·주소·번호는 모두 가짜)"""

from pathlib import Path

import concur_auto as ca

# 처음 읽기: 날짜·금액은 맞게 읽었지만 가게 이름·메뉴가 깨짐(찾을 글자가 하나도 안 맞음)
GARBLED = """/쇼60/名춈[영수증]-재발행
가나도!|지ㅉ1개 / 123-45-67891 / 홍길동
전남 목포시 가나로 1 가나내시피개
2026-07-21 13:29:33
상 품 명 단 가 수 량 금 액
생도|시피7H 9,000 2 18,000
합 계 금 액 18,000"""

# 크게 다시 읽기: 가게 이름·메뉴는 보이지만 금액 하나를 잘못 읽음(이 금액은 쓰지 않아야 함)
ZOOMED = """[영수증]-재발행
가나돼지찌개 / 123-45-67891 / 홍길동
전남 목포시 가나로 1
2026-07-21 13:29:33
생돼지찌개 9,000 2 18,000
합 계 금 액 180,000"""

STILL_GARBLED = """[영수증]
가나도!|지ㅉ1개
2026-07-21 13:29:33
합 계 18,000"""

ENHANCED = """가나식당
2026-07-21 13:29:33
우럭매운탕 14,000 2 28,000
합계 28,000"""

PARKING = """주차영수증
주 차 장 명 : 가나주차장
입 차 시 간 : 2026-07-25 16:36:31
출 차 시 간 : 2026-07-25 19:21:52
주 차 요 금 : 6,100"""

OTHER_MONTH = """[영수증]
가나도!|지ㅉ1개
2026-09-02 12:00:00
합 계 9,000"""


def run(tmp_path, monkeypatch, texts, zoomed, use_pre=False, enhanced=None):
    src = tmp_path / "src"
    src.mkdir()
    for name in texts:
        (src / name).write_bytes(b"x")
    calls = []

    class FakeOcr:
        def __init__(self, temp_dir):
            pass

        def read(self, path):
            return enhanced if path.name.startswith("enh_") else texts[path.name]

        def read_zoomed(self, path):
            calls.append(path.name)
            return zoomed.get(path.name, "")

    def fake_enhanced_copy(cfg, orig, work_dir, log):
        work_dir.mkdir(parents=True, exist_ok=True)
        copy = work_dir / f"enh_{orig.stem}.jpg"
        copy.write_bytes(b"x")
        return copy

    monkeypatch.setattr(ca, "WindowsOcr", FakeOcr)
    monkeypatch.setattr(ca, "make_enhanced_copy", fake_enhanced_copy)
    cfg = ca.default_settings()
    cfg.update({"영수증_원본_폴더": str(src), "보정본_폴더": str(src / "p"), "영수증_보정_사용": use_pre})
    results = tmp_path / "결과"
    results.mkdir()
    lines = []
    log = ca.Log(tmp_path / "log", sink=lines.append)
    recs = ca.build_receipts(cfg, ca.TargetMonth(2026, 7), ca.DEFAULT_HOSPITALS, results, tmp_path / "w", log)
    return {r.original.name: r for r in recs}, calls, lines


def test_uncategorized_receipt_read_again_for_category(tmp_path, monkeypatch):
    texts = {"2026-07-21_132933.jpg": GARBLED, "2026-07-25_192152.jpg": PARKING, "2026-09-02_120000.jpg": OTHER_MONTH}
    by, calls, lines = run(tmp_path, monkeypatch, texts, {"2026-07-21_132933.jpg": ZOOMED})
    meal = by["2026-07-21_132933.jpg"]
    assert meal.category == "식대"                                   # 예전: 분류 없음(크게 읽지 않았음)
    assert meal.total == 18000 and meal.main_date == "2026-07-21"     # 금액·날짜는 처음 읽은 그대로(크게 읽은 180,000 아님)
    assert meal.use
    assert calls == ["2026-07-21_132933.jpg"]        # 분류가 있는 주차 영수증·다른 달 영수증은 다시 읽지 않음
    assert by["2026-07-25_192152.jpg"].category == "주차"
    assert any("분류를 못 찾은 이 달 영수증 1장을 다시 읽어 그중 1장의 분류를 찾았습니다" in l for l in lines)


def test_enhanced_read_when_zoomed_read_has_no_category(tmp_path, monkeypatch):
    first = STILL_GARBLED.replace("18,000", "28,000").replace("07-21", "07-24")
    texts = {"2026-07-24_125700.jpg": first}
    by, calls, lines = run(tmp_path, monkeypatch, texts, {"2026-07-24_125700.jpg": first}, use_pre=True,
                           enhanced=ENHANCED.replace("07-21", "07-23"))       # 보정본이 날짜를 다르게 읽어도 쓰지 않음
    rec = by["2026-07-24_125700.jpg"]
    assert calls == ["2026-07-24_125700.jpg"]
    assert rec.category == "식대" and rec.total == 28000 and rec.main_date == "2026-07-24"
    assert not list((tmp_path / "w").glob("enh_*"))                  # 보정 사본은 지움


def test_category_still_missing_is_reported(tmp_path, monkeypatch):
    texts = {"2026-07-21_132933.jpg": STILL_GARBLED}
    by, calls, lines = run(tmp_path, monkeypatch, texts, {"2026-07-21_132933.jpg": STILL_GARBLED})
    rec = by["2026-07-21_132933.jpg"]
    assert rec.category == "" and rec.total == 18000
    assert any("그중 0장의 분류를 찾았습니다(못 찾은 것은 영수증 확인 창에서 분류를 골라 주세요)" in l for l in lines)

    # 두 번째 실행은 크게 읽은 결과를 기억해 두어 다시 읽지 않음
    calls.clear()
    log = ca.Log(tmp_path / "log2", sink=lambda s: None)
    cfg = ca.default_settings()
    cfg.update({"영수증_원본_폴더": str(tmp_path / "src"), "보정본_폴더": str(tmp_path / "p"), "영수증_보정_사용": False})
    ca.build_receipts(cfg, ca.TargetMonth(2026, 7), ca.DEFAULT_HOSPITALS, tmp_path / "결과", tmp_path / "w", log)
    assert calls == []


def test_ocr_misread_jjigae_counts_as_meal():
    """'찌개'를 '씨개'로 읽어도 식대(판 8 기본 찾을 글자), 매운탕·돼지도"""
    cfg = ca.default_settings()
    for text in ("가나도!|지씨개 2026-07-21 18,000", "우럭매운탕 14,000 2 28,000", "생돼지 2인 18,000"):
        rec = ca.analyze_receipt(Path("IMG_1.HEIC"), Path("IMG_1.HEIC"), text, cfg, [], None)
        assert rec.category == "식대", text


def test_version_8_words_added_to_meal_once(tmp_path):
    settings, table = tmp_path / ca.SETTINGS_FILE_NAME, tmp_path / ca.CATEGORY_FILE_NAME
    cfg = ca.default_settings()
    cfg["_기본_찾을_글자_판"] = 7
    ca.save_settings(settings, cfg)
    ca.save_category_rows(table, [["식대", "찌개, 한식, 식당, 매운탕", "Meal", "For meal", "Other Staff Expenses", ""],
                                  ["카페", "카페", "Cafe", "For wifi", "Other Staff Expenses", ""]])
    assert ca.upgrade_category_table(table, settings) == len(ca.MEAL_V8_WORDS) - 1      # 이미 있는 매운탕은 빼고
    by = {r[ca.C_NAME]: r for r in ca.load_category_rows(table, ca.default_settings())}
    assert by["식대"][ca.C_KEYWORDS] == ", ".join(["찌개", "한식", "식당", "매운탕"]
                                                 + [w for w in ca.MEAL_V8_WORDS if w != "매운탕"])
    assert by["카페"][ca.C_KEYWORDS] == "카페"
    assert ca.upgrade_category_table(table, settings) == 0                              # 한 번만
