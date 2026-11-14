# -*- coding: utf-8 -*-
"""식대·숙박 기본 찾을 글자: 새로 만든 자동입력표에 들어 있고, 예전 자동입력표에는 한 번만 더해짐(사용자 글자·문구는 그대로)"""

from pathlib import Path

import pytest

import concur_auto as ca


def detect(text):
    cfg = ca.default_settings()
    return ca.detect_category(ca.compact(text), cfg["키워드"], cfg["분류_우선순위"])


@pytest.mark.parametrize("text, category", [
    ("김치찌개 1 8,000 합계 8,000", "식대"),
    ("구내식당 식권 5,000", "식대"),
    ("짜장면 짬뽕 탕수육 OO반점", "식대"),
    ("맘스터치 싸이버거 세트", "식대"),
    ("목포 비즈니스 HOTEL 객실료 70,000", "숙박"),                 # 영어는 대소문자 상관없이
    ("Ramada Hotel & Suites", "숙박"),
    ("야놀자 숙박비 55,000 입실 15:00", "숙박"),
    ("호텔 레스토랑 스테이크 45,000", "식대"),                     # 호텔 안 식당은 식대(위쪽 항목 먼저)
    ("스타벅스 아메리카노 4,500", "카페"),
    ("메가MGC커피 아이스 라떼", "카페"),
    ("STARBUCKS COFFEE KOREA", "카페"),
    ("OO타워 PARKING 무인정산 입출차", "주차"),
    ("부산 해운대구 우동 OO주차장 주차요금 3,000", "주차"),          # 지명 '우동'은 식대로 안 봄
    ("다이소 합계 9,000", ""),
])
def test_default_words(text, category):
    assert detect(text) == category


def test_new_table_has_user_items(tmp_path):
    rows = ca.load_category_rows(tmp_path / ca.CATEGORY_FILE_NAME, ca.default_settings())
    by = {r[ca.C_NAME]: r for r in rows}
    assert [r[ca.C_NAME] for r in rows] == ["식대", "숙박", "카페", "주차"]
    assert by["식대"][ca.C_KEYWORDS].startswith("찌개, 한식, 식당, ")
    assert by["숙박"][ca.C_KEYWORDS].startswith("모텔, 호텔, ") and by["숙박"][ca.C_TYPE] == "Accommodation"
    assert by["숙박"][ca.C_COMMENT] == "Accommodation fee for service in {hospital}"
    assert by["숙박"][ca.C_PURPOSE] == "For service in hospital"


def test_old_table_gets_words_once(tmp_path):
    settings, table = tmp_path / ca.SETTINGS_FILE_NAME, tmp_path / ca.CATEGORY_FILE_NAME
    ca.save_settings(settings, ca.default_settings())
    ca.save_category_rows(table, [
        ["식대", "찌개, 한식, 식당", "Overtime meal fee for service in {hospital}", "For Overtime meal", "Other Staff Expenses", ""],
        ["숙박", "모텔, 호텔", "Accomodation fee for service in hospital", "For service in hospital", "Accommodation", ""],
        ["카페", "메리카, 카페", "Cafe for wifi in {hospital}", "For wifi", "Other Staff Expenses", ""],
        ["주차", "주차, 입차", "Parking fee for service in {hospital}", "For service", "Parking & Toll", ""],
        ["톨게이트", "통행료", "Toll fee", "For service", "Parking & Toll", ""]])
    lines = []
    assert ca.upgrade_category_table(table, settings, lines.append) > 0 and "기본 찾을 글자" in lines[0]
    assert any("Accomodation → Accommodation" in l for l in lines)
    by = {r[ca.C_NAME]: r for r in ca.load_category_rows(table, ca.default_settings())}
    meal = [k.strip() for k in by["식대"][ca.C_KEYWORDS].split(",")]
    assert meal[:3] == ["찌개", "한식", "식당"] and "김밥" in meal and "버터" not in meal   # 사용자가 지운 글자는 안 넣음
    assert "hotel" in by["숙박"][ca.C_KEYWORDS] and "객실" in by["숙박"][ca.C_KEYWORDS]
    assert by["숙박"][ca.C_COMMENT] == "Accommodation fee for service in hospital"        # 사용자 문구는 철자만 고침
    cafe = [k.strip() for k in by["카페"][ca.C_KEYWORDS].split(",")]
    assert cafe[:2] == ["메리카", "카페"] and "스타벅스" in cafe and "coffee" in cafe
    assert "parking" in by["주차"][ca.C_KEYWORDS] and by["주차"][ca.C_KEYWORDS].startswith("주차, 입차, ")
    assert by["톨게이트"][ca.C_KEYWORDS] == "통행료"                                          # 사용자 항목은 그대로
    assert ca.load_settings(settings)[0]["_기본_찾을_글자_판"] == ca.DEFAULT_WORDS_VERSION

    rows = ca.load_category_rows(table, ca.default_settings())                           # 사용자가 글자를 지워도
    rows[0][ca.C_KEYWORDS] = "찌개, 한식, 식당, 김밥"
    ca.save_category_rows(table, rows)
    assert ca.upgrade_category_table(table, settings) == 0                               # 다시 더하지 않음(한 번만)
    assert ca.load_category_rows(table, ca.default_settings())[0][ca.C_KEYWORDS] == "찌개, 한식, 식당, 김밥"


def test_version_1_users_get_only_cafe_parking(tmp_path):
    settings, table = tmp_path / ca.SETTINGS_FILE_NAME, tmp_path / ca.CATEGORY_FILE_NAME
    cfg = ca.default_settings()
    cfg["_기본_찾을_글자_판"] = 1                                                          # 식대·숙박 글자는 이미 받음
    ca.save_settings(settings, cfg)
    ca.save_category_rows(table, [["식대", "찌개, 김밥", "Meal", "For meal", "Other Staff Expenses", ""],
                                  ["카페", "카페", "Cafe", "For wifi", "Other Staff Expenses", ""]])
    ca.upgrade_category_table(table, settings)
    by = {r[ca.C_NAME]: r for r in ca.load_category_rows(table, ca.default_settings())}
    assert by["식대"][ca.C_KEYWORDS] == "찌개, 김밥"                                       # 지운 식대 글자는 다시 안 넣음
    assert "숙박" not in by                                                               # 숙박 넣기도 판 1에서 끝남
    assert "스타벅스" in by["카페"][ca.C_KEYWORDS]


def test_old_table_without_accommodation(tmp_path):
    settings, table = tmp_path / ca.SETTINGS_FILE_NAME, tmp_path / ca.CATEGORY_FILE_NAME
    ca.save_settings(settings, ca.default_settings())
    ca.save_category_rows(table, [["식대", "버터", "Meal", "For meal", "Other Staff Expenses", ""],
                                  ["주차", "주차", "Parking", "For service", "Parking & Toll", ""]])
    ca.upgrade_category_table(table, settings)
    rows = ca.load_category_rows(table, ca.default_settings())
    assert [r[ca.C_NAME] for r in rows] == ["식대", "숙박", "주차"]                        # 식대 아래에 숙박
    assert rows[1][ca.C_TYPE] == "Accommodation" and rows[0][ca.C_KEYWORDS].startswith("버터, 찌개")


def test_spelling_fixed_once_for_version_2_users(tmp_path):
    settings, table = tmp_path / ca.SETTINGS_FILE_NAME, tmp_path / ca.CATEGORY_FILE_NAME
    cfg = ca.default_settings()
    cfg["_기본_찾을_글자_판"] = 2                                                          # 글자 더하기는 이미 받음
    ca.save_settings(settings, cfg)
    ca.save_category_rows(table, [["숙박", "모텔, 호텔", "Accomodation fee for service in hospital",
                                   "For service in hospital", "Accommodation", ""]])
    lines = []
    assert ca.upgrade_category_table(table, settings, lines.append) == 0                 # 더한 글자는 없음
    row = ca.load_category_rows(table, ca.default_settings())[0]
    assert row[ca.C_COMMENT] == "Accommodation fee for service in hospital"
    assert row[ca.C_KEYWORDS] == "모텔, 호텔" and lines == ["자동 입력표 문구의 철자를 고쳤습니다(Accomodation → Accommodation, 1곳)."]
