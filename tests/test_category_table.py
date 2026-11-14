# -*- coding: utf-8 -*-
"""자동입력표.csv: 분류(주차·카페·식대 …)를 엑셀에서 더하고 고치기"""

from datetime import datetime
from pathlib import Path

import pytest

import concur_auto as ca


def table(tmp_path, rows=None):
    """자동입력표를 만들고(rows가 있으면 그 내용으로) 설정에 반영"""
    cfg = ca.default_settings()
    path = tmp_path / ca.CATEGORY_FILE_NAME
    if rows is not None:
        ca.write_csv(path, ca.CATEGORY_HEADER, rows)
    lines = []
    order = ca.apply_category_table(cfg, path, lines.append)
    return cfg, order, path, lines


def analyze(cfg, text, hospitals=()):
    p = Path("IMG_0001.JPG")
    return ca.analyze_receipt(p, p, text, cfg, list(hospitals), datetime(2026, 11, 3, 9, 0))


def test_first_table_has_current_settings(tmp_path):
    cfg, order, path, lines = table(tmp_path)
    assert order == ["식대", "카페", "주차"]
    rows = [r for r in ca.read_csv_dicts(path)]
    assert [r["분류"] for r in rows] == ["식대", "카페", "주차"]               # '(분류 없음)' 줄은 이제 없음
    parking = rows[2]
    assert parking["영수증에서 찾을 글자(쉼표로 구분)"] == "주차, 입차, 출차, 주자"
    assert parking["Comment"] == "Parking fee for service in {hospital}"
    assert parking["Expense Type이 Undefined일 때"] == "Parking & Toll"
    assert cfg["키워드"]["주차"] == ["주차", "입차", "출차", "주자"]
    assert lines == ["자동 입력표: 분류 3개 (식대, 카페, 주차)"]


def test_add_new_category_and_change_existing_text(tmp_path):
    rows = [
        ["톨게이트", "통행료, 톨게이트, 하이패스", "Toll fee for service at {병원}", "For service", "Parking & Toll", ""],
        ["식대", "버터, 식당, 김밥", "Overtime meal fee for service in {hospital}", "For Overtime meal", "Other Staff Expenses", ""],
        ["카페", "메리카, 리카노, 카페", "Cafe for wifi in {hospital}", "For wifi", "Other Staff Expenses", ""],
        ["주차", "주차, 입차, 출차", "Parking fee for customer visit in {hospital}", "Customer visit", "Parking & Toll", ""],
        ["(분류 없음)", "", "", "", "Miscellaneous", ""],
    ]
    cfg, order, _, _ = table(tmp_path, rows)
    assert order == ["톨게이트", "식대", "카페", "주차"]

    toll = analyze(cfg, "남광주 톨게이트 통행료 2,500원")
    assert toll.category == "톨게이트" and toll.is_receipt
    assert ca.texts_for(toll, cfg) == {"Comment": "Toll fee for service", "Business Purpose": "For service"}
    toll.hospital_en = "Chosun University Hospital"
    assert ca.texts_for(toll, cfg)["Comment"] == "Toll fee for service at Chosun University Hospital"

    parking = analyze(cfg, "주차장 입차 출차 3,000원")
    parking.hospital_en = "Wonkwang University Hospital"
    assert ca.texts_for(parking, cfg) == {"Comment": "Parking fee for customer visit in Wonkwang University Hospital",
                                          "Business Purpose": "Customer visit"}
    assert analyze(cfg, "김밥천국 합계 8,000").category == "식대"            # 새로 넣은 찾을 글자
    assert cfg["Undefined일_때_선택할_유형"] == {"톨게이트": "Parking & Toll",          # 예전 '(분류 없음)' 줄은 무시
                                             "식대": "Other Staff Expenses", "카페": "Other Staff Expenses",
                                             "주차": "Parking & Toll"}


def test_new_category_can_be_typed_in_receipt_list(tmp_path):
    rows = [["톨게이트", "통행료", "Toll fee for service in {hospital}", "For service", "Parking & Toll", ""]]
    cfg, order, _, _ = table(tmp_path, rows)
    rec = analyze(cfg, "합계 2,500 카드")
    row = {"사용(O/X)": "O", "분류": "톨게이트", "[자동] 분류": "", "날짜": "", "[자동] 날짜": "",
           "금액": "", "[자동] 금액": "", "병원 영문": "", "[자동] 병원 영문": ""}
    ca.apply_user_edits(rec, row, categories=order)
    assert rec.category == "톨게이트"
    row["분류"] = "택시"                                                   # 표에 없는 분류는 빈칸
    rec2 = analyze(cfg, "합계 2,500 카드")
    ca.apply_user_edits(rec2, row, categories=order)
    assert rec2.category == ""


def test_empty_table_keeps_settings(tmp_path):
    cfg, order, _, lines = table(tmp_path, [["", "", "", "", "", ""]])
    assert order == ["식대", "카페", "주차"] and cfg["입력_문구"]["주차"]["Business Purpose"] == "For service"
    assert "분류가 없어" in lines[0]


def test_table_saved_by_excel_in_ansi(tmp_path):
    path = tmp_path / ca.CATEGORY_FILE_NAME
    text = ",".join(ca.CATEGORY_HEADER) + "\n주차,주차,Parking fee in {hospital},For service,Parking & Toll,\n"
    path.write_bytes(text.encode("cp949"))
    cfg = ca.default_settings()
    assert ca.apply_category_table(cfg, path) == ["주차"]


def test_run_job_reads_table(tmp_path):
    src = tmp_path / "src"
    src.mkdir()
    base = tmp_path / "app"
    base.mkdir()
    ca.save_settings(base / ca.SETTINGS_FILE_NAME, {"영수증_원본_폴더": str(src), "결과_자동_열기": False})
    ca.write_csv(base / ca.CATEGORY_FILE_NAME, ca.CATEGORY_HEADER,
                 [["톨게이트", "통행료", "Toll fee", "For service", "Parking & Toll", ""]])

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
    assert any("자동 입력표: 분류 1개 (톨게이트)" in line for line in ui.lines)
