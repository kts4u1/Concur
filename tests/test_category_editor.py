# -*- coding: utf-8 -*-
"""자동 입력 편집 창: 읽기·저장·미리보기·확인할 것, 그리고 (창을 띄울 수 있으면) 실제 창"""

import os

import pytest

import concur_auto as ca


def test_load_and_save_rows(tmp_path):
    cfg = ca.default_settings()
    path = tmp_path / ca.CATEGORY_FILE_NAME
    rows, default_type = ca.load_category_rows(path, cfg)
    assert [r[0] for r in rows] == ["식대", "카페", "주차"] and default_type == "Other Staff Expenses"
    rows.append(["Taxi", "택시, 콜택시", "Taxi fee for service in {hospital}", "For service", "Taxi", ""])
    rows.append(["", "버려짐", "", "", "", ""])                      # 이름이 빈 줄은 저장하지 않음
    ca.save_category_rows(path, rows, "Other")
    again, default_type = ca.load_category_rows(path, cfg)
    assert [r[0] for r in again] == ["식대", "카페", "주차", "Taxi"] and default_type == "Other"
    assert ca.apply_category_table(cfg, path) == ["식대", "카페", "주차", "Taxi"]
    assert cfg["Undefined일_때_선택할_유형"]["분류없음"] == "Other"


def test_preview():
    text = ca.category_preview("Toll fee for service in {hospital}", "For service")
    assert "Comment: Toll fee for service in Chosun University Hospital" in text
    assert "Comment: Toll fee for service\n" in text                  # 병원을 못 찾으면 ' in …' 뺌
    assert "Business Purpose: (비어 있음)" in ca.category_preview("Fuel", "")


def test_problems():
    rows = [["주차", "주차, 입차", "Parking", "For service", "", ""],
            ["Taxi", "택시", "", "", "Taxi", ""],
            ["톨게이트", "통행료, 주차", "Toll", "For service", "", ""],
            ["주차", "", "", "", "", ""]]
    problems = ca.category_problems(rows)
    assert "'Taxi': 찾을 글자는 있는데 Comment·Business Purpose가 비어 있습니다." in problems
    assert "'주차' 글자가 '주차'와 '톨게이트'에 다 있습니다(위쪽 '주차'로 맞춰집니다)." in problems
    assert "'주차' 항목이 두 번 있습니다(위쪽 것만 씁니다)." in problems
    assert ca.category_problems([["주차", "주차", "Parking", "", "", ""]]) == []


# ---------- 실제 창 (tkinter와 화면이 있을 때만) ----------

tkinter = pytest.importorskip("tkinter") if os.environ.get("DISPLAY") else None


@pytest.mark.skipif(tkinter is None, reason="창을 띄울 화면이 없음")
def test_editor_window(tmp_path):
    cfg = ca.default_settings()
    ca.save_settings(tmp_path / ca.SETTINGS_FILE_NAME, cfg)
    ca.add_expense_types_to_table(tmp_path / ca.CATEGORY_FILE_NAME, ["Taxi", "Hotel", "Fuel"], cfg)
    app = ca.GuiApp(tmp_path)
    app.mb.askyesno = lambda *a, **k: True                          # 확인 창은 '예'로
    try:
        app.open_category_table()
        ed = app.editor
        app.open_category_table()                                    # 두 번 눌러도 창은 하나
        assert app.editor is ed
        names = [ed.rows[i][0] for i in ed.visible]
        assert names == ["식대", "카페", "주차", "Taxi", "Hotel", "Fuel"]
        assert ed.vars["name"].get() == "식대"                         # 처음엔 설정된 첫 항목
        ed.select_visible(names.index("Taxi"))
        ed.vars["keywords"].set("택시, 콜택시")
        ed.vars["comment"].set("Taxi fee for service in {hospital}")
        ed.vars["purpose"].set("For service")
        ed.on_edit()
        assert ed.listbox.get(3).startswith("● Taxi") and ed.dirty
        assert "Taxi fee for service in Chosun University Hospital" in ed.preview.cget("text")
        ed.move(-1)                                                  # Taxi를 주차 위로
        assert [r[0] for r in ed.rows][:4] == ["식대", "카페", "Taxi", "주차"]
        ed.only_on.set(True)
        ed.refresh_list()
        assert [ed.rows[i][0] for i in ed.visible] == ["식대", "카페", "Taxi", "주차"]
        ed.only_on.set(False)
        ed.filter.set("hot")
        ed.refresh_list()
        assert [ed.rows[i][0] for i in ed.visible] == ["Hotel"]
        ed.filter.set("")
        ed.refresh_list()
        ed.select_visible([ed.rows[i][0] for i in ed.visible].index("Fuel"))
        ed.remove()
        ed.add_named("주유")
        ed.vars["keywords"].set("주유, 휘발유")
        ed.vars["comment"].set("Fuel for service")
        ed.on_edit()
        assert ed.save() and not ed.dirty
        rows, _ = ca.load_category_rows(tmp_path / ca.CATEGORY_FILE_NAME, cfg)
        assert [r[0] for r in rows] == ["식대", "카페", "Taxi", "주차", "Hotel", "주유"]
        assert rows[2][1] == "택시, 콜택시"
        ed.close()
        assert not ed.win.winfo_exists()
    finally:
        app.root.destroy()


def find_button(widget, text):
    for child in widget.winfo_children():
        if getattr(child, "cget", None) and child.winfo_class() in ("TButton", "Button") and child.cget("text") == text:
            return child
        found = find_button(child, text)
        if found is not None:
            return found
    return None


@pytest.mark.skipif(tkinter is None, reason="창을 띄울 화면이 없음")
def test_korean_name_and_concur_expense_type(tmp_path):
    cfg = ca.default_settings()
    cfg["Concur_경비_유형"] = ca.KNOWN_EXPENSE_TYPES + ["Taxi"]           # [화면 조사]로 읽어 온 것처럼
    ca.save_settings(tmp_path / ca.SETTINGS_FILE_NAME, cfg)
    app = ca.GuiApp(tmp_path)
    try:
        app.open_category_table()
        ed = app.editor
        assert ed.name_entry.winfo_class() == "TEntry"                      # 항목 이름은 직접 적는 칸
        etype_box = next(w for w in ed.name_entry.master.winfo_children() if w.winfo_class() == "TCombobox")
        for t in ("Car mileage", "Meals & Drinks", "Other", "Training", "Parking & Toll", "Taxi"):
            assert t in etype_box.cget("values")                            # Expense Type은 Concur 목록에서 고르기

        find_button(ed.win, "새 항목").invoke()                              # [새 항목] → 바로 이름 적기
        assert ed.rows[ed.current][:5] == ["새 항목", "", "", "", ""]
        assert ed.vars["name"].get() == "새 항목"
        assert ed.add() == "새 항목 2"                                        # 또 누르면 이름이 겹치지 않게
        ed.remove()
        ed.select_visible([ed.rows[i][0] for i in ed.visible].index("새 항목"))
        ed.vars["name"].set("택시")
        ed.vars["keywords"].set("택시, 콜택시")
        ed.vars["comment"].set("Taxi fee for service in {hospital}")
        ed.vars["purpose"].set("For service")
        ed.on_edit()
        assert "Expense Type: Other Staff Expenses (비어 있어 아래 기본값)" in ed.preview.cget("text")
        ed.vars["etype"].set("Taxi")
        etype_box.event_generate("<<ComboboxSelected>>")
        assert "Expense Type: Taxi" in ed.preview.cget("text")
        assert ed.listbox.get(ed.visible.index(ed.current)) == "● 택시   → Taxi"
        assert ed.save()
        rows, _ = ca.load_category_rows(tmp_path / ca.CATEGORY_FILE_NAME, cfg)
        taxi = next(r for r in rows if r[0] == "택시")
        assert taxi[:5] == ["택시", "택시, 콜택시", "Taxi fee for service in {hospital}", "For service", "Taxi"]
        assert "새 항목" not in [r[0] for r in rows]
        ed.close()
    finally:
        app.root.destroy()
