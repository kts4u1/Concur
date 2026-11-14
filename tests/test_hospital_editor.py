# -*- coding: utf-8 -*-
"""병원 목록 편집 창(엑셀 대신): 찾을 글자 → 영문 이름, 자동으로 만든 영문 이름 확인, 저장 전 확인할 것"""

import os

import pytest

import concur_auto as ca


def test_hospital_problems():
    rows = [["원광대", "Wonkwang University Hospital", ""], ["원광대", "Other", ""], ["순천향", "", ""], ["대", "X", ""]]
    problems = ca.hospital_problems(rows)
    assert "'원광대' 찾을 글자가 두 번 있습니다(위쪽 'Wonkwang University Hospital'만 씁니다)." in problems
    assert "'순천향': Concur에 쓸 영문 이름이 비어 있습니다." in problems
    assert "'대': 찾을 글자가 너무 짧아 다른 글자에도 맞을 수 있습니다." in problems
    assert ca.hospital_problems([["조선대", "Chosun University Hospital", ""]]) == []


tkinter = pytest.importorskip("tkinter") if os.environ.get("DISPLAY") else None


@pytest.mark.skipif(tkinter is None, reason="창을 띄울 화면이 없음")
def test_hospital_editor_window(tmp_path):
    ca.save_settings(tmp_path / ca.SETTINGS_FILE_NAME, ca.default_settings())
    path = tmp_path / ca.HOSPITAL_FILE_NAME
    ca.write_csv(path, ca.HOSPITAL_HEADER, [
        ["원광대", "Wonkwang University Hospital", "영수증에서 확인"],
        ["목포한국", "Mokpo Hankook Hospital", "영수증에서 확인"],
        ["전남대", "Jeonnam University Hospital", "엑셀에서 자동 변환: 전남대학교병원 (영문 이름 확인)"],
        ["순천향", "", "직접 적음"]])                                            # 영문 이름이 빈 줄도 지우지 않음
    app = ca.GuiApp(tmp_path)
    app.mb.askyesno = lambda *a, **k: True
    try:
        ed = app.open_hospitals()
        assert app.open_hospitals() is ed                                       # 두 번 눌러도 창은 하나
        names = [ed.rows[i][0] for i in ed.visible]
        assert names == ["원광대", "목포한국", "전남대", "순천향"]
        tags = {ed.rows[int(i)][0]: ed.tree.item(i, "tags")[0] for i in ed.tree.get_children()}
        assert tags == {"원광대": "on", "목포한국": "on", "전남대": "warn", "순천향": "off"}

        ed.select_visible(names.index("전남대"))
        assert "자동으로 만든 영문 이름" in ed.check_label.cget("text") and ed.check_button.winfo_manager() == "grid"
        ed.vars["en"].set("Chonnam National University Hospital")              # 고치면 저절로 확인됨
        ed.on_edit()
        assert ed.rows[ed.current][2] == ca.HOSP_USER_NOTE and ed.tree.item(str(ed.current), "tags")[0] == "on"
        assert ed.check_label.cget("text") == ""

        ed.add()
        ed.vars["key"].set("조선대")
        ed.vars["en"].set("Chosun University Hospital")
        ed.on_edit()
        ed.filter.set("조선")
        ed.refresh_list()
        assert [ed.rows[i][0] for i in ed.visible] == ["조선대"]
        ed.filter.set("")
        ed.refresh_list()
        assert ed.save() and not ed.dirty
        saved = {r[0]: r[1] for r in (list(x.values()) for x in ca.read_csv_dicts(path))}
        assert saved["전남대"] == "Chonnam National University Hospital" and saved["조선대"] == "Chosun University Hospital"
        assert "순천향" in saved
        assert ["조선대", "Chosun University Hospital"] in ca.load_hospitals(path)   # 다음 실행부터 씀
        ed.close()
        assert not ed.win.winfo_exists()
    finally:
        app.root.destroy()
