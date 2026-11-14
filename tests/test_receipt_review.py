# -*- coding: utf-8 -*-
"""영수증 확인 창(엑셀 대신): ①이 끝나면 영수증을 사진과 함께 보며 사용 여부·날짜·금액·분류·병원·도시를 고르거나 고치고,
고친 내용은 영수증 목록 파일에 저장되어 다음 실행에 그대로 쓰입니다."""

import os
from datetime import datetime
from pathlib import Path

import pytest

import concur_auto as ca

NOV = ca.TargetMonth(2026, 11)
TEXTS = {"IMG_0001.JPG": "원광대병원 주차 2026.11.02 합계 2,000",
         "IMG_0002.JPG": "다이소 2026.11.05 합계 9,000",
         "IMG_0003.JPG": "스타벅스 아메리카노 2026.11.05 합계 4,500"}


class FakeOcr:
    def __init__(self, work_dir):
        pass

    def read(self, path):
        return TEXTS.get(Path(path).name, "")


def photo(path: Path, when=datetime(2026, 11, 5, 9, 0)):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"\xff\xd8\xff\xd9" + path.name.encode())
    os.utime(path, (when.timestamp(), when.timestamp()))
    return path


@pytest.fixture
def made(tmp_path, monkeypatch):
    src = tmp_path / "src"
    for name in TEXTS:
        photo(src / name)
    monkeypatch.setattr(ca, "WindowsOcr", FakeOcr)
    cfg = ca.default_settings()
    cfg.update({"영수증_원본_폴더": str(src), "영수증_보정_사용": False})
    results = tmp_path / "결과"
    results.mkdir()

    def build():
        return {r.original.name: r for r in ca.build_receipts(cfg, NOV, [], results, tmp_path / "work", lambda s="": None)}

    return cfg, build, ca.list_csv_path(results, NOV), src


def test_edits_in_review_are_used_next_time(made):
    cfg, build, path, src = made
    first = build()
    assert first["IMG_0002.JPG"].category == "" and first["IMG_0003.JPG"].use
    rows = ca.read_csv_dicts(path)
    by = {r["파일"]: r for r in rows}
    assert Path(by["IMG_0001.JPG"]["사진 위치"]) == src / "IMG_0001.JPG"       # 사진 위치도 저장
    assert ca.review_status(by["IMG_0002.JPG"]) == "⚠ 확인"                    # 분류 없음
    assert ca.review_status(by["IMG_0001.JPG"]) == "✓"

    by["IMG_0002.JPG"]["분류"] = "식대"                     # 창에서 고른 것처럼
    by["IMG_0002.JPG"]["금액"] = "9,500"
    by["IMG_0003.JPG"]["사용(O/X)"] = "X"
    by["IMG_0001.JPG"]["도시"] = "Mokpo"
    assert ca.review_status(by["IMG_0002.JPG"]) == "✎ 고침" and ca.review_status(by["IMG_0003.JPG"]) == "✕ 뺌"
    ca.save_review_rows(path, rows)

    again = build()
    assert again["IMG_0002.JPG"].category == "식대" and again["IMG_0002.JPG"].total == 9500
    assert not again["IMG_0003.JPG"].use
    assert ca.city_name(again["IMG_0001.JPG"].city) == "Mokpo"


def test_extra_comment_kept_and_added_to_comment(made):
    cfg, build, path, src = made
    build()
    rows = ca.read_csv_dicts(path)
    by = {r["파일"]: r for r in rows}
    by["IMG_0001.JPG"]["추가 Comment"] = "SO 1234567"                          # 영수증 확인 창에서 적은 것처럼
    assert ca.review_status(by["IMG_0001.JPG"]) == "✎ 고침"
    ca.save_review_rows(path, rows)
    again = build()
    rec = again["IMG_0001.JPG"]
    assert rec.extra_comment == "SO 1234567"
    assert ca.texts_for(rec, cfg)["Comment"].endswith(" SO 1234567")         # Concur Comment 맨 뒤
    kept = {r["파일"]: r for r in ca.read_csv_dicts(path)}["IMG_0001.JPG"]
    assert kept["추가 Comment"] == "SO 1234567"                                 # 다시 만들어도 남아 있음


def test_unchanged_value_is_not_an_edit():
    row = {"사용(O/X)": "O", "[자동] 사용": "O", "날짜": "2026-11-05", "[자동] 날짜": "2026-11-05",
           "금액": "4,500", "[자동] 금액": "4500", "분류": "카페", "[자동] 분류": "카페", "확인 필요": ""}
    assert not ca.review_row_edited(row) and ca.review_status(row) == "✓"     # 4,500과 4500은 같은 금액


def test_find_photo_by_name_when_location_missing(made):
    cfg, build, path, src = made
    build()
    row = dict(ca.read_csv_dicts(path)[0])
    row["사진 위치"] = str(src / "옮겨짐" / row["파일"])                        # 예전 위치에 없으면
    assert ca.find_receipt_photo(row, cfg, {}) == src / row["파일"]           # 영수증 폴더에서 이름으로 찾음
    row["파일"] = "없는사진.jpg"
    assert ca.find_receipt_photo(row, cfg, {}) is None


def test_ocr_mode_opens_review_instead_of_excel(tmp_path, monkeypatch):
    src = tmp_path / "src"
    photo(src / "IMG_0001.JPG")
    base = tmp_path / "app"
    base.mkdir()
    ca.save_settings(base / ca.SETTINGS_FILE_NAME, {"영수증_원본_폴더": str(src), "영수증_보정_사용": False})
    monkeypatch.setattr(ca, "WindowsOcr", FakeOcr)
    seen = {"review": [], "open": [], "info": []}

    class UI:
        def log(self, line):
            pass

        def info(self, title, msg):
            seen["info"].append(title)

        def open(self, path):
            seen["open"].append(path)

        def review(self, path, wait=False):
            seen["review"].append((Path(path).name, wait))
            return True

    ca.run_job("ocr", "2026-11", UI(), base)
    assert seen == {"review": [("영수증목록_2026-11.csv", False)], "open": [], "info": []}


# ---------- 실제 창 (tkinter와 화면이 있을 때만) ----------

tkinter = pytest.importorskip("tkinter") if os.environ.get("DISPLAY") else None


@pytest.mark.skipif(tkinter is None, reason="창을 띄울 화면이 없음")
def test_review_window(made, tmp_path):
    cfg, build, path, src = made
    build()
    base = tmp_path / "app"
    base.mkdir()
    ca.save_settings(base / ca.SETTINGS_FILE_NAME, cfg)
    app = ca.GuiApp(base)
    try:
        closed = []
        ca.GuiBridge(app.q).review(path)                                      # 작업 스레드에서 보낸 것처럼
        app.poll()
        rv = app.review_win
        assert rv.win.winfo_exists() and len(rv.tree.get_children()) == 3
        assert "확인 필요 1장" in rv.summary.cget("text")
        iid = next(i for i in rv.tree.get_children() if rv.rows[int(i)]["파일"] == "IMG_0002.JPG")
        rv.tree.selection_set(iid)
        rv.on_select()
        assert rv.vars["cat"].get() == rv.NO_CATEGORY_TEXT and rv.vars["amount"].get() == "9,000"
        assert "식대" in rv.cat_box.cget("values") and "9,000" in rv.amount_box.cget("values")
        assert rv.file_label.cget("text") == "IMG_0002.JPG"
        rv.vars["cat"].set("식대")
        rv.on_edit()
        assert rv.tree.item(iid, "values")[0] == "✎ 고침" and rv.dirty
        rv.vars["date"].set("11월 5일")
        rv.on_edit()
        assert "2026-10-27처럼" in rv.status.cget("text")                     # 날짜 모양이 틀리면 알려 줌
        rv.vars["date"].set("2026-11-05")
        rv.vars["memo"].set("SO 1234567")                                     # 서비스 오더 번호 등
        rv.on_edit()
        assert rv.save() and not rv.dirty
        saved = {r["파일"]: r for r in ca.read_csv_dicts(path)}["IMG_0002.JPG"]
        assert saved["분류"] == "식대" and saved["추가 Comment"] == "SO 1234567"
        rv.only_check.set(True)
        rv.refresh_list()
        assert rv.tree.get_children() == ()                                   # 확인 필요가 남지 않음
        rv.close()
        assert not rv.win.winfo_exists()

        rv2 = app.open_review(path, on_close=lambda: closed.append(1))        # ②의 '입력 전 확인'처럼 닫을 때 알림
        rv2.close()
        assert closed == [1]
    finally:
        app.root.destroy()
