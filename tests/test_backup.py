# -*- coding: utf-8 -*-
"""백업·되살리기: 자동 입력 규칙·병원 목록·설정·영수증 목록을 zip으로 묶고 되살리기,
바뀐 것이 있을 때만 자동 백업, 새 폴더에 받은 프로그램을 처음 켤 때 되살리기"""

import json
import os
import zipfile
from pathlib import Path

import pytest

import concur_auto as ca


def make_base(base: Path, city="Gwangju (ex Kwangju), KOREA, REPUBLIC OF", rule="톨게이트"):
    base.mkdir(parents=True, exist_ok=True)
    cfg = ca.default_settings()
    cfg["기본_도시"] = city
    ca.save_settings(base / ca.SETTINGS_FILE_NAME, cfg)
    rows, default_type = ca.load_category_rows(base / ca.CATEGORY_FILE_NAME, cfg)
    rows.append([rule, "통행료, 톨게이트", "Toll fee for service in {hospital}", "For service", "Parking & Toll", ""])
    ca.save_category_rows(base / ca.CATEGORY_FILE_NAME, rows, default_type)
    (base / ca.HOSPITAL_FILE_NAME).write_text("찾을 글자,영문 이름\n원광대,Wonkwang University Hospital\n", encoding="utf-8-sig")
    (base / "결과").mkdir(exist_ok=True)
    (base / "결과" / "영수증목록_2026-11.csv").write_text("파일,금액\nIMG_0001.JPG,3000\n", encoding="utf-8-sig")
    (base / "결과" / "ocr_cache.json").write_text("{}", encoding="utf-8")
    (base / "결과" / "화면.png").write_bytes(b"png")                     # 사진·캡처는 백업하지 않음
    (base / "실행기록").mkdir(exist_ok=True)
    (base / "실행기록" / "실행기록_1.txt").write_text("log", encoding="utf-8")
    return base


def test_backup_and_restore_into_new_folder(tmp_path):
    old = make_base(tmp_path / "예전 프로그램")
    folder = tmp_path / "백업"
    info = ca.make_backup(old, folder / "a.zip", "직접 백업")
    assert sorted(info["파일"]) == sorted([ca.SETTINGS_FILE_NAME, ca.CATEGORY_FILE_NAME, ca.HOSPITAL_FILE_NAME,
                                           "결과/영수증목록_2026-11.csv", "결과/ocr_cache.json"])
    assert info["프로그램"] == ca.APP_VERSION

    new = tmp_path / "새 프로그램"
    new.mkdir()
    names = ca.restore_backup(folder / "a.zip", new, folder)
    assert sorted(names) == sorted(info["파일"])
    for name in names:
        assert (new / name).read_bytes() == (old / name).read_bytes()
    rows, _ = ca.load_category_rows(new / ca.CATEGORY_FILE_NAME, ca.default_settings())
    assert "톨게이트" in [r[0] for r in rows]                               # 만든 규칙이 새 프로그램에 그대로
    assert not (new / "실행기록").exists() and not (new / "결과" / "화면.png").exists()


def test_auto_backup_only_when_changed_and_keeps_recent(tmp_path, monkeypatch):
    base = make_base(tmp_path / "app")
    folder = tmp_path / "백업"
    first = ca.auto_backup(base, folder)
    assert first is not None and first.name.startswith(ca.AUTO_BACKUP_PREFIX)
    assert ca.auto_backup(base, folder) is None                           # 바뀐 것이 없으면 만들지 않음
    (base / ca.HOSPITAL_FILE_NAME).write_text("찾을 글자,영문 이름\n목포한국,Mokpo Hankook Hospital\n", encoding="utf-8-sig")
    assert ca.auto_backup(base, folder) is not None

    monkeypatch.setattr(ca, "AUTO_BACKUP_KEEP", 3)
    for i in range(5):
        (base / "결과" / "ocr_cache.json").write_text(json.dumps({"n": i}), encoding="utf-8")
        ca.auto_backup(base, folder)
    autos = [p for p, _ in ca.list_backups(folder) if p.name.startswith(ca.AUTO_BACKUP_PREFIX)]
    assert len(autos) == 3                                                # 최근 것만 남김
    assert json.loads(zipfile.ZipFile(autos[0]).read("결과/ocr_cache.json")) == {"n": 4}


def test_restore_keeps_undo_backup(tmp_path):
    base = make_base(tmp_path / "app", rule="톨게이트")
    folder = tmp_path / "백업"
    ca.make_backup(base, folder / "직접백업_old.zip", "직접 백업")
    rows, d = ca.load_category_rows(base / ca.CATEGORY_FILE_NAME, ca.default_settings())
    rows.append(["택시", "택시", "Taxi fee", "For service", "Taxi", ""])
    ca.save_category_rows(base / ca.CATEGORY_FILE_NAME, rows, d)        # 그 뒤에 규칙을 더 만듦

    ca.restore_backup(folder / "직접백업_old.zip", base, folder)
    names = [r[0] for r in ca.load_category_rows(base / ca.CATEGORY_FILE_NAME, ca.default_settings())[0]]
    assert "택시" not in names                                             # 예전 상태로
    undo, info = ca.list_backups(folder)[0]
    assert info["메모"] == "되살리기 전"
    assert ca.backup_label(undo, info)[1] == "자동 (되살리기 전)"
    ca.restore_backup(undo, base, folder)                                 # 되돌리기
    names = [r[0] for r in ca.load_category_rows(base / ca.CATEGORY_FILE_NAME, ca.default_settings())[0]]
    assert "택시" in names


def test_restore_refuses_strange_files(tmp_path):
    folder = tmp_path / "백업"
    folder.mkdir()
    plain = folder / "그냥.zip"
    with zipfile.ZipFile(plain, "w") as z:
        z.writestr("자동입력표.csv", "x")
    with pytest.raises(RuntimeError, match="이 프로그램이 만든 백업 파일이 아닙니다"):
        ca.restore_backup(plain, tmp_path / "app", folder)

    bad = folder / "bad.zip"
    names = ["../밖.csv", "결과/../../밖.json", "결과/하위/x.csv", "concur_auto.py", "결과/영수증목록_2026-11.csv"]
    with zipfile.ZipFile(bad, "w") as z:
        z.writestr(ca.BACKUP_INFO_NAME, json.dumps({"파일": names}))
        for n in names:
            z.writestr(n, "x")
    base = tmp_path / "app"
    base.mkdir()
    assert ca.restore_backup(bad, base, folder) == ["결과/영수증목록_2026-11.csv"]   # 알려진 개인 파일만
    assert not (tmp_path / "밖.csv").exists() and not (base / "concur_auto.py").exists()


def test_backup_folder_choice(tmp_path, monkeypatch):
    monkeypatch.delenv("CONCUR_BACKUP_DIR")
    assert ca.backup_dir({"백업_폴더": str(tmp_path / "내 백업")}) == tmp_path / "내 백업"
    assert ca.backup_dir({"백업_폴더": ""}) == Path.home() / "Documents" / "Concur 경비 자동 입력 백업"


def test_nothing_to_back_up(tmp_path):
    with pytest.raises(RuntimeError, match="백업할 것이 없습니다"):
        ca.make_backup(tmp_path, tmp_path / "x.zip")
    assert ca.auto_backup(tmp_path, tmp_path / "백업") is None


# ---------- 실제 창 (tkinter와 화면이 있을 때만) ----------

tkinter = pytest.importorskip("tkinter") if os.environ.get("DISPLAY") else None


@pytest.mark.skipif(tkinter is None, reason="창을 띄울 화면이 없음")
def test_new_folder_offers_restore_and_backup_tab(tmp_path, backup_folder, monkeypatch):
    import tkinter.messagebox as mb
    old = make_base(tmp_path / "예전 프로그램", city="Iksan, KOREA, REPUBLIC OF")
    ca.auto_backup(old, backup_folder)
    asked = []
    monkeypatch.setattr(mb, "askyesno", lambda *a, **k: asked.append(a[0]) or True)
    monkeypatch.setattr(mb, "showinfo", lambda *a, **k: asked.append(a[0]))

    new = tmp_path / "새 프로그램"
    new.mkdir()
    app = ca.GuiApp(new)                                                  # 설정 파일이 없는 새 폴더
    try:
        assert asked[0] == "예전 설정 되살리기" and app.restored
        assert app.city.get() == "Iksan, KOREA, REPUBLIC OF"              # 예전 설정이 그대로
        tabs = [app.notebook.tab(t, "text").strip() for t in app.notebook.tabs()]
        assert tabs == ["실행", "영수증 불러오기", "백업·되살리기"]

        tab = app.backups
        manual = tab.backup_now()
        assert manual.name.startswith("직접백업_") and tab.tree.get_children()
        app.city.set("Mokpo, KOREA, REPUBLIC OF")
        app.save_form()
        app.open_category_table()                                         # 편집 창을 저장하면 자동 백업
        ed = app.editor
        ed.select_visible([ed.rows[i][0] for i in ed.visible].index("톨게이트"))
        ed.vars["keywords"].set("통행료, 톨게이트, 하이패스")
        ed.on_edit()
        before = len(ca.list_backups(backup_folder))
        assert ed.save()
        assert len(ca.list_backups(backup_folder)) == before + 1

        assert not tab.restore(manual, ask=False)                         # 편집 창이 열려 있으면 안 됨
        assert asked[-1] == "되살리기"
        ed.close()
        assert tab.restore(manual, ask=False)
        assert app.city.get() == "Iksan, KOREA, REPUBLIC OF"              # 창에도 바로 반영
        rows, _ = ca.load_category_rows(new / ca.CATEGORY_FILE_NAME, ca.default_settings())
        assert next(r for r in rows if r[0] == "톨게이트")[1] == "통행료, 톨게이트"
        app.on_close()                                                    # 끌 때도 자동 백업(바뀐 게 있으면)
    except Exception:
        app.root.destroy()
        raise
