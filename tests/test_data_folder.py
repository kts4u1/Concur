# -*- coding: utf-8 -*-
"""설정·목록 파일을 프로그램 옆(예: 바탕화면)이 아니라 문서\\Concur 경비 자동 입력에 두기"""

import json
import os
import zipfile
from pathlib import Path

import concur_auto as ca


def test_data_dir_is_documents_folder(monkeypatch):
    monkeypatch.delenv("CONCUR_DATA_DIR")
    assert ca.data_dir() == ca.documents_dir() / "Concur 경비 자동 입력"
    assert ca.data_dir() != ca.app_dir()
    monkeypatch.setenv("CONCUR_DATA_DIR", "/somewhere/else")
    assert ca.data_dir() == Path("/somewhere/else")


def make_old(desktop: Path):
    """예전 판이 바탕화면(exe 옆)에 만든 파일들"""
    imported = desktop / ca.IMPORTED_FOLDER_NAME
    cfg = ca.default_settings()
    cfg["영수증_추가_폴더"] = [str(imported), "C:/Users/me/OneDrive/Pictures"]
    ca.save_settings(desktop / ca.SETTINGS_FILE_NAME, cfg)
    (desktop / ca.CATEGORY_FILE_NAME).write_text("분류\n식대\n", encoding="utf-8-sig")
    (desktop / ca.HOSPITAL_FILE_NAME).write_text("OCR에서 찾을 글자\n원광대\n", encoding="utf-8-sig")
    (desktop / "결과").mkdir()
    (desktop / "결과" / "영수증목록_2026-10.csv").write_text("파일\n", encoding="utf-8-sig")
    (desktop / "실행기록").mkdir()
    imported.mkdir()
    (imported / "a.jpg").write_bytes(b"jpg")
    (desktop / "ConcurAuto.exe").write_bytes(b"exe")
    (desktop / "내 문서.txt").write_text("바탕화면의 다른 파일", encoding="utf-8")
    return imported


def test_old_files_move_off_desktop(tmp_path):
    desktop, data = tmp_path / "Desktop", tmp_path / "Documents" / "Concur 경비 자동 입력"
    desktop.mkdir()
    make_old(desktop)
    notes = ca.move_personal_files(desktop, data)
    for name in (ca.SETTINGS_FILE_NAME, ca.CATEGORY_FILE_NAME, ca.HOSPITAL_FILE_NAME, "결과", "실행기록", ca.IMPORTED_FOLDER_NAME):
        assert (data / name).exists() and not (desktop / name).exists()
    assert (data / "결과" / "영수증목록_2026-10.csv").is_file() and (data / ca.IMPORTED_FOLDER_NAME / "a.jpg").is_file()
    assert sorted(f.name for f in desktop.iterdir()) == ["ConcurAuto.exe", "내 문서.txt"]     # 다른 파일은 그대로
    cfg = json.loads((data / ca.SETTINGS_FILE_NAME).read_text(encoding="utf-8"))
    assert cfg["영수증_추가_폴더"] == [str(data / ca.IMPORTED_FOLDER_NAME), "C:/Users/me/OneDrive/Pictures"]
    assert len(notes) == 1 and "옮겼습니다" in notes[0] and str(data) in notes[0]
    loaded, created = ca.load_settings(data / ca.SETTINGS_FILE_NAME)
    assert not created and data / ca.IMPORTED_FOLDER_NAME in ca.receipt_folders(loaded)


def test_existing_files_in_documents_are_kept(tmp_path):
    desktop, data = tmp_path / "Desktop", tmp_path / "data"
    desktop.mkdir()
    make_old(desktop)
    data.mkdir()
    ca.save_settings(data / ca.SETTINGS_FILE_NAME, {"기본_도시": "Mokpo, KOREA, REPUBLIC OF"})
    notes = ca.move_personal_files(desktop, data)
    assert json.loads((data / ca.SETTINGS_FILE_NAME).read_text(encoding="utf-8")) == {"기본_도시": "Mokpo, KOREA, REPUBLIC OF"}
    assert (desktop / ca.SETTINGS_FILE_NAME).exists()                       # 덮어쓰지 않고 남겨 둠
    assert (data / ca.CATEGORY_FILE_NAME).exists() and not (desktop / ca.CATEGORY_FILE_NAME).exists()
    assert any("이미 같은 이름" in n and ca.SETTINGS_FILE_NAME in n for n in notes)


def test_nothing_to_move(tmp_path):
    folder = tmp_path / "same"
    folder.mkdir()
    make_old(folder)
    assert ca.move_personal_files(folder, folder) == []                    # 같은 폴더면 그대로
    empty = tmp_path / "empty"
    empty.mkdir()
    assert ca.move_personal_files(empty, tmp_path / "data") == [] and not (tmp_path / "data").exists()


def test_old_backup_folder_follows_documents(tmp_path, monkeypatch):
    monkeypatch.delenv("CONCUR_BACKUP_DIR")
    home, docs = tmp_path / "home", tmp_path / "OneDrive" / "Documents"
    monkeypatch.setattr(Path, "home", classmethod(lambda cls: home))
    monkeypatch.setattr(ca, "documents_dir", lambda: docs)
    old = home / "Documents" / ca.BACKUP_FOLDER_NAME
    old.mkdir(parents=True)
    with zipfile.ZipFile(old / "자동백업_20261114_100000.zip", "w") as z:
        z.writestr("x.txt", "x")
    notes = ca.move_old_backup_folder()
    assert (docs / ca.BACKUP_FOLDER_NAME / "자동백업_20261114_100000.zip").is_file() and not old.exists()
    assert notes and "옮겼습니다" in notes[0]
    assert ca.move_old_backup_folder() == []                                # 두 번째에는 할 일 없음
