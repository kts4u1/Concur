# -*- coding: utf-8 -*-
"""영수증 불러오기: 여러 폴더(하위 폴더 포함), 같은 사진은 한 번만, 안드로이드·카카오톡 파일 이름 날짜,
사진 직접 가져오기, 이 PC에서 찾은 곳, 그리고 (창을 띄울 수 있으면) [영수증 불러오기] 탭"""

import os
from datetime import datetime
from pathlib import Path

import pytest

import concur_auto as ca

NOV = ca.TargetMonth(2026, 11)


def photo(path: Path, data: bytes = b"\xff\xd8\xff\xd9", when=datetime(2026, 11, 5, 9, 0)) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)
    ts = when.timestamp()
    os.utime(path, (ts, ts))
    return path


@pytest.mark.parametrize("name, when", [
    ("20261103_123456.jpg", datetime(2026, 11, 3, 12, 34, 56)),                 # 갤럭시
    ("IMG_20261103_123456.jpg", datetime(2026, 11, 3, 12, 34, 56)),             # 안드로이드
    ("PXL_20261103_123456789.jpg", datetime(2026, 11, 3, 12, 34, 56)),          # 픽셀
    ("KakaoTalk_20261103_123456789.jpg", datetime(2026, 11, 3, 12, 34, 56)),    # 카카오톡(보낸 시각)
    ("Screenshot_20261103-123456_Gallery.jpg", datetime(2026, 11, 3, 12, 34, 56)),
    ("2026-10-22_123342.jpg", datetime(2026, 10, 22, 12, 33, 42)),              # 예전 단축어
    ("IMG_0001.JPG", None),
    ("receipt_99920261103_123456.jpg", None),                                   # 숫자 중간은 날짜로 안 봄
])
def test_dates_in_phone_file_names(name, when):
    assert ca.file_name_datetime(name) == when


def test_old_settings_mean_one_folder():
    cfg = ca.default_settings()
    cfg.update({"영수증_원본_폴더": "C:/A"})
    assert ca.receipt_folders(cfg) == [Path("C:/A")]                  # 예전 설정 그대로
    cfg["영수증_추가_폴더"] = ["C:/B", "C:/A", "", "C:/C"]
    assert ca.receipt_folders(cfg) == [Path("C:/A"), Path("C:/B"), Path("C:/C")]
    ca.set_receipt_folders(cfg, ["C:/B", "C:/C"])                      # 첫 폴더를 빼면 다음 폴더가 기본
    assert cfg["영수증_원본_폴더"] == "C:/B" and cfg["영수증_추가_폴더"] == ["C:/C"]
    ca.set_receipt_folders(cfg, [])
    assert ca.receipt_folders(cfg) == []
    with pytest.raises(RuntimeError, match="영수증 불러오기"):
        ca.list_originals(cfg, NOV)


def test_photos_from_several_folders(tmp_path):
    icloud, galaxy, kakao = tmp_path / "iCloud", tmp_path / "OneDrive" / "Samsung Gallery", tmp_path / "카카오톡 받은 파일"
    photo(icloud / "IMG_0001.JPG")
    photo(galaxy / "DCIM" / "Camera" / "20261104_101010.jpg", b"galaxy")              # 하위 폴더
    photo(kakao / "KakaoTalk_20261106_090000123.jpg", b"kakao")
    photo(kakao / "IMG_0001.JPG")                                                     # 아이폰 사진을 또 받음 → 한 번만
    photo(galaxy / ".thumbnails" / "20261104_101010.jpg", b"thumb")                   # 숨김 폴더는 안 봄
    photo(icloud / "Concur_Expenses_Processed" / "IMG_0001_보정됨.png", b"old")       # 예전 보정본
    photo(galaxy / "20261010_101010.jpg", b"october")                                 # 처리할 달 전
    lines = []
    cfg = {"영수증_원본_폴더": str(icloud), "영수증_추가_폴더": [str(galaxy), str(kakao), str(tmp_path / "없는 폴더")]}
    found = ca.list_originals(cfg, NOV, {}, lines.append)
    assert [p.name for p, _ in found] == ["20261104_101010.jpg", "IMG_0001.JPG", "KakaoTalk_20261106_090000123.jpg"]
    assert found[1][0].parent == icloud                                               # 먼저 넣은 폴더의 사진을 씀
    assert any("없는 폴더" in l and "건너뜁니다" in l for l in lines)
    assert any("같은 사진이 있어 한 번만" in l and "1장" in l for l in lines)

    counts = dict(ca.count_photos_per_folder(cfg, NOV, {}))
    assert counts[icloud] == 1 and counts[galaxy] == 1 and counts[kakao] == 2 and counts[tmp_path / "없는 폴더"] is None


def test_same_name_different_photo_is_kept_and_warned(tmp_path):
    a, b = tmp_path / "a", tmp_path / "b"
    photo(a / "IMG_0001.JPG", b"one")
    photo(b / "IMG_0001.JPG", b"other photo")
    lines = []
    found = ca.list_originals({"영수증_원본_폴더": str(a), "영수증_추가_폴더": [str(b)]}, NOV, {}, lines.append)
    assert len(found) == 2
    assert any("이름이 같은 다른 사진" in l for l in lines)


def test_import_files_copies_with_dates(tmp_path):
    usb = tmp_path / "USB"
    one = photo(usb / "20261103_123456.jpg", b"receipt one", datetime(2026, 11, 3, 12, 35))
    two = photo(usb / "sub" / "20261103_123456.jpg", b"another receipt", datetime(2026, 11, 3, 13, 0))
    heic = photo(usb / "IMG_0002.HEIC", b"heic")
    memo = usb / "memo.txt"
    memo.write_text("x")
    dest = tmp_path / "app" / ca.IMPORTED_FOLDER_NAME
    r = ca.import_receipt_files([one, two, heic, memo], dest)
    assert r["copied"] == ["20261103_123456.jpg", "20261103_123456_2.jpg", "IMG_0002.HEIC"]   # 이름이 겹치면 _2
    assert r["skipped"] == ["memo.txt"]
    assert (dest / "20261103_123456_2.jpg").read_bytes() == b"another receipt"
    assert int((dest / "20261103_123456.jpg").stat().st_mtime) == int(one.stat().st_mtime)  # 날짜 정보 유지
    again = ca.import_receipt_files([one, two], dest)
    assert again["copied"] == [] and again["same"] == 2                                      # 두 번 가져와도 그대로
    assert one.exists()                                                                      # 원본은 그대로


def test_suggested_folders(tmp_path):
    home = tmp_path / "home"
    od = tmp_path / "OneDrive"
    for d in (home / "iCloudDrive" / "Concur_Expense", home / "Pictures" / "iCloud Photos" / "Photos",
              od / "사진" / "Samsung Gallery", home / "Documents" / "카카오톡 받은 파일", home / "Downloads"):
        d.mkdir(parents=True)
    found = ca.suggested_receipt_folders(home, {"OneDrive": str(od)})
    labels = [label for label, _ in found]
    assert labels == ["iCloud Drive 영수증 폴더 (아이폰)", "iCloud 사진 (아이폰)", "OneDrive 갤럭시 사진 (Samsung Gallery)",
                      "카카오톡 받은 파일", "다운로드"]
    assert dict(found)["OneDrive 갤럭시 사진 (Samsung Gallery)"] == od / "사진" / "Samsung Gallery"
    assert ca.suggested_receipt_folders(tmp_path / "빈 집", {}) == []


def test_build_receipts_uses_all_folders(tmp_path, monkeypatch):
    icloud, kakao = tmp_path / "iCloud", tmp_path / "kakao"
    photo(icloud / "IMG_0001.JPG", b"a")
    photo(kakao / "KakaoTalk_20261106_090000123.jpg", b"b")
    texts = {"IMG_0001.JPG": "주차 2026.11.05 합계 3,000", "KakaoTalk_20261106_090000123.jpg": "카페 2026.11.06 합계 4,500"}

    class FakeOcr:
        def __init__(self, work_dir):
            pass

        def read(self, path):
            return texts.get(Path(path).name, "")

    monkeypatch.setattr(ca, "WindowsOcr", FakeOcr)
    cfg = ca.default_settings()
    cfg.update({"영수증_원본_폴더": str(icloud), "영수증_추가_폴더": [str(kakao)], "영수증_보정_사용": False})
    results = tmp_path / "결과"
    results.mkdir()
    recs = ca.build_receipts(cfg, NOV, [], results, tmp_path / "work", lambda s="": None)
    assert sorted((r.original.name, r.total) for r in recs) == [("IMG_0001.JPG", 3000),
                                                                ("KakaoTalk_20261106_090000123.jpg", 4500)]


# ---------- 실제 창 (tkinter와 화면이 있을 때만) ----------

tkinter = pytest.importorskip("tkinter") if os.environ.get("DISPLAY") else None


def button_texts(widget):
    out = []
    for child in widget.winfo_children():
        if child.winfo_class() in ("TButton", "Button"):
            out.append(child.cget("text"))
        out.extend(button_texts(child))
    return out


@pytest.mark.skipif(tkinter is None, reason="창을 띄울 화면이 없음")
def test_receipt_sources_tab(tmp_path):
    icloud = tmp_path / "iCloud"
    photo(icloud / "IMG_0001.JPG")
    cfg = ca.default_settings()
    cfg["영수증_원본_폴더"] = str(icloud)
    ca.save_settings(tmp_path / ca.SETTINGS_FILE_NAME, cfg)
    app = ca.GuiApp(tmp_path)
    try:
        tabs = [app.notebook.tab(t, "text").strip() for t in app.notebook.tabs()]
        assert tabs[:2] == ["실행", "영수증 불러오기"]
        src = app.sources
        assert src.folders == [str(icloud)]
        assert app.folder_summary.cget("text") == "1곳: iCloud"
        buttons = button_texts(app.notebook.nametowidget(app.notebook.tabs()[0]))
        assert buttons == ["① 영수증 분석만", "② 전체 실행",                     # ③·중지·결과 폴더 열기는 없앰
                           "영수증 확인", "병원 목록 편집", "자동 입력 편집"]

        usb = tmp_path / "USB"
        r = src.import_files([photo(usb / "20261104_101010.jpg", b"usb"), usb / "없는.jpg"])
        assert r["copied"] == ["20261104_101010.jpg"] and "1장 가져옴" in src.import_status.cget("text")
        imported = str(tmp_path / ca.IMPORTED_FOLDER_NAME)
        assert src.folders == [str(icloud), imported]                        # 가져온 영수증 폴더가 목록에 들어감
        saved, _ = ca.load_settings(tmp_path / ca.SETTINGS_FILE_NAME)
        assert ca.receipt_folders(saved) == [icloud, Path(imported)]          # 바로 저장됨
        assert app.folder_summary.cget("text") == "2곳: iCloud, 가져온 영수증"

        assert not src.add_folder(str(icloud))                               # 같은 폴더는 한 번만
        app.month.set("2026-11")
        src.count_photos(wait=True)
        app.root.update()
        app.poll()
        values = [src.tree.item(i, "values") for i in src.tree.get_children()]
        assert [v[1] for v in values] == ["1장", "1장"]

        src.tree.selection_set("0")
        src.remove_selected(ask=False)                                       # 기본 폴더를 빼면 다음 폴더가 기본
        saved, _ = ca.load_settings(tmp_path / ca.SETTINGS_FILE_NAME)
        assert saved["영수증_원본_폴더"] == imported and saved["영수증_추가_폴더"] == []
        app.save_form()
        assert ca.receipt_folders(ca.load_settings(tmp_path / ca.SETTINGS_FILE_NAME)[0]) == [Path(imported)]
    finally:
        app.root.destroy()
