# -*- coding: utf-8 -*-
"""단축어 없이 찍은 사진(IMG_1234.HEIC 등)을 쓰는 기능 테스트:
촬영 날짜로 지난달 사진 고르기, HEIC 읽기·JPEG 변환, 영수증이 아닌 사진 빼기"""

import os
from datetime import datetime
from pathlib import Path

import pytest

import concur_auto as ca


# ---------- 영수증인지 판단 ----------

def judge(text):
    cfg = ca.default_settings()
    p = Path("IMG_0001.JPG")
    return ca.analyze_receipt(p, p, text, cfg, ca.OLD_DEFAULT_HOSPITALS, datetime(2026, 10, 5, 12, 0))


@pytest.mark.parametrize("text", [
    "합계 4,500",
    "2026.10.05 12,000",
    "아메리카노 4,500",                     # 분류 키워드 + 금액
    "카드승인 결제금액",                     # 금액을 못 읽어도 영수증 글자 두 개
    "건양대병원 주차 입차 3,000원",
])
def test_receipts_are_kept(text):
    rec = judge(text)
    assert rec.is_receipt and rec.use


@pytest.mark.parametrize("text", ["", "서울특별시 종로구", "스타벅스", "1,000", "합계"])
def test_other_photos_are_excluded(text):
    rec = judge(text)
    assert not rec.is_receipt and not rec.use
    assert ca.receipt_issues(rec, [rec]) == []


def test_taken_date_is_used_when_ocr_has_no_date():
    rec = judge("합계 4,500")
    assert rec.dates == ["2026-10-05"] and not rec.date_from_ocr


# ---------- 지난달 사진 고르기 ----------

def _set_mtime(path, dt):
    ts = dt.timestamp()
    os.utime(path, (ts, ts))


def _jpeg(path, taken=None, size=(40, 20)):
    from PIL import Image
    exif = Image.Exif()
    if taken:
        exif.get_ifd(0x8769)[0x9003] = taken
    Image.new("RGB", size, "white").save(path, exif=exif)


def test_photos_from_month_start_to_60_days_after(tmp_path):
    pytest.importorskip("PIL")
    now = datetime(2026, 11, 12, 9, 0)
    (tmp_path / "2026-10-22_123342.jpg").write_bytes(b"x")           # 단축어 파일 이름
    _set_mtime(tmp_path / "2026-10-22_123342.jpg", now)
    _jpeg(tmp_path / "IMG_0001.JPG", "2026:10:07 09:30:00")           # 사진 속 촬영 날짜
    _set_mtime(tmp_path / "IMG_0001.JPG", now)                         # iCloud가 내려받은 시각
    _jpeg(tmp_path / "IMG_0002.JPG", "2026:11:01 08:00:00")           # 다음 달에 찍은 사진도 봄
    _set_mtime(tmp_path / "IMG_0002.JPG", now)
    (tmp_path / "2026-09-30_235959.jpg").write_bytes(b"x")            # 처리할 달 전에 찍은 사진은 제외
    (tmp_path / "IMG_0003.PNG").write_bytes(b"x")                      # 촬영 날짜 없음 → 파일 날짜
    _set_mtime(tmp_path / "IMG_0003.PNG", datetime(2026, 10, 15, 18, 0))
    (tmp_path / "IMG_0004.JPG").write_bytes(b"x")                      # 9월 파일은 열지도 않음
    _set_mtime(tmp_path / "IMG_0004.JPG", datetime(2026, 9, 3))
    (tmp_path / "2026-10-01_000000_보정됨.png").write_bytes(b"x")      # 보정본은 제외
    (tmp_path / "memo.txt").write_text("x")

    cache = {}
    found = ca.list_originals({"영수증_원본_폴더": str(tmp_path)}, ca.TargetMonth(2026, 10), cache)
    assert [(p.name, t) for p, t in found] == [
        ("IMG_0001.JPG", datetime(2026, 10, 7, 9, 30)),
        ("IMG_0003.PNG", datetime(2026, 10, 15, 18, 0)),
        ("2026-10-22_123342.jpg", datetime(2026, 10, 22, 12, 33, 42)),
        ("IMG_0002.JPG", datetime(2026, 11, 1, 8, 0)),
    ]
    assert "IMG_0004.JPG" not in cache            # 오래된 파일은 열어 보지 않음
    assert cache["IMG_0001.JPG"]["taken"] == "2026-10-07T09:30:00"


def test_taken_date_cache_is_reused(tmp_path, monkeypatch):
    p = tmp_path / "IMG_0001.JPG"
    p.write_bytes(b"x")
    st = p.stat()
    cache = {"IMG_0001.JPG": {"size": st.st_size, "mtime": int(st.st_mtime), "taken": "2026-10-07T09:30:00"}}
    monkeypatch.setattr(ca, "exif_taken_at", lambda path: pytest.fail("캐시를 써야 합니다"))
    assert ca.photo_taken_at(p, st, cache) == datetime(2026, 10, 7, 9, 30)


def test_photo_window_crosses_year_and_is_configurable(tmp_path):
    for name in ("2026-11-30_235959.jpg", "2026-12-31_235959.jpg", "2027-02-28_120000.jpg",
                 "2027-03-01_235959.jpg", "2027-03-02_000000.jpg"):
        (tmp_path / name).write_bytes(b"x")
    cfg = {"영수증_원본_폴더": str(tmp_path)}
    found = ca.list_originals(cfg, ca.TargetMonth(2026, 12))          # 12월 + 60일 = 3월 1일까지
    assert [p.name for p, _ in found] == ["2026-12-31_235959.jpg", "2027-02-28_120000.jpg", "2027-03-01_235959.jpg"]
    cfg["영수증_촬영_허용_일수"] = 0
    assert [p.name for p, _ in ca.list_originals(cfg, ca.TargetMonth(2026, 12))] == ["2026-12-31_235959.jpg"]


# ---------- HEIC ----------

def test_heic_read_date_and_upload_as_jpeg(tmp_path):
    pytest.importorskip("cv2")
    pillow_heif = pytest.importorskip("pillow_heif")
    from PIL import Image
    pillow_heif.register_heif_opener()
    src = tmp_path / "IMG_0005.HEIC"
    exif = Image.Exif()
    exif.get_ifd(0x8769)[0x9003] = "2026:10:09 13:05:00"
    Image.new("RGB", (60, 30), "white").save(src, exif=exif)

    assert ca.exif_taken_at(src) == datetime(2026, 10, 9, 13, 5)
    img = ca.imread_unicode(src)
    assert img.shape == (30, 60, 3)

    rec = ca.Receipt(original=src, ocr_file=src)
    out = ca.prepare_upload_file(rec, ca.default_settings(), tmp_path / "w")
    assert out.suffix == ".jpg" and out.exists()
    with Image.open(out) as im:
        assert im.format == "JPEG" and im.size == (60, 30)


def test_jpeg_under_limit_is_uploaded_as_is(tmp_path):
    pytest.importorskip("cv2")
    pytest.importorskip("PIL")
    src = tmp_path / "IMG_0006.JPG"
    _jpeg(src)
    rec = ca.Receipt(original=src, ocr_file=src)
    assert ca.prepare_upload_file(rec, ca.default_settings(), tmp_path / "w") == src


# ---------- 영수증 목록 ----------

def test_list_puts_excluded_photos_last_and_user_can_bring_back(tmp_path, monkeypatch):
    src = tmp_path / "src"
    src.mkdir()
    texts = {"2026-10-05_090000.jpg": "가족사진", "2026-10-06_090000.jpg": "합계 4,500",
             "2026-10-07_090000.jpg": "합계 9,000"}
    for name in texts:
        (src / name).write_bytes(b"x")

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
        ca.build_receipts(cfg, tm, ca.OLD_DEFAULT_HOSPITALS, results, tmp_path / "w", log)
        list_path = ca.list_csv_path(results, tm)
        rows = ca.read_csv_dicts(list_path)
        assert [r["파일"] for r in rows] == ["2026-10-06_090000.jpg", "2026-10-07_090000.jpg", "2026-10-05_090000.jpg"]
        last = rows[-1]
        assert (last["사용(O/X)"], last["[자동] 사용"], last["확인 필요"]) == ("X", "X", ca.NOT_RECEIPT_NOTE)
        assert rows[0]["촬영 시각"] == "2026-10-06 09:00"

        last["사용(O/X)"] = "O"                      # 사실은 영수증이었음
        ca.write_csv(list_path, ca.LIST_HEADER, [[r[h] for h in ca.LIST_HEADER] for r in rows])
        recs = ca.build_receipts(cfg, tm, ca.OLD_DEFAULT_HOSPITALS, results, tmp_path / "w", log)
        family = next(r for r in recs if r.original.name == "2026-10-05_090000.jpg")
        assert family.use and family.edited == ["사용"]
        assert ca.read_csv_dicts(list_path)[0]["파일"] == "2026-10-05_090000.jpg"   # 이제 촬영 순서대로 맨 위
    finally:
        log.close()


def test_photo_is_read_as_is_and_only_unreadable_receipts_are_corrected(tmp_path, monkeypatch):
    pytest.importorskip("cv2")
    pytest.importorskip("PIL")
    src = tmp_path / "src"
    src.mkdir()
    for name in ("2026-10-05_090000.jpg", "2026-10-06_090000.jpg", "2026-10-07_090000.jpg"):
        _jpeg(src / name, size=(30, 30))
    texts = {"2026-10-05_090000.jpg": "2026.10.05 합계 3,000 카드",      # 잘 읽힘 → 보정 안 함
             "2026-10-06_090000.jpg": "합계 카드 결제",                   # 금액·날짜를 못 읽음 → 보정해서 다시
             "pre_2026-10-06_090000.jpg": "2026.10.06 합계 4,500 카드",
             "2026-10-07_090000.jpg": "풍경"}                           # 영수증 아님 → 보정 안 함
    read = []
    work = tmp_path / "w"

    class FakeOcr:
        def __init__(self, temp_dir):
            pass

        def read(self, path):
            read.append(path.name)
            assert path.exists()
            return texts[path.name]

    monkeypatch.setattr(ca, "WindowsOcr", FakeOcr)
    cfg = ca.default_settings()
    cfg.update({"영수증_원본_폴더": str(src), "보정본_폴더": str(src / "p"), "영수증_보정_사용": True})
    results = tmp_path / "결과"
    results.mkdir()
    lines = []
    log = ca.Log(tmp_path / "log", sink=lines.append)
    try:
        recs = ca.build_receipts(cfg, ca.TargetMonth(2026, 10), ca.OLD_DEFAULT_HOSPITALS, results, work, log)
        assert read == ["2026-10-05_090000.jpg", "2026-10-06_090000.jpg", "pre_2026-10-06_090000.jpg", "2026-10-07_090000.jpg"]
        assert not (src / "p").exists()                                  # 보정본 폴더(큰 png)를 만들지 않음
        assert list(work.glob("*")) == []                                # 보정 사본은 읽고 바로 지움
        fixed = next(r for r in recs if r.original.name == "2026-10-06_090000.jpg")
        assert fixed.total == 4500 and fixed.main_date == "2026-10-06" and fixed.ocr_file.name == "2026-10-06_090000.jpg"
        assert any("보정해서 다시 읽은 영수증 1장, 그중 더 잘 읽힌 것 1장" in l for l in lines)

        read.clear()                                                     # 두 번째 실행은 기억해 둔 결과를 씀
        ca.build_receipts(cfg, ca.TargetMonth(2026, 10), ca.OLD_DEFAULT_HOSPITALS, results, work, log)
        assert read == []
    finally:
        log.close()


def test_correction_kept_only_when_it_reads_more(tmp_path, monkeypatch):
    pytest.importorskip("cv2")
    pytest.importorskip("PIL")
    src = tmp_path / "src"
    src.mkdir()
    _jpeg(src / "2026-10-06_090000.jpg", size=(30, 30))
    texts = {"2026-10-06_090000.jpg": "합계 4,500 카드", "pre_2026-10-06_090000.jpg": "합"}   # 보정본이 더 못 읽음

    class FakeOcr:
        def __init__(self, temp_dir):
            pass

        def read(self, path):
            return texts[path.name]

    monkeypatch.setattr(ca, "WindowsOcr", FakeOcr)
    cfg = ca.default_settings()
    cfg.update({"영수증_원본_폴더": str(src), "보정본_폴더": str(src / "p")})
    results = tmp_path / "결과"
    results.mkdir()
    log = ca.Log(tmp_path / "log", sink=lambda s: None)
    try:
        recs = ca.build_receipts(cfg, ca.TargetMonth(2026, 10), [], results, tmp_path / "w", log)
    finally:
        log.close()
    assert recs[0].total == 4500 and recs[0].ocr_text == "합계 4,500 카드"


def test_ocr_temp_copy_is_jpg_and_deleted(tmp_path):
    pytest.importorskip("cv2")
    pytest.importorskip("PIL")
    photo = tmp_path / "IMG_0001.JPG"
    _jpeg(photo, size=(40, 20))
    ocr = ca.WindowsOcr.__new__(ca.WindowsOcr)          # Windows OCR 없이 임시 사본만 확인
    ocr.max_dim, ocr.temp_dir = 10000, tmp_path / "tmp"
    seen = []

    async def fake_read(path):
        seen.append((path.name, path.exists()))
        return "ok"

    ocr._read = fake_read
    assert ocr.read(photo) == "ok"
    assert seen == [("ocr_IMG_0001.jpg", True)]
    assert list((tmp_path / "tmp").glob("*")) == []


def test_old_processed_folder_notice(tmp_path):
    folder = tmp_path / "Concur_Expenses_Processed"
    folder.mkdir()
    (folder / "2026-10-05_090000_보정됨.png").write_bytes(b"x" * 2048)
    note = ca.old_processed_note({"보정본_폴더": str(folder)})
    assert "보정 사진 1개" in note and "지워도 됩니다" in note
    assert ca.old_processed_note({"보정본_폴더": str(tmp_path / "없음")}) == ""
