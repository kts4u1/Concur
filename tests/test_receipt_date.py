# -*- coding: utf-8 -*-
"""어느 달 영수증인지: 영수증에 적힌 날짜가 최우선, 촬영 날짜는 참고만"""

from datetime import datetime
from pathlib import Path

import concur_auto as ca

OCT = ca.TargetMonth(2026, 10)


def placed(text, taken, name="IMG_0001.JPG"):
    p = Path(name)
    rec = ca.analyze_receipt(p, p, text, ca.default_settings(), ca.DEFAULT_HOSPITALS, taken)
    ca.place_in_month(rec, OCT)
    return rec


def test_receipt_photographed_a_month_later_belongs_to_receipt_month():
    rec = placed("2026.10.27 합계 12,000", datetime(2026, 11, 25, 10, 0))
    assert rec.use and rec.skip_reason == ""
    assert rec.main_date == "2026-10-27"
    assert not any(i.startswith("날짜") for i in ca.receipt_issues(rec, [rec]))


def test_receipt_date_beats_photo_date():
    rec = placed("2026.11.02 합계 12,000", datetime(2026, 10, 30, 10, 0))
    assert not rec.use and rec.skip_reason == "다른 달 영수증(2026-11-02)(자동 제외)"


def test_month_date_is_chosen_when_receipt_has_several_dates():
    rec = placed("출력 2026-11-03\n거래일 2026.10.28 합계 7,000", datetime(2026, 11, 3, 9, 0))
    assert rec.use and rec.main_date == "2026-10-28"


def test_unreadable_date_uses_photo_date_only_when_taken_in_month():
    rec = placed("합계 4,500 카드", datetime(2026, 10, 9, 12, 0))
    assert rec.use and rec.main_date == "2026-10-09"
    assert ca.receipt_issues(rec, [rec])[0] == "날짜 못 읽음(촬영 날짜 사용)"

    rec = placed("합계 4,500 카드", datetime(2026, 11, 9, 12, 0))
    assert not rec.use and rec.skip_reason == ca.UNDATED_NOTE


def test_find_receipt_prefers_receipt_date_over_photo_date():
    by_photo = placed("2026.10.03 합계 4,500", datetime(2026, 10, 5, 9, 0), "a.jpg")   # 10/5에 찍은 10/3 영수증
    by_receipt = placed("2026.10.05 합계 4,500", datetime(2026, 11, 1, 9, 0), "b.jpg")
    assert "2026-10-05" in by_photo.dates
    assert ca.find_receipt([by_photo, by_receipt], "2026-10-05", 4500) is by_receipt
    by_receipt.used_for = "1번"
    assert ca.find_receipt([by_photo, by_receipt], "2026-10-05", 4500) is by_photo   # 촬영 날짜는 참고 후보


def test_writing_a_date_brings_receipt_back(tmp_path, monkeypatch):
    src = tmp_path / "src"
    src.mkdir()
    texts = {"2026-11-09_120000.jpg": "합계 4,500 카드",        # 날짜를 못 읽은 영수증을 11월에 찍음
             "2026-11-10_120000.jpg": "풍경"}                    # 영수증 아님
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
    try:
        ca.build_receipts(cfg, OCT, ca.DEFAULT_HOSPITALS, results, tmp_path / "w", log)
        list_path = ca.list_csv_path(results, OCT)
        rows = ca.read_csv_dicts(list_path)
        assert [(r["파일"], r["사용(O/X)"], r["확인 필요"]) for r in rows] == [
            ("2026-11-09_120000.jpg", "X", ca.UNDATED_NOTE),
            ("2026-11-10_120000.jpg", "X", ca.NOT_RECEIPT_NOTE),
        ]
        rows[0]["날짜"] = "2026-10-15"          # 사용자가 영수증 날짜를 적음 (사용 칸은 그대로 X)
        rows[1]["날짜"] = "2026-10-20"          # 영수증 아님으로 빠졌던 것도 날짜를 적으면 영수증으로 봄
        ca.write_csv(list_path, ca.LIST_HEADER, [[r[h] for h in ca.LIST_HEADER] for r in rows])

        recs = ca.build_receipts(cfg, OCT, ca.DEFAULT_HOSPITALS, results, tmp_path / "w", log)
        assert [(r.original.name, r.use, r.main_date) for r in recs] == [
            ("2026-11-09_120000.jpg", True, "2026-10-15"),
            ("2026-11-10_120000.jpg", True, "2026-10-20"),
        ]
        rows = ca.read_csv_dicts(list_path)
        assert [r["사용(O/X)"] for r in rows] == ["O", "O"]
    finally:
        log.close()


def test_default_folder_finds_similar_name(tmp_path):
    drive = tmp_path / "iCloudDrive"
    (drive / "Concur_Expense").mkdir(parents=True)
    assert ca.default_receipt_folder(tmp_path) == drive / "Concur_Expense"
    (drive / "Concur_Expenses").mkdir()
    assert ca.default_receipt_folder(tmp_path) == drive / "Concur_Expenses"


def test_default_folder_when_nothing_found(tmp_path):
    assert ca.default_receipt_folder(tmp_path) == tmp_path / "iCloudDrive" / "Concur_Expenses"


def test_missing_folder_message_suggests_similar_folder(tmp_path):
    (tmp_path / "Concur_Expense").mkdir()
    try:
        ca.list_originals({"영수증_원본_폴더": str(tmp_path / "Concur_Expenses")}, OCT)
    except RuntimeError as e:
        assert "Concur_Expense" in str(e) and "[찾기]" in str(e)
    else:
        raise AssertionError("폴더가 없으면 알려야 합니다")
