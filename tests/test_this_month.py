# -*- coding: utf-8 -*-
"""이번 달(아직 끝나지 않은 달)도 처리할 수 있는지"""

from datetime import date

import concur_auto as ca

TODAY = date(2026, 11, 12)
NOV = ca.TargetMonth(2026, 11)


def test_month_list_starts_with_this_month_but_default_stays_last_month():
    assert ca.recent_months(2, TODAY, include_current=True) == ["2026-11", "2026-10", "2026-09"]   # 이번 달 + 지난 2개월
    assert ca.recent_months(2, TODAY) == ["2026-10", "2026-09"]
    assert ca.resolve_target_month(None, TODAY).prefix == "2026-10"
    assert ca.recent_months(1, date(2027, 1, 3), include_current=True) == ["2027-01", "2026-12"]


def test_unfinished_month():
    assert ca.is_unfinished_month(NOV, TODAY)
    assert ca.is_unfinished_month(ca.TargetMonth(2026, 12), TODAY)
    assert not ca.is_unfinished_month(ca.TargetMonth(2026, 10), TODAY)
    assert not ca.is_unfinished_month(ca.TargetMonth(2025, 12), TODAY)


def test_this_month_report_dates():
    assert NOV.report_name("{month_name} Expenses") == "November Expenses"
    assert (NOV.first_day_us, NOV.last_day_us) == ("11/01/2026", "11/30/2026")


def test_processing_this_month(tmp_path, monkeypatch):
    src = tmp_path / "src"
    src.mkdir()
    texts = {"2026-10-31_200000.jpg": "2026.10.31 합계 8,000",       # 11월 전에 찍음 → 읽지 않음
             "2026-11-03_090000.jpg": "2026.11.03 주차 합계 5,000",
             "2026-11-05_090000.jpg": "2026.10.30 합계 3,000",        # 10월 영수증을 11월에 찍음
             "2026-11-10_090000.jpg": "합계 4,500 카드"}               # 날짜 못 읽음 → 촬영 날짜(11월) 사용
    for name in texts:
        (src / name).write_bytes(b"x")
    read = []

    class FakeOcr:
        def __init__(self, temp_dir):
            pass

        def read(self, path):
            read.append(path.name)
            return texts[path.name]

    monkeypatch.setattr(ca, "WindowsOcr", FakeOcr)
    cfg = ca.default_settings()
    cfg.update({"영수증_원본_폴더": str(src), "보정본_폴더": str(src / "p"), "영수증_보정_사용": False})
    results = tmp_path / "결과"
    results.mkdir()
    lines = []
    log = ca.Log(tmp_path / "log", sink=lines.append)
    try:
        recs = ca.build_receipts(cfg, NOV, ca.OLD_DEFAULT_HOSPITALS, results, tmp_path / "w", log)
    finally:
        log.close()
    assert "2026-10-31_200000.jpg" not in read
    assert [(r.original.name, r.use, r.main_date) for r in recs] == [
        ("2026-11-03_090000.jpg", True, "2026-11-03"),
        ("2026-11-05_090000.jpg", False, "2026-10-30"),
        ("2026-11-10_090000.jpg", True, "2026-11-10"),
    ]
    assert recs[1].skip_reason == "다른 달 영수증(2026-10-30)(자동 제외)"
    assert (results / "영수증목록_2026-11.csv").exists()
    assert ca.find_receipt(recs, "2026-11-03", 5000) is recs[0]
    assert ca.find_receipt(recs, "2026-10-30", 3000) is None       # 11월 리포트에는 붙이지 않음


def test_run_job_logs_unfinished_month_note(tmp_path, monkeypatch):
    src = tmp_path / "src"
    src.mkdir()
    base = tmp_path / "app"
    base.mkdir()
    ca.save_settings(base / ca.SETTINGS_FILE_NAME, {"영수증_원본_폴더": str(src), "결과_자동_열기": False})
    this_month = ca.recent_months(1, include_current=True)[0]

    class UI:
        def __init__(self):
            self.lines, self.infos = [], []

        def log(self, line):
            self.lines.append(line)

        def info(self, title, msg):
            self.infos.append(title)

        def open(self, path):
            pass

    ui = UI()
    ca.run_job("ocr", this_month, ui, base)
    assert any("아직 끝나지 않은 달입니다" in line for line in ui.lines)
    assert ui.infos == ["영수증 분석 완료"]
