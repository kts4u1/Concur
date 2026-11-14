# -*- coding: utf-8 -*-
"""② 전체 실행이 끝난 뒤: 창에서는 경비처리 파일을 엑셀로 열지 않고 확인할 경비를 안내에 바로 보여 줌. 아이콘."""

import concur_auto as ca

RESULTS = [
    {"순번": 1, "Concur 날짜": "11/04/2026", "Concur 금액": "3,100", "가게 이름": "", "영수증": "IMG_5619.HEIC",
     "분류": "주차", "병원": "", "병원 근거": "", "도시": "", "처리": "", "확인 필요": "", "오류": ""},
    {"순번": 2, "Concur 날짜": "11/02/2026", "Concur 금액": "2,000", "가게 이름": "", "영수증": "",
     "분류": "", "병원": "", "병원 근거": "", "도시": "", "처리": "", "확인 필요": "영수증 없음", "오류": ""},
]


class FakeBot:
    def __init__(self, cfg, log, ui, stop_flag):
        pass

    def start(self):
        pass

    def wait_login(self, need_home=True):
        pass

    def find_report(self, tm):
        return None

    def create_report(self, tm):
        pass

    def add_available_expenses(self, tm):
        pass

    def open_first_expense(self):
        pass

    def process_expenses(self, receipts, work_dir, stop_flag):
        return [dict(r) for r in RESULTS]

    def wait_until_closed(self, stop_flag=None):
        pass

    def stop(self):
        pass


class UI:
    def __init__(self, window):
        self.opened, self.infos = [], []
        if window:
            self.review = lambda path, wait=False: True          # 창(GuiBridge)에만 있는 기능

    def log(self, line):
        pass

    def info(self, title, msg):
        self.infos.append((title, msg))

    def confirm(self, title, msg):
        return True

    def yes_no_cancel(self, title, msg):
        return True

    def open(self, path):
        self.opened.append(path.name)


def run(tmp_path, monkeypatch, window):
    src = tmp_path / "src"
    src.mkdir()
    base = tmp_path / "app"
    base.mkdir()
    ca.save_settings(base / ca.SETTINGS_FILE_NAME, {"영수증_원본_폴더": str(src), "결과_자동_열기": True})
    monkeypatch.setattr(ca, "ConcurBot", FakeBot)
    ui = UI(window)
    ca.run_job("full", "2026-11", ui, base)
    return ui


def test_window_does_not_open_results_file(tmp_path, monkeypatch):
    ui = run(tmp_path, monkeypatch, window=True)
    assert not any(name.startswith("경비처리_") for name in ui.opened)
    title, msg = ui.infos[-1]
    assert title == "완료" and "확인 필요 1건" in msg
    assert "· 11/02/2026 2,000원: 영수증 없음" in msg                 # 파일을 안 열어도 무엇을 볼지 알 수 있음
    assert (tmp_path / "app" / "결과" / "경비처리_2026-11.csv").is_file()   # 파일은 그대로 남김


def test_command_prompt_still_opens_results_file(tmp_path, monkeypatch):
    ui = run(tmp_path, monkeypatch, window=False)
    assert "경비처리_2026-11.csv" in ui.opened


def test_check_lines_limit():
    checks = [{"Concur 날짜": f"11/{i:02d}/2026", "Concur 금액": "1,000", "확인 필요": "영수증 없음"} for i in range(1, 12)]
    text = ca.check_lines(checks, limit=3)
    assert text.count("· 11/") == 3 and "그 밖에 8건" in text
    assert ca.check_lines([]) == ""


def test_vendor_text_without_checkbox_label():
    rows = [["Select expense, Parking & Toll, KRW 2,000, date, 11/02/2026", "11/02/2026", "KAKAO T PARKING", "2,000"]]
    assert ca.parse_entry_rows(rows) == [("2026-11-02", 2000, "11/02/2026 KAKAO T PARKING 2,000")]


def test_icon_file_ships_with_program():
    ico = ca.resource_path(ca.ICON_FILE_NAME)
    assert ico.is_file() and ico.read_bytes()[:4] == b"\x00\x00\x01\x00"     # Windows 아이콘 파일
    ca.set_window_icon(None)                                                   # Windows가 아니면 아무것도 안 함
