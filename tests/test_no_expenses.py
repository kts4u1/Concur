# -*- coding: utf-8 -*-
"""넣을 경비가 없을 때(No Available Expenses, 그 달 경비 없음, 빈 리포트): 오류 없이 끝내고 Concur 창은 그대로"""

import os
import time
from pathlib import Path
from urllib.parse import urlparse

import pytest

import concur_auto as ca

sync_api = pytest.importorskip("playwright.sync_api")

REPORT = "https://us2.concursolutions.com/nui/expense/reports/R1"

EMPTY_REPORT = f"""<!doctype html><meta charset="utf-8"><h1>November Expenses</h1><p>Not Submitted</p>
<div class="report-empty-state-actions__card-body" onclick="location.href='{REPORT}?modal=add-expense&view=available_expenses'">
  Add <b>Available Expenses</b></div>"""

NO_AVAILABLE = """<!doctype html><meta charset="utf-8"><h2>Add Expense to Report</h2>
<a>Manually Create Expense</a><a>Select from Available Expenses (0)</a>
<h2>No Available Expenses</h2><p>New/Incoming expenses will be added to this list.</p>
<button disabled>Add To Report</button><button>Cancel</button>"""

AVAILABLE = """<!doctype html><meta charset="utf-8"><h2>Add Expense to Report</h2>
<div role="row"><input type="checkbox" class="sapcnqr-checkbox__input"><span>Select all</span></div>
{rows}
<button onclick="location.href='{report}?done=1'">Add To Report</button>"""

ROW = '<div role="row"><input type="checkbox" class="sapcnqr-checkbox__input"><span>{d}</span><span>KRW 4,500</span></div>'

FILLED_REPORT = """<!doctype html><meta charset="utf-8"><h1>November Expenses</h1>
<div class="entries-list"><div role="row"><input type="checkbox"><span>11/03/2026</span></div></div>"""


class Site:
    def __init__(self, available_dates=None):
        self.available_dates = available_dates

    def handle(self, route, request):
        u = urlparse(request.url)
        if "view=available_expenses" in u.query:
            if self.available_dates is None:
                body = NO_AVAILABLE
            else:
                body = AVAILABLE.format(rows="".join(ROW.format(d=d) for d in self.available_dates), report=REPORT)
        elif "done=1" in u.query:
            body = FILLED_REPORT
        else:
            body = EMPTY_REPORT
        route.fulfill(status=200, content_type="text/html; charset=utf-8", body=body)


@pytest.fixture
def browser():
    exe = os.environ.get("PW_CHROMIUM") or "/opt/pw-browsers/chromium"
    with sync_api.sync_playwright() as p:
        b = p.chromium.launch(executable_path=exe if Path(exe).exists() else None)
        yield b
        b.close()


def bot_on(browser, site):
    ctx = browser.new_context()
    ctx.route("**/*", site.handle)
    page = ctx.new_page()
    page.set_default_timeout(5000)
    lines = []
    bot = ca.ConcurBot(ca.default_settings(), lambda s="": lines.append(s))
    bot.page, bot.ctx, bot.timeout_ms = page, ctx, 5000
    page.goto(REPORT)
    return bot, lines


def test_no_available_expenses_ends_without_error(browser):
    bot, lines = bot_on(browser, Site(available_dates=None))
    started = time.time()
    with pytest.raises(ca.NoExpenses, match="사용 가능한 경비가 없습니다"):
        bot.add_available_expenses(ca.TargetMonth(2026, 11))
    assert time.time() - started < 5                          # 예전에는 60초 기다린 뒤 오류
    assert "view=available_expenses" in bot.page.url        # 창은 그 화면 그대로


def test_no_expenses_in_that_month(browser):
    bot, lines = bot_on(browser, Site(available_dates=["10/05/2026", "10/30/2026"]))
    with pytest.raises(ca.NoExpenses, match="2026-11 경비가 없습니다"):
        bot.add_available_expenses(ca.TargetMonth(2026, 11))


def test_expenses_of_that_month_are_added(browser):
    bot, lines = bot_on(browser, Site(available_dates=["10/30/2026", "11/03/2026"]))
    bot.add_available_expenses(ca.TargetMonth(2026, 11))
    assert "사용 가능한 경비 중 2026-11 경비 1건 선택" in lines
    assert bot.page.url.endswith("done=1")


def test_empty_report_in_edit_mode(browser):
    bot, lines = bot_on(browser, Site())
    with pytest.raises(ca.NoExpenses, match="이 리포트에는 경비가 없습니다"):
        bot.open_first_expense()


def test_run_job_shows_notice_and_keeps_window(tmp_path, monkeypatch):
    src = tmp_path / "src"
    src.mkdir()
    base = tmp_path / "app"
    base.mkdir()
    ca.save_settings(base / ca.SETTINGS_FILE_NAME, {"영수증_원본_폴더": str(src), "결과_자동_열기": False})
    calls = []

    class FakeBot:
        def __init__(self, cfg, log, ui, stop_flag):
            self.page = None

        def start(self):
            calls.append("start")

        def wait_login(self, need_home=True):
            calls.append("login")

        def find_report(self, tm):
            calls.append("find")
            return None

        def create_report(self, tm):
            calls.append("report")

        def add_available_expenses(self, tm):
            raise ca.NoExpenses("Concur에 사용 가능한 경비가 없습니다(0건).")

        def wait_until_closed(self, stop_flag=None):
            calls.append("기다림(창 열어 둠)")

        def stop(self):
            calls.append("stop")

    class UI:
        infos, lines = [], []

        def log(self, line):
            self.lines.append(line)

        def info(self, title, msg):
            self.infos.append((title, msg))

        def confirm(self, title, msg):
            return True

        def yes_no_cancel(self, title, msg):
            return False

        def open(self, path):
            pass

    monkeypatch.setattr(ca, "ConcurBot", FakeBot)
    ui = UI()
    ca.run_job("full", "2026-11", ui, base)
    assert [t for t, _ in ui.infos] == ["넣을 경비 없음"]            # '오류'가 아님
    assert "Concur 창은 그대로 두니" in ui.infos[0][1]
    assert calls == ["start", "login", "find", "report", "기다림(창 열어 둠)", "stop"]
    assert not any(line.startswith("[") and "오류:" in line for line in ui.lines)
