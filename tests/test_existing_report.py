# -*- coding: utf-8 -*-
"""② 전체 실행: Concur에 같은 이름의 리포트(예: Monthly Expense for Nov)가 있으면 이어서 넣을지 묻기
(예전 ③ '만든 리포트에 입력만' 버튼을 없애고 ②로 합침)"""

import os
from pathlib import Path
from urllib.parse import urlparse

import pytest

import concur_auto as ca

NOV = ca.TargetMonth(2026, 11)
BASE = "https://us2.concursolutions.com"


# ---------- run_job: 찾았을 때 / 못 찾았을 때 ----------

class FakeBot:
    calls, found, empty = [], None, False

    def __init__(self, cfg, log, ui, stop_flag):
        self.page = None

    def start(self):
        pass

    def wait_login(self, need_home=True):
        pass

    def find_report(self, tm):
        FakeBot.calls.append("find")
        return FakeBot.found

    def open_report(self, item):
        FakeBot.calls.append(f"open:{item}")

    def report_is_empty(self):
        return FakeBot.empty

    def go_home(self):
        FakeBot.calls.append("home")

    def create_report(self, tm):
        FakeBot.calls.append("create")

    def add_available_expenses(self, tm):
        FakeBot.calls.append("add")

    def open_first_expense(self):
        FakeBot.calls.append("first")

    def process_expenses(self, receipts, work_dir, stop_flag):
        return []

    def wait_until_closed(self, stop_flag=None):
        pass

    def stop(self):
        pass


class UI:
    def __init__(self, answer):
        self.answer, self.asked, self.lines = answer, [], []

    def log(self, line):
        self.lines.append(line)

    def info(self, title, msg):
        pass

    def confirm(self, title, msg):
        return True

    def yes_no_cancel(self, title, msg):
        self.asked.append((title, msg))
        return self.answer

    def open(self, path):
        pass


def run(tmp_path, monkeypatch, found, answer, empty=False):
    src = tmp_path / "src"
    src.mkdir()
    base = tmp_path / "app"
    base.mkdir()
    ca.save_settings(base / ca.SETTINGS_FILE_NAME, {"영수증_원본_폴더": str(src), "결과_자동_열기": False})
    FakeBot.calls, FakeBot.found, FakeBot.empty = [], found, empty
    monkeypatch.setattr(ca, "ConcurBot", FakeBot)
    ui = UI(answer)
    ca.run_job("full", "2026-11", ui, base)
    return FakeBot.calls, ui


def test_no_report_yet_creates_one(tmp_path, monkeypatch):
    calls, ui = run(tmp_path, monkeypatch, found=None, answer=True)
    assert calls == ["find", "create", "add", "first"] and ui.asked == []


def test_existing_report_continue(tmp_path, monkeypatch):
    calls, ui = run(tmp_path, monkeypatch, found="R1", answer=True)
    assert ui.asked[0][0] == "이미 만든 리포트" and "Monthly Expense for Nov" in ui.asked[0][1]
    assert calls == ["find", "open:R1", "first"]                     # 이미 든 경비에 이어서(새로 만들지 않음)
    assert any("이어서 입력합니다" in l for l in ui.lines)


def test_existing_empty_report_gets_expenses(tmp_path, monkeypatch):
    calls, _ = run(tmp_path, monkeypatch, found="R1", answer=True, empty=True)
    assert calls == ["find", "open:R1", "add", "first"]              # 빈 리포트면 그 달 경비를 넣음


def test_existing_report_but_make_new(tmp_path, monkeypatch):
    calls, _ = run(tmp_path, monkeypatch, found="R1", answer=False)
    assert calls == ["find", "home", "create", "add", "first"]


def test_existing_report_cancel(tmp_path, monkeypatch):
    calls, ui = run(tmp_path, monkeypatch, found="R1", answer=None)
    assert calls == ["find"] and any("중단" in l for l in ui.lines)


# ---------- 가짜 Concur에서 리포트 찾기 ----------

sync_api = pytest.importorskip("playwright.sync_api")

LIST = """<!doctype html><meta charset="utf-8"><h1>Manage Expenses</h1><h2>Active Reports</h2>
<div class="card" onclick="location.href='/nui/expense/reports/OLD'"><span>Monthly Expense for Oct</span></div>
{card}
<div>Create New Report</div>"""
CARD = """<div class="card" onclick="location.href='/nui/expense/reports/R11'"><span>Monthly Expense for Nov</span></div>"""
REPORT = """<!doctype html><meta charset="utf-8"><h1>Monthly Expense for Nov</h1>
<div class="entries-list"><div role="row"><input type="checkbox"><span>11/03/2026</span></div></div>"""
HOME = """<!doctype html><meta charset="utf-8"><a>Create Report</a>"""


@pytest.fixture
def bot_on():
    exe = os.environ.get("PW_CHROMIUM") or "/opt/pw-browsers/chromium"
    with sync_api.sync_playwright() as p:
        browser = p.chromium.launch(executable_path=exe if Path(exe).exists() else None)

        def make(card):
            ctx = browser.new_context()

            def handle(route, request):
                path = urlparse(request.url).path
                body = REPORT if "/reports/" in path else LIST.format(card=card) if path == "/nui/expense" else HOME
                route.fulfill(status=200, content_type="text/html; charset=utf-8", body=body)

            ctx.route("**/*", handle)
            page = ctx.new_page()
            page.set_default_timeout(5000)
            cfg = ca.default_settings()
            cfg["리포트_찾기_대기_초"] = 1
            bot = ca.ConcurBot(cfg, lambda s="": None)
            bot.page, bot.ctx, bot.timeout_ms = page, ctx, 5000
            page.goto(BASE + "/home")
            return bot

        yield make
        browser.close()


def test_find_and_open_existing_report(bot_on):
    bot = bot_on(CARD)
    item = bot.find_report(NOV)
    assert item is not None
    bot.open_report(item)
    assert bot.page.url.endswith("/nui/expense/reports/R11")            # 10월이 아니라 11월 리포트
    assert bot.report_is_empty() is False


def test_no_report_goes_back_home(bot_on):
    bot = bot_on("")
    assert bot.find_report(NOV) is None
    assert urlparse(bot.page.url).path == "/home"                       # 새로 만들 수 있게 홈 화면으로
