# -*- coding: utf-8 -*-
"""리포트 만들기: 이름·Comment = 'Monthly Expense for Nov'(달만 바뀜), Business Purpose = 'service', 기간 = 그 달 1일~말일"""

import json
import os
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import pytest

import concur_auto as ca


@pytest.mark.parametrize("month, abbr", [(1, "Jan"), (5, "May"), (9, "Sep"), (11, "Nov"), (12, "Dec")])
def test_report_name_follows_month(month, abbr):
    fmt = ca.default_settings()["리포트_이름_형식"]
    assert ca.TargetMonth(2026, month).report_name(fmt) == f"Monthly Expense for {abbr}"


def test_old_report_name_setting_is_replaced(tmp_path):
    s = tmp_path / ca.SETTINGS_FILE_NAME
    s.write_text(json.dumps({"리포트_이름_형식": "{month_name} Expenses"}), encoding="utf-8")
    cfg, _ = ca.load_settings(s)
    assert cfg["리포트_이름_형식"] == "Monthly Expense for {month_abbr}"
    assert cfg["리포트_Business_Purpose"] == "service"
    s.write_text(json.dumps({"리포트_이름_형식": "{year}-{month} 경비"}), encoding="utf-8")
    assert ca.load_settings(s)[0]["리포트_이름_형식"] == "{year}-{month} 경비"      # 직접 바꾼 값은 그대로


sync_api = pytest.importorskip("playwright.sync_api")

BASE = "https://us2.concursolutions.com"
HOME = """<!doctype html><meta charset="utf-8"><h1>SAP Concur</h1>
<a href="/nui/expense?modal=header&context=new-report">Create Report</a>"""

FORM = """<!doctype html><meta charset="utf-8"><h2>Create New Report</h2>
<label for="name">Report Name *</label><input id="name">
<label for="policy">Policy *</label><select id="policy"><option>Korea Policy</option><option>Other Policy</option></select>
<label for="custom11-date-input-field-input">From Date *</label><input id="custom11-date-input-field-input">
<label for="custom12-date-input-field-input">To Date *</label><input id="custom12-date-input-field-input">
<label for="businessPurpose">Business Purpose</label><input id="businessPurpose">
<label for="cost">Cost Object</label><input id="cost" value="미리 들어간 값">
<label for="comment">Comment</label><textarea id="comment"></textarea>
<button onclick="go()">Create Report</button>
<script>
function go() {
  const v = id => document.getElementById(id).value;
  location.href = '/nui/expense/reports/R1?' + new URLSearchParams({name: v('name'), policy: v('policy'),
    start: v('custom11-date-input-field-input'), end: v('custom12-date-input-field-input'),
    purpose: v('businessPurpose'), cost: v('cost'), comment: v('comment')});
}
</script>"""


@pytest.fixture
def browser():
    exe = os.environ.get("PW_CHROMIUM") or "/opt/pw-browsers/chromium"
    with sync_api.sync_playwright() as p:
        b = p.chromium.launch(executable_path=exe if Path(exe).exists() else None)
        yield b
        b.close()


def create_on(browser, form):
    ctx = browser.new_context()

    def handle(route, request):
        path = urlparse(request.url).path
        body = form if path == "/nui/expense" else HOME if path in ("/", "/home") else "<p>report</p>"
        route.fulfill(status=200, content_type="text/html; charset=utf-8", body=body)

    ctx.route("**/*", handle)
    page = ctx.new_page()
    page.set_default_timeout(5000)
    lines = []
    bot = ca.ConcurBot(ca.default_settings(), lambda s="": lines.append(s))
    bot.page, bot.ctx, bot.timeout_ms = page, ctx, 5000
    page.goto(BASE + "/home")
    bot.create_report(ca.TargetMonth(2026, 11))
    return {k: v[0] for k, v in parse_qs(urlparse(page.url).query).items()}, lines


def test_create_report_fills_like_user_screen(browser):
    got, lines = create_on(browser, FORM)
    assert got["name"] == "Monthly Expense for Nov"
    assert got["comment"] == "Monthly Expense for Nov"                 # Comment = 리포트 이름
    assert got["purpose"] == "service"
    assert (got["start"], got["end"]) == ("11/01/2026", "11/30/2026")
    assert (got["policy"], got["cost"]) == ("Korea Policy", "미리 들어간 값")   # Concur가 넣은 값은 그대로
    assert any("Business Purpose=service" in line for line in lines)


def test_report_dates_found_by_from_to_labels(browser):
    form = FORM.replace("custom11-date-input-field-input", "d1").replace("custom12-date-input-field-input", "d2")
    got, _ = create_on(browser, form)                                  # 위치가 바뀌어도 From Date/To Date 칸 이름으로
    assert (got["start"], got["end"]) == ("11/01/2026", "11/30/2026")
