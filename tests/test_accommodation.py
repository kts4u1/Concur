# -*- coding: utf-8 -*-
"""숙박(Accommodation): Date Range 1박(새벽 영수증은 전날부터)을 달력에서 차례로 눌러 넣고,
저장 뒤 '필수 정보가 빠졌다'는 창에서 [Yes]를 누르기. (Itemizations는 test_itemization.py)"""

import os
from datetime import date, datetime
from pathlib import Path

import pytest

import concur_auto as ca


# ---------- 영수증 시각과 숙박 기간 ----------

@pytest.mark.parametrize("text, want", [
    ("거래일시 2026-11-03 02:15:33", (2, 15)),
    ("2026.11.02 23:10", (23, 10)),
    ("26/11/03 오전 01:20", (1, 20)),
    ("승인일시: 2026/11/03 PM 11:20", (23, 20)),
    ("2026-11-03\n02:40:11", (2, 40)),                       # 날짜 바로 다음 줄
    ("입실 22:30 퇴실 11:00", (22, 30)),
    ("영업시간 05:00~24:00\n합계 70,000", None),               # 영업시간은 영수증 시각이 아님
    ("모텔 합계 70,000", None),
])
def test_receipt_time(text, want):
    assert ca.find_receipt_time(text) == want


def rec_of(text, day, taken=None):
    return ca.Receipt(original=Path("a.jpg"), ocr_file=Path("a.jpg"), ocr_text=text, dates=[day], taken=taken)


def test_stay_dates_rules():
    D = date
    assert ca.stay_dates(rec_of("모텔 2026-11-02 21:15", "2026-11-02"), None)[:2] == (D(2026, 11, 2), D(2026, 11, 3))
    # 야근 뒤 새벽에 받은 영수증 → 전날부터
    assert ca.stay_dates(rec_of("모텔 2026-11-03 02:15", "2026-11-03"), None) == (D(2026, 11, 2), D(2026, 11, 3), "")
    assert ca.stay_dates(rec_of("2026-11-03 06:10", "2026-11-03"), None)[:2] == (D(2026, 11, 3), D(2026, 11, 4))
    assert ca.stay_dates(rec_of("2026-11-03 06:10", "2026-11-03"), None, cutoff_hour=7)[:2] == (D(2026, 11, 2), D(2026, 11, 3))
    # 달이 바뀌는 날
    assert ca.stay_dates(rec_of("2026-12-01 03:00", "2026-12-01"), None)[:2] == (D(2026, 11, 30), D(2026, 12, 1))
    # 영수증에 시각이 없으면 같은 날 찍은 사진 시각, 다른 날 찍었으면 영수증 날짜부터(확인 필요)
    assert ca.stay_dates(rec_of("모텔", "2026-11-03", datetime(2026, 11, 3, 1, 5)), None)[:2] == (D(2026, 11, 2), D(2026, 11, 3))
    start, end, note = ca.stay_dates(rec_of("모텔", "2026-11-03", datetime(2026, 11, 20, 1, 5)), None)
    assert (start, end) == (D(2026, 11, 3), D(2026, 11, 4)) and "시각" in note
    # 영수증이 없으면 거래일부터
    assert ca.stay_dates(None, "2026-11-30") == (D(2026, 11, 30), D(2026, 12, 1), "숙박 기간 확인(영수증이 없어 거래일부터 1박)")


# ---------- 가짜 Concur 화면 ----------

CALENDAR = """
<div id="cal" style="display:none; border:1px solid #888; width:300px">
  <div class="head"><button class="nav" id="prev" {prev}>{prevtext}</button>
    <span id="title"></span><button class="nav" id="next" {next}>{nexttext}</button></div>
  <div id="days"></div>
</div>
<script>
window.steps = window.steps || [];
const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
let view = new Date(2026, 10, 1), start = null;
const pad = n => String(n).padStart(2, '0');
const fmt = d => pad(d.getMonth() + 1) + '/' + pad(d.getDate()) + '/' + d.getFullYear();
function draw() {{
  document.getElementById('title').textContent = MONTHS[view.getMonth()] + ' ' + view.getFullYear();
  const first = new Date(view.getFullYear(), view.getMonth(), 1);
  const cur = new Date(first); cur.setDate(1 - first.getDay());
  let html = '';
  for (let i = 0; i < 42; i++) {{
    const d = new Date(cur); const out = d.getMonth() !== view.getMonth();
    html += '<button class="day' + (out && {mark_outside} ? ' outside' : '') + '" {label}data-t="' + d.getTime() + '" onclick="pick(this)">' + d.getDate() + '</button>';
    cur.setDate(cur.getDate() + 1);
  }}
  document.getElementById('days').innerHTML = html;
}}
function pick(b) {{
  const d = new Date(parseInt(b.dataset.t, 10));
  if (!start) {{ start = d; steps.push('start ' + fmt(d)); return; }}
  document.getElementById('range').value = fmt(start) + ' - ' + fmt(d);
  steps.push('range');
  document.getElementById('nights').textContent = Math.round((d - start) / 864e5);
  start = null; document.getElementById('cal').style.display = 'none';
}}
document.getElementById('prev').onclick = () => {{ view = new Date(view.getFullYear(), view.getMonth() - 1, 1); draw(); }};
document.getElementById('next').onclick = () => {{ view = new Date(view.getFullYear(), view.getMonth() + 1, 1); draw(); }};
function openCal() {{ document.getElementById('cal').style.display = 'block'; draw(); }}
</script>"""


def calendar(aria=True, mark_outside=True):
    return CALENDAR.format(
        prev='aria-label="Previous Month"' if aria else "", next='aria-label="Next Month"' if aria else "",
        prevtext="" if aria else "", nexttext="" if aria else "",
        mark_outside="true" if mark_outside else "false",
        label="aria-label=\"' + MONTHS[d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear() + '\" " if aria else "")


def range_page(aria=True, mark_outside=True):
    return f"""<!doctype html><meta charset="utf-8">
<label for="range">Date Range</label><input id="range" placeholder="MM/DD/YYYY - MM/DD/YYYY" readonly onclick="openCal()">
Nights: <b id="nights">0</b>
<label for="transactionDate-date-input-field-input">Transaction Date</label>
<input id="transactionDate-date-input-field-input" value="11/03/2026">
{calendar(aria, mark_outside)}
<style>.day {{ width:40px }} .nav {{ width:30px; height:20px }}</style>"""


@pytest.fixture
def browser():
    sync_api = pytest.importorskip("playwright.sync_api")
    exe = os.environ.get("PW_CHROMIUM") or "/opt/pw-browsers/chromium"
    with sync_api.sync_playwright() as p:
        b = p.chromium.launch(executable_path=exe if Path(exe).exists() else None)
        yield b
        b.close()


def bot_on(browser, html, cfg=None):
    pg = browser.new_page()
    pg.set_default_timeout(3000)
    pg.set_content(html)
    bot = ca.ConcurBot(cfg or ca.default_settings(), lambda s="": None)
    bot.page, bot.timeout_ms, bot.upload_wait_s, bot.upload_done_s = pg, 3000, 2, 2
    return bot, pg


@pytest.mark.parametrize("aria, mark_outside", [(True, True), (False, False)])
@pytest.mark.parametrize("start, end", [
    (date(2026, 11, 2), date(2026, 11, 3)),
    (date(2026, 11, 30), date(2026, 12, 1)),                 # 다음 달로 넘겨서 누름
    (date(2026, 10, 31), date(2026, 11, 1)),                 # 지난달로 갔다가 다시 이번 달
])
def test_pick_dates_on_calendar(browser, aria, mark_outside, start, end):
    bot, pg = bot_on(browser, range_page(aria, mark_outside))
    bot.pick_date_range(pg.locator("#range"), start, end)
    assert pg.input_value("#range") == f"{start:%m/%d/%Y} - {end:%m/%d/%Y}"
    assert pg.evaluate("steps") == [f"start {start:%m/%d/%Y}", "range"]          # 앞의 날 → 다음 날을 차례로 누름
    assert pg.inner_text("#nights") == "1"


def test_typed_when_no_calendar(browser):
    bot, pg = bot_on(browser, """<label for="r">Date Range</label><input id="r" placeholder="MM/DD/YYYY - MM/DD/YYYY">""")
    bot.pick_date_range(pg.locator("#r"), date(2026, 11, 2), date(2026, 11, 3))
    assert pg.input_value("#r") == "11/02/2026 - 11/03/2026"


EXPENSE = """<!doctype html><meta charset="utf-8">
<p>Expense Type *</p><div class="field__control"><span name="expenseType" onclick="openList(this)">Undefined</span></div>
<div id="rangeRow" style="display:none"><label for="range">Date Range</label>
  <input id="range" placeholder="MM/DD/YYYY - MM/DD/YYYY" readonly onclick="openCal()"> Nights: <b id="nights">0</b></div>
<label for="transactionDate-date-input-field-input">Transaction Date</label>
<input id="transactionDate-date-input-field-input" value="11/03/2026">
<label for="businessPurpose">Business Purpose</label><input id="businessPurpose" oninput="note('purpose')">
<label for="comment">Comment</label><textarea id="comment" oninput="note('comment')"></textarea>
<div id="pop" style="display:none"><div id="items"></div></div>
<script>
window.steps = [];
function note(s) { if (steps[steps.length - 1] !== s) steps.push(s); }
let target = null;
function openList(el) { target = el; document.getElementById('pop').style.display = 'block';
  document.getElementById('items').innerHTML = ['Accommodation', 'Parking & Toll'].map(t =>
    '<div class="sapcnqr-listbox-item__wrapper" onclick="pickType(this)">' + t + '</div>').join(''); }
function pickType(el) { target.textContent = el.textContent; note('type'); document.getElementById('pop').style.display = 'none';
  setTimeout(() => { document.getElementById('rangeRow').style.display = el.textContent === 'Accommodation' ? 'block' : 'none'; }, 300); }
</script>
<style>span[name] { display:inline-block; min-width:220px; min-height:18px; border:1px solid #999 }</style>
""" + calendar()

ACCOM = ["숙박", "모텔, 호텔", "Accommodation fee for service in hospital", "For service in hospital", "Accommodation", "",
         "", "", "", ""]


def test_accommodation_expense_gets_date_range(browser, tmp_path):
    cfg = ca.default_settings()
    table = tmp_path / ca.CATEGORY_FILE_NAME
    ca.save_category_rows(table, [ACCOM])
    ca.apply_category_table(cfg, table)
    bot, pg = bot_on(browser, EXPENSE, cfg)
    rec = rec_of("OO모텔 거래일시 2026-11-03 02:15 합계 70,000", "2026-11-03")
    rec.category = "숙박"
    checks = []
    done = bot.fill_expense(rec, None, checks)
    assert done[:3] == ["Expense Type=Accommodation", "Date Range=11/02/2026 - 11/03/2026", "Business Purpose 입력"]
    assert pg.evaluate("steps") == ["type", "start 11/02/2026", "range", "purpose", "comment"]   # 화면 위에서 아래 순서
    assert pg.input_value("#range") == "11/02/2026 - 11/03/2026" and checks == []


def test_other_expense_has_no_date_range(browser, tmp_path):
    bot, pg = bot_on(browser, EXPENSE)
    pg.evaluate("document.querySelector('[name=expenseType]').textContent = 'Parking & Toll'")
    done = bot.fill_expense(rec_of("주차 2026-11-03 02:15", "2026-11-03"), None, [])
    assert not any(d.startswith("Date Range") for d in done) and pg.input_value("#range") == ""


def test_date_range_already_filled_is_kept(browser):
    bot, pg = bot_on(browser, EXPENSE)
    pg.evaluate("""() => { document.querySelector('[name=expenseType]').textContent = 'Accommodation';
        document.getElementById('rangeRow').style.display = 'block';
        document.getElementById('range').value = '11/01/2026 - 11/03/2026'; }""")
    done = bot.fill_expense(rec_of("2026-11-03 02:15", "2026-11-03"), None, [])
    assert pg.input_value("#range") == "11/01/2026 - 11/03/2026" and not any("Date Range" in d for d in done)


# ---------- 저장 뒤 '필수 정보가 빠졌다'는 창 ----------

NEXT_PAGE = """<!doctype html><meta charset="utf-8">
<span id="prog">1 of 2</span> <button onclick="next()">Next Expense</button> <button onclick="save()">Save Expense</button>
<div id="dlg" role="dialog" style="display:none"><b>Error</b>
  <p>This expense has been saved, but it is missing required information or has an error that needs to be resolved
     before the report can be submitted. Would you like to make corrections now?</p>
  <button onclick="answer('yes')">Yes</button><button onclick="answer('no')">No</button></div>
<div id="alerts" style="display:none">Alerts: Itemizations are required for this entry.</div>
<script>
window.steps = []; window.missing = %s;
function ask() { document.getElementById('dlg').style.display = 'block'; }
function next() { steps.push('next'); if (missing) ask(); else document.getElementById('prog').textContent = '2 of 2'; }
function save() { steps.push('save'); if (missing) ask(); }
function answer(a) { steps.push(a); document.getElementById('dlg').style.display = 'none';
  if (a === 'yes') document.getElementById('alerts').style.display = 'block';
  else if (document.getElementById('prog').textContent === '1 of 2') document.getElementById('prog').textContent = '2 of 2'; }
</script>"""


def test_missing_info_dialog_yes_then_continue(browser):
    bot, pg = bot_on(browser, NEXT_PAGE % "true")
    checks = []
    bot.go_next(1, 2, checks)
    assert pg.evaluate("steps") == ["next", "yes", "next", "no"]       # [Yes] → (Itemizations 탭이 없음) → 다음 경비로
    assert pg.inner_text("#prog") == "2 of 2" and checks == ["필수 정보 직접 확인(Concur 알림 참고)"]


def test_last_expense_save_dialog(browser):
    bot, pg = bot_on(browser, NEXT_PAGE % "true")
    pg.evaluate("document.getElementById('prog').textContent = '2 of 2'")
    checks = []
    bot.go_next(2, 2, checks)
    assert pg.evaluate("steps") == ["save", "yes", "save", "no"] and checks == ["필수 정보 직접 확인(Concur 알림 참고)"]


def test_no_dialog_goes_straight_on(browser):
    bot, pg = bot_on(browser, NEXT_PAGE % "false")
    checks = []
    bot.go_next(1, 2, checks)
    assert pg.evaluate("steps") == ["next"] and checks == [] and pg.inner_text("#prog") == "2 of 2"
