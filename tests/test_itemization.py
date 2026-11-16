# -*- coding: utf-8 -*-
"""숙박 Itemizations (사용자 설명 순서): 저장 → '필수 정보 빠짐' 창 [Yes] → [Itemizations] → Room Rate에 금액 →
[Save Itemization] → 목록에서 항목 체크 → [Edit] → 빈 LCRT·Comment 채우기 → [Save Itemization] → 다음 경비."""

import os
from pathlib import Path

import pytest

import concur_auto as ca

PAGE = """<!doctype html><meta charset="utf-8">
<div id="expense">
  <span id="prog">1 of 2</span> <button aria-label="Next Expense" onclick="next()">›</button>
  <button onclick="saveExp()">Save Expense</button>
  <div role="tablist"><button role="tab" onclick="tab('details')">Details</button>
    <button role="tab" onclick="tab('items')">Itemizations</button></div>
  <div id="details"><label for="comment">Comment</label><textarea id="comment">Accommodation fee</textarea></div>
  <div id="itemsList" style="display:none; position:relative">
    <div style="display:flex; gap:20px"><h3 id="head">Itemizations (0)</h3>
      <label><input type="checkbox" aria-label="Select All">Select All</label></div>
    <button id="edit" onclick="edit()" disabled>Edit</button>
    <div id="rows"></div>
  </div>
</div>
<div id="add" style="display:none">
  <a href="#" onclick="back(); return false">Expense</a>
  <h2>Add Itemization</h2><button onclick="saveAdd()">Save Itemization</button>
  {rows}
</div>
<div id="itemForm" style="display:none">
  <h2>Itemization</h2><button onclick="saveItem()">Save Itemization</button>
  <label for="businessPurpose">Business Purpose</label><input id="businessPurpose" value="For service">
  <p>Legal compliance registration tool (LCRT) *</p>
  <div id="custom32"><span class="sapcnqr-select-field__input" onclick="openList(this)"></span></div>
  <label for="comment">Comment</label><textarea id="comment"></textarea>
</div>
<div id="alerts" style="display:none; position:absolute; top:0; left:0; width:600px; height:260px; background:#fff; border:1px solid #333">
  <div><h3>Alerts</h3><button aria-label="Close" onclick="hideAlerts()">X</button></div><p id="alertText"></p></div>
<div id="dlg" role="dialog" style="display:none"><b>Error</b>
  <p>This expense has been saved, but it is missing required information or has an error that needs to be resolved
     before the report can be submitted. Would you like to make corrections now?</p>
  <button onclick="answer('yes')">Yes</button><button onclick="answer('no')">No</button></div>
<div id="pop" style="display:none"><div class="sapcnqr-listbox-item__wrapper" onclick="pick(this)">NA</div></div>
<script>
window.steps = []; let itemized = false, itemOk = false, target = null;
const show = (id, on) => { document.getElementById(id).style.display = on ? 'block' : 'none'; };
const done = () => itemized && itemOk;
function next() { steps.push('next'); if (!done()) show('dlg', 1); else document.getElementById('prog').textContent = '2 of 2'; }
function saveExp() { steps.push('save'); if (!done()) show('dlg', 1); }
function answer(a) { steps.push(a); show('dlg', 0);
  if (a === 'yes') alerts('Error: Itemizations are required for this entry.');
  else document.getElementById('prog').textContent = '2 of 2'; }
function alerts(t) { document.getElementById('alertText').textContent = t; show('alerts', 1); }
function hideAlerts() { steps.push('close alerts'); show('alerts', 0); }
function tab(w) { steps.push('tab ' + w);
  if (w === 'items') { show('details', 0); if (!itemized) { show('expense', 0); show('add', 1); } else show('itemsList', 1); } }
function saveAdd() { const r = document.getElementById('rate'); const v = r ? r.value.replace(/,/g, '') : '';
  steps.push('save add ' + v); if (v === '70000') itemized = true;
  show('add', 0); show('expense', 1); show('itemsList', 1); draw();
  alerts('Error: This expense has itemizations with missing required fields. Missing required field: LCRT'); }
function draw() { document.getElementById('head').textContent = 'Itemizations (' + (itemized ? 1 : 0) + ')';
  document.getElementById('rows').innerHTML = itemized ? '<div class="row" style="display:flex; gap:20px">'
    + '<input type="checkbox" onchange="chk(this)"><span onclick="openItem()">Accommodation</span><span>11/02/2026</span>'
    + '<span>' + (itemOk ? '' : 'Errors: 1') + '</span><span>KRW 70,000</span></div>' : ''; }
function back() { steps.push('back'); show('add', 0); show('itemForm', 0); show('expense', 1); show('details', 1); }
function chk(c) { steps.push('check'); document.getElementById('edit').disabled = !c.checked; }
function edit() { steps.push('edit'); show('expense', 0); show('itemForm', 1); }
function openItem() { steps.push('open item'); show('expense', 0); show('itemForm', 1); }
function openList(el) { target = el; show('pop', 1); }
function pick(el) { target.textContent = el.textContent; show('pop', 0); }
function saveItem() { const l = document.querySelector('#itemForm .sapcnqr-select-field__input').textContent;
  const c = document.querySelector('#itemForm #comment').value; steps.push('save item ' + l + ' | ' + c);
  itemOk = !!l && !!c; show('itemForm', 0); show('expense', 1); show('itemsList', 1); draw(); }
</script>
<style>.sapcnqr-select-field__input { display:inline-block; min-width:200px; min-height:18px; border:1px solid #999 }</style>"""

ROWS = """<div style="display:flex"><div style="width:120px">Date</div><div style="width:160px">Room Rate <span>*</span></div>
  <div style="width:160px">Room Tax</div><div style="width:160px">Tax 2</div></div>
<div style="display:flex"><div style="width:120px">11/02/2026</div><input id="rate" style="width:150px; margin-right:10px">
  <input id="tax" style="width:150px; margin-right:10px"><input style="width:150px"></div>"""


@pytest.fixture
def browser():
    sync_api = pytest.importorskip("playwright.sync_api")
    exe = os.environ.get("PW_CHROMIUM") or "/opt/pw-browsers/chromium"
    with sync_api.sync_playwright() as p:
        b = p.chromium.launch(executable_path=exe if Path(exe).exists() else None)
        yield b
        b.close()


def bot_on(browser, rows=ROWS, page=PAGE):
    pg = browser.new_page()
    pg.set_default_timeout(3000)
    pg.set_content(page.replace("{rows}", rows))
    bot = ca.ConcurBot(ca.default_settings(), lambda s="": None)
    bot.page, bot.timeout_ms, bot.upload_wait_s, bot.upload_done_s = pg, 3000, 2, 2
    bot.screenshot = lambda name: None
    bot.last_values = {"Business Purpose": "For service in hospital", "LCRT": "NA",
                       "Comment": "Accommodation fee for service in Mokpo Hankook Hospital"}
    return bot, pg


def test_itemization_in_user_order(browser):
    bot, pg = bot_on(browser)
    checks = []
    bot.go_next(1, 2, checks, 70000)
    assert pg.evaluate("steps") == [
        "next", "yes", "close alerts", "tab items", "save add 70000",          # Room Rate에 금액 → 저장
        "close alerts", "open item",                                            # 오류가 뜬 Accommodation을 눌러 들어감
        "save item NA | Accommodation fee for service in Mokpo Hankook Hospital",   # 빈 LCRT·Comment 채우고 저장
        "next"]                                                                  # 다 채워져 그대로 다음 경비로
    assert pg.input_value("#itemForm #businessPurpose") == "For service"      # 이미 있는 값은 그대로
    assert pg.inner_text("#prog") == "2 of 2" and checks == []


def test_room_rate_not_found_still_moves_on(browser):
    bot, pg = bot_on(browser, rows="<p>No dates</p><div>Room Rate</div>")
    checks = []
    bot.go_next(1, 2, checks, 70000)
    steps = pg.evaluate("steps")
    assert steps[:4] == ["next", "yes", "close alerts", "tab items"] and steps[-3:] == ["back", "next", "no"]
    assert checks == ["Itemizations 직접 넣기"] and pg.inner_text("#prog") == "2 of 2"


def test_dialog_without_itemizations_tab(browser):
    bot, pg = bot_on(browser)
    pg.evaluate("document.querySelector('[role=tablist]').remove()")
    checks = []
    bot.go_next(1, 2, checks, 70000)
    assert checks == ["필수 정보 직접 확인(Concur 알림 참고)"] and pg.evaluate("steps")[-2:] == ["next", "no"]


def test_column_input_under_header(browser):
    bot, pg = bot_on(browser)
    pg.evaluate("document.getElementById('add').style.display = 'block'")
    assert pg.evaluate(ca.JS_COLUMN_INPUT, ["Room Rate"])
    assert pg.get_attribute("[data-concur-auto]", "id") == "rate"              # Room Tax 칸이 아니라 Room Rate 칸


def test_itemization_falls_back_to_check_and_edit(browser):
    """항목 글자를 눌러도 안 열리면 예전처럼 체크 → [Edit]"""
    bot, pg = bot_on(browser, page=PAGE.replace(' onclick="openItem()"', ""))
    checks = []
    bot.go_next(1, 2, checks, 70000)
    steps = pg.evaluate("steps")
    assert "check" in steps and "edit" in steps and "open item" not in steps
    assert steps[-2] == "save item NA | Accommodation fee for service in Mokpo Hankook Hospital" and checks == []


def test_itemization_lcrt_found_by_label(browser):
    """Itemization 수정 화면의 LCRT 칸이 경비 화면과 다른 id여도 칸 이름(LCRT)으로 찾아 NA를 고름"""
    page = PAGE.replace('<div id="custom32"><span class="sapcnqr-select-field__input" onclick="openList(this)"></span></div>',
                        '<div id="lcrtBox"><span class="sapcnqr-select-field__input" onclick="openList(this)"></span></div>')
    assert page != PAGE
    bot, pg = bot_on(browser, page=page)
    checks = []
    bot.go_next(1, 2, checks, 70000)
    assert "save item NA | Accommodation fee for service in Mokpo Hankook Hospital" in pg.evaluate("steps") and checks == []



def test_late_missing_info_dialog_still_answered(browser):
    """저장 뒤 '필수 정보 빠짐' 창이 5초 늦게 떠도 [Yes]를 누르고 Itemizations를 넣음(실제 화면에서 못 누르고 지나쳤음)"""
    page = PAGE.replace("function next() { steps.push('next'); if (!done()) show('dlg', 1);",
                        "function next() { steps.push('next'); if (!done()) setTimeout(() => show('dlg', 1), 5000);")
    assert page != PAGE
    bot, pg = bot_on(browser, page=page)
    checks = []
    bot.go_next(1, 2, checks, 70000)
    steps = pg.evaluate("steps")
    assert steps[:2] == ["next", "yes"] and "save add 70000" in steps and checks == []


# 실제 Concur(2026-11-16 실행 기록): 경비 화면을 열어도 뒤에 리포트 목록이 가려진 채 그대로 있음
# (Accommodation·Errors 글자와 체크박스, 도구 모음의 Edit). 예전에는 이 가려진 줄을 눌러 Itemization이 안 열렸음.
BEHIND = """<div id="report" style="position:fixed; inset:0; z-index:1">
  <h2>Expenses (3)</h2><div class="sapcnqr-toolbar__toolbar"><button onclick="steps.push('report edit')">Edit</button></div>
  <div style="margin-top:160px">
  <div class="row" style="display:flex; gap:20px"><input type="checkbox" onchange="steps.push('report check')">
    <span onclick="steps.push('report open')">Accommodation</span><span>10/20/2026</span><span>Errors: 2 Error</span>
    <span>KRW 75,000</span></div></div></div>
"""
OVER = """<style>#expense, #add, #itemForm { position:fixed; inset:0; z-index:10; background:#fff; overflow:auto }
#alerts, #dlg, #pop { position:fixed; z-index:20; background:#fff }</style>"""


def test_itemization_ignores_report_list_behind(browser):
    page = PAGE.replace('<div id="expense">', BEHIND + OVER + '<div id="expense">', 1)
    bot, pg = bot_on(browser, page=page)
    checks = []
    bot.go_next(1, 2, checks, 70000)
    steps = pg.evaluate("steps")
    assert not [x for x in steps if x.startswith("report")]                    # 뒤에 깔린 리포트 줄·Edit는 안 누름
    assert "open item" in steps and checks == []
    assert "save item NA | Accommodation fee for service in Mokpo Hankook Hospital" in steps


def test_check_and_edit_ignores_report_edit_behind(browser):
    page = PAGE.replace('<div id="expense">', BEHIND + OVER + '<div id="expense">', 1).replace(' onclick="openItem()"', "")
    bot, pg = bot_on(browser, page=page)
    checks = []
    bot.go_next(1, 2, checks, 70000)
    steps = pg.evaluate("steps")
    assert "check" in steps and "edit" in steps and not [x for x in steps if x.startswith("report")]
    assert checks == []


# 실제 Itemization 수정 화면(2026-11-16 사용자 화면): LCRT 목록이 'Search by Text' 칸·안내 글·NA/Registration Number가 있는
# 다른 모양이라 설정의 '드롭다운 항목' 위치로는 못 찾았음. 항목 줄의 체크박스에는 화면 낭독용 글이 숨어 있음.
SEARCH_POP = """<div id="pop" style="display:none"><input placeholder="Search by Text">
  <p>This field is missing required information.</p>
  <ul role="listbox"><li role="option" onclick="pick(this)">NA</li>
    <li role="option" onclick="pick(this)">Registration Number</li></ul></div>"""
OLD_POP = """<div id="pop" style="display:none"><div class="sapcnqr-listbox-item__wrapper" onclick="pick(this)">NA</div></div>"""
SR_ROW = ("'<label><input type=\"checkbox\" onchange=\"chk(this)\"><span style=\"position:absolute; width:1px; height:1px;"
          " overflow:hidden; clip:rect(0 0 0 0)\">Select Accommodation 10/20/2026 KRW 70,000</span></label>")


def test_lcrt_search_list_and_hidden_row_text(browser):
    page = PAGE.replace(OLD_POP, SEARCH_POP).replace(
        "'<input type=\"checkbox\" onchange=\"chk(this)\">", SR_ROW)
    assert SEARCH_POP in page and "Select Accommodation" in page
    page = page.replace('<div id="expense">', BEHIND.replace("Errors: 2 Error", "Errors: 2 Error</span><span>NA") + OVER
                        + '<div id="expense">', 1)
    bot, pg = bot_on(browser, page=page)
    checks = []
    bot.go_next(1, 2, checks, 70000)
    steps = pg.evaluate("steps")
    assert "open item" in steps                                                 # 숨은 글이 있어도 Accommodation을 누름
    assert "save item NA | Accommodation fee for service in Mokpo Hankook Hospital" in steps   # 모양이 다른 목록에서도 NA
    assert not [x for x in steps if x.startswith("report")] and checks == []


# 실제 Itemization 수정 화면(2026-11-16 사용자 설명): LCRT 칸을 눌러도 '필수 정보' 안내만 뜨고, 칸 오른쪽의 아래 화살표(⌄)를
# 눌러야 'Search by Text' 목록이 열림 → NA. (경비 화면의 LCRT처럼 옆에 지우기 X 단추도 있음)
ARROW_FIELD = """<div id="lcrtField" style="display:flex; width:420px; border:1px solid red">
  <input class="lcrt-in" role="combobox" readonly style="flex:1" onclick="steps.push('lcrt box')">
  <button class="combo-arrow" style="width:30px" onclick="steps.push('arrow'); openList(document.querySelector('.lcrt-in'))">v</button>
  <button aria-label="Clear" style="width:30px" onclick="steps.push('clear')">X</button></div>"""


def test_lcrt_opened_with_arrow(browser):
    page = (PAGE.replace(OLD_POP, SEARCH_POP)
            .replace('<div id="custom32"><span class="sapcnqr-select-field__input" onclick="openList(this)"></span></div>', ARROW_FIELD)
            .replace("function pick(el) { target.textContent = el.textContent;",
                     "function pick(el) { if (target.tagName === 'INPUT') target.value = el.textContent; else target.textContent = el.textContent;")
            .replace("const l = document.querySelector('#itemForm .sapcnqr-select-field__input').textContent;",
                     "const l = document.querySelector('#itemForm .lcrt-in').value;"))
    assert "lcrt-in" in page and "target.value" in page and "'#itemForm .lcrt-in'" in page
    bot, pg = bot_on(browser, page=page)
    checks = []
    bot.go_next(1, 2, checks, 70000)
    steps = pg.evaluate("steps")
    assert "arrow" in steps and "clear" not in steps                            # 지우기(X)가 아니라 화살표
    assert "save item NA | Accommodation fee for service in Mokpo Hankook Hospital" in steps and checks == []
