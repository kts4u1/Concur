# -*- coding: utf-8 -*-
"""톨게이트 새 경비(2026-11-15 사용자 화면·설명): 카드 거래가 아니라서 리포트의 [Add Expense] → [Manually Create Expense]
→ 목록에서 Parking & Toll → 새 경비 화면에서 Transaction Date = 마지막 이용일, Amount = 사용내역 총액, 나머지 칸,
사용내역·영수증 사진을 PDF로 묶지 않고 차례차례 올림 → [Save Expense]."""

import os
from datetime import datetime
from pathlib import Path

import pytest

import concur_auto as ca

REPORT = """<!doctype html><meta charset="utf-8">
<div id="report"><a href="#">Expense Reports</a><h1>Monthly Expense for Oct</h1><h2 id="count">Expenses (0)</h2>
  <div id="empty" style="display:none"><h2>No Expenses</h2>
    <div class="card" onclick="openDialog()"><b>Manually Create Expense</b><p>Enter expense details</p></div>
    <div class="card"><b>Available Expenses (0)</b><p>Select from your list of expenses</p></div></div>
  <div class="sapcnqr-toolbar__toolbar">
    <button id="addbtn" onclick="document.getElementById('menu').style.display='block'">Add Expense</button>
    <button onclick="editSel()">Edit</button></div>
  <div id="menu" role="menu" style="display:none">
    <div role="menuitem" onclick="openDialog()">Manually Create Expense</div>
    <div role="menuitem">Select from Available Expenses (0)</div></div>
  <div class="entries-list" id="list"></div></div>
<div id="dialog" role="dialog" style="display:none"><h2>Add Expense to Report</h2>
  <input placeholder="Search" oninput="filter(this.value)">
  <div id="types"></div></div>
<div id="form" style="display:none"><h1>New Expense</h1><button onclick="save()">Save Expense</button>
  <p>Expense Type *</p><div class="field__control"><span name="expenseType" id="type"></span></div>
  <label for="transactionDate-date-input-field-input">Transaction Date</label>
  <input id="transactionDate-date-input-field-input" value="11/02/2026" oninput="note('date')">
  <label for="businessPurpose">Business Purpose</label><input id="businessPurpose" oninput="note('purpose')">
  <p>City *</p><span id="location" onclick="openList('city', this)">Seoul, KOREA, REPUBLIC OF</span>
  <label for="transactionAmount">Amount</label><input id="transactionAmount" oninput="note('amount')">
  <p>LCRT *</p><div id="custom32"><span class="sapcnqr-select-field__input" onclick="openList('lcrt', this)"></span></div>
  <p>Receipt Status *</p><span name="receiptType" onclick="openList('receipt', this)">None Selected</span>
  <label for="vendor">Enter Vendor Name</label><input id="vendor" name="vendorName" oninput="note('vendor')">
  <label for="comment">Comment</label><textarea id="comment" oninput="note('comment')"></textarea>
  <div id="pane"><p id="drop">Upload or drag-and-drop your receipts here.</p>
    <button id="up" onclick="document.getElementById('f').click()">Upload New Receipt</button>
    <button id="choose" onclick="openAvail()">Choose From Available Receipts</button>
    <div id="viewer" style="display:none"><span id="vname"></span><iframe id="pdf" style="display:none" width="40" height="20"></iframe>
      <button>Remove</button><button id="addr" onclick="openAvail()">Add</button><button>Open</button></div></div>
  <input type="file" id="f" style="display:none" onchange="toLibrary(this)" MULTIPLE>
  <div id="avail" style="display:none"><h3>Attach Receipt</h3><span>Available Receipts</span> <span>Receipts in Report</span>
    <button aria-label="Close Available Receipts" onclick="closeAvail()">X</button>
    <button id="up2" onclick="document.getElementById('f2').click()">Upload New Receipt</button>
    <input type="file" id="f2" style="display:none" onchange="toLibrary(this)" MULTIPLE>
    <div id="cards" style="height:150px; overflow-y:auto" onscroll="more(this)"></div></div>
  <div id="dlg" style="display:none"><h3>Receipt Image</h3><div><div><span id="dname"></span></div>
    <div><div onclick="show('dlg', 0)">Delete</div><div class="btn" id="dbtn" onclick="attachDlg()">Attach</div></div></div>
    <button onclick="show('dlg', 0)">Close</button></div></div>
<div id="pop" style="display:none"><div id="items"></div></div>
<script>
window.steps = []; window.uploads = []; window.saved = null; window.expenses = []; let editing = null;
window.library = LIBRARY; let current = null, shownCards = 4;
const fresh = {}, SLOW = false, BGTHUMB = false, NONAME = false, ICONS = false;
const PIX = 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7';
function toLibrary(inp) { const names = Array.from(inp.files).map(f => f.name); library.push(...names);
  if (SLOW) { names.forEach(n => { fresh[n] = Date.now() + 1500; }); setTimeout(drawCards, 1600); }
  else if (BGTHUMB) names.forEach(n => { fresh[n] = 1; });
  steps.push('library ' + names.join(',')); inp.value = ''; drawCards(); }
function more(box) { if (box.scrollTop + box.clientHeight >= box.scrollHeight - 5 && shownCards < library.length) {
  const top = box.scrollTop; shownCards += 2; drawCards(); box.scrollTop = top; } }
function thumbOf(n) { if (fresh[n] && Date.now() < fresh[n]) return '<div class="spin">Processing...</div>';   // 막 올린 사진
  if (BGTHUMB && fresh[n]) return '<div class="thumb" style="width:60px;height:40px;background-image:url(' + PIX + ')" onclick="openDlg(&quot;' + n + '&quot;)"></div>';
  return '<img src="' + PIX + '" width="60" height="40" onclick="openDlg(&quot;' + n + '&quot;)">'; }
function drawCards() { document.getElementById('cards').innerHTML = library.slice(0, shownCards).map(n => '<div class="card" style="height:60px"><div>' + n +
  '</div>' + (ICONS ? '<img src="' + PIX + '" width="12" height="12" onclick="steps.push(&quot;delete icon&quot;)">' : '') + thumbOf(n) +
  '<button aria-label="Attach receipt" onclick="steps.push(&quot;card icon&quot;)">i</button></div>').join(''); }
function openAvail() { shownCards = 4; show('avail', 1); show('pane', 0); drawCards(); }
function closeAvail() { show('avail', 0); show('pane', 1); }
function openDlg(n) { current = n; document.getElementById('dname').textContent = NONAME ? '' : n;
  document.getElementById('dbtn').textContent = (uploads.length || window.had) ? 'Add' : 'Attach'; show('dlg', 1); }
function resetPane() { uploads = []; window.had = 0; ['drop', 'up', 'choose'].forEach(id => show(id, 1)); show('viewer', 0); }
function attachDlg() { uploads.push(current); steps.push((uploads.length > 1 || window.had ? 'add ' : 'attach ') + current); show('dlg', 0);
  closeAvail(); ['drop', 'up', 'choose'].forEach(id => show(id, 0)); show('viewer', 1);
  document.getElementById('vname').textContent = window.had ? document.getElementById('vname').textContent : uploads[0];
  show('pdf', uploads.length + (window.had || 0) > 1); }
const TYPES = ['M card parking & toll', 'M card taxi', 'Parking & Toll', 'Public Transport', 'Taxi'];
const lists = {city: ['Gwangju (ex Kwangju), KOREA, REPUBLIC OF', 'Seoul, KOREA, REPUBLIC OF'], lcrt: ['NA'],
               receipt: ['No Receipt', 'Receipt']};
let target = null, which = null;
function note(s) { if (steps[steps.length - 1] !== s) steps.push(s); }
const show = (id, on) => { document.getElementById(id).style.display = on ? 'block' : 'none'; };
function drawTypes(list) { document.getElementById('types').innerHTML = list.map(t =>
  '<div class="t" onclick="pickType(this)">' + t + '</div>').join(''); }
function openDialog() { steps.push('manual'); show('menu', 0); show('dialog', 1); drawTypes(TYPES); }
function filter(q) { drawTypes(TYPES.filter(t => t.toLowerCase().includes(q.toLowerCase()))); }
function pickType(el) { steps.push('type ' + el.textContent); show('dialog', 0); show('report', 0); show('form', 1);
  editing = null; resetPane();
  document.getElementById('type').textContent = el.textContent; }
function openList(w, el) { which = w; target = el; show('pop', 1); document.getElementById('items').innerHTML =
  lists[w].map(t => '<div class="sapcnqr-listbox-item__wrapper" onclick="pick(this)">' + t + '</div>').join(''); }
function pick(el) { target.textContent = el.textContent; note(which); show('pop', 0); }
if (EMPTY) { document.getElementById('empty').style.display = 'block'; document.getElementById('addbtn').style.display = 'none'; }
function drawRows() { document.getElementById('list').innerHTML = expenses.map((e, i) =>
  '<div role="row"><span class="sapcnqr-dropzone"><input type="checkbox" class="sapcnqr-checkbox__input" data-i="' + i +
  '"></span><span>' + e.date + '</span><span>' + e.type + '</span><span>KRW ' +
  Number(String(e.amount).replace(/,/g, '')).toLocaleString('en-US') + '</span></div>').join(''); }
function editSel() { const cb = document.querySelector('#list input:checked'); if (!cb) return;
  editing = +cb.dataset.i; steps.push('edit'); show('report', 0); show('form', 1); resetPane();
  const had = expenses[editing].uploads || []; window.had = had.length;            // 이미 붙은 영수증은 그대로 보임
  if (had.length) { ['drop', 'up', 'choose'].forEach(id => show(id, 0)); show('viewer', 1); show('addr', 1);
    if (document.getElementById('opn')) show('opn', 1);
    document.getElementById('vname').textContent = had[0]; show('pdf', had.length > 1); } }
function save() { steps.push('save');
  if (editing !== null) { saved = expenses[editing]; saved.uploads = saved.uploads.concat(uploads); }
  else { saved = {type: document.getElementById('type').textContent, date: document.getElementById('transactionDate-date-input-field-input').value,
           amount: document.getElementById('transactionAmount').value, comment: document.getElementById('comment').value,
           purpose: document.getElementById('businessPurpose').value, city: document.getElementById('location').textContent,
           lcrt: document.querySelector('#custom32 span').textContent,
           receipt: document.querySelector('[name=receiptType]').textContent,
           vendor: (document.getElementById('vendor') || {}).value, uploads: uploads.slice()};
         expenses.push(saved); }
  uploads = []; editing = null;
  show('form', 0); show('report', 1); document.getElementById('count').textContent = 'Expenses (' + expenses.length + ')';
  show('empty', 0); show('addbtn', 1); drawRows(); }
</script>
<style>span[name], #location, .sapcnqr-select-field__input { display:inline-block; min-width:220px; min-height:18px; border:1px solid #999 }
.t { padding: 4px }</style>"""

TOLL = ["톨게이트", "통행료, 하이패스, 사용내역", "Toll fee for service", "For service", "Parking & Toll", "", "", "", "", ""]


def report_html(multiple=True, empty=False, library=(), slow=False, bg=False, noname=False, icons=False, behind=False,
                slow_close=False, no_choose_new=False, save_busy=False, covered_save=False, real_pane=False,
                presave_no_add=False):
    """slow: 막 올린 사진은 1.5초 동안 'Processing...'(사진 없음, 눌러도 안 열림)
    bg: 막 올린 사진의 미리보기가 img가 아니라 배경 그림(div) / noname: 'Receipt Image' 창에 파일 이름이 없음
    icons: 카드마다 사진 앞에 작은 아이콘 그림(누르면 'delete icon')
    behind: 실제 Concur처럼 경비 화면을 열어도 리포트 화면이 뒤에 가려진 채 남고, 영수증 칸의 [Add]는 아이콘 글자가 붙음
    slow_close: 사진을 붙이면 보관함의 카드는 바로 사라지지만 보관함은 1.5초 뒤에 닫힘
    no_choose_new: 저장하기 전의 새 경비 화면에는 [Choose From Available Receipts]가 없음(저장한 뒤 다시 열어야 붙임)
    save_busy: 사진을 붙인 뒤 4초 동안은 [Save Expense]를 눌러도 저장되지 않음(사진을 합치는 중)
    covered_save: 위에 고정된 머리띠(회사 로고 줄)가 있고, 사진을 붙이면 화면이 내려가 [Save Expense]가 머리띠 밑에 가려짐
    real_pane: 실제 Concur처럼 보관함을 열어도 경비 영수증 칸이 그 밑에 그대로 깔려 있고(보관함이 덮음), 경비 칸의
               [Upload New Receipt](#upload-receipt-button)는 보관함 밖에 보여 누를 수 있으며, 그것으로 올리면 사진이
               보관함이 아니라 경비에 바로 붙고('direct …') 보관함이 닫힘(2026-11-16 실행 기록)
    presave_no_add: 실제 Concur처럼 저장하기 전의 새 경비는 사진을 붙여도 그 아래에 [Remove]만 있고 [Add]·[Open]이 없음
               (저장한 뒤 다시 열면 생김, 2026-11-16 23:26 사용자 화면)"""
    js = lambda on: "true" if on else "false"
    html = REPORT
    if behind:
        html = (html.replace("show('report', 0);", "")
                .replace('<button id="addr" onclick="openAvail()">Add</button>',
                         '<button id="addr" onclick="openAvail()"><i class="mi">add</i> Add</button>')
                + "<style>#form { position:fixed; inset:0; z-index:10; background:#fff; overflow:auto }"
                  " #dialog, #pop { position:fixed; top:0; left:0; z-index:30; background:#fff }</style>")
        assert html.count("show('report', 0)") == 0 and "mi" in html
    if covered_save:
        assert "uploads.push(current);" in html
        html = ('<div id="appbar" style="position:fixed; top:0; left:0; right:0; height:60px; background:#fff; z-index:50">LOGO</div>'
                + html.replace("uploads.push(current);", "uploads.push(current); setTimeout(() => window.scrollTo(0, 90), 50);")
                + "<style>body { padding-top: 70px } #form { min-height: 2500px }</style>")
    if save_busy:
        old = "function save() { steps.push('save');"
        assert old in html and "uploads.push(current);" in html
        html = (html.replace(old, "function save() { if (Date.now() < (window.busyUntil || 0)) { steps.push('save ignored'); return; } steps.push('save');")
                .replace("uploads.push(current);", "uploads.push(current); window.busyUntil = Date.now() + 4000;"))
    if no_choose_new:
        old = "editing = null; resetPane();"
        assert old in html
        html = html.replace(old, "editing = null; resetPane(); show('choose', 0);")
    if real_pane:
        old_open = "function openAvail() { shownCards = 4; show('avail', 1); show('pane', 0); drawCards(); }"
        old_up = '<button id="up" onclick="document.getElementById(\'f\').click()">Upload New Receipt</button>'
        old_f = '<input type="file" id="f" style="display:none" onchange="toLibrary(this)" MULTIPLE>'
        assert old_open in html and old_up in html and old_f in html
        html = (html.replace(old_open, "function openAvail() { shownCards = 4; show('avail', 1); drawCards(); }")
                .replace(old_up, old_up.replace('id="up"', 'id="upload-receipt-button"'))
                .replace(old_f, old_f.replace("toLibrary(this)", "attachDirect(this)"))
                .replace("function resetPane() {", "function attachDirect(inp) { Array.from(inp.files).forEach(f => {\n"
                         "  uploads.push(f.name); steps.push('direct ' + f.name); }); inp.value = ''; closeAvail();\n"
                         "  ['drop', 'upload-receipt-button', 'choose'].forEach(id => show(id, 0)); show('viewer', 1);\n"
                         "  document.getElementById('vname').textContent = uploads[0]; show('pdf', uploads.length > 1); }\n"
                         "function resetPane() {")
                .replace("['drop', 'up', 'choose']", "['drop', 'upload-receipt-button', 'choose']")
                + "<style>#form { position: relative } #avail { position: absolute; inset: 0; z-index: 20; background: #fff }"
                  " #upload-receipt-button { position: relative; z-index: 30 } #dlg { position: relative; z-index: 40; background: #fff }</style>")
        assert "'up'" not in html
    if presave_no_add:
        old_btns = '<button id="addr" onclick="openAvail()">Add</button><button>Open</button>'
        assert old_btns in html and "function pickType(el) {" in html and "editing = +cb.dataset.i;" in html
        html = (html.replace(old_btns, '<button id="addr" onclick="openAvail()">Add</button><button id="opn">Open</button>')
                .replace("function pickType(el) {", "function pickType(el) { window.presave = true;")
                .replace("editing = +cb.dataset.i;", "editing = +cb.dataset.i; window.presave = false;")
                .replace("show('pdf', uploads.length > 1);",
                         "show('pdf', uploads.length > 1); show('addr', !window.presave); show('opn', !window.presave);")
                .replace("show('pdf', uploads.length + (window.had || 0) > 1);",
                         "show('pdf', uploads.length + (window.had || 0) > 1); show('addr', !window.presave); show('opn', !window.presave);"))
    if slow_close:
        old = "show('dlg', 0);\n  closeAvail();"
        assert old in html
        html = html.replace(old, "show('dlg', 0);\n  document.getElementById('cards').innerHTML = ''; setTimeout(closeAvail, 1500);")
    return (html.replace(" MULTIPLE>", " multiple>" if multiple else ">")
            .replace("if (EMPTY)", "if (true)" if empty else "if (false)")
            .replace("LIBRARY", "[" + ",".join(f"'{n}'" for n in library) + "]")
            .replace("SLOW = false, BGTHUMB = false, NONAME = false, ICONS = false",
                     f"SLOW = {js(slow)}, BGTHUMB = {js(bg)}, NONAME = {js(noname)}, ICONS = {js(icons)}"))


@pytest.fixture
def browser():
    sync_api = pytest.importorskip("playwright.sync_api")
    exe = os.environ.get("PW_CHROMIUM") or "/opt/pw-browsers/chromium"
    with sync_api.sync_playwright() as p:
        b = p.chromium.launch(executable_path=exe if Path(exe).exists() else None)
        yield b
        b.close()


def open_report(browser, **kw):
    pg = browser.new_page()
    pg.set_default_timeout(3000)
    html = report_html(**kw)
    pg.route("**/report", lambda route: route.fulfill(body=html, content_type="text/html; charset=utf-8"))
    return pg


@pytest.fixture
def page(browser):
    return open_report(browser)


def photos(tmp_path):
    from PIL import Image
    out = []
    for name in ("IMG_5648.jpg", "IMG_5651.jpg", "IMG_5652.jpg"):
        f = tmp_path / name
        Image.new("RGB", (300, 400), (240, 240, 240)).save(f, quality=90)
        out.append(f)
    return out


def test_manual_toll_expense(page, tmp_path):
    pytest.importorskip("cv2")
    cfg = ca.default_settings()
    table = tmp_path / ca.CATEGORY_FILE_NAME
    ca.save_category_rows(table, [TOLL])
    ca.apply_category_table(cfg, table)
    stmt, s1, s2 = photos(tmp_path)
    rec = ca.Receipt(original=stmt, ocr_file=stmt, bundle=[s1, s2], total=61500, dates=["2026-10-22"],
                     last_use="2026-10-22", category="톨게이트")
    bot = ca.ConcurBot(cfg, lambda s="": None)
    bot.page, bot.timeout_ms, bot.upload_wait_s, bot.upload_done_s = page, 3000, 3, 3
    bot.report_url = "http://concur.test/report"
    bot.screenshot = lambda name: None
    rows = bot.create_manual_expenses([rec], tmp_path / "w", start_no=5)
    saved = page.evaluate("saved")
    assert saved["type"] == "Parking & Toll"                                       # 'M card parking & toll'이 아님
    assert saved["date"] == "10/22/2026" and saved["amount"].replace(",", "") == "61500"
    assert saved["comment"] == "Toll fee for service" and saved["purpose"] == "For service"
    assert saved["lcrt"] == "NA" and saved["receipt"] == "Receipt" and saved["city"].startswith("Gwangju")
    assert saved["uploads"] == ["IMG_5648.jpg", "IMG_5651.jpg", "IMG_5652.jpg"]    # PDF가 아니라 사진을 차례대로
    steps = page.evaluate("steps")
    assert steps[:2] == ["manual", "type Parking & Toll"] and steps[-1] == "save"
    one = "library IMG_5648.jpg,IMG_5651.jpg,IMG_5652.jpg"                         # 보관함에 한 번에 여러 장 올림
    assert steps.index("date") < steps.index("amount") < steps.index(one) < steps.index("save")
    assert [x for x in steps if x.startswith(("attach", "add "))] == [            # 첫 장 Attach, 다음 장은 Add
        "attach IMG_5648.jpg", "add IMG_5651.jpg", "add IMG_5652.jpg"]
    assert "card icon" not in steps                                                # 'Receipt Image' 창의 [Attach]를 누름
    # 이 가짜 화면은 저장 전 화면에서 올린 사진이 보관함으로만 감 → 저장한 뒤 다시 열어 보관함에서 붙임(2026-11-16)
    assert steps.count("save") == 2 and steps.index("save") < steps.index("edit") < steps.index("attach IMG_5648.jpg")
    assert saved["vendor"] == ""                                                   # Vendor Name은 숙박만
    assert rows[0]["순번"] == 5 and rows[0]["Concur 날짜"] == "10/22/2026" and rows[0]["확인 필요"] == ""
    assert "영수증 첨부(IMG_5648.jpg, IMG_5651.jpg, IMG_5652.jpg)" in rows[0]["처리"] and rec.used_for
    assert page.inner_text("#count") == "Expenses (1)"


def make_bot(page, tmp_path):
    cfg = ca.default_settings()
    table = tmp_path / ca.CATEGORY_FILE_NAME
    ca.save_category_rows(table, [TOLL])
    ca.apply_category_table(cfg, table)
    bot = ca.ConcurBot(cfg, lambda s="": None)
    bot.page, bot.timeout_ms, bot.upload_wait_s, bot.upload_done_s = page, 3000, 3, 3
    bot.report_url = "http://concur.test/report"
    bot.screenshot = lambda name: None
    return bot


def toll_rec(tmp_path):
    stmt, s1, s2 = photos(tmp_path)
    return ca.Receipt(original=stmt, ocr_file=stmt, bundle=[s1, s2], total=61500, dates=["2026-10-22"],
                      last_use="2026-10-22", category="톨게이트")


def test_one_at_a_time_when_chooser_takes_one(browser, tmp_path):
    pytest.importorskip("cv2")
    pg = open_report(browser, multiple=False)
    rows = make_bot(pg, tmp_path).create_manual_expenses([toll_rec(tmp_path)], tmp_path / "w")
    assert pg.evaluate("saved")["uploads"] == ["IMG_5648.jpg", "IMG_5651.jpg", "IMG_5652.jpg"]
    assert [s for s in pg.evaluate("steps") if s.startswith("library")] == [
        "library IMG_5648.jpg", "library IMG_5651.jpg", "library IMG_5652.jpg"]          # 한 장씩 차례로 올림
    assert "영수증 첨부(IMG_5648.jpg, IMG_5651.jpg, IMG_5652.jpg)" in rows[0]["처리"]


def attach_via_library(pg, tmp_path):
    """보관함(Available Receipts) 방식만 따로 봄(카드 경비·저장한 뒤 다시 연 경비에 쓰는 방식): 새 경비 화면에서 바로
    attach_from_available → (다 붙였는지, done, checks)"""
    bot = make_bot(pg, tmp_path)
    bot.back_to_report()
    bot.open_manual_expense("Parking & Toll")
    done, checks = [], []
    ok = bot.attach_from_available(photos(tmp_path), done, checks)
    return ok, done, checks


def test_already_uploaded_receipts_are_picked_not_uploaded_again(browser, tmp_path):
    """예전 실행에서 보관함(Available Receipts)에 올라간 사진은 다시 올리지 않고 골라서 붙임"""
    pytest.importorskip("cv2")
    pg = open_report(browser, library=("IMG_5648.jpg", "IMG_5651.jpg"))
    ok, done, checks = attach_via_library(pg, tmp_path)
    steps = pg.evaluate("steps")
    assert [s for s in steps if s.startswith("library")] == ["library IMG_5652.jpg"]       # 없는 것만 올림
    assert pg.evaluate("uploads") == ["IMG_5648.jpg", "IMG_5651.jpg", "IMG_5652.jpg"]
    assert ok and checks == []


def test_single_upload_that_only_reaches_library_is_picked(browser, tmp_path):
    """카드 경비처럼 한 장을 [Upload New Receipt]로 올렸는데 보관함에만 들어가면 골라서 붙임"""
    pytest.importorskip("cv2")
    pg = open_report(browser)
    bot = make_bot(pg, tmp_path)
    bot.back_to_report()
    bot.open_manual_expense("Parking & Toll")
    f = photos(tmp_path)[0]
    done, checks = [], []
    bot.attach_receipt(f, done, checks)
    assert pg.evaluate("uploads") == ["IMG_5648.jpg"] and checks == []
    assert done == ["영수증 첨부(IMG_5648.jpg)"]
    bot.attach_receipt(photos(tmp_path), done, checks)                             # 첫 장만 붙어 있으면 나머지를 올려 붙임
    assert pg.evaluate("uploads") == ["IMG_5648.jpg", "IMG_5651.jpg", "IMG_5652.jpg"] and checks == []
    assert "이미 첨부됨(IMG_5648.jpg)" in done


def test_fresh_uploads_wait_and_never_open_neighbour(browser, tmp_path):
    """막 올린 사진(2026-11-16 실제 실행: 10:37에 올라갔지만 안 붙음): 준비 중인 카드는 사진이 없어서 예전에는 위로 올라가
    옆 카드(다른 영수증)의 사진을 눌렀음 → 준비될 때까지 기다리고, 카드 밖으로는 올라가지 않음"""
    pytest.importorskip("cv2")
    pg = open_report(browser, slow=True, bg=True, library=("IMG_5000.jpg", "IMG_5001.jpg"))
    rows = make_bot(pg, tmp_path).create_manual_expenses([toll_rec(tmp_path)], tmp_path / "w")
    assert pg.evaluate("saved")["uploads"] == ["IMG_5648.jpg", "IMG_5651.jpg", "IMG_5652.jpg"]   # 옆 사진(IMG_5000)이 아님
    assert rows[0]["확인 필요"] == ""


def test_dialog_button_next_to_delete_without_file_name(browser, tmp_path):
    """'Receipt Image' 창에 파일 이름이 안 보여도 [Delete] 옆의 [Attach]/[Add]를 누름(경비 칸의 Remove·Add·Open이 아님).
    카드에 작은 아이콘 그림이 있어도 가장 큰 그림(영수증 사진)을 누름"""
    pytest.importorskip("cv2")
    pg = open_report(browser, noname=True, icons=True)
    rows = make_bot(pg, tmp_path).create_manual_expenses([toll_rec(tmp_path)], tmp_path / "w")
    assert "delete icon" not in pg.evaluate("steps")                               # 카드의 작은 아이콘이 아니라 사진을 누름
    assert pg.evaluate("saved")["uploads"] == ["IMG_5648.jpg", "IMG_5651.jpg", "IMG_5652.jpg"]
    assert [x for x in pg.evaluate("steps") if x.startswith(("attach", "add "))] == [
        "attach IMG_5648.jpg", "add IMG_5651.jpg", "add IMG_5652.jpg"]
    assert rows[0]["확인 필요"] == ""


def test_report_behind_and_add_with_icon_text(browser, tmp_path):
    """실제 화면(2026-11-16 실행 기록): 경비 화면 뒤에 리포트 화면이 깔려 있고, 두 번째 사진부터 누를 영수증 칸의 [Add]는
    이름이 정확히 'Add'가 아니어서 못 찾았음 → 사진 아래 Remove · Add · Open의 Add를 찾아 누름"""
    pytest.importorskip("cv2")
    pg = open_report(browser, behind=True)
    rows = make_bot(pg, tmp_path).create_manual_expenses([toll_rec(tmp_path)], tmp_path / "w")
    assert pg.evaluate("saved")["uploads"] == ["IMG_5648.jpg", "IMG_5651.jpg", "IMG_5652.jpg"]
    assert [x for x in pg.evaluate("steps") if x.startswith(("attach", "add "))] == [
        "attach IMG_5648.jpg", "add IMG_5651.jpg", "add IMG_5652.jpg"]
    assert rows[0]["확인 필요"] == ""


def test_panel_closing_slowly_after_add(browser, tmp_path):
    """2026-11-16 실제 실행: 두 번째 사진을 [Add]로 붙인 뒤 보관함이 닫히는 동안 세 번째 사진을 찾아 '없다'고 하고 넘어감
    → 보관함이 닫힐 때까지 기다린 뒤 영수증 칸의 [Add]로 다시 열어 세 장 모두 붙임"""
    pytest.importorskip("cv2")
    pg = open_report(browser, behind=True, slow_close=True)
    rows = make_bot(pg, tmp_path).create_manual_expenses([toll_rec(tmp_path)], tmp_path / "w")
    assert pg.evaluate("saved")["uploads"] == ["IMG_5648.jpg", "IMG_5651.jpg", "IMG_5652.jpg"]
    assert rows[0]["확인 필요"] == ""


def test_attach_after_save_when_new_expense_cannot_attach(browser, tmp_path):
    """저장하기 전의 새 경비 화면에서 사진을 붙일 수 없으면, 예전처럼 저장한 뒤 다시 열어 붙임"""
    pytest.importorskip("cv2")
    pg = open_report(browser, no_choose_new=True)
    rows = make_bot(pg, tmp_path).create_manual_expenses([toll_rec(tmp_path)], tmp_path / "w")
    steps = pg.evaluate("steps")
    assert pg.evaluate("saved")["uploads"] == ["IMG_5648.jpg", "IMG_5651.jpg", "IMG_5652.jpg"]
    assert steps.index("save") < steps.index("edit") and steps.count("save") == 2
    assert rows[0]["확인 필요"] == ""


def test_save_pressed_again_when_ignored_after_photos(browser, tmp_path):
    """2026-11-16 실제 화면: 사진 세 장을 붙인 새 경비에서 [Save Expense] 뒤 그대로 멈춤 → 그대로면 다시 눌러 저장"""
    pytest.importorskip("cv2")
    pg = open_report(browser, save_busy=True)
    bot = make_bot(pg, tmp_path)
    bot.save_dialog_wait_s, bot.save_retry_s = 0.5, 0.5
    rows = bot.create_manual_expenses([toll_rec(tmp_path)], tmp_path / "w")
    steps = pg.evaluate("steps")
    assert "save ignored" in steps and steps[-1] == "save"
    assert pg.evaluate("saved")["uploads"] == ["IMG_5648.jpg", "IMG_5651.jpg", "IMG_5652.jpg"]
    assert rows[0]["확인 필요"] == "" and pg.inner_text("#count") == "Expenses (1)"


def test_save_button_under_fixed_top_bar(browser, tmp_path):
    """2026-11-16 실제 화면: 사진을 붙이고 나면 화면이 내려가 [Save Expense]가 위쪽 회사 로고 머리띠 밑에 가려져 못 누름
    → 단추가 보이게 올린 뒤 누름"""
    pytest.importorskip("cv2")
    pg = open_report(browser, covered_save=True)
    rows = make_bot(pg, tmp_path).create_manual_expenses([toll_rec(tmp_path)], tmp_path / "w")
    assert rows[0]["오류"] == "" and rows[0]["확인 필요"] == ""
    assert pg.evaluate("saved")["uploads"] == ["IMG_5648.jpg", "IMG_5651.jpg", "IMG_5652.jpg"]
    assert pg.evaluate("steps")[-1] == "save" and pg.inner_text("#count") == "Expenses (1)"


def test_upload_goes_to_panel_not_expense_pane(browser, tmp_path):
    """2026-11-16 실행 기록: 보관함에 없는 사진을 올릴 때 보관함 밑에 깔린 경비 칸의 [Upload New Receipt]를 눌러 사진이
    경비에 바로 붙고(그 사진 한 장만), 그 이름을 보관함 카드로 잘못 봐서 나머지를 못 붙였음
    → 보관함 안의 [Upload New Receipt]로 올리고 차례대로 붙임"""
    pytest.importorskip("cv2")
    pg = open_report(browser, real_pane=True, library=("IMG_5648.jpg", "IMG_5652.jpg"))
    ok, done, checks = attach_via_library(pg, tmp_path)
    steps = pg.evaluate("steps")
    assert not [s for s in steps if s.startswith("direct")]                        # 경비 칸의 업로드 단추를 누르지 않음
    assert [s for s in steps if s.startswith("library")] == ["library IMG_5651.jpg"]    # 없는 사진만 보관함에
    assert [x for x in steps if x.startswith(("attach", "add "))] == [
        "attach IMG_5648.jpg", "add IMG_5651.jpg", "add IMG_5652.jpg"]
    assert pg.evaluate("uploads") == ["IMG_5648.jpg", "IMG_5651.jpg", "IMG_5652.jpg"]
    assert ok and checks == []


def test_resume_after_a_middle_photo_was_attached_directly(browser, tmp_path):
    """②를 다시 돌렸을 때 가운데 사진(IMG_5651) 한 장만 붙어 있는 경비(지난 실행에서 바로 붙음): 나머지 두 장을 [Add]로
    붙이고 순서 확인을 남김(예전에는 첫 사진이 아니라고 '이미 첨부됨'으로 건너뜀)"""
    pytest.importorskip("cv2")
    pg = open_report(browser, real_pane=True, library=("IMG_5648.jpg", "IMG_5652.jpg"))
    pg.goto("http://concur.test/report")
    pg.evaluate("() => { openDialog(); pickType(document.querySelector('.t')); }")
    pg.evaluate("""() => { uploads = ['IMG_5651.jpg']; ['drop', 'upload-receipt-button', 'choose'].forEach(id => show(id, 0));
                           show('viewer', 1); document.getElementById('vname').textContent = 'IMG_5651.jpg'; }""")
    bot = make_bot(pg, tmp_path)
    s1, s2, s3 = photos(tmp_path)
    done, checks = [], []
    ok = bot.attach_from_available([s1, s2, s3], done, checks)
    assert ok and pg.evaluate("uploads") == ["IMG_5651.jpg", "IMG_5648.jpg", "IMG_5652.jpg"]
    assert any("순서 확인" in c for c in checks) and "이미 첨부됨(IMG_5651.jpg)" in done


def test_presave_photos_uploaded_together(browser, tmp_path):
    """2026-11-16 23:26 실행 기록·사용자 화면: 저장 전 새 경비는 사진을 한 장 붙이면 그 아래에 [Remove]만 있고 [Add]가 없어
    세 경비 모두 첫 사진만 붙었음 → 경비 칸의 [Upload New Receipt]로 여러 장을 한 번에 올려 차례대로 붙이고 한 번 저장"""
    pytest.importorskip("cv2")
    pg = open_report(browser, real_pane=True, presave_no_add=True)
    rows = make_bot(pg, tmp_path).create_manual_expenses([toll_rec(tmp_path)], tmp_path / "w")
    steps = pg.evaluate("steps")
    assert [s for s in steps if s.startswith("direct")] == [
        "direct IMG_5648.jpg", "direct IMG_5651.jpg", "direct IMG_5652.jpg"]          # 한 번에, 차례대로
    assert pg.evaluate("saved")["uploads"] == ["IMG_5648.jpg", "IMG_5651.jpg", "IMG_5652.jpg"]
    assert steps.count("save") == 1 and "edit" not in steps and steps[-1] == "save"
    assert rows[0]["확인 필요"] == "" and "영수증 첨부(IMG_5648.jpg, IMG_5651.jpg, IMG_5652.jpg)" in rows[0]["처리"]


def test_presave_one_photo_then_add_after_save(browser, tmp_path):
    """파일 고르기 창이 한 장만 받으면 저장 전에는 첫 장만 붙이고, 저장한 뒤 다시 열어 [Add]로 나머지를 차례대로 붙임"""
    pytest.importorskip("cv2")
    pg = open_report(browser, real_pane=True, presave_no_add=True, multiple=False)
    rows = make_bot(pg, tmp_path).create_manual_expenses([toll_rec(tmp_path)], tmp_path / "w")
    steps = pg.evaluate("steps")
    assert pg.evaluate("expenses")[0]["uploads"] == ["IMG_5648.jpg", "IMG_5651.jpg", "IMG_5652.jpg"]
    assert steps.index("direct IMG_5648.jpg") < steps.index("save") < steps.index("edit") < steps.index("add IMG_5651.jpg")
    assert steps.index("add IMG_5651.jpg") < steps.index("add IMG_5652.jpg") and steps.count("save") == 2
    assert rows[0]["확인 필요"] == ""


def test_presave_upload_goes_to_library_then_attach_after_save(browser, tmp_path):
    """저장 전 화면에서 올린 사진이 경비에 안 붙고 보관함으로만 가면(2026-11-15 네 번째 실제 확인) 저장한 뒤 다시 열어
    보관함에서 차례대로 붙임(다시 올리지 않음)"""
    pytest.importorskip("cv2")
    pg = open_report(browser, presave_no_add=True)
    rows = make_bot(pg, tmp_path).create_manual_expenses([toll_rec(tmp_path)], tmp_path / "w")
    steps = pg.evaluate("steps")
    assert [s for s in steps if s.startswith("library")] == ["library IMG_5648.jpg,IMG_5651.jpg,IMG_5652.jpg"]
    assert [x for x in steps if x.startswith(("attach", "add "))] == [
        "attach IMG_5648.jpg", "add IMG_5651.jpg", "add IMG_5652.jpg"]
    assert steps.index("save") < steps.index("edit") and steps.count("save") == 2
    assert pg.evaluate("expenses")[0]["uploads"] == ["IMG_5648.jpg", "IMG_5651.jpg", "IMG_5652.jpg"]
    assert rows[0]["확인 필요"] == ""


REPORT_BEHIND_SCROLLED = """<!doctype html><meta charset="utf-8">
<div id="report"><h1>Monthly Expense for Oct</h1><button id="addbtn">Add Expense</button><div style="height:200px"></div></div>
<div id="form" style="position:absolute; top:0; left:0; right:0; min-height:4000px; z-index:10; background:#fff">
  <div style="height:90px"></div><button id="save" onclick="window.pressed = (window.pressed || 0) + 1">Save Expense</button>
  <div style="height:3800px">Receipts</div></div>
<div id="bar" style="position:fixed; top:0; left:0; right:0; height:60px; background:#eee; z-index:50">LOGO</div>"""


def test_report_check_does_not_scroll_or_misread_hidden_report(browser, tmp_path):
    """2026-11-16 사용자 화면: 저장을 기다리며 '리포트로 돌아왔나'를 볼 때 뒤에 깔린 리포트의 [Add Expense]를 보이게 옮기느라
    화면이 위아래로 흔들렸고, 화면 밖으로 밀린 그 단추를 '보인다'고 봐서 저장이 안 됐는데 끝났다고 넘어감(숙박에서 멈춤).
    → 화면을 옮기지 않고, 경비 화면([Save Expense])이 열려 있으면 리포트가 아님. 내려간 화면의 [Save Expense]는 올려서 누름."""
    pg = browser.new_page()
    pg.set_content(REPORT_BEHIND_SCROLLED)
    pg.evaluate("window.scrollTo(0, 1500)")
    bot = make_bot(pg, tmp_path)
    assert bot._report_shown() is False
    assert pg.evaluate("window.scrollY") == 1500                                   # 화면을 옮기지 않음
    bot._press_top("경비 저장")
    assert pg.evaluate("window.pressed") == 1                                       # 올려서 누름
    pg.evaluate("document.getElementById('form').style.display = 'none'")
    assert bot._report_shown() is True


def test_receipt_stem():
    assert ca.receipt_stem("IMG_5648_업로드용.jpg") == "IMG_5648"
    assert ca.receipt_stem("IMG_5648.HEIC") == "IMG_5648" and ca.receipt_stem("KakaoTalk_1.jpg") == "KakaoTalk_1"


def test_empty_report_uses_manual_card(browser, tmp_path):
    pytest.importorskip("cv2")
    pg = open_report(browser, empty=True)
    rows = make_bot(pg, tmp_path).create_manual_expenses([toll_rec(tmp_path)], tmp_path / "w")
    assert pg.evaluate("steps")[:2] == ["manual", "type Parking & Toll"]           # 'No Expenses'의 왼쪽 카드
    assert pg.evaluate("saved")["amount"].replace(",", "") == "61500" and rows[0]["오류"] == ""


def test_manual_expense_needs_expense_type(page, tmp_path):
    stmt, s1, s2 = photos(tmp_path)
    rec = ca.Receipt(original=stmt, ocr_file=stmt, bundle=[s1], total=61500, dates=["2026-10-22"], category="")
    bot = ca.ConcurBot(ca.default_settings(), lambda s="": None)
    bot.page, bot.timeout_ms = page, 3000
    bot.screenshot = lambda name: None
    rows = bot.create_manual_expenses([rec], tmp_path / "w")
    assert "Expense Type이 없어" in rows[0]["오류"] and not rec.used_for


# ---------- ② 전체 실행: 카드 경비가 없어도 톨게이트는 새 경비로 ----------
from test_toll_bundle import TEXTS, FakeOcr       # noqa: E402  (사용내역 + 영수증 인쇄물 2장)


class FakeBot:
    made = []

    def __init__(self, cfg, log, ui, stop_flag):
        self.page = None

    def start(self):
        pass

    def wait_login(self, need_home=True):
        pass

    def find_report(self, tm):
        return None

    def create_report(self, tm):
        pass

    def add_available_expenses(self, tm):
        raise ca.NoExpenses("Concur에 사용 가능한 경비가 없습니다(0건).")

    def create_manual_expenses(self, recs, work_dir, start_no=1):
        FakeBot.made = [(r.original.name, [p.name for p in r.bundle], r.total, r.main_date, start_no) for r in recs]
        for r in recs:
            r.used_for = "새 경비"
        return [{"순번": start_no, "Concur 날짜": "10/22/2026", "Concur 금액": "61,500", "가게 이름": "", "영수증": recs[0].original.name,
                 "분류": "톨게이트", "병원": "", "병원 근거": "", "도시": "", "처리": "새 경비", "확인 필요": "", "오류": ""}]

    def wait_until_closed(self, stop_flag=None):
        pass

    def stop(self):
        pass


class UI:
    def __init__(self):
        self.infos = []

    def log(self, line):
        pass

    def info(self, title, msg):
        self.infos.append((title, msg))

    def confirm(self, title, msg):
        return True

    def yes_no_cancel(self, title, msg):
        return True

    def open(self, path):
        pass


def test_full_run_creates_toll_expense_without_card_expenses(tmp_path, monkeypatch):
    src = tmp_path / "src"
    src.mkdir()
    for name in TEXTS:
        p = src / name
        p.write_bytes(b"\xff\xd8\xff\xd9" + name.encode())
        t = datetime(2026, 11, 15, 10, 0).timestamp()
        os.utime(p, (t, t))
    base = tmp_path / "app"
    base.mkdir()
    cfg = ca.default_settings()
    cfg.update({"영수증_원본_폴더": str(src), "결과_자동_열기": False, "영수증_보정_사용": False})
    ca.save_settings(base / ca.SETTINGS_FILE_NAME, cfg)
    ca.save_category_rows(base / ca.CATEGORY_FILE_NAME, [TOLL])
    monkeypatch.setattr(ca, "WindowsOcr", FakeOcr)
    monkeypatch.setattr(ca, "ConcurBot", FakeBot)
    FakeBot.made = []
    ui = UI()
    ca.run_job("full", "2026-10", ui, base)
    assert FakeBot.made == [("IMG_5648.JPG", ["IMG_5651.JPG", "IMG_5652.JPG"], 61500, "2026-10-22", 1)]
    assert ui.infos[-1][0] == "완료" and "경비 1건" in ui.infos[-1][1]



# ---------- 카드 경비가 없는 숙박 영수증도 새 경비로(사용자 요청 2026-11-15) ----------
def test_manual_accommodation_with_vendor(browser, tmp_path):
    pytest.importorskip("cv2")
    from PIL import Image
    html = report_html().replace("'Taxi'];", "'Taxi', 'Accommodation'];")
    pg = browser.new_page()
    pg.set_default_timeout(3000)
    pg.route("**/report", lambda route: route.fulfill(body=html, content_type="text/html; charset=utf-8"))
    slip, inv = tmp_path / "IMG_5646.jpg", tmp_path / "IMG_5647.jpg"
    for f in (slip, inv):
        Image.new("RGB", (300, 400), (240, 240, 240)).save(f, quality=90)
    bot = make_bot(pg, tmp_path)
    bot.cfg["Undefined일_때_선택할_유형"]["숙박"] = "Accommodation"
    rec = ca.Receipt(original=slip, ocr_file=slip, bundle=[inv], total=75000, dates=["2026-10-20"], category="숙박",
                     vendor="가나 1호점")
    rows = bot.create_manual_expenses([rec], tmp_path / "w")
    saved = pg.evaluate("saved")
    assert saved["type"] == "Accommodation" and saved["date"] == "10/20/2026" and saved["vendor"] == "Gana No.1"   # 영어로
    assert "Vendor Name 확인(Gana No.1: 한글 이름을 로마자로 바꿈)" in rows[0]["확인 필요"]
    assert saved["uploads"] == ["IMG_5646.jpg", "IMG_5647.jpg"] and rows[0]["오류"] == ""      # 영수증 → 인보이스


def test_full_run_creates_unmatched_receipts_without_asking(tmp_path, monkeypatch):
    from test_stay_bundle import NOL_SLIP
    src = tmp_path / "src"
    src.mkdir()
    p = src / "IMG_5646.JPG"
    p.write_bytes(b"\xff\xd8\xff\xd9")
    t = datetime(2026, 11, 15, 10, 0).timestamp()
    os.utime(p, (t, t))
    base = tmp_path / "app"
    base.mkdir()
    cfg = ca.default_settings()
    cfg.update({"영수증_원본_폴더": str(src), "결과_자동_열기": False, "영수증_보정_사용": False})
    ca.save_settings(base / ca.SETTINGS_FILE_NAME, cfg)

    class Ocr(FakeOcr):
        def read(self, path):
            return NOL_SLIP

    monkeypatch.setattr(ca, "WindowsOcr", Ocr)
    monkeypatch.setattr(ca, "ConcurBot", FakeBot)
    FakeBot.made = []
    asked = []

    class AskUI(UI):
        def confirm(self, title, msg):
            asked.append((title, msg))
            return True

    ui = AskUI()
    ca.run_job("full", "2026-10", ui, base)
    assert FakeBot.made == [("IMG_5646.JPG", [], 75000, "2026-10-20", 1)]       # 넣을 카드 경비가 없으면 묻지 않고 만듦
    assert not any(a[0] == "카드 경비가 없는 영수증" for a in asked)



def test_cards_below_need_scrolling(browser, tmp_path):
    """보관함 목록은 아래로 내려야 나머지 카드가 나타남(실제 화면) — 내리며 찾아 다시 올리지 않고 붙임"""
    pytest.importorskip("cv2")
    others = tuple(f"IMG_90{i:02d}_업로드용.jpg" for i in range(8))
    pg = open_report(browser, library=others + ("IMG_5648.jpg", "IMG_5651.jpg", "IMG_5652.jpg"))
    ok, done, checks = attach_via_library(pg, tmp_path)
    steps = pg.evaluate("steps")
    assert not [s for s in steps if s.startswith("library")]                       # 이미 있는 사진은 다시 안 올림
    assert pg.evaluate("uploads") == ["IMG_5648.jpg", "IMG_5651.jpg", "IMG_5652.jpg"]
    assert ok and checks == []


def test_resume_when_only_first_receipt_attached(browser, tmp_path):
    """첫 장만 붙고 멈춘 경비는 ②를 다시 돌리면 나머지만 [Add]로 붙임(이미 여러 장 붙어 있으면 건너뜀)"""
    pytest.importorskip("cv2")
    pg = open_report(browser, library=("IMG_5648.jpg", "IMG_5651.jpg", "IMG_5652.jpg"))
    bot = make_bot(pg, tmp_path)
    bot.back_to_report()
    bot.open_manual_expense("Parking & Toll")
    pg.evaluate("current = 'IMG_5648.jpg'; attachDlg()")                             # 지난번에 첫 장만 붙음
    files = photos(tmp_path)
    done, checks = [], []
    bot.attach_receipt(files, done, checks, pick=True)
    assert pg.evaluate("uploads") == ["IMG_5648.jpg", "IMG_5651.jpg", "IMG_5652.jpg"] and checks == []
    assert done[0] == "이미 첨부됨(IMG_5648.jpg)"
    bot.attach_receipt(files, done, checks, pick=True)                              # 다 붙은 뒤에는 건너뜀
    assert pg.evaluate("uploads") == ["IMG_5648.jpg", "IMG_5651.jpg", "IMG_5652.jpg"] and done[-1] == "이미 첨부됨(건너뜀)"



def test_slow_loading_cards_are_waited_for(browser, tmp_path):
    """[Add]·[Choose From Available Receipts]로 연 보관함의 사진 목록이 늦게 뜨면(실제 화면) 다 뜰 때까지 기다렸다 고름
    — 예전에는 바로 못 찾아 같은 사진을 또 올리고 두 번째 사진부터 못 붙였음"""
    pytest.importorskip("cv2")
    html = report_html(library=("IMG_5648.jpg", "IMG_5651.jpg", "IMG_5652.jpg")).replace(
        "function openAvail() { shownCards = 4; show('avail', 1); show('pane', 0); drawCards(); }",
        "function openAvail() { shownCards = 4; show('avail', 1); show('pane', 0);"
        " document.getElementById('cards').innerHTML = ''; setTimeout(drawCards, 1500); }")
    assert "setTimeout(drawCards" in html
    pg = browser.new_page()
    pg.set_default_timeout(3000)
    pg.route("**/report", lambda route: route.fulfill(body=html, content_type="text/html; charset=utf-8"))
    ok, done, checks = attach_via_library(pg, tmp_path)
    steps = pg.evaluate("steps")
    assert not [s for s in steps if s.startswith("library")]                       # 다시 안 올림
    assert pg.evaluate("uploads") == ["IMG_5648.jpg", "IMG_5651.jpg", "IMG_5652.jpg"]
    assert ok and checks == []
