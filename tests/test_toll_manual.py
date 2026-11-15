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
  <button id="addbtn" onclick="document.getElementById('menu').style.display='block'">Add Expense</button>
  <div id="menu" role="menu" style="display:none">
    <div role="menuitem" onclick="openDialog()">Manually Create Expense</div>
    <div role="menuitem">Select from Available Expenses (0)</div></div></div>
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
  <label for="comment">Comment</label><textarea id="comment" oninput="note('comment')"></textarea>
  <div id="pane"><button id="up" onclick="document.getElementById('f').click()">Upload New Receipt</button></div>
  <input type="file" id="f" style="display:none" onchange="uploaded(this)" MULTIPLE></div>
<div id="pop" style="display:none"><div id="items"></div></div>
<script>
window.steps = []; window.uploads = []; window.saved = null;
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
  document.getElementById('type').textContent = el.textContent; }
function openList(w, el) { which = w; target = el; show('pop', 1); document.getElementById('items').innerHTML =
  lists[w].map(t => '<div class="sapcnqr-listbox-item__wrapper" onclick="pick(this)">' + t + '</div>').join(''); }
function pick(el) { target.textContent = el.textContent; note(which); show('pop', 0); }
function uploaded(inp) { const names = Array.from(inp.files).map(f => f.name); uploads.push(...names);
  steps.push('upload ' + names.join(',')); document.getElementById('up').style.display = 'none'; inp.value = ''; }
if (EMPTY) { document.getElementById('empty').style.display = 'block'; document.getElementById('addbtn').style.display = 'none'; }
function save() { steps.push('save');
  saved = {type: document.getElementById('type').textContent, date: document.getElementById('transactionDate-date-input-field-input').value,
           amount: document.getElementById('transactionAmount').value, comment: document.getElementById('comment').value,
           purpose: document.getElementById('businessPurpose').value, city: document.getElementById('location').textContent,
           lcrt: document.querySelector('#custom32 span').textContent,
           receipt: document.querySelector('[name=receiptType]').textContent, uploads: uploads.slice()};
  show('form', 0); show('report', 1); document.getElementById('count').textContent = 'Expenses (1)';
  show('empty', 0); show('addbtn', 1); }
</script>
<style>span[name], #location, .sapcnqr-select-field__input { display:inline-block; min-width:220px; min-height:18px; border:1px solid #999 }
.t { padding: 4px }</style>"""

TOLL = ["톨게이트", "통행료, 하이패스, 사용내역", "Toll fee for service", "For service", "Parking & Toll", "", "", "", "", ""]


def report_html(multiple=True, empty=False):
    return (REPORT.replace(" MULTIPLE>", " multiple>" if multiple else ">")
            .replace("if (EMPTY)", "if (true)" if empty else "if (false)"))


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
    one = "upload IMG_5648.jpg,IMG_5651.jpg,IMG_5652.jpg"                          # 파일 고르기 창에서 한 번에 여러 장
    assert steps.index("date") < steps.index("amount") < steps.index(one)
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
    assert [s for s in pg.evaluate("steps") if s.startswith("upload")] == [
        "upload IMG_5648.jpg", "upload IMG_5651.jpg", "upload IMG_5652.jpg"]             # 한 장씩 차례로
    assert "영수증 더 첨부(IMG_5651.jpg, IMG_5652.jpg)" in rows[0]["처리"]


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
