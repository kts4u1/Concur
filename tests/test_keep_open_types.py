# -*- coding: utf-8 -*-
"""작업이 끝나도 Edge 창을 열어 두기, Concur Expense Type 목록을 읽어 자동 입력표에 더하기"""

import os
import threading
import time
from datetime import datetime
from pathlib import Path

import pytest

import concur_auto as ca

sync_api = pytest.importorskip("playwright.sync_api")

TYPES = ["Car mileage", "Parking & Toll", "Meals & Drinks", "Other", "Training"] + [f"Type {i:02d}" for i in range(40)]

# Concur처럼 보이는 부분만 그려 주는(내리면 다음 항목이 생기는) Expense Type 목록
DROPDOWN = """<!doctype html><meta charset="utf-8"><p>Expense Type *</p>
<div class="field__control"><span name="expenseType" onclick="openList()">Undefined</span></div>
<div id="pop" style="display:none"><input placeholder="Search for an expense type">
  <div style="font-weight:bold">Recently Used</div>
  <div id="box" style="height:150px; overflow:auto; position:relative"><div id="spacer" style="position:relative"></div></div>
</div>
<script>
const all = %s, H = 30;
function render() {
  const box = document.getElementById('box'), sp = document.getElementById('spacer');
  sp.style.height = (all.length * H) + 'px';
  const start = Math.floor(box.scrollTop / H), end = Math.min(all.length, start + Math.ceil(150 / H) + 1);
  sp.innerHTML = all.slice(start, end).map((t, k) => '<div class="sapcnqr-listbox-item__wrapper" style="position:absolute;top:'
    + ((start + k) * H) + 'px;height:' + H + 'px">' + t + '</div>').join('');
}
function openList() { document.getElementById('pop').style.display = 'block'; render(); }
document.getElementById('box').addEventListener('scroll', render);
document.addEventListener('keydown', e => { if (e.key === 'Escape') document.getElementById('pop').style.display = 'none'; });
</script>""" % repr(TYPES)


@pytest.fixture
def browser():
    exe = os.environ.get("PW_CHROMIUM") or "/opt/pw-browsers/chromium"
    with sync_api.sync_playwright() as p:
        b = p.chromium.launch(executable_path=exe if Path(exe).exists() else None)
        yield b
        b.close()


def make_bot(page):
    bot = ca.ConcurBot(ca.default_settings(), lambda s="": None)
    bot.page = page
    bot.ctx = page.context
    return bot


def test_read_all_expense_types_from_long_list(browser):
    page = browser.new_page()
    page.set_content(DROPDOWN)
    bot = make_bot(page)
    assert bot.read_dropdown_options("Expense Type") == TYPES
    assert not page.is_visible("#pop")                     # 다 읽고 닫음


def test_inspect_screen_lists_expense_types(browser, tmp_path):
    page = browser.new_page()
    page.set_content(DROPDOWN)
    bot = make_bot(page)
    text = bot.inspect_screen(tmp_path / "화면조사.txt").read_text(encoding="utf-8")
    assert f"[Expense Type 목록] {len(TYPES)}개" in text and "  Meals & Drinks" in text
    assert bot.last_expense_types == TYPES


def test_no_dropdown_means_no_types(browser):
    page = browser.new_page()
    page.set_content("<p>다른 화면</p>")
    assert make_bot(page).read_dropdown_options("Expense Type") == []


def test_expense_types_added_to_table_once(tmp_path):
    cfg = ca.default_settings()
    path = tmp_path / ca.CATEGORY_FILE_NAME
    types = ["Car mileage", "Parking & Toll", "Meals & Drinks", "Other Staff Expenses", "Undefined", "Training"]
    assert ca.add_expense_types_to_table(path, types, cfg) == 3      # 이미 쓰는 Parking & Toll·Other Staff Expenses 빼고
    rows = ca.read_csv_dicts(path)
    assert [r["분류"] for r in rows] == ["식대", "카페", "주차", "Car mileage", "Meals & Drinks", "Training"]
    new = rows[4]
    assert (new["영수증에서 찾을 글자(쉼표로 구분)"], new["Expense Type이 Undefined일 때"]) == ("", "Meals & Drinks")
    assert ca.add_expense_types_to_table(path, types, cfg) == 0      # 두 번째는 그대로

    order = ca.apply_category_table(cfg, path)
    assert order == ["식대", "카페", "주차", "Car mileage", "Meals & Drinks", "Training"]
    p = Path("IMG_0001.JPG")
    rec = ca.analyze_receipt(p, p, "주차장 입차 3,000원", cfg, [], datetime(2026, 11, 3, 9, 0))
    assert rec.category == "주차"                                    # 찾을 글자가 빈 줄은 영수증에 안 맞춰짐


def test_wait_until_closed_returns_when_window_closed_or_stopped(browser):
    ctx = browser.new_context()
    page = ctx.new_page()
    bot = make_bot(page)
    stop = threading.Event()
    stop.set()
    started = time.time()
    bot.wait_until_closed(stop)                                      # [중지]를 누르면 바로 끝
    assert time.time() - started < 1
    page.close()
    started = time.time()
    bot.wait_until_closed()                                          # 창이 닫혀 있으면 바로 끝
    assert time.time() - started < 1


def test_wait_until_closed_waits_while_window_open(browser):
    page = browser.new_context().new_page()
    bot = make_bot(page)
    checks = {"n": 0}
    real_wait = page.wait_for_timeout

    def wait_then_user_closes(ms):          # 세 번째 확인 때 사용자가 창을 닫음
        checks["n"] += 1
        real_wait(100)
        if checks["n"] == 3:
            page.close()

    page.wait_for_timeout = wait_then_user_closes
    bot.wait_until_closed()
    assert checks["n"] == 3                 # 창이 열려 있는 동안은 계속 기다림


def test_remember_expense_types_in_settings(tmp_path):
    s = tmp_path / ca.SETTINGS_FILE_NAME
    assert ca.remember_expense_types(s, ["Car mileage", "Fuel", "Hotel", "Undefined"]) == 2
    assert ca.load_settings(s)[0]["Concur_경비_유형"][-2:] == ["Fuel", "Hotel"]
    assert ca.remember_expense_types(s, ["Fuel", "Taxi"]) == 0


def test_known_types_added_to_old_settings_file(tmp_path):
    s = tmp_path / ca.SETTINGS_FILE_NAME
    old = ["Car mileage", "Meals & Drinks", "Other", "Other Staff Expenses", "Parking & Toll", "Training", "Taxi"]
    s.write_text(__import__("json").dumps({"Concur_경비_유형": old}), encoding="utf-8")
    types = ca.load_settings(s)[0]["Concur_경비_유형"]
    assert types[:7] == old                                           # 예전 목록(직접 읽어 온 Taxi 포함)은 그대로
    for t in ("Accommodation", "M card taxi - Covered recipient", "Business Meal & Drinks (with clients)"):
        assert types.count(t) == 1                                    # 사용자가 보내 준 항목이 더해짐
    assert types.count("Meals & Drinks") == 1


SIMILAR = """<!doctype html><meta charset="utf-8"><p>Expense Type</p>
<span name="expenseType" onclick="document.getElementById('pop').style.display='block'">Undefined</span>
<div id="pop" style="display:none">%s</div>
<script>document.querySelectorAll('.sapcnqr-listbox-item__wrapper').forEach(el => el.onclick = () => {
  document.querySelector('[name=expenseType]').textContent = el.textContent;
  document.getElementById('pop').style.display = 'none'; });</script>""" % "".join(
    f'<div class="sapcnqr-listbox-item__wrapper">{t}</div>' for t in (
        "Accommodation Covered Recipient", "M card accommodation", "M card accommodation - Covered recipient",
        "Accommodation", "M card business meals and drinks (with clients)", "M card business meals & drinks - Covered recipient",
        "M card meals and drinks", "Meals & Drinks"))


@pytest.mark.parametrize("name", ["Accommodation", "Meals & Drinks", "M card business meals & drinks - Covered recipient"])
def test_similar_expense_type_names_chosen_exactly(browser, name):
    page = browser.new_page()
    page.set_default_timeout(3000)
    page.set_content(SIMILAR)
    bot = make_bot(page)
    trigger, value = bot.find_select(bot.sel("Expense Type"))
    assert value == "Undefined"
    bot.choose_option(trigger, name)
    assert page.inner_text("[name=expenseType]") == name               # 이름이 비슷한 다른 항목을 고르지 않음


def test_known_expense_types_from_user_captures():
    types = ca.KNOWN_EXPENSE_TYPES
    assert len(types) == len(set(types)) == 47                        # 사용자가 캡처한 47개, 겹치는 이름 없이
    for t in ("M Card Service Fee (Airfare Booking)", "M card parking & toll", "Trade show and Marcom Expenses",
              "Phone & Fax", "Taxi", "Parking & Toll", "Other Staff Expenses"):
        assert t in types
    assert all(t == t.strip() and "  " not in t for t in types)
