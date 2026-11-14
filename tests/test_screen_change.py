# -*- coding: utf-8 -*-
"""Concur 화면이 바뀌었을 때: 위치(id)가 바뀌어도 칸 이름(라벨)으로 찾기, 설정 '화면_요소'로 고치기, [화면 조사] 파일"""

import json
import os
from pathlib import Path

import pytest

import concur_auto as ca

sync_api = pytest.importorskip("playwright.sync_api")

OLD_IDS = """<!doctype html><meta charset="utf-8"><h1>Parking &amp; Toll KRW 99,000</h1>
<label for="transactionDate-date-input-field-input">Transaction Date</label>
<input id="transactionDate-date-input-field-input" value="11/02/2026">
<label for="transactionAmount">Amount</label><input id="transactionAmount" value="99,000">
<label for="comment">Comment</label><textarea id="comment"></textarea>
<label for="businessPurpose">Business Purpose</label><input id="businessPurpose">
<button>Save Expense</button><button>Upload New Receipt</button>"""

# Concur가 바뀌어 id가 모두 달라졌지만 칸 이름은 그대로인 화면
NEW_IDS = """<!doctype html><meta charset="utf-8"><h1>Parking &amp; Toll KRW 99,000</h1>
<label for="f1">Transaction Date *</label><input id="f1" name="txnDate" value="11/02/2026">
<label for="f2">Amount</label><input id="f2" name="amt" value="99,000">
<label for="f3">Comment</label><textarea id="f3" name="cmt"></textarea>
<label for="f4">Business Purpose</label><input id="f4" name="bp">
<button>Save Expense</button>"""


@pytest.fixture
def page():
    exe = os.environ.get("PW_CHROMIUM") or "/opt/pw-browsers/chromium"
    with sync_api.sync_playwright() as p:
        b = p.chromium.launch(executable_path=exe if Path(exe).exists() else None)
        pg = b.new_page()
        pg.set_default_timeout(3000)
        yield pg
        b.close()


def make_bot(page, cfg=None):
    bot = ca.ConcurBot(cfg or ca.default_settings(), lambda s="": None)
    bot.page = page
    bot.timeout_ms = 3000
    bot.upload_wait_s = 1
    return bot


def test_fields_found_by_saved_position(page):
    page.set_content(OLD_IDS)
    bot = make_bot(page)
    assert bot.field("경비 날짜").input_value() == "11/02/2026"
    assert bot.field("경비 금액").input_value() == "99,000"


def test_fields_found_by_label_when_position_changed(page):
    page.set_content(NEW_IDS)
    bot = make_bot(page)
    assert bot.field("경비 날짜").get_attribute("id") == "f1"       # 'Transaction Date *'도 칸 이름으로 찾음
    assert bot.field("Comment").get_attribute("id") == "f3"
    rec = ca.Receipt(original=Path("a.jpg"), ocr_file=Path("a.jpg"), category="주차", hospital_en="Chosun University Hospital")
    done = bot.fill_expense(rec, None, [])
    assert done == ["Business Purpose 입력", "Comment 입력"]                # 화면 위에서 아래 순서
    assert page.input_value("#f3") == "Parking fee for service in Chosun University Hospital"
    assert page.input_value("#f4") == "For service"


def test_position_fixed_in_settings(page, tmp_path):
    page.set_content(NEW_IDS.replace('<label for="f3">Comment</label>', ""))     # 칸 이름도 없어진 경우
    s = tmp_path / ca.SETTINGS_FILE_NAME
    s.write_text(json.dumps({"화면_요소": {"Comment": "#f3"}}, ensure_ascii=False), encoding="utf-8")
    cfg, _ = ca.load_settings(s)
    assert cfg["화면_요소"]["Comment"] == "#f3" and cfg["화면_요소"]["경비 금액"] == "#transactionAmount"
    assert make_bot(page, cfg).field("Comment").get_attribute("id") == "f3"


def test_missing_field_says_what_to_fix(page):
    page.set_content("<p>빈 화면</p>")
    bot = make_bot(page)
    with pytest.raises(RuntimeError) as e:
        bot.field("Comment", timeout_s=0.5)
    assert "'Comment' 칸을 찾지 못했습니다" in str(e.value) and "화면_요소" in str(e.value) and "화면 조사" in str(e.value)


def test_inspect_screen_file(page, tmp_path):
    page.set_content(NEW_IDS)
    bot = make_bot(page)
    path = bot.inspect_screen(tmp_path / "화면조사_test.txt")
    text = path.read_text(encoding="utf-8")
    assert "주소:" in text and "[이 화면의 입력칸]" in text
    assert "input/ | #f2 | amt | Amount | 99,000" in text
    assert "textarea/ | #f3 | cmt | Comment |" in text
    assert "Save Expense" in text
    line = next(l for l in text.splitlines() if l.strip().startswith("경비 금액"))
    assert line.rstrip().endswith("없음")                                 # 저장된 위치로는 못 찾음 → 고칠 곳
    assert path.with_suffix(".png").exists()


def test_report_list_with_custom_position(page):
    page.set_content("""<div class="report-entries"><div role="row"><input type="checkbox"><span>STARBUCKS</span>
      <span>11/04/2026</span><span>KRW 4,500</span></div></div>""")
    rows = ca.parse_entry_rows(page.evaluate(ca.JS_READ_ROWS, ".report-entries"))
    assert rows == [("2026-11-04", 4500, "STARBUCKS 11/04/2026 KRW 4,500")]
