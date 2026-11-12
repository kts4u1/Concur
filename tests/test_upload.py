# -*- coding: utf-8 -*-
"""영수증 첨부: 영수증 칸이 늦게 떠도, 버튼이 영어(Upload New Receipt)여도 올리는지.
기본 병원 목록(광주·전라권)으로 바꾸기, 서울(본사 주소)은 뒤로 미루기."""

import os
from datetime import datetime
from pathlib import Path

import pytest

import concur_auto as ca

sync_api = pytest.importorskip("playwright.sync_api")

PAGE = """<!doctype html><meta charset="utf-8">
<label>Comment<input id="comment"></label><label>Business Purpose<input id="businessPurpose"></label>
<div id="pane">{pane}</div>
<input type="file" id="f" style="display:none" onchange="uploaded(this)">
<script>
const later = {later!r};
if (later) setTimeout(() => {{ document.getElementById('pane').innerHTML = later; }}, {delay});
function uploaded(inp) {{
  const keep = {keep};
  document.getElementById('pane').innerHTML = (keep ? document.getElementById('pane').innerHTML : '')
    + '<img alt="receipt"><span id="name">' + inp.files[0].name + '</span>';
}}
</script>"""

UPLOAD_EN = """<button onclick="document.getElementById('f').click()">Upload New Receipt</button>
<button>Choose From Available Receipts</button>"""


@pytest.fixture
def page():
    exe = os.environ.get("PW_CHROMIUM") or "/opt/pw-browsers/chromium"
    with sync_api.sync_playwright() as p:
        b = p.chromium.launch(executable_path=exe if Path(exe).exists() else None)
        pg = b.new_page()
        pg.set_default_timeout(5000)
        yield pg
        b.close()


@pytest.fixture
def receipt_file(tmp_path):
    p = tmp_path / "IMG_0099.JPG"
    p.write_bytes(b"\xff\xd8\xff\xd9")
    return p


def make_bot(page):
    lines = []
    bot = ca.ConcurBot(ca.default_settings(), lambda s="": lines.append(s))
    bot.page = page
    bot.upload_wait_s, bot.upload_done_s = 6, 3
    return bot, lines


def show(page, pane="", later="", delay=0, keep="false"):
    page.set_content(PAGE.format(pane=pane, later=later, delay=delay, keep=keep))


def test_upload_button_in_english_appears_late(page, receipt_file):
    show(page, later=UPLOAD_EN, delay=2500)            # 영수증 칸이 2.5초 뒤에 뜸 (예전에는 1.5초만 기다리고 건너뜀)
    bot, _ = make_bot(page)
    done, checks = [], []
    bot.attach_receipt(receipt_file, done, checks)
    assert done == ["영수증 첨부(IMG_0099.JPG)"] and checks == []
    assert page.inner_text("#name") == "IMG_0099.JPG"


def test_already_attached_receipt_is_skipped(page, receipt_file):
    show(page, pane='<img alt="receipt"><button>Detach Receipt</button>')
    bot, _ = make_bot(page)
    done, checks = [], []
    bot.attach_receipt(receipt_file, done, checks)
    assert done == ["이미 첨부됨(건너뜀)"] and checks == []


def test_no_upload_button_is_reported_not_silently_skipped(page, receipt_file):
    show(page, pane="<p>빈 칸</p>")
    page.evaluate("document.getElementById('f').remove()")
    bot, lines = make_bot(page)
    bot.upload_wait_s = 1
    done, checks = [], []
    bot.attach_receipt(receipt_file, done, checks)
    assert done == [] and checks == ["영수증 첨부 버튼을 못 찾음(첨부됐는지 확인)"]


def test_hidden_file_input_is_used_when_button_text_unknown(page, receipt_file):
    show(page, pane="<button>Something else</button>")
    bot, _ = make_bot(page)
    bot.upload_wait_s = 1
    done, checks = [], []
    bot.attach_receipt(receipt_file, done, checks)
    assert done == ["영수증 첨부(IMG_0099.JPG)"]
    assert page.inner_text("#name") == "IMG_0099.JPG"


def test_upload_not_confirmed_is_reported(page, receipt_file):
    show(page, pane=UPLOAD_EN, keep="true")            # 올린 뒤에도 버튼이 남아 있음
    bot, _ = make_bot(page)
    done, checks = [], []
    bot.attach_receipt(receipt_file, done, checks)
    assert done == ["영수증 첨부(IMG_0099.JPG)"] and checks == ["영수증이 올라갔는지 확인"]


def test_fill_expense_attaches_after_typing(page, receipt_file):
    show(page, later=UPLOAD_EN, delay=1500)
    bot, _ = make_bot(page)
    rec = ca.Receipt(original=receipt_file, ocr_file=receipt_file, category="주차")
    checks = []
    done = bot.fill_expense(rec, receipt_file, checks)
    assert "Comment 입력" in done and "영수증 첨부(IMG_0099.JPG)" in done
    assert page.input_value("#comment") == "Parking fee for service"


# ---------- 기본 병원 목록: 영수증에서 확인한 병원만 ----------

def test_old_default_rows_are_removed_and_only_confirmed_hospitals_stay(tmp_path):
    p = tmp_path / ca.HOSPITAL_FILE_NAME
    old_rows = [pair + ["기본 목록(동료분 흐름)"] for pair in ca.OLD_DEFAULT_HOSPITALS]
    ca.write_csv(p, ca.HOSPITAL_HEADER, old_rows + [
        ["전남대", "Chonnam National University Hospital", "기본 목록(광주·전라권)"],   # 잠깐 넣었던 기본 줄
        ["원광대", "Wongwang University Hospital", "영수증에서 찾음: 원광대병원 (영문 이름 확인)"],
        ["조선대", "Chosun Univ. Hospital", "직접 넣음"],
    ])
    rows = ca.load_hospital_rows(p)
    assert [r[:2] for r in rows] == [
        ["목포한국", "Mokpo Hankook Hospital"],
        ["원광대", "Wonkwang University Hospital"],        # 자동 영문명 → 확인한 공식 영문명
        ["조선대", "Chosun Univ. Hospital"],               # 사용자가 넣은 줄은 그대로
    ]
    assert ca.load_hospital_rows(p) == rows                  # 두 번째부터는 그대로


def test_new_hospital_file_has_only_confirmed_hospitals(tmp_path):
    p = tmp_path / ca.HOSPITAL_FILE_NAME
    rows = ca.load_hospital_rows(p)
    assert rows == [["원광대", "Wonkwang University Hospital", "영수증에서 확인"],
                    ["목포한국", "Mokpo Hankook Hospital", "영수증에서 확인"]]


# ---------- 서울(본사 주소)은 뒤로 ----------

@pytest.mark.parametrize("text, city", [
    ("(주)하이파킹 서울특별시 강남구 테헤란로 / 광주광역시 동구 제봉로 42", "Gwangju (ex Kwangju)"),
    ("서울특별시 중구 세종대로 110", "Seoul"),
])
def test_seoul_is_used_only_when_no_other_city(text, city):
    assert ca.city_name(ca.detect_city(ca.compact(text), ca.DEFAULT_CITIES)) == city


def test_parking_receipt_with_seoul_head_office_gets_local_city():
    p = Path("IMG_0100.JPG")
    text = "하이파킹 본사 서울특별시 강남구\n조선대학교병원 주차장 광주광역시 동구\n2026.11.02 주차요금 99,000"
    rec = ca.analyze_receipt(p, p, text, ca.default_settings(), [["조선대", "Chosun University Hospital"]],
                             datetime(2026, 11, 2, 9, 0))
    assert (rec.total, rec.hospital_en, ca.city_name(rec.city)) == (99000, "Chosun University Hospital", "Gwangju (ex Kwangju)")
