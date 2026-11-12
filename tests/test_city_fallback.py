# -*- coding: utf-8 -*-
"""영수증 주소의 지역으로 도시 고르기(시가 없는 주소, 늘어난 도시 목록),
도시·Expense Type을 Concur 목록에서 못 골라도 영수증 첨부까지 계속하기"""

import os
from datetime import datetime
from pathlib import Path

import pytest

import concur_auto as ca


@pytest.mark.parametrize("text, city", [
    ("전북 익산 무왕로 895", "Iksan"),                       # '시'가 없는 주소
    ("전남 영암군 삼호읍", "Yeongam"),                       # 새로 넣은 군
    ("전남진도 고군면", "Jindo"),
    ("대전 유성구 대학로 99", "Daejeon (ex Taejon)"),         # 광역시 + 구
    ("(주)하이파킹 서울 강남구 / 전남 나주 빛가람로", "Naju"),   # 서울(본사)은 뒤로
    ("경기 광주 오포읍", ""),                                # 경기도 광주는 아님 → 기본 도시
    ("스타벅스 합계 4,500", ""),
])
def test_city_from_receipt_address(text, city):
    assert ca.city_name(ca.detect_city(ca.compact(text), ca.DEFAULT_CITIES)) == city


def test_receipt_city_overrides_default_gwangju():
    p = Path("IMG_0001.JPG")
    rec = ca.analyze_receipt(p, p, "원광대병원 주차정산기\n전북 익산 무왕로 895\n2026.11.05 주차요금 3,000",
                             ca.default_settings(), ca.DEFAULT_HOSPITALS, datetime(2026, 11, 5, 9, 0))
    assert ca.city_name(rec.city) == "Iksan"
    assert ca.city_to_choose(rec, "", ca.default_settings()["기본_도시"]) == "Iksan, KOREA, REPUBLIC OF"
    assert ca.city_to_choose(rec, "Gwangju (ex Kwangju), KOREA, REPUBLIC OF", ca.default_settings()["기본_도시"]) \
        == "Iksan, KOREA, REPUBLIC OF"


# ---------- 못 골라도 계속 ----------

sync_api = pytest.importorskip("playwright.sync_api")

PAGE = """<!doctype html><meta charset="utf-8">
<label for="comment">Comment</label><textarea id="comment"></textarea>
<label for="businessPurpose">Business Purpose</label><input id="businessPurpose">
<p>Expense Type</p><span name="expenseType" onclick="openList('type')">Undefined</span>
<p>City</p><span id="location" onclick="openList('city')"></span>
<div id="pop" style="display:none"><input id="q" oninput="filter()"><div id="items"></div></div>
<div id="pane"><button onclick="document.getElementById('f').click()">Upload New Receipt</button></div>
<input type="file" id="f" style="display:none" onchange="document.getElementById('pane').innerHTML='<span id=name>'+this.files[0].name+'</span>'">
<script>
const lists = {type: ["Parking & Toll", "Meals & Drinks", "Other"],
               city: ["Gwangju (ex Kwangju), KOREA, REPUBLIC OF", "Iksan, KOREA, REPUBLIC OF", "Mokpo, KOREA, REPUBLIC OF"]};
let which = null;
function show(list) { document.getElementById('items').innerHTML = list.map(t =>
  '<div class="sapcnqr-listbox-item__wrapper" onclick="pick(this)">' + t + '</div>').join(''); }
function openList(w) { which = w; document.getElementById('pop').style.display = 'block'; document.getElementById('q').value = '';
  show(lists[w]); document.getElementById('q').focus(); }
function filter() { const q = document.getElementById('q').value.toLowerCase(); show(lists[which].filter(t => t.toLowerCase().includes(q))); }
function pick(el) { document.querySelector(which === 'type' ? '[name=expenseType]' : '#location').textContent = el.textContent;
  document.getElementById('pop').style.display = 'none'; }
document.addEventListener('keydown', e => { if (e.key === 'Escape') document.getElementById('pop').style.display = 'none'; });
</script>
<style>span[name], #location { display:inline-block; min-width:220px; min-height:18px; border:1px solid #999 }</style>"""


@pytest.fixture
def page():
    exe = os.environ.get("PW_CHROMIUM") or "/opt/pw-browsers/chromium"
    with sync_api.sync_playwright() as p:
        b = p.chromium.launch(executable_path=exe if Path(exe).exists() else None)
        pg = b.new_page()
        pg.set_default_timeout(3000)
        yield pg
        b.close()


def test_missing_city_and_type_do_not_stop_receipt_upload(page, tmp_path):
    photo = tmp_path / "IMG_0002.JPG"
    photo.write_bytes(b"\xff\xd8\xff\xd9")
    page.set_content(PAGE)
    cfg = ca.default_settings()
    cfg["Undefined일_때_선택할_유형"]["톨게이트"] = "Tolls"                # Concur 목록에 없는 이름
    lines = []
    bot = ca.ConcurBot(cfg, lambda s="": lines.append(s))
    bot.page, bot.timeout_ms, bot.upload_wait_s, bot.upload_done_s = page, 3000, 3, 3
    rec = ca.Receipt(original=photo, ocr_file=photo, category="톨게이트", city="Yeongam, KOREA, REPUBLIC OF")
    checks = []
    done = bot.fill_expense(rec, photo, checks)
    assert "Expense Type(Tolls) 확인" in checks                          # 못 골랐다고 남기고
    assert "도시=Gwangju (ex Kwangju)(기본)" in done                      # 영암이 목록에 없으면 기본 광주
    assert "도시 'Yeongam'를 Concur 목록에서 못 찾아 기본 도시로" in checks
    assert "영수증 첨부(IMG_0002.JPG)" in done                           # 그래도 영수증은 붙임
    assert page.inner_text("#location") == "Gwangju (ex Kwangju), KOREA, REPUBLIC OF"
    assert page.inner_text("#name") == "IMG_0002.JPG"


def test_receipt_city_is_chosen(page, tmp_path):
    page.set_content(PAGE)
    bot = ca.ConcurBot(ca.default_settings(), lambda s="": None)
    bot.page, bot.timeout_ms = page, 3000
    rec = ca.Receipt(original=Path("a.jpg"), ocr_file=Path("a.jpg"), category="주차", city="Iksan, KOREA, REPUBLIC OF")
    done = bot.fill_expense(rec, None, [])
    assert "도시=Iksan" in done and page.inner_text("#location") == "Iksan, KOREA, REPUBLIC OF"
    assert "Expense Type=Parking & Toll" in done
