# -*- coding: utf-8 -*-
"""도시 자동 선택(기본 광주), 쉼표 없는 금액·두 자리 연도 읽기, 영수증 글자에서 병원 이름 찾기"""

import json
import os
from datetime import datetime
from pathlib import Path

import pytest

import concur_auto as ca

CITIES = ca.DEFAULT_CITIES


# ---------- 금액·날짜 ----------

@pytest.mark.parametrize("text, amounts", [
    ("주차요금 6000", [6000]),
    ("결제금액: 6,0OO", [6000]),                 # OCR이 0을 O로 읽음
    ("합계 6, 000원", [6000]),                   # 쉼표 뒤 띄어쓰기
    ("결제일시 2026-11-03 14:22 승인번호 12345678", []),
    ("아메리카노 4,500 합계 4,500", [4500]),
])
def test_amounts(text, amounts):
    assert ca.find_amounts(text) == amounts


@pytest.mark.parametrize("text, dates", [
    ("26.11.03 14:22", ["2026-11-03"]),
    ("거래일시 20261103 142233", ["2026-11-03"]),
    ("승인번호 20261103", []),                   # 날짜 글자 없이 8자리 숫자는 날짜로 보지 않음
    ("21.500.000", []),
    ("2026.11.03 / 26.11.03", ["2026-11-03"]),
])
def test_dates(text, dates):
    assert ca.find_ocr_dates(text) == dates


def test_receipt_with_bare_amount_is_kept():
    p = Path("IMG_0001.JPG")
    rec = ca.analyze_receipt(p, p, "목포한국병원 주차장\n주차요금 6000\n26.11.03 14:22", ca.default_settings(),
                             ca.OLD_DEFAULT_HOSPITALS, datetime(2026, 11, 20, 9, 0))
    assert rec.is_receipt and rec.total == 6000 and rec.ocr_dates == ["2026-11-03"]


def test_date_only_match_for_unreadable_amount():
    def r(name, date, total):
        return ca.Receipt(original=Path(name), ocr_file=Path(name), dates=[date], total=total,
                          amounts=[total] if total else [])
    only = r("a.jpg", "2026-11-03", 0)
    assert ca.find_receipt_by_date_only([only, r("b.jpg", "2026-11-03", 4500)], "2026-11-03") is only
    assert ca.find_receipt_by_date_only([only, r("c.jpg", "2026-11-03", 0)], "2026-11-03") is None   # 두 장이면 모름
    assert ca.find_receipt_by_date_only([r("d.jpg", "2026-11-03", 4500)], "2026-11-03") is None


# ---------- 영수증 글자에서 병원 ----------

@pytest.mark.parametrize("text, name", [
    ("의료법인 행복의료재단 목포한국병원\n전남 목포시 영산로 483", "목포한국병원"),
    ("(주)목포한국병원 주차장", "목포한국병원"),
    ("원광대학교 병원 주차장", "원광대학교병원"),
    ("원광대병원 주차정산기", "원광대병원"),
    ("병원주차장 이용 감사합니다", ""),
    ("행복동물병원", ""),
    ("스타벅스 광주점", ""),
])
def test_find_hospital_name_in_text(text, name):
    assert ca.find_hospital_name_in_text(text) == name


@pytest.mark.parametrize("name, key, en", [
    ("원광대병원", "원광대", "Wongwang University Hospital"),       # 엑셀의 '원광대학교병원'과 같은 찾을 글자
    ("목포한국병원", "목포한국", "Mokpo Hanguk Hospital"),
    ("광주기독병원", "광주기독", "Gwangju Christian Hospital"),
])
def test_found_hospital_key_and_english(name, key, en):
    assert (ca.hospital_key(name), ca.hospital_english_name(name)) == (key, en)


def test_hospital_not_in_list_is_found_and_added(tmp_path, monkeypatch):
    src = tmp_path / "src"
    src.mkdir()
    texts = {"2026-11-03_090000.jpg": "목포한국병원 주차장\n전남 목포시 영산로 483\n2026.11.03 주차요금 6000",
             "2026-11-03_120000.jpg": "2026.11.03 아메리카노 4,500",
             "2026-11-05_090000.jpg": "원광대병원 주차정산\n전북 익산시 무왕로 895\n2026.11.05 합계 3,100"}
    for name in texts:
        (src / name).write_bytes(b"x")

    class FakeOcr:
        def __init__(self, temp_dir):
            pass

        def read(self, path):
            return texts[path.name]

    monkeypatch.setattr(ca, "WindowsOcr", FakeOcr)
    cfg = ca.default_settings()
    cfg.update({"영수증_원본_폴더": str(src), "보정본_폴더": str(src / "p"), "영수증_보정_사용": False})
    results = tmp_path / "결과"
    results.mkdir()
    csv_path = tmp_path / ca.HOSPITAL_FILE_NAME
    ca.write_csv(csv_path, ca.HOSPITAL_HEADER, [["건양대", "Konyang University Hospital", "테스트"]])   # 두 병원이 없는 목록
    lines = []
    log = ca.Log(tmp_path / "log", sink=lines.append)
    tm = ca.TargetMonth(2026, 11)
    try:
        recs = ca.build_receipts(cfg, tm, ca.load_hospitals(csv_path), results, tmp_path / "w", log, csv_path)
        assert [(r.hospital_en, r.hospital_src, ca.city_name(r.city)) for r in recs] == [
            ("Mokpo Hanguk Hospital", "영수증 글자", "Mokpo"),
            ("Mokpo Hanguk Hospital", "같은 날 영수증", "Mokpo"),        # 카페 영수증: 같은 날 주차 영수증의 병원·도시
            ("Wongwang University Hospital", "영수증 글자", "Iksan"),
        ]
        assert recs[0].total == 6000
        added = [r for r in ca.load_hospital_rows(csv_path) if r[2].startswith("영수증에서 찾음")]
        assert [(r[0], r[1]) for r in added] == [("목포한국", "Mokpo Hanguk Hospital"), ("원광대", "Wongwang University Hospital")]
        assert any("새 병원 2곳" in line for line in lines)

        # 다시 실행: 이제 목록에서 찾고, 같은 병원을 또 더하지 않음
        recs = ca.build_receipts(cfg, tm, ca.load_hospitals(csv_path), results, tmp_path / "w", log, csv_path)
        assert recs[0].hospital_new == "" and recs[0].hospital_en == "Mokpo Hanguk Hospital"
        assert len([r for r in ca.load_hospital_rows(csv_path) if r[2].startswith("영수증에서 찾음")]) == 2

        # 목록에서 도시를 고치면 그 도시를 씀
        list_path = ca.list_csv_path(results, tm)
        rows = ca.read_csv_dicts(list_path)
        assert [r["도시"] for r in rows] == ["Mokpo", "Mokpo", "Iksan"]
        rows[1]["도시"] = "광주"
        ca.write_csv(list_path, ca.LIST_HEADER, [[r[h] for h in ca.LIST_HEADER] for r in rows])
        recs = ca.build_receipts(cfg, tm, ca.load_hospitals(csv_path), results, tmp_path / "w", log, csv_path)
        assert recs[1].city == "Gwangju (ex Kwangju), KOREA, REPUBLIC OF" and "도시" in recs[1].edited
    finally:
        log.close()


# ---------- 도시 ----------

@pytest.mark.parametrize("text, hospital_key, city", [
    ("광주광역시 동구 제봉로 42", "", "Gwangju (ex Kwangju)"),
    ("전북특별자치도 익산시 무왕로 895", "", "Iksan"),
    ("전남 목포시 영산로 483 / 본사 서울특별시 강남구", "", "Mokpo"),     # 먼저 나온 주소(가게 주소)
    ("경기도 광주시 오포읍", "", ""),                                     # 경기도 광주는 아님
    ("", "목포한국", "Mokpo"),                                           # 주소가 없으면 병원 이름 앞의 도시
    ("", "순천향대", ""),
    ("스타벅스", "", ""),
])
def test_detect_city(text, hospital_key, city):
    assert ca.city_name(ca.detect_city(ca.compact(text), CITIES, hospital_key)) == city


def test_resolve_city():
    assert ca.resolve_city("광주", CITIES) == "Gwangju (ex Kwangju), KOREA, REPUBLIC OF"
    assert ca.resolve_city("gwangju", CITIES) == "Gwangju (ex Kwangju), KOREA, REPUBLIC OF"
    assert ca.resolve_city("Mokpo", CITIES) == "Mokpo, KOREA, REPUBLIC OF"
    assert ca.resolve_city("Yeongam", CITIES) == "Yeongam, KOREA, REPUBLIC OF"
    assert ca.resolve_city("", CITIES) == ""


def test_city_to_choose():
    gwangju, mokpo = CITIES["광주"], CITIES["목포"]
    r = ca.Receipt(original=Path("a"), ocr_file=Path("a"), city=mokpo)
    assert ca.city_to_choose(r, "", gwangju) == mokpo                       # 비어 있으면 영수증 도시
    assert ca.city_to_choose(r, gwangju, gwangju) == mokpo                  # 다르면 영수증 도시로
    assert ca.city_to_choose(r, "Mokpo, KOREA, REPUBLIC OF", gwangju) is None
    no_city = ca.Receipt(original=Path("b"), ocr_file=Path("b"))
    assert ca.city_to_choose(no_city, "", gwangju) == gwangju               # 못 찾으면 기본 광주
    assert ca.city_to_choose(None, "", gwangju) == gwangju                  # 영수증 없는 경비도 기본 도시
    assert ca.city_to_choose(no_city, "Seoul, KOREA, REPUBLIC OF", gwangju) == gwangju   # 영수증이 있으면 기본 광주로 맞춤
    assert ca.city_to_choose(None, "Seoul, KOREA, REPUBLIC OF", gwangju) is None       # 영수증 없는 경비는 그대로


def test_default_city_is_gwangju_and_old_settings_are_moved_once(tmp_path):
    assert ca.default_settings()["기본_도시"] == "Gwangju (ex Kwangju), KOREA, REPUBLIC OF"
    p = tmp_path / ca.SETTINGS_FILE_NAME
    p.write_text(json.dumps({"기본_도시": ca.OLD_DEFAULT_CITY}, ensure_ascii=False), encoding="utf-8")
    cfg, _ = ca.load_settings(p)
    assert cfg["기본_도시"].startswith("Gwangju")
    # 그 뒤 사용자가 일부러 대전으로 바꾸면 그대로 둠
    saved = json.loads(p.read_text(encoding="utf-8"))
    saved["기본_도시"] = ca.OLD_DEFAULT_CITY
    p.write_text(json.dumps(saved, ensure_ascii=False), encoding="utf-8")
    assert ca.load_settings(p)[0]["기본_도시"] == ca.OLD_DEFAULT_CITY


# ---------- Concur 도시 드롭다운 ----------

sync_api = pytest.importorskip("playwright.sync_api")

DROPDOWN = """<!doctype html><meta charset="utf-8">
<div class="field__control"><span id="location" onclick="openList()"></span></div>
<div id="pop" style="display:none">
  <input id="q" oninput="filter()">
  <div id="items"></div>
</div>
<script>
const recent = ["Daejeon (ex Taejon), KOREA, REPUBLIC OF", "Gwangju (ex Kwangju), KOREA, REPUBLIC OF",
                "Jeonju, KOREA, REPUBLIC OF", "Iksan, KOREA, REPUBLIC OF", "Mokpo, KOREA, REPUBLIC OF"];
const all = recent.concat(["Kunsan (Gunsan), KOREA, REPUBLIC OF", "Naju, KOREA, REPUBLIC OF"]);
function show(list) {
  document.getElementById('items').innerHTML = list.map(t =>
    '<div class="sapcnqr-listbox-item__wrapper" onclick="pick(this)">' + t + '</div>').join('');
}
function openList() { document.getElementById('pop').style.display = 'block'; document.getElementById('q').value = '';
  show(recent); document.getElementById('q').focus(); }
function filter() { const q = document.getElementById('q').value.toLowerCase();
  show(all.filter(t => t.toLowerCase().includes(q))); }
function pick(el) { document.getElementById('location').textContent = el.textContent;
  document.getElementById('pop').style.display = 'none'; }
</script>
<style>#location{display:inline-block;min-width:200px;min-height:20px;border:1px solid #999}</style>"""


@pytest.fixture
def page():
    exe = os.environ.get("PW_CHROMIUM") or "/opt/pw-browsers/chromium"
    with sync_api.sync_playwright() as p:
        b = p.chromium.launch(executable_path=exe if Path(exe).exists() else None)
        yield b.new_page()
        b.close()


@pytest.mark.parametrize("want, shown", [
    ("Gwangju (ex Kwangju), KOREA, REPUBLIC OF", "Gwangju (ex Kwangju), KOREA, REPUBLIC OF"),   # 최근 목록에 있음
    ("Naju, KOREA, REPUBLIC OF", "Naju, KOREA, REPUBLIC OF"),                                   # 검색해서 고름
    ("Gunsan, KOREA, REPUBLIC OF", "Kunsan (Gunsan), KOREA, REPUBLIC OF"),                      # Concur 이름이 조금 달라도
])
def test_choose_city(page, want, shown):
    page.set_content(DROPDOWN)
    bot = ca.ConcurBot(ca.default_settings(), lambda s="": None)
    bot.page = page
    trigger, value = bot.find_select("#location")
    assert value == ""
    bot.choose_city(trigger, want)
    assert page.inner_text("#location") == shown


def test_choose_city_not_found(page):
    page.set_content(DROPDOWN)
    bot = ca.ConcurBot(ca.default_settings(), lambda s="": None)
    bot.page = page
    trigger, _ = bot.find_select("#location")
    with pytest.raises(RuntimeError, match="Yeongam"):
        bot.choose_city(trigger, "Yeongam, KOREA, REPUBLIC OF")
