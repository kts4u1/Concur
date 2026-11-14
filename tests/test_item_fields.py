# -*- coding: utf-8 -*-
"""자동 입력 항목마다 Concur 경비 화면의 칸(Expense Type·Business Purpose·City·LCRT·Receipt Status·
Personal Expense·Comment)을 정하고, Concur에 화면 위에서 아래 순서대로 넣기. '병원 없음'은 확인할 것에서 뺌."""

import os
from datetime import date
from pathlib import Path

import pytest

import concur_auto as ca

TOLL = ["톨게이트", "통행료, 톨게이트", "Toll fee for service in {hospital}", "For service", "Parking & Toll", "",
        "Mokpo, KOREA, REPUBLIC OF", "NA", "Receipt", "O"]


def cfg_with(tmp_path, rows):
    cfg = ca.default_settings()
    path = tmp_path / ca.CATEGORY_FILE_NAME
    ca.save_category_rows(path, rows)
    ca.apply_category_table(cfg, path)
    return cfg


def test_hospital_missing_is_not_a_check():
    rec = ca.Receipt(original=Path("IMG_5617.HEIC"), ocr_file=Path("IMG_5617.HEIC"), category="주차",
                     total=2000, dates=[date(2026, 11, 2)], date_from_ocr=True)
    rec.use = True
    assert rec.hospital_en == "" and ca.receipt_issues(rec, [rec]) == []      # 병원이 없어도 묻지 않음


def test_old_table_still_works(tmp_path):
    path = tmp_path / ca.CATEGORY_FILE_NAME
    old_header = ["분류", "영수증에서 찾을 글자(쉼표로 구분)", "Comment", "Business Purpose", "Expense Type이 Undefined일 때", "설명"]
    ca.write_csv(path, old_header, [["주차", "주차", "Parking fee in {hospital}", "For service", "Parking & Toll", ""],
                                    ["(분류 없음)", "", "", "", "Other Staff Expenses", ""]])
    cfg = ca.default_settings()
    assert ca.apply_category_table(cfg, path) == ["주차"]
    assert cfg["항목_추가_입력"]["주차"] == {"도시": "", "LCRT": "", "Receipt Status": "", "Personal Expense": False}
    rows = ca.load_category_rows(path, cfg)
    assert rows == [["주차", "주차", "Parking fee in {hospital}", "For service", "Parking & Toll", "", "", "", "", ""]]
    ca.save_category_rows(path, rows)
    assert list(ca.read_csv_dicts(path)[0].keys()) == ca.CATEGORY_HEADER         # 저장하면 새 칸이 생김


def test_city_priority():
    rec = ca.Receipt(original=Path("a.jpg"), ocr_file=Path("a.jpg"), city="Iksan, KOREA, REPUBLIC OF")
    gwangju = "Gwangju (ex Kwangju), KOREA, REPUBLIC OF"
    assert ca.city_to_choose(rec, "Seoul, KOREA, REPUBLIC OF", gwangju) == "Iksan, KOREA, REPUBLIC OF"
    assert ca.city_to_choose(rec, "Seoul", gwangju, "Mokpo, KOREA, REPUBLIC OF") == "Mokpo, KOREA, REPUBLIC OF"   # 항목에 정한 도시
    rec.edited.append("도시")                                                 # 영수증 목록에서 직접 고친 도시가 먼저
    assert ca.city_to_choose(rec, "Seoul", gwangju, "Mokpo, KOREA, REPUBLIC OF") == "Iksan, KOREA, REPUBLIC OF"


# ---------- 가짜 Concur 경비 화면(사용자 화면과 같은 칸 순서) ----------

PAGE = """<!doctype html><meta charset="utf-8"><h1>Parking &amp; Toll KRW 2,000</h1>
<p>Expense Type *</p><div class="field__control"><span name="expenseType" onclick="openList('type', this)">Undefined</span></div>
<p>Transaction Date *</p><input id="transactionDate-date-input-field-input" value="11/02/2026">
<label for="businessPurpose">Business Purpose</label><input id="businessPurpose" oninput="note('purpose')">
<p>City *</p><span id="location" onclick="openList('city', this)">Seoul, KOREA, REPUBLIC OF</span>
<label for="transactionAmount">Amount</label><input id="transactionAmount" value="2,000" readonly>
<p>LCRT *</p><div id="custom32"><span class="sapcnqr-select-field__input" onclick="openList('lcrt', this)"></span></div>
<p>Receipt Status *</p><span name="receiptType" onclick="openList('receipt', this)">No Receipt</span>
<input type="checkbox" id="personal" onchange="note('personal')"><label for="personal">Personal Expense (do not reimburse)</label>
<label for="comment">Comment</label><textarea id="comment" oninput="note('comment')"></textarea>
<div id="pop" style="display:none"><input id="q" oninput="filter()"><div id="items"></div></div>
<div id="pane"><button onclick="document.getElementById('f').click()">Upload New Receipt</button></div>
<input type="file" id="f" style="display:none" onchange="note('upload'); document.getElementById('pane').innerHTML='<span id=name>'+this.files[0].name+'</span>'">
<script>
window.steps = [];
function note(s) { if (steps[steps.length - 1] !== s) steps.push(s); }
const lists = {type: ["Meals & Drinks", "Parking & Toll", "Taxi"],
               city: ["Gwangju (ex Kwangju), KOREA, REPUBLIC OF", "Mokpo, KOREA, REPUBLIC OF", "Seoul, KOREA, REPUBLIC OF"],
               lcrt: ["NA", "HCP-0001"], receipt: ["No Receipt", "Receipt"]};
let which = null, target = null;
function show(list) { document.getElementById('items').innerHTML = list.map(t =>
  '<div class="sapcnqr-listbox-item__wrapper" onclick="pick(this)">' + t + '</div>').join(''); }
function openList(w, el) { which = w; target = el; document.getElementById('pop').style.display = 'block';
  document.getElementById('q').value = ''; show(lists[w]); document.getElementById('q').focus(); }
function filter() { const q = document.getElementById('q').value.toLowerCase(); show(lists[which].filter(t => t.toLowerCase().includes(q))); }
function pick(el) { target.textContent = el.textContent; note(which); document.getElementById('pop').style.display = 'none'; }
document.addEventListener('keydown', e => { if (e.key === 'Escape') document.getElementById('pop').style.display = 'none'; });
</script>
<style>span[name], #location, .sapcnqr-select-field__input { display:inline-block; min-width:220px; min-height:18px; border:1px solid #999 }</style>"""


@pytest.fixture
def page():
    sync_api = pytest.importorskip("playwright.sync_api")      # 브라우저가 필요한 테스트만 건너뜀
    exe = os.environ.get("PW_CHROMIUM") or "/opt/pw-browsers/chromium"
    with sync_api.sync_playwright() as p:
        b = p.chromium.launch(executable_path=exe if Path(exe).exists() else None)
        pg = b.new_page()
        pg.set_default_timeout(3000)
        pg.set_content(PAGE)
        yield pg
        b.close()


def make_bot(page, cfg):
    bot = ca.ConcurBot(cfg, lambda s="": None)
    bot.page, bot.timeout_ms, bot.upload_wait_s, bot.upload_done_s = page, 3000, 3, 3
    return bot


def test_fills_top_to_bottom_with_item_values(page, tmp_path):
    photo = tmp_path / "IMG_5617.jpg"
    photo.write_bytes(b"\xff\xd8\xff\xd9")
    cfg = cfg_with(tmp_path, [TOLL])
    rec = ca.Receipt(original=photo, ocr_file=photo, category="톨게이트", hospital_en="Mokpo Hankook Hospital",
                     city="Iksan, KOREA, REPUBLIC OF")
    checks = []
    done = make_bot(page, cfg).fill_expense(rec, photo, checks)
    assert page.evaluate("steps") == ["type", "purpose", "city", "lcrt", "receipt", "personal", "comment", "upload"]
    assert done == ["Expense Type=Parking & Toll", "Business Purpose 입력", "도시=Mokpo", "LCRT=NA",
                    "Receipt Status=Receipt", "Personal Expense 체크", "Comment 입력", "영수증 첨부(IMG_5617.jpg)"]
    assert page.inner_text("[name=expenseType]") == "Parking & Toll"
    assert page.inner_text("#location") == "Mokpo, KOREA, REPUBLIC OF"            # 항목에 정한 도시
    assert page.is_checked("#personal")
    assert page.input_value("#comment") == "Toll fee for service in Mokpo Hankook Hospital"
    assert checks == []                                                          # 자동 입력표대로 고른 것은 묻지 않음


def test_no_matching_item_leaves_expense_type(page, tmp_path):
    cfg = cfg_with(tmp_path, [TOLL])
    rec = ca.Receipt(original=Path("a.jpg"), ocr_file=Path("a.jpg"), category="")
    checks = []
    done = make_bot(page, cfg).fill_expense(rec, None, checks)
    assert page.inner_text("[name=expenseType]") == "Undefined"                  # 기본 Expense Type은 없앰
    assert "Expense Type 직접 고르기(맞는 자동 입력 항목 없음)" in checks
    assert done == ["도시=Gwangju (ex Kwangju)", "LCRT=NA", "Receipt Status=Receipt"]   # 나머지는 기본값
    assert not page.is_checked("#personal") and page.input_value("#comment") == ""


def test_values_already_right_are_left_alone(page, tmp_path):
    page.evaluate("""() => { document.querySelector('[name=expenseType]').textContent = 'Taxi';
        document.querySelector('#custom32 span').textContent = 'NA';
        document.querySelector('[name=receiptType]').textContent = 'Receipt'; }""")
    cfg = cfg_with(tmp_path, [TOLL[:6]])
    rec = ca.Receipt(original=Path("a.jpg"), ocr_file=Path("a.jpg"), category="톨게이트")
    make_bot(page, cfg).fill_expense(rec, None, [])
    assert page.evaluate("steps") == ["purpose", "city", "comment"]             # 이미 맞는 칸은 건드리지 않음
    assert page.inner_text("[name=expenseType]") == "Taxi"                      # 카드로 들어온 Expense Type은 그대로


# ---------- 편집 창 (tkinter와 화면이 있을 때만) ----------

tkinter = pytest.importorskip("tkinter") if os.environ.get("DISPLAY") else None


def all_texts(widget):
    out = []
    for child in widget.winfo_children():
        try:
            out.append(str(child.cget("text")))
        except Exception:
            pass
        out.extend(all_texts(child))
    return out


@pytest.mark.skipif(tkinter is None, reason="창을 띄울 화면이 없음")
def test_editor_fields_in_concur_order(tmp_path):
    cfg = ca.default_settings()
    ca.save_settings(tmp_path / ca.SETTINGS_FILE_NAME, cfg)
    app = ca.GuiApp(tmp_path)
    try:
        app.open_category_table()
        ed = app.editor
        right = ed.name_entry.master
        labels = [w.cget("text") for w in sorted(right.grid_slaves(column=0), key=lambda w: int(w.grid_info()["row"]))
                  if w.winfo_class() == "TLabel"]
        assert labels == ["항목 이름", "영수증에서 찾을 글자", "Concur 경비 화면에 넣을 값 (화면 위에서 아래 순서대로 입력)",
                          "Expense Type", "Business Purpose", "City", "LCRT", "Receipt Status", "Personal Expense",
                          "Comment", "미리보기"]
        assert not any("어느 항목에도 안 맞을 때" in t for t in all_texts(ed.win))      # 아래 기본 Expense Type 칸은 없앰

        ed.select_visible([ed.rows[i][0] for i in ed.visible].index("주차"))
        assert ed.vars["city"].get() == ca.CITY_FROM_RECEIPT and ed.vars["lcrt"].get() == "NA"
        assert ed.vars["receipt"].get() == "Receipt" and not ed.personal.get()
        ed.vars["city"].set("Mokpo, KOREA, REPUBLIC OF")
        ed.city_box.event_generate("<<ComboboxSelected>>")
        ed.personal.set(True)
        ed.on_edit()
        text = ed.preview.cget("text")
        assert "City: Mokpo, KOREA, REPUBLIC OF" in text and "Personal Expense: 체크" in text
        assert text.index("Expense Type") < text.index("Business Purpose") < text.index("City") < text.index("Comment")
        assert ed.save()
        parking = next(r for r in ca.load_category_rows(tmp_path / ca.CATEGORY_FILE_NAME, cfg) if r[0] == "주차")
        assert parking[ca.C_CITY:] == ["Mokpo, KOREA, REPUBLIC OF", "NA", "Receipt", "O"]

        ed.vars["city"].set(ca.CITY_FROM_RECEIPT)                             # 다시 '영수증 주소에서 찾기'로
        ed.on_edit()
        assert ed.rows[ed.current][ca.C_CITY] == ""
        ed.dirty = False
        ed.close()
    finally:
        app.root.destroy()
