# -*- coding: utf-8 -*-
"""'사용 가능한 경비'에서 지난달 경비만 고르는 JavaScript(JS_SELECT_MONTH) 테스트.
PAD 흐름은 '10/'만 보고 골라 11/10 같은 다른 달 경비까지 넣었으므로 그 경우를 확인합니다."""

import os
from pathlib import Path

import pytest

import concur_auto as ca

sync_api = pytest.importorskip("playwright.sync_api")

ROWS = ["10/05/2026", "11/10/2026", "1/10/2026", "10/31/2025", "10/31/2026", "01/31/2026"]
HTML = "<div role='grid'><div role='row'><input type='checkbox' class='sapcnqr-checkbox__input'>Select all</div>" + "".join(
    f"<div role='row'><input type='checkbox' class='sapcnqr-checkbox__input' data-date='{d}'"
    f"{' checked' if d == '10/31/2026' else ''}><span>{d}</span><span><b>KRW</b> 4,500</span></div>"
    for d in ROWS) + "</div>"


@pytest.fixture(scope="module")
def page():
    exe = os.environ.get("PW_CHROMIUM") or "/opt/pw-browsers/chromium"
    with sync_api.sync_playwright() as p:
        browser = p.chromium.launch(executable_path=exe if Path(exe).exists() else None)
        yield browser.new_page()
        browser.close()


def checked_dates(page):
    return page.eval_on_selector_all("input[data-date]:checked", "els => els.map(e => e.dataset.date)")


def test_selects_only_target_month(page):
    page.set_content(HTML)
    r = page.evaluate(ca.JS_SELECT_MONTH, {"month": 10, "year": 2026})
    assert (r["matched"], r["clicked"]) == (2, 1)
    assert checked_dates(page) == ["10/05/2026", "10/31/2026"]


def test_single_digit_month_does_not_match_november(page):
    page.set_content(HTML.replace(" checked", ""))
    r = page.evaluate(ca.JS_SELECT_MONTH, {"month": 1, "year": 2026})
    assert r["matched"] == 2
    assert checked_dates(page) == ["1/10/2026", "01/31/2026"]
