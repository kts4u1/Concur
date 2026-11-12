# -*- coding: utf-8 -*-
"""Concur 칸에 글 넣기: 한꺼번에 넣고, 칸이 안 받으면 한 글자씩(예전 방식)"""

import os
import time
from pathlib import Path

import pytest

import concur_auto as ca

sync_api = pytest.importorskip("playwright.sync_api")

LONG = "Overtime meal fee for service in Wonkwang University Hospital " * 3


@pytest.fixture
def page():
    exe = os.environ.get("PW_CHROMIUM") or "/opt/pw-browsers/chromium"
    with sync_api.sync_playwright() as p:
        b = p.chromium.launch(executable_path=exe if Path(exe).exists() else None)
        pg = b.new_page()
        pg.set_default_timeout(5000)
        yield pg
        b.close()


def bot(page, how="한꺼번에"):
    cfg = ca.default_settings()
    cfg["글자_입력_방식"] = how
    b = ca.ConcurBot(cfg, lambda s="": None)
    b.page = page
    return b


def test_text_is_put_in_at_once(page):
    page.set_content('<textarea id="comment">예전 글</textarea><input id="next">')
    started = time.time()
    bot(page).type_text(page.locator("#comment"), LONG)
    assert page.input_value("#comment") == LONG
    assert time.time() - started < 1.0                      # 한 글자씩이면 180자 × 0.01초 이상


def test_field_that_only_takes_key_presses_gets_typed(page):
    # 키를 누를 때만 받아들이는 칸(한꺼번에 넣으면 지워 버림) → 한 글자씩으로 다시
    page.set_content("""<input id="d" onkeydown="window.k=true" oninput="if(!window.k) this.value=''"><input id="next">""")
    bot(page).type_text(page.locator("#d"), "11/01/2026")
    assert page.input_value("#d") == "11/01/2026"


def test_old_one_by_one_setting(page):
    page.set_content('<input id="c"><input id="next">')
    page.evaluate("() => { window.keys = 0; document.getElementById('c').addEventListener('keydown', () => window.keys++); }")
    bot(page, "한 글자씩").type_text(page.locator("#c"), "For service")
    assert page.input_value("#c") == "For service"
    assert page.evaluate("window.keys") >= len("For service")   # 한 글자씩 눌렀음
