# -*- coding: utf-8 -*-
"""이메일 링크 자동 로그인 테스트.
가짜 Concur 로그인 화면(이메일 → [다음] → [이메일 링크로 로그인])과 가짜 Outlook으로 확인합니다."""

import os
import sys
import threading
import types
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import pytest

import concur_auto as ca

# ---------- 메일에서 링크 찾기 ----------

MAIL = """<html><body>
<div style="background:#2a7ac0"><h2>SAP Concur 로그인 요청</h2></div>
<p>안녕하세요? 홍길동 님,</p>
<a href="https://www.concur.com/privacy">개인정보 처리방침</a>
<table><tr><td><a href="{href}" style="color:#fff"><span>SAP Concur</span> <b>로그인</b></a></td></tr></table>
<p>이 링크는 한 시간 동안 활성화되며 한 번만 사용할 수 있습니다.</p>
</body></html>"""


def test_extract_login_link():
    href = "https://www.concursolutions.com/nui/signin/emaillink?token=abc&amp;lang=ko"
    assert ca.extract_login_link(MAIL.format(href=href), "SAP Concur 로그인") == \
        "https://www.concursolutions.com/nui/signin/emaillink?token=abc&lang=ko"


def test_extract_login_link_wrapped_by_mail_security():
    href = "https://nam11.safelinks.protection.outlook.com/?url=https%3A%2F%2Fwww.concursolutions.com%2Fx&amp;data=1"
    assert ca.extract_login_link(MAIL.format(href=href), "SAP Concur 로그인").startswith(
        "https://nam11.safelinks.protection.outlook.com/?url=https%3A%2F%2Fwww.concursolutions.com")


def test_extract_login_link_none():
    assert ca.extract_login_link("<a href='https://example.com'>다른 링크</a>", "SAP Concur 로그인") is None
    assert ca.extract_login_link("", "SAP Concur 로그인") is None


# ---------- 가짜 Outlook ----------

class FakeMail:
    def __init__(self, subject, minutes_ago, href, now, sender="noreply@concursolutions.com", body=None):
        self.Class = 43
        self.Subject = subject
        self.SenderEmailAddress = sender
        # pywin32는 PC 시각에 UTC 표시를 붙여 돌려줌 → 프로그램은 tzinfo를 떼고 비교
        self.ReceivedTime = (now - timedelta(minutes=minutes_ago)).replace(tzinfo=timezone.utc)
        self.HTMLBody = (body or MAIL).format(href=href)


class FakeItems:
    def __init__(self, mails):
        self.mails, self.i = list(mails), 0

    def Sort(self, key, descending):
        self.mails.sort(key=lambda m: m.ReceivedTime, reverse=descending)

    def GetFirst(self):
        self.i = 0
        return self.mails[0] if self.mails else None

    def GetNext(self):
        self.i += 1
        return self.mails[self.i] if self.i < len(self.mails) else None


class FakeInbox:
    def __init__(self, mails, arriving=()):
        self.mails, self.arriving, self.reads = list(mails), list(arriving), 0

    @property
    def Items(self):
        self.reads += 1
        if self.reads >= 2 and self.arriving:       # 두 번째로 볼 때 새 메일이 도착
            self.mails += self.arriving
            self.arriving = []
        return FakeItems(self.mails)


@pytest.fixture
def fake_outlook(monkeypatch):
    box = {}
    pythoncom = types.ModuleType("pythoncom")
    pythoncom.CoInitialize = lambda: None
    pythoncom.CoUninitialize = lambda: None
    client = types.ModuleType("win32com.client")

    class App:
        def GetNamespace(self, name):
            return types.SimpleNamespace(GetDefaultFolder=lambda n: box["inbox"])

    client.Dispatch = lambda name: App()
    win32com = types.ModuleType("win32com")
    win32com.client = client
    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    return box


def test_outlook_picks_newest_login_mail_after_request(fake_outlook):
    now = datetime.now().replace(microsecond=0)
    fake_outlook["inbox"] = FakeInbox([
        FakeMail("SAP Concur 로그인 요청", 30, "https://c.example/old", now),     # 예전 요청 메일
        FakeMail("주간 회의", 0, "https://c.example/meeting", now, sender="boss@example.com"),
        FakeMail("SAP Concur 로그인 요청", 0, "https://c.example/new", now),
    ])
    lines = []
    link = ca.find_login_link_in_outlook({"로그인_메일_대기_초": 0}, now, lines.append)
    assert link == "https://c.example/new"


def test_outlook_waits_for_mail_to_arrive(fake_outlook):
    now = datetime.now().replace(microsecond=0)
    fake_outlook["inbox"] = FakeInbox([FakeMail("SAP Concur 로그인 요청", 30, "https://c.example/old", now)],
                                      arriving=[FakeMail("SAP Concur 로그인 요청", 0, "https://c.example/new", now)])
    link = ca.find_login_link_in_outlook({"로그인_메일_대기_초": 5}, now, lambda s: None, poll_s=0.01)
    assert link == "https://c.example/new"


def test_outlook_gives_up_after_timeout(fake_outlook):
    now = datetime.now().replace(microsecond=0)
    fake_outlook["inbox"] = FakeInbox([FakeMail("SAP Concur 로그인 요청", 30, "https://c.example/old", now)])
    lines = []
    assert ca.find_login_link_in_outlook({"로그인_메일_대기_초": 0}, now, lines.append) is None
    assert any("찾지 못했습니다" in line for line in lines)


def test_outlook_missing_package(monkeypatch):
    monkeypatch.setitem(sys.modules, "win32com", None)
    lines = []
    assert ca.find_login_link_in_outlook({}, datetime.now(), lines.append) is None
    assert any("pywin32" in line for line in lines)


# ---------- 가짜 Concur 로그인 화면 ----------

sync_api = pytest.importorskip("playwright.sync_api")

SIGNIN = """<!doctype html><meta charset="utf-8"><h1>로그인</h1>
<label for="u">사용자 이름, 검증된 이메일 주소 또는 SSO 코드</label>
<input id="u" type="text" value="{value}">
<button onclick="location.href='https://www.concursolutions.com/nui/signin/hrd?email='+encodeURIComponent(document.getElementById('u').value)">다음</button>
<label><input type="checkbox" checked> 로그인 정보 저장</label>"""

HRD = """<!doctype html><meta charset="utf-8"><h1>로그인</h1>
<p>다음으로 로그인:</p>
<button>회사 SSO</button><button>SAP Concur 암호</button><p>OR</p>
<button onclick="document.cookie='linkreq=1; path=/'; document.getElementById('m').textContent='이메일을 확인하세요'">이메일 링크로 로그인</button>
<p id="m"></p>"""

HOME = """<!doctype html><meta charset="utf-8"><h1>안녕하세요?</h1>
<div role="link" tabindex="0">경비 리포트 생성</div><h2>사용 가능한 경비(4)</h2>"""

REDIRECT = """<!doctype html><meta charset="utf-8"><script>location.replace({url!r})</script>"""

MAGIC = "https://www.concursolutions.com/nui/signin/emaillink?token=abc"


class FakeConcur:
    def __init__(self, logged_in=False, saved_email=""):
        self.logged_in, self.saved_email = logged_in, saved_email
        self.typed_email = None

    def handle(self, route, request):
        u = urlparse(request.url)
        if u.netloc == "us2.concursolutions.com" and u.path == "/home":
            body = HOME if self.logged_in else REDIRECT.format(url="https://www.concursolutions.com/nui/signin")
        elif u.path == "/nui/signin":
            body = SIGNIN.format(value=self.saved_email)
        elif u.path == "/nui/signin/hrd":
            self.typed_email = parse_qs(u.query).get("email", [""])[0]
            body = HRD
        elif u.path == "/nui/signin/emaillink":
            self.logged_in = True
            body = REDIRECT.format(url="https://us2.concursolutions.com/home")
        else:
            body = "not found"
        route.fulfill(status=200, content_type="text/html; charset=utf-8", body=body)


class UI:
    def __init__(self, paste=None):
        self.paste, self.asked = paste, []

    def ask_text(self, title, msg):
        self.asked.append(title)
        return self.paste


@pytest.fixture
def browser():
    exe = os.environ.get("PW_CHROMIUM") or "/opt/pw-browsers/chromium"
    with sync_api.sync_playwright() as p:
        b = p.chromium.launch(executable_path=exe if Path(exe).exists() else None)
        yield b
        b.close()


def make_bot(browser, site, cfg_changes=None, ui=None, mail_link=MAGIC, stop_flag=None):
    ctx = browser.new_context()
    ctx.route("**/*", site.handle)
    cfg = ca.default_settings()
    cfg.update(cfg_changes or {})
    lines = []
    bot = ca.ConcurBot(cfg, lambda s="": lines.append(s), ui, stop_flag)
    bot.page = ctx.new_page()
    bot.poll_s = 0.05
    bot.login_timeout_s = 20
    calls = []

    def fake_find(cfg_, since, log, check_stop):
        calls.append(since)
        return mail_link

    bot.find_mail_link = fake_find
    return bot, ctx, calls, lines


def link_requested(ctx):
    return any(c["name"] == "linkreq" for c in ctx.cookies())


def test_already_logged_in_just_continues(browser):
    site = FakeConcur(logged_in=True)
    bot, ctx, calls, lines = make_bot(browser, site)
    bot.wait_login()
    assert calls == [] and not link_requested(ctx)
    assert bot.page.url == "https://us2.concursolutions.com/home"


def test_saved_email_next_then_email_link(browser):
    site = FakeConcur(saved_email="saved@example.com")
    bot, ctx, calls, lines = make_bot(browser, site)
    bot.wait_login()
    assert site.typed_email == "saved@example.com"
    assert link_requested(ctx) and len(calls) == 1
    assert bot.page.url == "https://us2.concursolutions.com/home"
    assert "로그인 확인" in lines


def test_empty_email_is_filled_from_settings(browser):
    site = FakeConcur(saved_email="")
    bot, ctx, calls, lines = make_bot(browser, site, {"로그인_이메일": "user@example.com"})
    bot.wait_login()
    assert site.typed_email == "user@example.com"
    assert bot.page.url == "https://us2.concursolutions.com/home"


def test_paste_link_when_mail_not_found(browser):
    site = FakeConcur(saved_email="saved@example.com")
    ui = UI(paste="  " + MAGIC + "  ")
    bot, ctx, calls, lines = make_bot(browser, site, ui=ui, mail_link=None)
    bot.wait_login()
    assert ui.asked == ["로그인 링크 붙여넣기"]
    assert bot.page.url == "https://us2.concursolutions.com/home"


def test_email_link_requested_only_once(browser):
    site = FakeConcur(saved_email="saved@example.com")
    ui = UI(paste=None)                               # 메일도 없고 붙여넣기도 취소
    bot, ctx, calls, lines = make_bot(browser, site, ui=ui, mail_link=None)
    bot.login_timeout_s = 1
    with pytest.raises(RuntimeError):
        bot.wait_login()
    assert len(calls) == 1 and ui.asked == ["로그인 링크 붙여넣기"]
    assert any("직접 로그인하세요" in line for line in lines)


def test_manual_mode_does_not_click(browser):
    site = FakeConcur(saved_email="saved@example.com")
    bot, ctx, calls, lines = make_bot(browser, site, {"로그인_방법": "직접"})
    bot.login_timeout_s = 1
    with pytest.raises(RuntimeError):
        bot.wait_login()
    assert site.typed_email is None and calls == []


def test_stop_button_stops_login_wait(browser):
    site = FakeConcur(saved_email="")
    stop = threading.Event()
    stop.set()
    bot, ctx, calls, lines = make_bot(browser, site, stop_flag=stop)
    with pytest.raises(ca.Cancelled):
        bot.wait_login()


def test_signin_screen_is_not_mistaken_for_logged_in(browser):
    site = FakeConcur()
    bot, ctx, calls, lines = make_bot(browser, site)
    bot.page.goto("https://www.concursolutions.com/nui/signin")
    assert bot.on_signin_screen() and not bot.is_logged_in()
    bot.page.goto("https://www.concursolutions.com/nui/signin/hrd")
    assert bot.on_signin_screen() and not bot.is_logged_in()


# ---------- 실제 사이트처럼 [다음] 버튼 모양이 다를 때 ----------

GO_HRD = "location.href='https://www.concursolutions.com/nui/signin/hrd?email='+encodeURIComponent(document.getElementById('u').value)"

SIGNIN_VARIANTS = {
    # 버튼이 <button>이 아니라 상자(div)로 만들어진 경우
    "div_button": f"""<!doctype html><meta charset="utf-8"><h1>로그인</h1>
<input id="u" type="text" aria-label="사용자 이름, 검증된 이메일 주소 또는 SSO 코드" value="{{value}}">
<div role="button" tabindex="0" onclick="{GO_HRD}"><span class="sapcnqr-button__text"> 다음 </span></div>""",
    # <input type=submit value="다음"> 인 경우
    "submit_input": f"""<!doctype html><meta charset="utf-8"><h1>로그인</h1>
<form onsubmit="event.preventDefault(); {GO_HRD}">
<label>사용자 이름, 검증된 이메일 주소 또는 SSO 코드<input id="u" type="email" value="{{value}}"></label>
<input type="submit" value="다음"></form>""",
    # 버튼 위에 투명한 막이 덮여 클릭이 안 되고 Enter만 되는 경우
    "covered_button": f"""<!doctype html><meta charset="utf-8"><h1>로그인</h1>
<form onsubmit="event.preventDefault(); {GO_HRD}" style="position:relative">
<input id="u" name="username" value="{{value}}">
<button type="submit">다음</button>
<div style="position:absolute; inset:0; z-index:9"></div></form>""",
}


class VariantConcur(FakeConcur):
    def __init__(self, signin_html, **kw):
        super().__init__(**kw)
        self.signin_html = signin_html

    def handle(self, route, request):
        if urlparse(request.url).path == "/nui/signin":
            route.fulfill(status=200, content_type="text/html; charset=utf-8",
                          body=self.signin_html.format(value=self.saved_email))
        else:
            super().handle(route, request)


@pytest.mark.parametrize("variant", list(SIGNIN_VARIANTS))
def test_next_button_variants(browser, variant):
    site = VariantConcur(SIGNIN_VARIANTS[variant], saved_email="saved@example.com")
    bot, ctx, calls, lines = make_bot(browser, site)
    bot.wait_login()
    assert site.typed_email == "saved@example.com"
    assert bot.page.url == "https://us2.concursolutions.com/home"


def test_unknown_signin_screen_is_written_to_log(browser):
    html = """<!doctype html><meta charset="utf-8"><h1>로그인</h1>
<input id="u" type="text" value="saved@example.com"><button>계속</button>"""
    site = VariantConcur(html, saved_email="saved@example.com")
    bot, ctx, calls, lines = make_bot(browser, site)
    bot.login_timeout_s = 2
    bot.diagnose_after_s = 0.2
    with pytest.raises(RuntimeError):
        bot.wait_login()
    assert any("화면의 버튼" in line and "계속" in line for line in lines)


# ---------- 로그인 화면이 영어로 바뀌었을 때 ----------

SIGNIN_EN = """<!doctype html><meta charset="utf-8"><h1>Sign In</h1>
<label for="u">Username, verified email address, or SSO code</label>
<input id="u" type="text" value="{value}">
<button onclick="location.href='https://www.concursolutions.com/nui/signin/hrd?email='+encodeURIComponent(document.getElementById('u').value)">Next</button>
<label><input type="checkbox" checked> Remember me</label><a href="#">Forgot username</a>"""

HRD_EN = """<!doctype html><meta charset="utf-8"><h1>Sign In</h1><p>Sign in with:</p>
<button>Company SSO</button><button>SAP Concur Password</button><p>OR</p>
<button onclick="document.cookie='linkreq=1; path=/'">{link_text}</button>"""


class EnglishConcur(FakeConcur):
    def __init__(self, link_text, **kw):
        super().__init__(**kw)
        self.link_text = link_text

    def handle(self, route, request):
        u = urlparse(request.url)
        if u.path == "/nui/signin":
            body = SIGNIN_EN.format(value=self.saved_email)
        elif u.path == "/nui/signin/hrd":
            self.typed_email = parse_qs(u.query).get("email", [""])[0]
            body = HRD_EN.format(link_text=self.link_text)
        else:
            return super().handle(route, request)
        route.fulfill(status=200, content_type="text/html; charset=utf-8", body=body)


@pytest.mark.parametrize("link_text", ["Sign in with email link", "Email me a sign-in link", "Send Email Link"])
def test_english_signin_screens(browser, link_text):
    site = EnglishConcur(link_text, saved_email="saved@example.com")
    bot, ctx, calls, lines = make_bot(browser, site)
    bot.wait_login()
    assert site.typed_email == "saved@example.com"
    assert link_requested(ctx) and len(calls) == 1
    assert bot.page.url == "https://us2.concursolutions.com/home"


def test_screen_text_alternatives():
    bot = ca.ConcurBot(ca.default_settings(), lambda s="": None)
    assert bot._alts("다음") == ["다음", "Next", "Continue"]
    assert bot._re("편집").match(" Edit ") and bot._re("편집").match("편집") and not bot._re("편집").match("편집하기")
    assert bot._re("리포트 생성", exact=False).search("경비 리포트 생성")


def test_old_saved_screen_texts_get_english_too(tmp_path):
    import json
    p = tmp_path / ca.SETTINGS_FILE_NAME
    p.write_text(json.dumps({"화면_문구": {"다음": "다음", "편집": "수정"}}, ensure_ascii=False), encoding="utf-8")
    cfg, _ = ca.load_settings(p)
    assert cfg["화면_문구"]["다음"] == "다음|Next|Continue"     # 예전 기본값 → 새 기본값
    assert cfg["화면_문구"]["편집"] == "수정"                   # 사용자가 직접 바꾼 값은 그대로


def test_edge_opens_in_korean_without_translate_popup():
    opts = ca.ConcurBot(ca.default_settings(), lambda s="": None).launch_options()
    assert opts["locale"] == "ko-KR" and opts["channel"] == "msedge"
    assert "--lang=ko-KR" in opts["args"] and "--disable-features=Translate" in opts["args"]


# ---------- 로그인 메일이 영어로 올 때 ----------

MAIL_EN = """<html><body><div><h2>Sign-In Request to SAP Concur</h2></div><p>Hello Hong Gildong,</p>
<p>You requested a one-time link to sign in to SAP Concur.</p>
<a href="https://www.concur.com/privacy">Privacy Policy</a>
<table><tr><td><a href="{href}"><span>Sign in to SAP Concur</span></a></td></tr></table>
<p>This link will be active for an hour and can only be used once.</p></body></html>"""


def test_extract_login_link_english_mail():
    assert ca.extract_login_link(MAIL_EN.format(href="https://c.example/en"), ca.default_settings()["로그인_메일_버튼"]) == "https://c.example/en"
    # 버튼 글자가 조금 달라도 'Concur'와 'Sign in'이 들어 있으면
    other = MAIL_EN.replace("Sign in to SAP Concur", "Log in to Concur").format(href="https://c.example/x")
    assert ca.extract_login_link(other, "SAP Concur 로그인|Sign in to SAP Concur") == "https://c.example/x"


def test_outlook_finds_english_login_mail(fake_outlook):
    now = datetime.now().replace(microsecond=0)
    fake_outlook["inbox"] = FakeInbox([
        FakeMail("Weekly meeting", 0, "https://c.example/meeting", now, sender="boss@example.com"),
        FakeMail("Sign-In Request to SAP Concur", 0, "https://c.example/en", now, body=MAIL_EN),
    ])
    cfg = dict(ca.default_settings(), 로그인_메일_대기_초=0)
    assert ca.find_login_link_in_outlook(cfg, now, lambda s: None) == "https://c.example/en"


def test_outlook_finds_login_mail_from_concur_even_if_subject_changes(fake_outlook):
    now = datetime.now().replace(microsecond=0)
    fake_outlook["inbox"] = FakeInbox([FakeMail("Anmeldeanfrage", 0, "https://c.example/de", now, body=MAIL_EN)])
    cfg = dict(ca.default_settings(), 로그인_메일_대기_초=0)
    assert ca.find_login_link_in_outlook(cfg, now, lambda s: None) == "https://c.example/de"


def test_old_saved_mail_settings_get_english_too(tmp_path):
    import json
    p = tmp_path / ca.SETTINGS_FILE_NAME
    p.write_text(json.dumps({"로그인_메일_제목": "SAP Concur 로그인 요청", "로그인_메일_버튼": "SAP Concur 로그인"},
                            ensure_ascii=False), encoding="utf-8")
    cfg, _ = ca.load_settings(p)
    assert cfg["로그인_메일_제목"] == "SAP Concur 로그인 요청|Sign-In Request to SAP Concur"
    assert cfg["로그인_메일_버튼"] == "SAP Concur 로그인|Sign in to SAP Concur"
