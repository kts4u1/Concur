# -*- coding: utf-8 -*-
"""가게 기억: 영수증의 사업자등록번호로 같은 가게를 알아보고, 사용자가 한 번 고른 분류·병원·도시를 다음부터 자동으로.
카드 가맹점 이름으로 분류하기."""

import os
from datetime import datetime
from pathlib import Path

import pytest

import concur_auto as ca

NOV = ca.TargetMonth(2026, 11)


def make_biz(prefix9: str) -> str:
    """끝자리 검사가 맞는 사업자등록번호 만들기(테스트용)"""
    d = [int(x) for x in prefix9]
    total = sum(a * b for a, b in zip(d, (1, 3, 7, 1, 3, 7, 1, 3, 5))) + (d[8] * 5) // 10
    digits = prefix9 + str((10 - total % 10) % 10)
    return f"{digits[:3]}-{digits[3:5]}-{digits[5:]}"


BIZ = make_biz("123456789")


def test_find_business_number():
    assert ca.valid_business_number(BIZ.replace("-", ""))
    assert ca.find_business_number(f"상호 다이소 사업자번호 : {BIZ} 대표 홍길동") == BIZ
    assert ca.find_business_number("사업자등록번호1234567890") == "123-45-67890"           # '사업자' 옆이면 그대로
    assert ca.find_business_number(f"주소 ... {BIZ} TEL") == BIZ                          # 끝자리 검사가 맞으면
    bad = BIZ[:-1] + str((int(BIZ[-1]) + 1) % 10)
    assert ca.find_business_number(f"승인번호 {bad}") == ""                               # 끝자리가 틀린 숫자는 아님
    assert ca.find_business_number("TEL 062-123-4567 카드 1234-56**-**** 2026-11-05") == ""


TEXTS = {"IMG_0001.JPG": f"다이소 목포점 사업자번호 {BIZ} 2026.11.03 합계 9,000",
         "IMG_0002.JPG": f"다이소 목포점 사업자번호 {BIZ} 2026.11.20 합계 5,000",
         "IMG_0003.JPG": "다른가게 2026.11.21 합계 3,000"}


class FakeOcr:
    def __init__(self, work_dir):
        pass

    def read(self, path):
        return TEXTS.get(Path(path).name, "")


@pytest.fixture
def setup(tmp_path, monkeypatch):
    base = tmp_path / "app"
    src = tmp_path / "src"
    src.mkdir(parents=True)
    for name in TEXTS:
        (src / name).write_bytes(name.encode())
        t = datetime(2026, 11, 22, 9, 0).timestamp()
        os.utime(src / name, (t, t))
    results = base / "결과"
    results.mkdir(parents=True)
    monkeypatch.setattr(ca, "WindowsOcr", FakeOcr)
    cfg = ca.default_settings()
    cfg.update({"영수증_원본_폴더": str(src), "영수증_보정_사용": False})

    def build():
        return {r.original.name: r for r in ca.build_receipts(cfg, NOV, [], results, tmp_path / "work", lambda s="": None)}

    return base, build, ca.list_csv_path(results, NOV)


def test_store_remembered_after_one_fix(setup):
    base, build, path = setup
    first = build()
    assert first["IMG_0001.JPG"].biz_no == BIZ and first["IMG_0002.JPG"].category == ""
    rows = ca.read_csv_dicts(path)
    by = {r["파일"]: r for r in rows}
    assert by["IMG_0001.JPG"]["사업자번호"] == BIZ
    by["IMG_0001.JPG"]["분류"] = "식대"                                            # 한 장만 고침(영수증 확인 창처럼)
    by["IMG_0001.JPG"]["병원 영문"] = "Mokpo Hankook Hospital"
    ca.save_review_rows(path, rows)
    assert ca.remember_stores(base / ca.STORE_MEMORY_FILE_NAME, rows) == 1
    memory = ca.load_store_memory(base / ca.STORE_MEMORY_FILE_NAME)
    assert memory[BIZ]["분류"] == "식대" and memory[BIZ]["병원 영문"] == "Mokpo Hankook Hospital"
    assert "도시" not in memory[BIZ]                                                # 고친 칸만 기억

    again = build()
    other = again["IMG_0002.JPG"]                                                   # 고치지 않은 같은 가게 영수증
    assert other.category == "식대" and other.hospital_en == "Mokpo Hankook Hospital"
    assert other.hospital_src == ca.HOSP_FROM_MEMORY and other.remembered
    assert again["IMG_0003.JPG"].category == ""                                     # 다른 가게는 그대로
    by = {r["파일"]: r for r in ca.read_csv_dicts(path)}
    assert by["IMG_0002.JPG"]["가게 기억"] == "O" and by["IMG_0002.JPG"]["[자동] 분류"] == "식대"


def test_build_learns_from_list_edits_made_in_excel(setup):
    base, build, path = setup
    build()
    rows = ca.read_csv_dicts(path)
    {r["파일"]: r for r in rows}["IMG_0001.JPG"]["분류"] = "식대"                   # 엑셀에서 고친 것처럼(기억은 안 함)
    ca.save_review_rows(path, rows)
    again = build()                                                                 # 다음 실행이 먼저 배움
    assert again["IMG_0002.JPG"].category == "식대"


def test_excluded_receipt_is_not_remembered(tmp_path):
    rows = [{"사업자번호": BIZ, "사용(O/X)": "X", "분류": "식대", "[자동] 분류": ""}]
    assert ca.remember_stores(tmp_path / "m.json", rows) == 0


@pytest.mark.parametrize("vendor, category", [
    ("STARBUCKS KOREA CO", "카페"), ("MEGA COFFEE GWANGJU", "카페"), ("KAKAO T PARKING", "주차"),
    ("LOTTE HOTEL MOKPO", "숙박"), ("(주)다이소", ""), ("", "")])
def test_category_from_card_vendor(vendor, category):
    assert ca.category_from_vendor(vendor, ca.default_settings()) == category


def test_store_memory_is_backed_up(tmp_path):
    base = tmp_path / "app"
    base.mkdir()
    (base / ca.STORE_MEMORY_FILE_NAME).write_text('{"x": {"분류": "식대"}}', encoding="utf-8")
    info = ca.make_backup(base, tmp_path / "b.zip")
    assert ca.STORE_MEMORY_FILE_NAME in info["파일"]
    new = tmp_path / "new"
    new.mkdir()
    assert ca.STORE_MEMORY_FILE_NAME in ca.restore_backup(tmp_path / "b.zip", new, tmp_path / "bk")
