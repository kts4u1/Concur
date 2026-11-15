# -*- coding: utf-8 -*-
"""톨게이트(2026-11-15 사용자 설명): 하이패스 '기간별 사용내역'의 총액이 경비 금액이고, 하이패스 영수증 인쇄물
(한 장에 여러 장)은 사용내역 순서대로의 영수증이라 같은 경비에 함께 올리기만 합니다. 영수증 합계도 비교합니다."""

import os
import re
from datetime import datetime
from pathlib import Path

import pytest

import concur_auto as ca

OCT = ca.TargetMonth(2026, 10)
STATEMENT = """기간별 사용내역
카드구분 전체 카드사 전체 카드번호 1234-56**-****-0000
조회 구분 사용일 기준 조회 기간 2026년 10월 01일 ~ 2026년 10월 31일
(총 13건 / 총액 61,500원)
번호 거래일시 구분 카드번호 별명 차종 입구 출구 이용차로 거래금액 청구일자
13 2026/10/01 16:17:55 선불 1234-56**-****-0000 1종 파하 가나 선불Hi-pass 4,700 -
12 2026/10/01 20:43:23 선불 1234-56**-****-0000 1종 가나 파하 선불Hi-pass 4,700 -
1 2026/10/22 19:06:26 선불 1234-56**-****-0000 1종 자차 마바 선불Hi-pass 4,300 -"""


def receipt(office, when, amount):
    return (f"하이패스는 빠르고 편리합니다 영수증 한국도로공사 {office}영업소 사업자번호 : 129-82-00103 "
            f"{when} 입구영업소 : 파하 1 종 {amount:,} 원(카드) 공급가액 : {amount:,} 원 SM 하이플러스카드\n")


SHEET1 = "".join(receipt(o, w, a) for o, w, a in [
    ("가나", "2026년10월01일 16시17분", 4700), ("파하", "2026년10월01일 20시43분", 4700),
    ("다라", "2026년10월07일 09시47분", 4300), ("파하", "2026년10월07일 16시03분", 4300),
    ("가나", "2026년10월08일 11시18분", 4700), ("파하", "2026년10월08일 14시44분", 4700),
    ("자차", "2026년10월13일 11시11분", 4300), ("마바", "2026년10월13일 15시24분", 4300),
    ("사아", "2026년10월20일 11시40분", 1200)])
SHEET2 = "".join(receipt(o, w, a) for o, w, a in [
    ("카타", "2026년10월20일 15시46분", 8000), ("파하", "2026년10월21일 11시10분", 7700),
    ("자차", "2026년10월22일 10시52분", 4300), ("마바", "2026년10월22일 19시06분", 4300)]) + "총 13 건/61500 원"


def test_recognize_statement_and_sheets():
    assert ca.toll_statement_info(STATEMENT) == {"total": 61500, "count": 13}
    one = ca.hipass_sheet_info(SHEET1)
    assert one["amounts"] == [4700, 4700, 4300, 4300, 4700, 4700, 4300, 4300, 1200] and one["footer"] is None
    assert ca.hipass_sheet_info(SHEET2)["footer"] == (13, 61500)
    assert ca.toll_statement_info(SHEET2) is None and ca.hipass_sheet_info(STATEMENT) is None
    assert ca.toll_statement_info("스타벅스 2026.10.07 합계 4,500") is None


def make(name, text, taken):
    return ca.analyze_receipt(Path(name), Path(name), text, ca.default_settings(), [], taken)


def recs():
    t = datetime(2026, 11, 15, 10, 0)
    return [make("IMG_5652.HEIC", SHEET2, t), make("IMG_5648.HEIC", STATEMENT, t), make("IMG_5651.HEIC", SHEET1, t),
            make("IMG_0001.JPG", "스타벅스 2026.10.07 합계 4,500", t)]


def test_bundle_statement_with_sheets():
    rs = recs()
    for r in rs:
        ca.place_in_month(r, OCT)
    keywords = {"톨게이트": ca.TOLL_DEFAULT_WORDS, "카페": ["스타벅스"]}
    assert ca.group_toll_bundles(rs, OCT, keywords, ["톨게이트", "카페"]) == 1
    sheet2, stmt, sheet1, cafe = rs
    assert stmt.use and stmt.total == 61500 and stmt.category == "톨게이트"
    assert [p.name for p in stmt.bundle] == ["IMG_5651.HEIC", "IMG_5652.HEIC"]          # 사용내역 순서대로
    assert stmt.period == ("2026-10-01", "2026-10-31")
    assert stmt.last_use == "2026-10-22" and stmt.main_date == "2026-10-22"            # 마지막 이용일(조회 기간 끝 10-31이 아님)
    assert stmt.bundle_note == "사용내역 13건 61,500원, 마지막 이용 2026-10-22, 영수증 인쇄물 2장 함께 첨부"
    assert not sheet1.use and not sheet2.use and "IMG_5648.HEIC" in sheet1.skip_reason
    assert sheet1.total == 0 and sheet2.total == 0                                   # 인쇄물 금액은 틀리게 읽혀 비움
    assert cafe.use and cafe.bundle == [] and cafe.total == 4500                           # 다른 영수증은 그대로
    assert ca.receipt_issues(stmt, rs) == []
    # Concur 경비: 사용내역 기간 안의 날짜 + 총액이면 맞춤
    assert ca.find_receipt(rs, "2026-10-31", 61500) is stmt
    assert ca.find_receipt(rs, "2026-11-02", 61500) is None


def test_sheet_amounts_not_compared():
    """인쇄물 금액은 작은 글자라 틀리게 읽혀(118,100·70,300) 비교하지 않고 첨부만(사용자 요청 2026-11-15)"""
    t = datetime(2026, 11, 15, 10, 0)
    wrong = make("B.jpg", SHEET2.replace("총 13 건/61500 원", "총 13 건/60500 원").replace("8,000 원(카드)", "118,100 원(카드)"), t)
    stmt = make("S.jpg", STATEMENT, t)
    for r in (stmt, wrong):
        ca.place_in_month(r, OCT)
    ca.group_toll_bundles([stmt, wrong], OCT, {"톨게이트": ca.TOLL_DEFAULT_WORDS}, ["톨게이트"])
    assert stmt.total == 61500 and stmt.bundle_issue == "" and ca.receipt_issues(stmt, [stmt, wrong]) == []
    assert [p.name for p in stmt.bundle] == ["B.jpg"] and wrong.total == 0


def test_sheets_alone_are_left_as_before():
    rs = [make("A.jpg", SHEET1, datetime(2026, 11, 15)), make("B.jpg", SHEET2, datetime(2026, 11, 15))]
    for r in rs:
        ca.place_in_month(r, OCT)
    assert ca.group_toll_bundles(rs, OCT) == 0 and all(r.use and not r.bundled_into for r in rs)


def test_upload_files_in_order(tmp_path):
    pytest.importorskip("cv2")
    from PIL import Image
    files = []
    for i, size in enumerate([(300, 400), (300, 400), (400, 300)]):
        f = tmp_path / f"IMG_{i}.jpg"
        Image.new("RGB", size, (240, 240, 240)).save(f, quality=90)
        files.append(f)
    rec = ca.Receipt(original=files[0], ocr_file=files[0], bundle=files[1:])
    assert ca.upload_files_for(rec, ca.default_settings(), tmp_path / "w") == files      # PDF로 묶지 않고 차례대로


TEXTS = {"IMG_5648.JPG": STATEMENT, "IMG_5651.JPG": SHEET1, "IMG_5652.JPG": SHEET2}


class FakeOcr:
    def __init__(self, work_dir):
        pass

    def read(self, path):
        return TEXTS[Path(path).name]

    def read_zoomed(self, path):
        return ""


def test_receipt_list_shows_bundle(tmp_path, monkeypatch):
    src = tmp_path / "src"
    src.mkdir()
    for name in TEXTS:
        p = src / name
        p.write_bytes(b"\xff\xd8\xff\xd9" + name.encode())
        t = datetime(2026, 11, 15, 10, 0).timestamp()
        os.utime(p, (t, t))
    monkeypatch.setattr(ca, "WindowsOcr", FakeOcr)
    cfg = ca.default_settings()
    table = tmp_path / ca.CATEGORY_FILE_NAME
    ca.save_category_rows(table, ca.category_rows_from_cfg(cfg) + [
        ["톨게이트", ", ".join(ca.TOLL_DEFAULT_WORDS), "Toll fee for service", "For service", "Parking & Toll", ""]])
    ca.apply_category_table(cfg, table)
    cfg.update({"영수증_원본_폴더": str(src), "영수증_보정_사용": False})
    results = tmp_path / "결과"
    results.mkdir()
    lines = []
    got = {r.original.name: r for r in ca.build_receipts(cfg, OCT, [], results, tmp_path / "w", lines.append)}
    assert got["IMG_5648.JPG"].use and got["IMG_5648.JPG"].category == "톨게이트"
    assert not got["IMG_5651.JPG"].use and not got["IMG_5652.JPG"].use
    rows = {r["파일"]: r for r in ca.read_csv_dicts(ca.list_csv_path(results, OCT))}
    assert rows["IMG_5648.JPG"]["함께 첨부"].startswith("IMG_5651.JPG, IMG_5652.JPG | 사용내역 13건 61,500원")
    assert rows["IMG_5648.JPG"]["금액"] == "61500" and ca.review_status(rows["IMG_5648.JPG"]) == "✓"
    assert rows["IMG_5648.JPG"]["날짜"] == "2026-10-22"
    assert "함께 첨부" in rows["IMG_5651.JPG"]["확인 필요"]
    assert any("톨게이트: IMG_5648.JPG 사용내역 13건" in l for l in lines)


def test_toll_item_gets_default_words(tmp_path):
    settings, table = tmp_path / ca.SETTINGS_FILE_NAME, tmp_path / ca.CATEGORY_FILE_NAME
    cfg = ca.default_settings()
    cfg["_기본_찾을_글자_판"] = 4
    ca.save_settings(settings, cfg)
    ca.save_category_rows(table, [["톨게이트", "톨게이트, 하이패스", "Toll fee", "For service", "Parking & Toll", ""]])
    ca.upgrade_category_table(table, settings)
    words = ca.load_category_rows(table, ca.default_settings())[0][ca.C_KEYWORDS]
    assert words.startswith("톨게이트, 하이패스, ") and "사용내역" in words and words.count("하이패스") == 1
