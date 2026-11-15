# -*- coding: utf-8 -*-
"""작은 글자 크게 읽기: Windows OCR은 큰 사진을 줄여서 읽는데, 영수증 여러 장을 한 장에 찍은 사진은 글자가 너무 작아져
날짜·금액이 틀리게 읽혔습니다(2026-11-15 사용자 화면: 하이패스 영수증 9장 인쇄물). 날짜·금액을 못 읽으면 사진을 조각내
키워서 다시 읽습니다. 연도 없이 '10월01일'로 찍힌 날짜도 읽습니다."""

import json
import os
from datetime import datetime
from pathlib import Path

import pytest

import concur_auto as ca


def test_tile_boxes_cover_whole_photo():
    boxes = ca.tile_boxes(3024, 4032, 1732)
    assert len(boxes) == 6
    assert all(x1 - x0 <= 1732 and y1 - y0 <= 1732 for x0, y0, x1, y1 in boxes)
    assert min(b[0] for b in boxes) == 0 and max(b[2] for b in boxes) == 3024
    assert min(b[1] for b in boxes) == 0 and max(b[3] for b in boxes) == 4032
    assert boxes[1][0] < boxes[0][2] and boxes[2][1] < boxes[0][3]              # 조각끼리 겹침(줄이 잘리지 않게)
    assert ca.tile_boxes(1000, 800, 1732) == [(0, 0, 1000, 800)]               # 작은 사진은 한 조각


def test_read_zoomed_reads_each_piece_and_cleans_up(tmp_path):
    cv2 = pytest.importorskip("cv2")
    np = pytest.importorskip("numpy")
    path = tmp_path / "IMG_5651.jpg"
    cv2.imwrite(str(path), np.full((4032, 3024, 3), 255, dtype=np.uint8))
    ocr = object.__new__(ca.WindowsOcr)
    ocr.max_dim, ocr.temp_dir = 2600, tmp_path / "임시"
    seen = []

    async def fake_read(p):
        h, w = cv2.imread(str(p)).shape[:2]
        seen.append((w, h))
        return f"조각 {len(seen)}"

    ocr._read = fake_read
    text = ocr.read_zoomed(path)
    assert text.splitlines() == [f"조각 {i}" for i in range(1, 7)]
    assert all(max(w, h) <= 2600 for w, h in seen) and seen[0] == (2598, 2598)  # 1.5배로 키워도 OCR 한도 안
    assert list((tmp_path / "임시").iterdir()) == []                           # 임시 조각은 지움


def test_dates_without_year():
    nov15 = datetime(2026, 11, 15)
    assert ca.find_ocr_dates("하이패스 영수증 교10월01일16시17분", nov15) == ["2026-10-01"]
    assert ca.find_ocr_dates("12월 30일", datetime(2027, 1, 5)) == ["2026-12-30"]       # 촬영보다 뒤 달이면 전해
    assert ca.find_ocr_dates("2026-10-05 결제 10월01일", nov15) == ["2026-10-05"]      # 연도까지 있는 날짜가 먼저
    assert ca.find_ocr_dates("10월01일") == []                                        # 기준 날짜가 없으면 안 씀
    assert ca.find_receipt_time("교10월01일16시17분") == (16, 17)


SHEET = "하이패스빠르고편리합니다영수증한국도로공사 사업자번호 교10원01일 1617 합계 118,100"
ZOOMED = ("하이패스 빠르고 편리합니다 영수증 한국도로공사 북김제영업소\n2026년10월01일 16시17분\n통행료 4,700원(카드)\n"
          "하이패스 빠르고 편리합니다 영수증 광주영업소\n2026년10월02일 08시05분\n통행료 4,300원(카드)")


class FakeOcr:
    zoom_calls = 0

    def __init__(self, work_dir):
        pass

    def read(self, path):
        return SHEET if "5651" in Path(path).name else "스타벅스 2026.10.07 합계 4,500"

    def read_zoomed(self, path):
        FakeOcr.zoom_calls += 1
        return ZOOMED


def test_small_text_read_again_bigger(tmp_path, monkeypatch):
    src = tmp_path / "src"
    src.mkdir()
    for name in ("IMG_5651.JPG", "IMG_5652.JPG"):
        p = src / name
        p.write_bytes(b"\xff\xd8\xff\xd9" + name.encode())
        t = datetime(2026, 11, 15, 10, 0).timestamp()
        os.utime(p, (t, t))
    monkeypatch.setattr(ca, "WindowsOcr", FakeOcr)
    FakeOcr.zoom_calls = 0
    cfg = ca.default_settings()
    cfg.update({"영수증_원본_폴더": str(src), "영수증_보정_사용": False})        # 보정을 꺼도 크게 읽기는 함
    results = tmp_path / "결과"
    results.mkdir()
    lines = []
    recs = {r.original.name: r for r in ca.build_receipts(cfg, ca.TargetMonth(2026, 10), [], results, tmp_path / "w", lines.append)}
    sheet = recs["IMG_5651.JPG"]
    assert sheet.use and sheet.main_date == "2026-10-01" and sheet.total == 4700   # 2026-11-15·118,100이 아님
    assert recs["IMG_5652.JPG"].main_date == "2026-10-07"
    assert FakeOcr.zoom_calls == 1                                              # 잘 읽힌 사진은 다시 읽지 않음
    cache = json.loads((results / "ocr_cache.json").read_text(encoding="utf-8"))
    assert cache["IMG_5651.JPG#크게"]["text"] == ZOOMED
    assert any("크게 다시 읽은 사진 1장" in l for l in lines)
    ca.build_receipts(cfg, ca.TargetMonth(2026, 10), [], results, tmp_path / "w", lines.append)
    assert FakeOcr.zoom_calls == 1                                              # 두 번째부터는 기억한 글자를 씀
