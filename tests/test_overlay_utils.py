import json
from pathlib import Path

import pytest

fitz = pytest.importorskip("fitz")

from pdf_text_overlay.overlay import (
    OverlaySettings,
    _apply_rotation,
    _convert_bbox_to_point,
    _normalize_text,
    apply_text_overlay,
)
from pdf_text_overlay.ocr_io import OCRPage, OCRWord, _extract_rect, load_ocr_json


def test_convert_bbox_to_point():
    x_pt, y_pt, font_size = _convert_bbox_to_point(
        bbox=(100, 200, 400, 50),
        width_px=2000,
        height_px=3000,
        width_pt=595.0,
        height_pt=842.0,
        baseline_ratio=0.2,
    )
    assert pytest.approx(x_pt, rel=1e-3) == 29.75
    expected_y = 842.0 - (200 + 50 * 0.2) * (842.0 / 3000)
    assert pytest.approx(y_pt, rel=1e-3) == expected_y
    assert pytest.approx(font_size, rel=1e-3) == 50 * (842.0 / 3000)


@pytest.mark.parametrize(
    "rotation,expected",
    [
        (0, (100.0, 200.0, 0)),
        (90, (642.0, 100.0, 90)),
        (180, (495.0, 642.0, 180)),
        (270, (200.0, 495.0, 270)),
    ],
)
def test_apply_rotation(rotation, expected):
    assert _apply_rotation(100.0, 200.0, 595.0, 842.0, rotation) == expected


def test_normalize_text_collapses_spaces():
    assert _normalize_text(" A\u200bB \n C ", keep_spaces=False) == "A B C"


def test_load_ocr_json(tmp_path: Path):
    payload = [
        {
            "page": 0,
            "width_px": 100,
            "height_px": 200,
            "lines": [{"text": "hello", "bbox": [0, 0, 20, 10]}],
            "words": [{"text": "hello", "bbox": [0, 0, 20, 10]}],
        }
    ]
    file_path = tmp_path / "ocr.json"
    file_path.write_text(json.dumps(payload), encoding="utf-8")
    pages = load_ocr_json(file_path)
    assert isinstance(pages[0], OCRPage)
    assert pages[0].lines[0].text == "hello"


@pytest.mark.parametrize("method", ["invisible", "opacity"])
def test_apply_text_overlay_adds_searchable_text(tmp_path: Path, method: str):
    base_pdf = tmp_path / "base.pdf"
    doc = fitz.open()
    doc.new_page(width=300, height=300)
    doc.save(base_pdf)
    doc.close()

    ocr_page = OCRPage(
        index=0,
        width_px=100,
        height_px=100,
        rotation=0,
        words=[OCRWord(text="Test", bbox=(10, 10, 40, 20))],
        lines=[],
    )

    out_pdf = tmp_path / f"searchable_{method}.pdf"
    settings = OverlaySettings(method=method, granularity="word")
    apply_text_overlay(base_pdf, [ocr_page], out_pdf, settings)

    result = fitz.open(out_pdf)
    try:
        extracted = result[0].get_text()
    finally:
        result.close()
    assert "Test" in extracted


class _FakeRect:
    def __init__(self, x: float, y: float, w: float, h: float) -> None:
        self.x = x
        self.y = y
        self.width = w
        self.height = h


class _FakeLine:
    def __init__(self) -> None:
        self.boundingRect = _FakeRect(1, 2, 3, 4)


def test_extract_rect_handles_camel_case():
    rect = _extract_rect(_FakeLine())
    assert rect is not None
    assert rect.x == 1
