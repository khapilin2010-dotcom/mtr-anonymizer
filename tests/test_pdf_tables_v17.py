import os
import hashlib
import re
from pathlib import Path

import fitz
import pytest

from MTR_Obezlichivatel import process_pdf
from mtr_core import Anonymizer


FONT = ("C:/Windows/Fonts/arial.ttf" if os.name == "nt"
        else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")


def _make_table_pdf(path: Path, scan: bool = False):
    doc = fitz.open()
    page = doc.new_page(width=1200, height=420)
    xs = [20, 60, 330, 560, 730, 910, 970, 1030, 1100, 1180]
    ys = [20, 80, 220, 360]
    for x in xs:
        page.draw_line((x, ys[0]), (x, ys[-1]), width=0.7)
    for y in ys:
        page.draw_line((xs[0], y), (xs[-1], y), width=0.7)
    headers = ["Поз.", "Наименование и техническая характеристика",
               "Тип, марка, обозначение документа, опросного листа",
               "Код продукции", "Поставщик", "Ед.", "Количество",
               "Масса", "Примечание"]
    rows = [
        ["1", "Клапан Армтел IP66 УХЛ1 Ex d IIC T6 DN100 PN1,6 МПа сталь 09Г2С 100х50 мм",
         "Комплектация по обосновывающему документу TEST.0001-АТТ.ОЛ1\nМодель TEST-M1 ГОСТ 12345",
         "1234567 (0)1)\n4143086 (0)1)\n2576244 (0)1)",
         'АО "ТЕСТОВЫЙ ЗАВОД", ИНН 1234567890', "шт.", "1", "10", ""],
        ["2", "Труба 57х3,5 сталь 20 давление 1,6 МПа температура -60...+100С",
         "ТУ 1234-567-890 TEST-M2", "Заявка № Z1234567 (1342)",
         'ООО "ДРУГОЙ ТЕСТОВЫЙ ЗАВОД", ИНН 123456789012', "м", "2", "20", ""],
    ]
    for col, value in enumerate(headers):
        page.insert_textbox((xs[col] + 2, 24, xs[col + 1] - 2, 76), value,
                            fontsize=6, fontname="dejavu", fontfile=FONT)
    for row_no, values in enumerate(rows):
        for col, value in enumerate(values):
            page.insert_textbox((xs[col] + 2, ys[row_no + 1] + 4,
                                 xs[col + 1] - 2, ys[row_no + 2] - 4), value,
                                fontsize=7 if col == 2 else 8,
                                fontname="dejavu", fontfile=FONT)
    keep_tokens = ("IP66", "УХЛ1", "Ex d IIC T6", "DN100", "PN1,6 МПа",
                   "09Г2С", "ГОСТ 12345", "100х50", "TEST.0001-АТТ.ОЛ1",
                   "TEST-M1", "ТУ 1234-567-890", "TEST-M2")
    keep_boxes = {token: [tuple(rect) for rect in page.search_for(token)] for token in keep_tokens}
    assert all(keep_boxes.values()), f"Synthetic source overflowed KEEP text: {keep_boxes}"
    source_text = " ".join(page.get_text().replace("\u00ad", "").split())
    assert "Комплектация по обосновывающему документу TEST.0001-АТТ.ОЛ1" in source_text
    redact_boxes = {
        "brand": [tuple(rect) for rect in page.search_for("Армтел")],
        "supplier_cell": [(xs[4], ys[1], xs[5], ys[2])],
    }
    if scan:
        pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False)
        image = pix.tobytes("png")
        scanned = fitz.open(); target = scanned.new_page(width=1200, height=420)
        target.insert_image(target.rect, stream=image)
        scanned.save(path); scanned.close(); doc.close()
    else:
        doc.save(path); doc.close()
    return xs, ys, keep_boxes, redact_boxes


def _assert_table_result(src, dst, report, xs, ys, keep_boxes, redact_boxes, ocr=False):
    assert report["table_pages"] == 1
    assert report["supplier_cells_redacted"] >= 2
    assert report["code_column_unchanged"] is True
    assert report["code_column_intersections"] == 0
    for redaction in map(fitz.Rect, report["redaction_rects"]):
        for keep in map(fitz.Rect, report["code_keep_rects"]):
            assert (redaction & keep).get_area() == 0
    source = fitz.open(src); source_page = source[0]
    result = fitz.open(dst); page = result[0]
    textpage = page.get_textpage_ocr(language="rus+eng", dpi=300, full=True) if ocr else None
    words = page.get_text("words", textpage=textpage, sort=True) if ocr else None
    source_textpage = source_page.get_textpage_ocr(language="rus+eng", dpi=300, full=True) if ocr else None
    source_words = source_page.get_text("words", textpage=source_textpage, sort=True) if ocr else None

    def region_text(rect):
        if not ocr:
            return page.get_textbox(rect)
        return " ".join(word[4] for word in words if fitz.Rect(*word[:4]).intersects(rect))

    def source_region_text(rect):
        if not ocr:
            return source_page.get_textbox(rect)
        return " ".join(word[4] for word in source_words if fitz.Rect(*word[:4]).intersects(rect))

    def normalized(value):
        return " ".join(value.replace("\u00ad", "").replace("\r", "\n").split())

    code_rect = fitz.Rect(xs[3], ys[1], xs[4], ys[-1])
    code_text = normalized(region_text(code_rect))
    source_code_text = normalized(source_region_text(code_rect))
    if not ocr:
        # Native text layers must remain byte-for-byte equivalent after basic
        # extraction whitespace normalization.
        assert code_text == source_code_text
        assert "Заявка № Z1234567 (1342)" in code_text
    else:
        # OCR may read an unchanged glyph Z as 7. Pixel identity below is the
        # authoritative KEEP check; OCR remains supporting presence evidence.
        for digits in ("1234567", "4143086", "2576244"):
            assert digits in source_code_text and digits in code_text
        before_crop = source_page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=code_rect, alpha=False)
        after_crop = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=code_rect, alpha=False)
        assert (before_crop.width, before_crop.height, before_crop.n) == (after_crop.width, after_crop.height, after_crop.n)
        assert hashlib.sha256(before_crop.samples).digest() == hashlib.sha256(after_crop.samples).digest()
    supplier_text = region_text(fitz.Rect(xs[4], ys[1], xs[5], ys[-1]))
    assert not supplier_text.strip()
    all_text = page.get_text("text", textpage=textpage)
    source_text = source_page.get_text("text", textpage=source_textpage)
    normalized_text = normalized(all_text)
    normalized_source_text = normalized(source_text)
    exact_native_keep = ("IP66", "УХЛ1", "Ex d IIC T6", "DN100", "PN1,6 МПа",
                         "09Г2С", "ГОСТ 12345", "100х50",
                         "Комплектация по обосновывающему документу", "TEST.0001-АТТ.ОЛ1")
    if not ocr:
        for keep in exact_native_keep:
            assert keep in normalized_source_text, f"Fixture source lacks {keep!r}"
            assert keep in normalized_text, f"Anonymized native PDF lost {keep!r}"
    else:
        # OCR spelling is non-authoritative. The visual glyph areas recorded
        # before rasterization must remain pixel-identical and untouched by
        # every physical redaction rectangle.
        for token, boxes in keep_boxes.items():
            for box in map(fitz.Rect, boxes):
                assert all((box & fitz.Rect(redaction)).get_area() == 0
                           for redaction in report["redaction_rects"]), token
                before = source_page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=box, alpha=False)
                after = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=box, alpha=False)
                assert hashlib.sha256(before.samples).digest() == hashlib.sha256(after.samples).digest(), token
    # PDF extractors may represent the visual hyphen as U+00AD. Compare the
    # model semantically after removing only separators and whitespace, first
    # proving that it existed in the source and then that it survived.
    source_model_key = re.sub(r"[\s\-\u00ad]", "", normalized_source_text)
    result_model_key = re.sub(r"[\s\-\u00ad]", "", normalized_text)
    assert "TESTM1" in source_model_key
    if not ocr:
        assert "TESTM1" in result_model_key
        assert "ТУ 1234-567-890 TEST-M2" in normalized_text
    assert "ТЕСТОВЫЙ ЗАВОД" not in all_text and "1234567890" not in all_text
    if ocr:
        # Non-KEEP sensitive glyphs must physically change, proving the test
        # does not pass merely because Tesseract failed to recognize them.
        for boxes in redact_boxes.values():
            for box in map(fitz.Rect, boxes):
                before = source_page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=box, alpha=False)
                after = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=box, alpha=False)
                assert hashlib.sha256(before.samples).digest() != hashlib.sha256(after.samples).digest()
    result.close(); source.close()


def test_native_table_columns_are_respected(tmp_path):
    src, dst = tmp_path / "table.pdf", tmp_path / "table_anon.pdf"
    xs, ys, keep_boxes, redact_boxes = _make_table_pdf(src)
    report = process_pdf(src, dst, Anonymizer())
    _assert_table_result(src, dst, report, xs, ys, keep_boxes, redact_boxes)
    repeat = tmp_path / "table_repeat.pdf"
    second = process_pdf(dst, repeat, Anonymizer())
    assert second.get("unchanged") is True
    assert "".join(p.get_text() for p in fitz.open(dst)) == "".join(p.get_text() for p in fitz.open(repeat))


def test_ocr_table_columns_are_respected(tmp_path):
    if not os.environ.get("MTR_REQUIRE_OCR"):
        pytest.skip("Real RU+EN OCR is mandatory in Windows CI")
    src, dst = tmp_path / "scan.pdf", tmp_path / "scan_anon.pdf"
    xs, ys, keep_boxes, redact_boxes = _make_table_pdf(src, scan=True)
    report = process_pdf(src, dst, Anonymizer())
    assert report["ocr_pages"] == 1 and report["ocr_failed_pages"] == 0
    _assert_table_result(src, dst, report, xs, ys, keep_boxes, redact_boxes, ocr=True)
