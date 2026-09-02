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


def _find_source_token_boxes(page, token):
    """Find visual token boxes despite platform-specific hyphen extraction."""
    hyphens = ("-", "\u00ad", "\u2010", "\u2011")
    variants = {token}
    if "-" in token:
        variants.update(token.replace("-", hyphen) for hyphen in hyphens)
    for variant in variants:
        boxes = page.search_for(variant)
        if boxes:
            return [tuple(box) for box in boxes]

    # PyMuPDF on Windows can expose a visually ordinary '-' as a soft hyphen
    # in extracted words while search_for() finds neither spelling. Match a
    # short word sequence canonically and retain its actual visual geometry.
    def canonical(value):
        for hyphen in hyphens[1:]:
            value = value.replace(hyphen, "-")
        return "".join(value.split())

    wanted = canonical(token)
    words = page.get_text("words", sort=True)
    matches = []
    for start in range(len(words)):
        combined = ""
        for end in range(start, min(start + 8, len(words))):
            combined += canonical(words[end][4])
            if combined == wanted:
                box = fitz.Rect(*words[start][:4])
                for word in words[start + 1:end + 1]:
                    box |= fitz.Rect(*word[:4])
                matches.append(tuple(box))
                break
            if len(combined) >= len(wanted):
                break
    return matches


def _make_table_pdf(path: Path, scan: bool = False):
    doc = fitz.open()
    # Keep the fixture deliberately roomy. Font metrics differ slightly
    # between DejaVu on Linux and Arial on the Windows CI runner, so a tight
    # textbox can silently omit its entire value on only one platform.
    page_width, page_height = 1600, 620
    page = doc.new_page(width=page_width, height=page_height)
    xs = [20, 65, 440, 850, 1050, 1270, 1330, 1400, 1480, 1580]
    ys = [20, 100, 340, 580]
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
    def insert_checked(rect, value, *, fontsize, label):
        remaining = page.insert_textbox(rect, value, fontsize=fontsize,
                                        fontname="dejavu", fontfile=FONT)
        if value and remaining < 0:
            raise AssertionError(
                f"Synthetic fixture textbox overflow for {label}: "
                f"remaining={remaining:.2f}, rect={tuple(rect)}, value={value!r}"
            )

    for col, value in enumerate(headers):
        insert_checked((xs[col] + 2, 24, xs[col + 1] - 2, 96), value,
                       fontsize=6, label=f"header column {col + 1}")
    for row_no, values in enumerate(rows):
        for col, value in enumerate(values):
            insert_checked((xs[col] + 2, ys[row_no + 1] + 4,
                            xs[col + 1] - 2, ys[row_no + 2] - 4), value,
                           fontsize=7 if col == 2 else 8,
                           label=f"row {row_no + 1}, column {col + 1}")
    keep_tokens = ("IP66", "УХЛ1", "Ex d IIC T6", "DN100", "PN1,6 МПа",
                   "09Г2С", "ГОСТ 12345", "100х50", "TEST.0001-АТТ.ОЛ1",
                   "TEST-M1", "ТУ 1234-567-890", "TEST-M2")
    keep_boxes = {token: _find_source_token_boxes(page, token) for token in keep_tokens}
    assert all(keep_boxes.values()), f"Synthetic source overflowed KEEP text: {keep_boxes}"
    source_text = page.get_text()
    for hyphen in ("\u00ad", "\u2010", "\u2011"):
        source_text = source_text.replace(hyphen, "-")
    source_text = " ".join(source_text.split())
    assert "Комплектация по обосновывающему документу TEST.0001-АТТ.ОЛ1" in source_text
    redact_boxes = {
        "brand": [tuple(rect) for rect in page.search_for("Армтел")],
        "supplier_cell": [(xs[4], ys[1], xs[5], ys[2])],
    }
    if scan:
        pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False)
        image = pix.tobytes("png")
        scanned = fitz.open(); target = scanned.new_page(width=page_width, height=page_height)
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
        for hyphen in ("\u00ad", "\u2010", "\u2011"):
            value = value.replace(hyphen, "-")
        return " ".join(value.replace("\r", "\n").split())

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
    diagnostics = report["ocr_table_diagnostics"][0]
    assert diagnostics["mode"] == "ocr-grid-columns", diagnostics
    assert len(diagnostics["vertical_boundaries"]) >= 10, diagnostics
    assert diagnostics["code_column_zone"] and diagnostics["supplier_column_zone"], diagnostics
    _assert_table_result(src, dst, report, xs, ys, keep_boxes, redact_boxes, ocr=True)
