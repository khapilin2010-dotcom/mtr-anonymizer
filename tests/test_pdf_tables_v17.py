import os
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
         "Модель TEST-M1 ГОСТ 12345 Комплектация по обосновывающему документу TEST.0001-АТТ.ОЛ1",
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
                                fontsize=8, fontname="dejavu", fontfile=FONT)
    if scan:
        pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False)
        image = pix.tobytes("png")
        scanned = fitz.open(); target = scanned.new_page(width=1200, height=420)
        target.insert_image(target.rect, stream=image)
        scanned.save(path); scanned.close(); doc.close()
    else:
        doc.save(path); doc.close()
    return xs, ys


def _assert_table_result(src, dst, report, xs, ys, ocr=False):
    assert report["table_pages"] == 1
    assert report["supplier_cells_redacted"] >= 2
    assert report["code_column_unchanged"] is True
    assert report["code_column_intersections"] == 0
    for redaction in map(fitz.Rect, report["redaction_rects"]):
        for keep in map(fitz.Rect, report["code_keep_rects"]):
            assert (redaction & keep).get_area() == 0
    result = fitz.open(dst); page = result[0]
    textpage = page.get_textpage_ocr(language="rus+eng", dpi=300, full=True) if ocr else None
    words = page.get_text("words", textpage=textpage, sort=True) if ocr else None

    def region_text(rect):
        if not ocr:
            return page.get_textbox(rect)
        return " ".join(word[4] for word in words if fitz.Rect(*word[:4]).intersects(rect))

    code_text = region_text(fitz.Rect(xs[3], ys[1], xs[4], ys[-1]))
    assert "1234567 (0)1)" in code_text
    assert "4143086 (0)1)" in code_text
    assert "2576244 (0)1)" in code_text
    assert "Заявка № Z1234567 (1342)" in code_text
    supplier_text = region_text(fitz.Rect(xs[4], ys[1], xs[5], ys[-1]))
    assert not supplier_text.strip()
    all_text = page.get_text("text", textpage=textpage)
    normalized_text = " ".join(all_text.split())
    for keep in ("IP66", "УХЛ1", "Ex d IIC T6", "DN100", "PN1,6 МПа",
                 "09Г2С", "ГОСТ 12345", "100х50", "TEST-M1",
                 "Комплектация по обосновывающему документу", "TEST.0001-АТТ.ОЛ1"):
        assert keep in normalized_text
    assert "ТУ 1234-567-890 TEST-M2" in normalized_text
    assert "ТЕСТОВЫЙ ЗАВОД" not in all_text and "1234567890" not in all_text
    result.close()


def test_native_table_columns_are_respected(tmp_path):
    src, dst = tmp_path / "table.pdf", tmp_path / "table_anon.pdf"
    xs, ys = _make_table_pdf(src)
    report = process_pdf(src, dst, Anonymizer())
    _assert_table_result(src, dst, report, xs, ys)
    repeat = tmp_path / "table_repeat.pdf"
    second = process_pdf(dst, repeat, Anonymizer())
    assert second.get("unchanged") is True
    assert "".join(p.get_text() for p in fitz.open(dst)) == "".join(p.get_text() for p in fitz.open(repeat))


def test_ocr_table_columns_are_respected(tmp_path):
    if not os.environ.get("MTR_REQUIRE_OCR"):
        pytest.skip("Real RU+EN OCR is mandatory in Windows CI")
    src, dst = tmp_path / "scan.pdf", tmp_path / "scan_anon.pdf"
    xs, ys = _make_table_pdf(src, scan=True)
    report = process_pdf(src, dst, Anonymizer())
    assert report["ocr_pages"] == 1 and report["ocr_failed_pages"] == 0
    _assert_table_result(src, dst, report, xs, ys, ocr=True)
