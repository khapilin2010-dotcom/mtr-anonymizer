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
                   "TEST-M1", "ТУ 1234-567-890", "TEST-M2",
                   "1234567 (0)1)", "4143086 (0)1)", "2576244 (0)1)",
                   "Заявка № Z1234567 (1342)")
    keep_boxes = {token: _find_source_token_boxes(page, token) for token in keep_tokens}
    assert all(keep_boxes.values()), f"Synthetic source overflowed KEEP text: {keep_boxes}"
    source_text = page.get_text()
    for hyphen in ("\u00ad", "\u2010", "\u2011"):
        source_text = source_text.replace(hyphen, "-")
    source_text = " ".join(source_text.split())
    assert "Комплектация по обосновывающему документу TEST.0001-АТТ.ОЛ1" in source_text
    supplier_zone = fitz.Rect(xs[4], ys[1], xs[5], ys[-1])
    supplier_glyphs = [tuple(word[:4]) for word in page.get_text("words", sort=True)
                       if fitz.Rect(*word[:4]).intersects(supplier_zone)]
    assert supplier_glyphs, "Synthetic source has no supplier glyph boxes"
    redact_boxes = {
        "brand": [tuple(rect) for rect in page.search_for("Армтел")],
        "supplier_cell": [(xs[4], ys[1], xs[5], ys[2])],
        "supplier_glyphs": supplier_glyphs,
        "supplier_grid": [
            # Inspect the grid stroke from the non-redacted side. Adjacent
            # white cell background is not part of the line invariant.
            (xs[4] - 0.55, ys[1], xs[4] - 0.15, ys[-1]),
            (xs[5] - 0.55, ys[1], xs[5] - 0.15, ys[-1]),
            *[(xs[4], y - 0.55, xs[5], y - 0.15) for y in ys[1:]],
        ],
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


def _pixel_diff(before, after, width, channels):
    changed_channels = [index for index, pair in enumerate(zip(before, after))
                        if pair[0] != pair[1]]
    if not changed_channels:
        return {"changed_pixels": 0, "bbox": None, "max_rgb_delta": 0}
    pixels = {index // channels for index in changed_channels}
    px = [pixel % width for pixel in pixels]
    py = [pixel // width for pixel in pixels]
    return {
        "changed_pixels": len(pixels),
        "bbox": (min(px), min(py), max(px), max(py)),
        "max_rgb_delta": max(abs(before[index] - after[index])
                             for index in changed_channels),
    }


def _grid_line_diagnostics(source_page, result_page, rect, orientation):
    """Describe exact pixels and structural continuity of one table line."""
    scale = 4
    matrix = fitz.Matrix(scale, scale)
    center_x = (rect.x0 + rect.x1) / 2
    center_y = (rect.y0 + rect.y1) / 2
    analysis_rect = (fitz.Rect(rect.x0, center_y - 2, rect.x1, center_y + 2)
                     if orientation == "horizontal" else
                     fitz.Rect(center_x - 2, rect.y0, center_x + 2, rect.y1))
    before = source_page.get_pixmap(matrix=matrix, clip=analysis_rect,
                                    colorspace=fitz.csGRAY, alpha=False)
    after = result_page.get_pixmap(matrix=matrix, clip=analysis_rect,
                                   colorspace=fitz.csGRAY, alpha=False)
    diff = _pixel_diff(bytes(before.samples), bytes(after.samples), before.width, before.n)
    diff["max_abs_delta"] = diff["max_rgb_delta"]
    changed_coordinates = []
    for py in range(before.height):
        for px in range(before.width):
            index = py * before.width + px
            if before.samples[index] != after.samples[index]:
                changed_coordinates.append((
                    round(analysis_rect.x0 + (px + 0.5) / scale, 3),
                    round(analysis_rect.y0 + (py + 0.5) / scale, 3),
                ))

    def structure(pix):
        dark = 160
        if orientation == "horizontal":
            runs = [[py for py in range(pix.height)
                     if pix.samples[py * pix.width + px] < dark]
                    for px in range(pix.width)]
            darkness = [sum(pix.samples[py * pix.width + px]
                            for px in range(pix.width))
                        for py in range(pix.height)]
            core_index = min(range(pix.height), key=darkness.__getitem__)
            core = bytes(pix.samples[core_index * pix.width:(core_index + 1) * pix.width])
        else:
            runs = [[px for px in range(pix.width)
                     if pix.samples[py * pix.width + px] < dark]
                    for py in range(pix.height)]
            darkness = [sum(pix.samples[py * pix.width + px]
                            for py in range(pix.height))
                        for px in range(pix.width)]
            core_index = min(range(pix.width), key=darkness.__getitem__)
            core = bytes(pix.samples[py * pix.width + core_index]
                         for py in range(pix.height))
        thickness = [len(run) for run in runs]
        present = [bool(run) for run in runs]
        max_gap = gap = 0
        for value in present:
            gap = 0 if value else gap + 1
            max_gap = max(max_gap, gap)
        return {
            "dark_pixel_count": sum(value < dark for value in pix.samples),
            "minimum_grayscale": min(pix.samples),
            "continuous_samples": sum(present),
            "total_samples": len(present),
            "max_gap_pixels": max_gap,
            "core_index": core_index,
            "core_sha256": hashlib.sha256(core).hexdigest(),
            "thickness_min": min(thickness),
            "thickness_max": max(thickness),
            "thickness_average": sum(thickness) / len(thickness),
        }

    coordinate_bbox = None
    if changed_coordinates:
        xx, yy = zip(*changed_coordinates)
        coordinate_bbox = (min(xx), min(yy), max(xx), max(yy))
    return {"orientation": orientation, "rect": tuple(rect),
            "analysis_rect": tuple(analysis_rect), "pixel_diff": diff,
            "changed_pixel_coordinates": changed_coordinates[:32],
            "changed_pixel_coordinate_count": len(changed_coordinates),
            "changed_pixel_coordinates_bbox": coordinate_bbox,
            "before": structure(before), "after": structure(after)}


def _save_and_redaction_controls(src, tmp_dir, code_rect, supplier_rect):
    """Distinguish save/recompression effects from image-pixel redaction."""
    matrix = fitz.Matrix(2, 2)
    original = fitz.open(src)
    original_pix = original[0].get_pixmap(matrix=matrix, clip=code_rect, alpha=False)
    original_samples = bytes(original_pix.samples)
    original.close()

    saved_path = tmp_dir / "control_saved_without_redaction.pdf"
    saved = fitz.open(src)
    saved.save(saved_path, garbage=4, deflate=True, clean=True)
    saved.close()
    reopened = fitz.open(saved_path)
    saved_pix = reopened[0].get_pixmap(matrix=matrix, clip=code_rect, alpha=False)
    reopened.close()

    redacted_path = tmp_dir / "control_supplier_redaction.pdf"
    controlled = fitz.open(src)
    page = controlled[0]
    # Stay well inside the supplier cell: this control proves whether
    # PDF_REDACT_IMAGE_PIXELS itself rewrites remote code-column pixels.
    far_supplier = fitz.Rect(supplier_rect.x0 + 10, supplier_rect.y0 + 10,
                             supplier_rect.x1 - 10, supplier_rect.y1 - 10)
    page.add_redact_annot(far_supplier, fill=(1, 1, 1), cross_out=False)
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_PIXELS,
                          graphics=fitz.PDF_REDACT_LINE_ART_NONE,
                          text=fitz.PDF_REDACT_TEXT_REMOVE)
    controlled.save(redacted_path, garbage=4, deflate=True, clean=True)
    controlled.close()
    reopened = fitz.open(redacted_path)
    redacted_pix = reopened[0].get_pixmap(matrix=matrix, clip=code_rect, alpha=False)
    reopened.close()

    width, channels = original_pix.width, original_pix.n
    return {
        "save_without_redaction": _pixel_diff(
            original_samples, bytes(saved_pix.samples), width, channels),
        "far_supplier_pixel_redaction": _pixel_diff(
            original_samples, bytes(redacted_pix.samples), width, channels),
    }


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
        controls = _save_and_redaction_controls(
            src, Path(dst).parent, code_rect,
            fitz.Rect(redact_boxes["supplier_cell"][0]),
        )
        assert controls["save_without_redaction"]["changed_pixels"] == 0, controls
        assert controls["far_supplier_pixel_redaction"]["changed_pixels"] == 0, controls
        code_diff = _pixel_diff(bytes(before_crop.samples), bytes(after_crop.samples),
                                before_crop.width, before_crop.n)
        glyph_diffs = {}
        for token in ("1234567 (0)1)", "4143086 (0)1)", "2576244 (0)1)",
                      "Заявка № Z1234567 (1342)"):
            glyph_diffs[token] = []
            for glyph_box in map(fitz.Rect, keep_boxes[token]):
                before_glyph = source_page.get_pixmap(
                    matrix=fitz.Matrix(2, 2), clip=glyph_box, alpha=False)
                after_glyph = page.get_pixmap(
                    matrix=fitz.Matrix(2, 2), clip=glyph_box, alpha=False)
                glyph_diffs[token].append(_pixel_diff(
                    bytes(before_glyph.samples), bytes(after_glyph.samples),
                    before_glyph.width, before_glyph.n,
                ))
        assert all(
            diff["changed_pixels"] == 0
            for token_diffs in glyph_diffs.values() for diff in token_diffs
        ), {"code_glyph_diffs": glyph_diffs,
            "closest_code_redaction": report["closest_code_redaction"]}

        def region_diff(rect):
            before = source_page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=rect, alpha=False)
            after = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=rect, alpha=False)
            return _pixel_diff(bytes(before.samples), bytes(after.samples),
                               before.width, before.n)

        border_diffs = {
            "left": region_diff(fitz.Rect(code_rect.x0, code_rect.y0,
                                           code_rect.x0 + 1, code_rect.y1)),
            "right": region_diff(fitz.Rect(code_rect.x1 - 1, code_rect.y0,
                                            code_rect.x1, code_rect.y1)),
            "top": region_diff(fitz.Rect(code_rect.x0, code_rect.y0,
                                          code_rect.x1, code_rect.y0 + 1)),
            "bottom": region_diff(fitz.Rect(code_rect.x0, code_rect.y1 - 1,
                                             code_rect.x1, code_rect.y1)),
        }
        assert code_diff["changed_pixels"] == 0, {
            "processed_code_column_diff": code_diff,
            "code_glyph_diffs": glyph_diffs,
            "code_grid_border_diffs": border_diffs,
            "controls": controls,
            "redaction_rects": report["redaction_rects"],
            "redaction_distances": report["code_column_redaction_distances"],
            "closest_code_redaction": report["closest_code_redaction"],
            "code_zone": tuple(code_rect),
        }
    supplier_text = region_text(fitz.Rect(xs[4], ys[1], xs[5], ys[-1]))
    supplier_residual_words = ([{"text": word[4], "bbox": tuple(word[:4]),
                                 "distance_from_code_column": word[0] - xs[4]}
                                for word in (words or [])
                                if fitz.Rect(*word[:4]).intersects(
                                    fitz.Rect(xs[4], ys[1], xs[5], ys[-1]))])
    supplier_evidence = {
        "supplier_text": supplier_text,
        "supplier_residual_words": supplier_residual_words,
        "supplier_column_boundary": xs[4],
        "supplier_redaction_rects": [
            rect for rect in report["redaction_rects"]
            if fitz.Rect(rect).intersects(fitz.Rect(xs[4], ys[1], xs[5], ys[-1]))
        ],
        "source_supplier_words": report["ocr_table_diagnostics"][0].get(
            "supplier_source_words", []) if ocr else [],
        "code_keep_guard": report["ocr_table_diagnostics"][0].get(
            "code_keep_guard") if ocr else None,
    }
    if not ocr:
        assert not supplier_text.strip(), supplier_evidence
    else:
        grid_boxes = list(map(fitz.Rect, redact_boxes["supplier_grid"]))
        artifact_grid_boxes = [fitz.Rect(grid.x0 - 0.7, grid.y0 - 0.7,
                                         grid.x1 + 0.7, grid.y1 + 0.7)
                               for grid in grid_boxes]

        def classify_residual_pixels(rect):
            result_pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=rect,
                                         colorspace=fitz.csGRAY, alpha=False)
            source_pix = source_page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=rect,
                                                colorspace=fitz.csGRAY, alpha=False)
            outside, supplier_remnants, changed_outside = [], [], []
            for py in range(result_pix.height):
                for px in range(result_pix.width):
                    index = py * result_pix.width + px
                    if result_pix.samples[index] >= 200:
                        continue
                    point = fitz.Point(rect.x0 + (px + 0.5) / 2,
                                       rect.y0 + (py + 0.5) / 2)
                    if not any(point in grid for grid in artifact_grid_boxes):
                        coordinate = (round(point.x, 3), round(point.y, 3))
                        outside.append(coordinate)
                        if any(point in fitz.Rect(box)
                               for box in redact_boxes["supplier_glyphs"]):
                            supplier_remnants.append(coordinate)
                        if source_pix.samples[index] != result_pix.samples[index]:
                            changed_outside.append(coordinate)
            outside_bbox = None
            if outside:
                xx, yy = zip(*outside)
                outside_bbox = (min(xx), min(yy), max(xx), max(yy))
            return {
                "dark_pixels_outside_grid": len(outside),
                "dark_pixels_outside_grid_coordinates": outside,
                "dark_pixels_outside_grid_bbox": outside_bbox,
                "source_supplier_glyph_remnant_pixels": supplier_remnants,
                "changed_dark_pixels_outside_grid": changed_outside,
            }

        residual_analysis = []
        for residual in supplier_residual_words:
            rect = fitz.Rect(residual["bbox"])
            before = source_page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=rect,
                                             alpha=False)
            after = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=rect,
                                    alpha=False)
            pixel_classification = classify_residual_pixels(rect)
            residual_analysis.append({
                **residual,
                "confidence": None,  # PyMuPDF OCR TextPage does not expose it.
                "source_ocr_words": [
                    {"text": word[4], "bbox": tuple(word[:4])}
                    for word in source_words
                    if fitz.Rect(*word[:4]).intersects(rect)
                ],
                "pixel_diff": _pixel_diff(bytes(before.samples), bytes(after.samples),
                                            before.width, before.n),
                "intersects_source_supplier_glyph": any(
                    (rect & fitz.Rect(box)).get_area() > 0
                    for box in redact_boxes["supplier_glyphs"]),
                "intersects_grid": any((rect & grid).get_area() > 0
                                       for grid in artifact_grid_boxes),
                **pixel_classification,
            })

        # OCR may label an intact table stroke as e.g. "in". It is harmless
        # only when off-grid antialias pixels are unchanged source pixels and
        # none belongs to any original supplier glyph. This pixel provenance
        # check is stricter than trusting the OCR word bbox, which can be much
        # wider than the actual dark pixels at a grid intersection.
        assert all(item["intersects_grid"] and
                   not item["source_supplier_glyph_remnant_pixels"] and
                   not item["changed_dark_pixels_outside_grid"]
                   for item in residual_analysis), {
            **supplier_evidence, "residual_analysis": residual_analysis,
        }
        sensitive = re.compile(
            r"(?i)(?:тестов|завод|инн|inn|ооо|оао|зао|пао|\bао\b|"
            r"1234567890|123456789012)"
        )
        assert not sensitive.search(supplier_text), {
            **supplier_evidence, "residual_analysis": residual_analysis,
        }

        # Security authority is the source glyph geometry: every original
        # supplier word must contain dark pixels before redaction and none
        # afterwards. This proves physical removal independent of OCR output.
        for glyph_box in map(fitz.Rect, redact_boxes["supplier_glyphs"]):
            before = source_page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=glyph_box,
                                             colorspace=fitz.csGRAY, alpha=False)
            after = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=glyph_box,
                                    colorspace=fitz.csGRAY, alpha=False)
            assert sum(value < 200 for value in before.samples) > 0
            assert sum(value < 200 for value in after.samples) == 0, {
                "source_supplier_glyph_bbox": tuple(glyph_box),
                "residual_analysis": residual_analysis,
            }

        # Grid lines remain visually and physically identical.
        grid_diagnostics = []
        supplier_redactions = [fitz.Rect(rect) for rect in report["redaction_rects"]
                               if fitz.Rect(rect).intersects(
                                   fitz.Rect(xs[4], ys[1], xs[5], ys[-1]))]
        for grid_no, grid_box in enumerate(grid_boxes):
            before = source_page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=grid_box,
                                             alpha=False)
            after = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=grid_box,
                                    alpha=False)
            orientation = "horizontal" if grid_box.width > grid_box.height else "vertical"
            diagnostic = _grid_line_diagnostics(source_page, page, grid_box, orientation)
            distances = []
            for redaction in supplier_redactions:
                dx = max(grid_box.x0 - redaction.x1,
                         redaction.x0 - grid_box.x1, 0.0)
                dy = max(grid_box.y0 - redaction.y1,
                         redaction.y0 - grid_box.y1, 0.0)
                distances.append((dx * dx + dy * dy) ** 0.5)
            diagnostic["supplier_redaction_distances"] = distances
            diagnostic["minimum_supplier_redaction_distance"] = (
                min(distances) if distances else None
            )
            diagnostic["grid_role"] = (
                "code_supplier_vertical_boundary" if grid_no == 0 else
                "supplier_unit_vertical_boundary" if grid_no == 1 else
                "supplier_horizontal_boundary"
            )
            grid_diagnostics.append(diagnostic)
            failure = {"failed_grid": diagnostic,
                       "supplier_redaction_rects": [tuple(rect)
                                                     for rect in supplier_redactions]}
            if orientation == "vertical":
                # Especially the CODE/SUPPLIER divider remains the stricter
                # byte-identical invariant: it is part of absolute code KEEP.
                assert bytes(before.samples) == bytes(after.samples), failure
            else:
                source_line, result_line = diagnostic["before"], diagnostic["after"]
                # PDF_REDACT_IMAGE_PIXELS may re-quantize a few antialias edge
                # samples near line intersections. The authoritative line is
                # nevertheless unchanged only if its complete dark core,
                # position, continuity, gaplessness and minimum thickness all
                # remain intact. No tolerance applies to these invariants.
                assert source_line["core_index"] == result_line["core_index"], failure
                assert source_line["core_sha256"] == result_line["core_sha256"], failure
                assert source_line["minimum_grayscale"] == result_line["minimum_grayscale"], failure
                assert source_line["max_gap_pixels"] == result_line["max_gap_pixels"] == 0, failure
                assert (source_line["continuous_samples"] == source_line["total_samples"] ==
                        result_line["continuous_samples"] == result_line["total_samples"]), failure
                assert result_line["thickness_min"] >= source_line["thickness_min"], failure
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
        for key in ("brand", "supplier_glyphs"):
            boxes = redact_boxes[key]
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
    assert diagnostics["grid_guards"] and report["grid_keep_rects"], diagnostics
    assert all("label" in item and "expanded_bbox" in item and
               "final_safe_bbox" in item and "near_grid" in item
               for item in report["ocr_redaction_diagnostics"]), report
    _assert_table_result(src, dst, report, xs, ys, keep_boxes, redact_boxes, ocr=True)
