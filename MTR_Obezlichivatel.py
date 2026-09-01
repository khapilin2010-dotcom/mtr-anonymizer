# -*- coding: utf-8 -*-
from __future__ import annotations

import csv
import os
import re
import sys
import threading
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from mtr_core import Anonymizer, keep_signature

APP_TITLE = "Обезличивание МТР"
APP_VERSION = "15.0 FINAL"


def resource_path(name: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / name


def norm_header(value) -> str:
    s = str(value or "").strip().lower().replace("ё", "е")
    s = re.sub(r"\s+", " ", s)
    return s


def find_header_row(ws, max_rows: int = 30):
    code_names = {"код autodocs", "код автодокс", "код", "autodocs", "код мтр", "код ресурса", "№ смет/ код ресурса", "№ смет / код ресурса"}
    name_names = {"наименование", "наименование мтр", "мтр", "описание"}
    factory_names = {"завод", "изготовитель", "производитель", "поставщик", "завод/изготовитель/поставщик"}
    anon_names = {"обезличенное наименование", "обезличенное", "результат"}
    status_names = {"статус", "статус обезличивания"}

    for row in range(1, min(ws.max_row, max_rows) + 1):
        vals = [norm_header(ws.cell(row, c).value) for c in range(1, min(ws.max_column, 100) + 1)]
        mapping = {}
        for c, h in enumerate(vals, 1):
            if h in code_names and "code" not in mapping:
                mapping["code"] = c
            if h in name_names and "name" not in mapping:
                mapping["name"] = c
            if h in factory_names and "factory" not in mapping:
                mapping["factory"] = c
            if h in anon_names and "anon" not in mapping:
                mapping["anon"] = c
            if h in status_names and "status" not in mapping:
                mapping["status"] = c
        if "name" in mapping and ("code" in mapping or "factory" in mapping):
            return row, mapping
    raise ValueError(
        "Не найдена строка заголовков. Нужны столбцы «Наименование» и хотя бы один из: «Код Autodocs/Код ресурса» / «Завод»."
    )


def select_excel_sheets(wb):
    """Pick the intended working sheet in estimate/resource workbooks.

    Priority requested for production files: exact "Готово", otherwise exact
    "Выборка оборудования". A helper sheet such as
    "выборка_оборудования_все" is intentionally excluded. For ordinary
    input workbooks, fall back to all sheets and let header detection decide.
    """
    def title_key(title):
        s = str(title or "").strip().lower().replace("ё", "е").replace("_", " ")
        return re.sub(r"\s+", " ", s)

    ready = [ws for ws in wb.worksheets if title_key(ws.title) == "готово"]
    if ready:
        return ready
    equipment = [ws for ws in wb.worksheets if title_key(ws.title) == "выборка оборудования"]
    if equipment:
        return equipment
    return list(wb.worksheets)


def process_excel(src: Path, dst: Path, az: Anonymizer, progress=None):
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import PatternFill, Font, Alignment
    except ImportError as e:
        raise RuntimeError("Для Excel требуется пакет openpyxl. Запустите install_and_run.bat.") from e

    keep_vba = src.suffix.lower() == ".xlsm"
    wb = load_workbook(src, keep_vba=keep_vba)
    report = {"rows": 0, "changed": 0, "green": 0, "yellow": 0, "sheets": 0,
              "keep_losses": 0, "idempotence_failures": 0, "residual_confirmed": 0}

    green_fill = PatternFill("solid", fgColor="D9EAD3")
    yellow_fill = PatternFill("solid", fgColor="FFF2CC")
    header_fill = PatternFill("solid", fgColor="D9EAF7")

    processed_any = False
    for ws in select_excel_sheets(wb):
        try:
            header_row, cols = find_header_row(ws)
        except ValueError:
            continue
        processed_any = True
        report["sheets"] += 1

        anon_col = cols.get("anon")
        if not anon_col:
            anon_col = ws.max_column + 1
            ws.cell(header_row, anon_col).value = "Обезличенное наименование"
        status_col = cols.get("status")
        if not status_col:
            status_col = max(ws.max_column, anon_col) + 1
            ws.cell(header_row, status_col).value = "Статус"

        for c in (anon_col, status_col):
            cell = ws.cell(header_row, c)
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        data_rows = max(ws.max_row - header_row, 0)
        for i, r in enumerate(range(header_row + 1, ws.max_row + 1), 1):
            name_cell = ws.cell(r, cols["name"])
            name = str(name_cell.value or "").strip()
            # Formula/error rows such as #DIV/0! are workbook calculations,
            # not MTR positions and must not enter anonymization statistics.
            if not name or name_cell.data_type == "e" or re.fullmatch(r"#[A-Z0-9/?!._-]+", name, re.I):
                continue
            code = str(ws.cell(r, cols.get("code", 0)).value or "").strip() if cols.get("code") else ""
            factory = str(ws.cell(r, cols.get("factory", 0)).value or "").strip() if cols.get("factory") else ""
            # Estimate workbooks often have a second header row with column
            # numbers (1/2/3/4). It is layout metadata, not an MTR position.
            if r <= header_row + 5 and code in {"1", "2"} and name in {"3", "4"}:
                continue
            result = az.anonymize(name, code, factory)
            # Built-in regression guards: they do not alter the result.
            before_keep = keep_signature(name)
            after_keep = keep_signature(result["text"])
            if before_keep != after_keep:
                report["keep_losses"] += 1
            repeat = az.anonymize(result["text"], code, factory)
            if repeat["text"] != result["text"]:
                report["idempotence_failures"] += 1
            if az.redaction_spans(result["text"], code, factory):
                report["residual_confirmed"] += 1

            ws.cell(r, anon_col).value = result["text"]
            ws.cell(r, status_col).value = result["status"]
            fill = green_fill if result["status"] == "ЗЕЛЁНЫЙ" else yellow_fill
            ws.cell(r, status_col).fill = fill
            if cols.get("code"):
                ws.cell(r, cols["code"]).fill = fill
            if cols.get("factory"):
                # The anonymized output must not retain producer details.
                ws.cell(r, cols["factory"]).value = ""

            report["rows"] += 1
            report["changed"] += int(result["changed"])
            if result["status"] == "ЗЕЛЁНЫЙ":
                report["green"] += 1
            else:
                report["yellow"] += 1

            if progress and (i % 100 == 0 or i == data_rows):
                progress(f"{ws.title}: обработано {i} из {data_rows}")

        ws.column_dimensions[ws.cell(header_row, anon_col).column_letter].width = min(max(ws.column_dimensions[ws.cell(header_row, cols["name"]).column_letter].width or 25, 40), 100)
        ws.column_dimensions[ws.cell(header_row, status_col).column_letter].width = 15

    if not processed_any:
        raise ValueError("Ни на одном листе не найдены подходящие столбцы для обработки.")

    wb.save(dst)
    return report


def process_csv(src: Path, dst: Path, az: Anonymizer, progress=None):
    raw = src.read_bytes()
    encoding = "utf-8-sig"
    try:
        text = raw.decode(encoding)
    except UnicodeDecodeError:
        encoding = "cp1251"
        text = raw.decode(encoding)
    sample = text[:5000]
    dialect = csv.Sniffer().sniff(sample, delimiters=";,\t,")
    rows = list(csv.reader(text.splitlines(), dialect))
    if not rows:
        raise ValueError("CSV пуст.")

    header_idx = None
    mapping = None
    for idx, row in enumerate(rows[:30]):
        heads = [norm_header(x) for x in row]
        m = {}
        for c, h in enumerate(heads):
            if h in {"код autodocs", "код автодокс", "код", "autodocs", "код мтр"} and "code" not in m:
                m["code"] = c
            if h in {"наименование", "наименование мтр", "мтр", "описание"} and "name" not in m:
                m["name"] = c
            if h in {"завод", "изготовитель", "производитель", "поставщик", "завод/изготовитель/поставщик"} and "factory" not in m:
                m["factory"] = c
        if "name" in m and ("code" in m or "factory" in m):
            header_idx, mapping = idx, m
            break
    if mapping is None:
        raise ValueError("В CSV не найдены столбцы «Наименование» и «Код Autodocs»/«Завод».")

    header = rows[header_idx]
    anon_col = len(header)
    status_col = anon_col + 1
    header.extend(["Обезличенное наименование", "Статус"])
    max_len = len(header)
    report = {"rows": 0, "changed": 0, "green": 0, "yellow": 0, "sheets": 1}

    for i in range(header_idx + 1, len(rows)):
        row = rows[i]
        row.extend([""] * (max_len - len(row)))
        name = row[mapping["name"]].strip() if mapping["name"] < len(row) else ""
        if not name:
            continue
        code = row[mapping["code"]].strip() if mapping.get("code") is not None and mapping["code"] < len(row) else ""
        factory = row[mapping["factory"]].strip() if mapping.get("factory") is not None and mapping["factory"] < len(row) else ""
        result = az.anonymize(name, code, factory)
        row[anon_col] = result["text"]
        row[status_col] = result["status"]
        if mapping.get("factory") is not None:
            row[mapping["factory"]] = ""
        report["rows"] += 1
        report["changed"] += int(result["changed"])
        report["green"] += int(result["status"] == "ЗЕЛЁНЫЙ")
        report["yellow"] += int(result["status"] == "ЖЁЛТЫЙ")
        if progress and i % 200 == 0:
            progress(f"CSV: обработано {i - header_idx} строк")

    with open(dst, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f, dialect=dialect)
        writer.writerows(rows)
    return report


def _block_code(text: str, az: Anonymizer) -> str:
    # Conservative: a registry number is treated as Autodocs code only when it
    # is explicitly labelled or appears at the beginning of a text block.
    text = str(text or "")
    labelled = re.search(r"(?i)код(?:\s+autodocs|\s+автодокс)?\s*[:№-]?\s*(-?\d{3,9})", text)
    if labelled and labelled.group(1) in az.registry:
        return labelled.group(1)
    head = text.lstrip()[:80]
    for token in re.findall(r"(?<!\d)-?\d{3,9}(?!\d)", head):
        if token in az.registry:
            return token
    return ""


def _same_row_code(block_rect, code_blocks):
    best = None
    cy = (block_rect.y0 + block_rect.y1) / 2
    for rect, code in code_blocks:
        overlap = max(0.0, min(block_rect.y1, rect.y1) - max(block_rect.y0, rect.y0))
        min_h = max(1.0, min(block_rect.height, rect.height))
        row_like = overlap / min_h >= 0.45 or abs(cy - (rect.y0 + rect.y1) / 2) <= 7.0
        if not row_like:
            continue
        dist = abs(cy - (rect.y0 + rect.y1) / 2) + 0.001 * abs(block_rect.x0 - rect.x0)
        if best is None or dist < best[0]:
            best = (dist, code)
    return best[1] if best else ""


def _same_row_factory(block_rect, factory_blocks):
    """Return a confirmed manufacturer context from another block on the same visual row."""
    best = None
    cy = (block_rect.y0 + block_rect.y1) / 2
    for rect, factory, inn in factory_blocks:
        overlap = max(0.0, min(block_rect.y1, rect.y1) - max(block_rect.y0, rect.y0))
        min_h = max(1.0, min(block_rect.height, rect.height))
        row_like = overlap / min_h >= 0.45 or abs(cy - (rect.y0 + rect.y1) / 2) <= 7.0
        if not row_like:
            continue
        dist = abs(cy - (rect.y0 + rect.y1) / 2) + 0.001 * abs(block_rect.x0 - rect.x0)
        if best is None or dist < best[0]:
            best = (dist, factory, inn)
    return (best[1], best[2]) if best else ("", "")


def _raw_block_text_map(block):
    """Reconstruct one PyMuPDF raw text block and map every char to its bbox/line."""
    parts = []
    cmap = []
    line_no = 0
    lines = block.get("lines", [])
    for li, line in enumerate(lines):
        for span in line.get("spans", []):
            for ch in span.get("chars", []):
                parts.append(ch.get("c", ""))
                cmap.append((line_no, tuple(ch.get("bbox", (0, 0, 0, 0)))))
        if li != len(lines) - 1:
            parts.append("\n")
            cmap.append(None)
        line_no += 1
    return "".join(parts), cmap


def _tight_redaction_rects(text: str, cmap, start: int, end: int, fitz):
    """Turn a character span into thin per-line redaction strips.

    Many engineering PDFs use font bboxes taller than the physical line pitch.
    Full-height redaction rectangles therefore touch the line above/below and
    PyMuPDF may delete neighbouring technical text.  A thin strip through the
    target glyphs still removes the text object, but cannot intersect adjacent
    lines.
    """
    groups = {}
    for i in range(max(0, start), min(end, len(cmap))):
        item = cmap[i]
        if item is None or not text[i].strip():
            continue
        line_no, bbox = item
        groups.setdefault(line_no, []).append(bbox)

    rects = []
    for line_no in sorted(groups):
        boxes = groups[line_no]
        x0 = min(b[0] for b in boxes)
        y0 = min(b[1] for b in boxes)
        x1 = max(b[2] for b in boxes)
        y1 = max(b[3] for b in boxes)
        h = max(0.5, y1 - y0)
        # Middle-lower band of the target glyph bbox. For the 10 pt Arial used
        # in the supplied specifications this sits fully between neighbouring
        # line bboxes; the ratio also behaves well for other common fonts.
        sy0 = y0 + h * 0.55
        sy1 = y0 + h * 0.62
        if sy1 <= sy0:
            sy1 = sy0 + 0.8
        rects.append(fitz.Rect(x0, sy0, x1, sy1))
    return rects


def process_pdf(src: Path, dst: Path, az: Anonymizer, progress=None):
    try:
        import fitz
    except ImportError as e:
        raise RuntimeError("Для PDF требуется пакет PyMuPDF. Запустите install_and_run.bat.") from e

    doc = fitz.open(src)
    if doc.needs_pass:
        doc.close()
        raise ValueError("PDF защищён паролем. Сначала снимите защиту с файла.")

    report = {"pages": len(doc), "redactions": 0, "blocks": 0, "matches": 0,
              "tu_continuations": 0, "context_models": 0, "ocr_pages": 0,
              "ocr_failed_pages": 0, "review": False}

    # OCR is page-level: native text pages stay untouched, image-only pages are
    # recognized in Russian + English and redacted on the original page image.
    tessdata = None
    for candidate in (
        str(resource_path("tessdata")),
        os.environ.get("TESSDATA_PREFIX", ""),
        r"C:\Program Files\Tesseract-OCR\tessdata",
        r"C:\Program Files (x86)\Tesseract-OCR\tessdata",
        "/usr/share/tesseract-ocr/5/tessdata",
        "/usr/share/tesseract-ocr/4.00/tessdata",
    ):
        if candidate and Path(candidate).exists():
            tessdata = candidate
            break
    pending_tu_continuation = False
    ocr_legal_form_re = re.compile(r"(?i)\b(?:ООО|АО|ЗАО|ОАО|ПАО|НПО|НПП|ФГУП|[ОO]{3}|[ОO][АA][ОO]|З[АA][ОO]|П[АA][ОO]|[АA][ОO]|FGUP)\b")
    for page_no, page in enumerate(doc, 1):
        native_text = page.get_text("text")
        ocr_page = len(native_text.strip()) < 12
        if ocr_page:
            try:
                kwargs = {"language": "rus+eng", "dpi": 300, "full": True}
                if tessdata:
                    kwargs["tessdata"] = tessdata
                textpage = page.get_textpage_ocr(**kwargs)
                raw = page.get_text("rawdict", textpage=textpage)
                report["ocr_pages"] += 1
                if progress:
                    progress(f"PDF: OCR rus+eng, страница {page_no} из {len(doc)}")
            except Exception as exc:
                report["ocr_failed_pages"] += 1
                if progress:
                    progress(f"PDF: OCR не выполнен на стр. {page_no}: {exc}")
                raw = {"blocks": []}
        else:
            raw = page.get_text("rawdict")
        text_blocks = []
        code_blocks = []
        factory_blocks = []

        for rb in raw.get("blocks", []):
            if rb.get("type") != 0:
                continue
            text, cmap = _raw_block_text_map(rb)
            if not text.strip():
                continue
            rect = fitz.Rect(*rb.get("bbox", (0, 0, 0, 0)))
            text_blocks.append((rect, text, cmap))
            ccode = _block_code(text, az)
            if ccode:
                code_blocks.append((rect, ccode))
            # Supplier / producer blocks frequently sit in a separate PDF
            # column from the model designation. Cache only confirmed contexts
            # (known INN / alias / verified marker) for same-row propagation.
            _ff, _fi = az.identify_in_text(text)
            if not _fi:
                _mf, _mi = az.identify_pdf_marker(text)
                if _mi:
                    _ff, _fi = _mf, _mi
            if _fi:
                factory_blocks.append((rect, _ff or az.manufacturer_name_by_inn.get(_fi, ""), _fi))

        page_rects = []
        # Some specifications split a TU number exactly at a page boundary:
        # page N ends with "ТУ 2531-" and page N+1 starts with
        # "002-53597015-12". The continuation is still part of the identifier
        # and must be removed, but only when the preceding page proves the
        # context.
        pending_for_this_page = pending_tu_continuation

        def _has_unfinished_tu(rect, txt):
            # A TU may be split by the physical page boundary even if table
            # extraction appends unit/quantity cells after the visible text.
            # Only consider lower-page blocks and a TU prefix whose immediate
            # continuation is not another digit in the same block.
            if rect.y0 < page.rect.height * 0.60:
                return False
            for mm in re.finditer(r"(?i)\bТУ\s+[0-9][0-9.\-/]*-", txt):
                tail = txt[mm.end():].lstrip(" \t\r\n")
                if not tail or not tail[:1].isdigit():
                    return True
            return False

        pending_tu_continuation = any(
            _has_unfinished_tu(rect, txt) for rect, txt, _cm in text_blocks
        )

        for block_index, (block_rect, text, cmap) in enumerate(text_blocks):
            report["blocks"] += 1
            code = _block_code(text, az) or _same_row_code(block_rect, code_blocks)
            factory, inn = az.identify_in_text(text)
            if code:
                resolved_factory, resolved_inn = az.resolve_factory(code, "")
                if resolved_inn:
                    factory, inn = resolved_factory, resolved_inn
            if not inn:
                row_factory, row_inn = _same_row_factory(block_rect, factory_blocks)
                if row_inn:
                    factory, inn = row_factory, row_inn
            if not inn:
                marker_factory, marker_inn = az.identify_pdf_marker(text)
                if marker_inn:
                    factory, inn = marker_factory, marker_inn

            spans = az.redaction_spans(text, code=code, factory=factory if inn else "")

            # OCR may interleave neighbouring table columns between the label
            # "ИНН" and its digits (for example: "ИНН м 2 0.074 ... 5904184047").
            # Once the manufacturer is already confirmed, its exact known INN
            # is an unambiguous identifier and must be physically redacted
            # wherever it appears in the same block.
            if inn and re.search(r"(?i)ИНН", text):
                for _inn_m in re.finditer(r"(?<!\d)" + re.escape(str(inn)) + r"(?!\d)", text):
                    spans.append((_inn_m.start(), _inn_m.end(), "known_inn"))

            # Supplier INN / KPP values are often split into a separate raw PDF
            # block on the following visual line. If the immediately preceding
            # block ends with the identifier label and this block contains only
            # the corresponding digits, redact the continuation as service data.
            if block_index > 0:
                _prev_rect, _prev_text, _prev_cmap = text_blocks[block_index - 1]

                # OCR can split the INN digits into the next raw block together
                # with unrelated table cells. Resolve the previous block's
                # manufacturer and redact only that exact known INN in the
                # current block; never redact the surrounding technical text.
                if re.search(r"(?i)ИНН", _prev_text):
                    _prev_code = _block_code(_prev_text, az) or _same_row_code(_prev_rect, code_blocks)
                    _pf, _pi = az.identify_in_text(_prev_text)
                    if _prev_code:
                        _rf, _ri = az.resolve_factory(_prev_code, "")
                        if _ri:
                            _pf, _pi = _rf, _ri
                    if not _pi:
                        _mf, _mi = az.identify_pdf_marker(_prev_text)
                        if _mi:
                            _pf, _pi = _mf, _mi
                    if _pi:
                        for _cross_m in re.finditer(r"(?<!\d)" + re.escape(str(_pi)) + r"(?!\d)", text):
                            spans.append((_cross_m.start(), _cross_m.end(), "inn_cross_block"))

                if re.search(r"(?i)\bИНН\s*$", _prev_text) and re.fullmatch(r"\s*\d{10,12}\s*", text):
                    spans.append((0, len(text), "inn_continuation"))
                elif re.search(r"(?i)\bКПП\s*$", _prev_text) and re.fullmatch(r"\s*\d{9}\s*", text):
                    spans.append((0, len(text), "kpp_continuation"))

                # A supplier name may wrap into the next raw block: e.g.
                # `ООО Калужский` / `электротехнический завод "КВТ", ИНН ...`.
                # If the previous block starts an organization and this block
                # completes it with an INN, both fragments are service data.
                _prev_org = ocr_legal_form_re.search(_prev_text)
                _this_inn = re.search(r"(?i)\bИНН\s*\d{10,12}\b", text)
                _near = abs(block_rect.y0 - _prev_rect.y1) <= 35 and not (block_rect.x1 < _prev_rect.x0 or block_rect.x0 > _prev_rect.x1 + 120)
                if _prev_org and _this_inn and _near:
                    # current continuation through the INN
                    spans.append((0, _this_inn.end(), "supplier_org_continuation"))

            # If this block starts a legal organization and the next block
            # immediately completes it with an INN, redact the organization
            # tail in this block from the legal form onwards.
            _org_here = ocr_legal_form_re.search(text)
            if _org_here and not re.search(r"(?i)\bИНН\b", text) and block_index + 1 < len(text_blocks):
                _next_rect, _next_text, _next_cmap = text_blocks[block_index + 1]
                _near_next = abs(_next_rect.y0 - block_rect.y1) <= 35 and not (_next_rect.x1 < block_rect.x0 or _next_rect.x0 > block_rect.x1 + 120)
                if _near_next and re.search(r"(?i)\bИНН\s*\d{10,12}\b", _next_text):
                    spans.append((_org_here.start(), len(text), "supplier_org_wrapped"))

            # Cross-page TU continuation, proven by the previous page only.
            # Limit this to the upper portion of the page and to a block that
            # consists only of a hyphenated numeric continuation.
            if pending_for_this_page and block_rect.y0 < 170:
                m_cont = re.fullmatch(r"\s*[0-9]{2,4}(?:-[0-9]{2,12}){1,4}\s*", text)
                if m_cont:
                    a, b = m_cont.span()
                    spans.append((a, b, "tu_continuation"))
                    report["tu_continuations"] += 1
                    pending_for_this_page = False

            # Layout-aware PDF fallback for KШГ catalogue designations. In
            # table PDFs the word KШГ and its catalogue number may be separated
            # in extraction order by unit/quantity columns even though they are
            # visually part of one item. If both occur in the same row block,
            # remove the confirmed firm designation and catalogue token while
            # leaving DN/PN/material/temperature text untouched.
            if re.search(r"(?i)(?<![\w])КШГ(?![\w])", text):
                model_pat = re.compile(r"(?<![\w])(?:71|79)\.[0-9]{3}\.[0-9]{3}\.[A-Za-zА-Яа-я]\.[0-9]+(?:\.[0-9]+)?(?![\w])")
                model_hits = list(model_pat.finditer(text))
                if model_hits:
                    for km in re.finditer(r"(?i)(?<![\w])КШГ(?![\w])", text):
                        spans.append((km.start(), km.end(), "context:КШГ"))
                    for mm in model_hits:
                        spans.append((mm.start(), mm.end(), "context:КШГ-model"))
                    report["context_models"] += len(model_hits)

            # De-duplicate and merge after adding contextual spans.
            spans = sorted(spans, key=lambda x: (x[0], x[1]))
            merged_spans = []
            for a, b, label in spans:
                if merged_spans and a <= merged_spans[-1][1]:
                    merged_spans[-1] = (merged_spans[-1][0], max(merged_spans[-1][1], b),
                                        merged_spans[-1][2] + ";" + label)
                else:
                    merged_spans.append((a, b, label))
            spans = merged_spans

            for a, b, _label in spans:
                rects = _tight_redaction_rects(text, cmap, a, b, fitz)
                if ocr_page and rects:
                    # OCR comes from an image. Thin strips delete a text object,
                    # but not the visible glyph pixels. Expand each matched strip
                    # to the full OCR glyph band before pixel redaction.
                    full_groups = {}
                    for ci in range(max(0, a), min(b, len(cmap))):
                        item = cmap[ci]
                        if item is None or not text[ci].strip():
                            continue
                        ln, bb = item
                        full_groups.setdefault(ln, []).append(bb)
                    rects = []
                    for ln in sorted(full_groups):
                        boxes = full_groups[ln]
                        x0=min(x[0] for x in boxes); y0=min(x[1] for x in boxes)
                        x1=max(x[2] for x in boxes); y1=max(x[3] for x in boxes)
                        rects.append(fitz.Rect(x0-0.8, y0-0.8, x1+0.8, y1+0.8))
                if not rects:
                    continue
                report["matches"] += 1
                for rect in rects:
                    # Avoid duplicate annotations from overlapping rules.
                    if any(
                        abs(rect.x0-r.x0) < 0.5 and abs(rect.y0-r.y0) < 0.5 and
                        abs(rect.x1-r.x1) < 0.5 and abs(rect.y1-r.y1) < 0.5
                        for r in page_rects
                    ):
                        continue
                    page.add_redact_annot(rect, fill=(1, 1, 1), cross_out=False)
                    page_rects.append(rect)
                    report["redactions"] += 1

        if page_rects:
            try:
                page.apply_redactions(
                    images=(fitz.PDF_REDACT_IMAGE_PIXELS if ocr_page else fitz.PDF_REDACT_IMAGE_NONE),
                    graphics=fitz.PDF_REDACT_LINE_ART_NONE,
                    text=fitz.PDF_REDACT_TEXT_REMOVE,
                )
            except TypeError:
                page.apply_redactions(images=0)
        if progress:
            progress(f"PDF: страница {page_no} из {len(doc)}")

    if report["redactions"] == 0:
        # For a scan, inability to make a safe automatic deletion is REVIEW,
        # not a fatal queue error. Save a separate copy and continue the batch.
        if report["ocr_pages"] or report["ocr_failed_pages"]:
            report["review"] = True
        else:
            doc.close()
            raise ValueError("В PDF не найдено фрагментов, подпадающих под правила обезличивания.")

    doc.save(dst, garbage=4, deflate=True, clean=True)
    doc.close()
    return report


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE} {APP_VERSION}")
        self.geometry("820x660")
        self.minsize(720, 600)
        self.configure(bg="#F4F8FC")
        self.file_paths: list[Path] = []
        self.last_outputs: list[Path] = []
        self.az = Anonymizer(resource_path("mtr_data.json.gz"))
        self._build_style()
        self._build_ui()

    def _build_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background="#F4F8FC")
        style.configure("Card.TFrame", background="#FFFFFF")
        style.configure("TLabel", background="#F4F8FC", foreground="#1F2D3D", font=("Segoe UI", 10))
        style.configure("Title.TLabel", background="#F4F8FC", foreground="#0A5AA8", font=("Segoe UI Semibold", 22))
        style.configure("Sub.TLabel", background="#F4F8FC", foreground="#5C6F82", font=("Segoe UI", 10))
        style.configure("Card.TLabel", background="#FFFFFF", foreground="#1F2D3D", font=("Segoe UI", 10))
        style.configure("Badge.TLabel", background="#E7F2FD", foreground="#0A5AA8", font=("Segoe UI Semibold", 9), padding=(9,4))
        style.configure("Primary.TButton", font=("Segoe UI Semibold", 11), padding=(16, 10), background="#0A6ACB", foreground="white")
        style.map("Primary.TButton", background=[("active", "#075AAE"), ("disabled", "#A9BDD1")])
        style.configure("Secondary.TButton", font=("Segoe UI", 10), padding=(12, 8))
        style.configure("Horizontal.TProgressbar", troughcolor="#DDE8F2", background="#0A6ACB", bordercolor="#DDE8F2", lightcolor="#0A6ACB", darkcolor="#0A6ACB")

    def _build_ui(self):
        outer = ttk.Frame(self, padding=24)
        outer.pack(fill="both", expand=True)

        top = ttk.Frame(outer)
        top.pack(fill="x")
        ttk.Label(top, text=APP_TITLE, style="Title.TLabel").pack(side="left")
        ttk.Label(top, text=f"База {self.az.version} • 138 630 МТР", style="Badge.TLabel").pack(side="right", pady=4)
        ttk.Label(outer, text="Удаление производителя, ТУ, фирменных моделей и артикулов с защитой технических характеристик.", style="Sub.TLabel").pack(anchor="w", pady=(4, 18))

        card = ttk.Frame(outer, style="Card.TFrame", padding=18)
        card.pack(fill="x")
        ttk.Label(card, text="Файл для обработки", style="Card.TLabel", font=("Segoe UI Semibold", 11)).pack(anchor="w")
        self.file_label = ttk.Label(card, text="Файл не выбран", style="Card.TLabel", foreground="#6D7F90")
        self.file_label.pack(anchor="w", pady=(6, 12))
        btns = ttk.Frame(card, style="Card.TFrame")
        btns.pack(fill="x")
        self.choose_btn = ttk.Button(btns, text="Выбрать файлы", command=self.choose_files, style="Secondary.TButton")
        self.choose_btn.pack(side="left")
        self.folder_btn = ttk.Button(btns, text="Выбрать папку", command=self.choose_folder, style="Secondary.TButton")
        self.folder_btn.pack(side="left", padx=(8, 0))
        self.run_btn = ttk.Button(btns, text="Обезличить", command=self.start_processing, style="Primary.TButton", state="disabled")
        self.run_btn.pack(side="left", padx=10)
        self.open_btn = ttk.Button(btns, text="Открыть папку результата", command=self.open_output_folder, style="Secondary.TButton", state="disabled")
        self.open_btn.pack(side="left")

        info = ttk.Frame(outer, padding=(0, 16, 0, 8))
        info.pack(fill="x")
        ttk.Label(info, text="Поддерживается: Excel .xlsx/.xlsm, CSV, текстовые и сканированные PDF (OCR RU+EN). Код Autodocs и ОЛ сохраняются. Неуверенные случаи получают жёлтый/REVIEW статус.", style="Sub.TLabel", wraplength=750).pack(anchor="w")

        self.progress = ttk.Progressbar(outer, mode="indeterminate")
        self.progress.pack(fill="x", pady=(4, 10))

        # Футер резервируется у нижней границы окна, чтобы подпись разработчика
        # оставалась видимой независимо от высоты журнала обработки.
        footer = ttk.Frame(outer)
        footer.pack(side="bottom", fill="x", pady=(12, 0))
        ttk.Label(footer, text="Разработал: Хапилин Виктор", style="Sub.TLabel").pack(side="left")
        ttk.Label(footer, text=f"v{APP_VERSION}", style="Sub.TLabel").pack(side="right")

        log_card = ttk.Frame(outer, style="Card.TFrame", padding=12)
        log_card.pack(fill="both", expand=True)
        ttk.Label(log_card, text="Ход обработки", style="Card.TLabel", font=("Segoe UI Semibold", 10)).pack(anchor="w", pady=(0, 6))
        self.log = tk.Text(log_card, height=10, wrap="word", bg="#FFFFFF", fg="#263746", relief="flat", font=("Consolas", 9), padx=4, pady=4)
        self.log.pack(fill="both", expand=True)
        self.log.configure(state="disabled")

    def log_msg(self, msg: str):
        def _write():
            self.log.configure(state="normal")
            self.log.insert("end", msg.rstrip() + "\n")
            self.log.see("end")
            self.log.configure(state="disabled")
        self.after(0, _write)

    def _set_files(self, paths):
        supported = {".xlsx", ".xlsm", ".csv", ".pdf"}
        clean = []
        seen = set()
        for p in paths:
            pp = Path(p)
            if pp.suffix.lower() not in supported or pp.stem.lower().endswith("_обезличено"):
                continue
            key = str(pp.resolve()) if pp.exists() else str(pp)
            if key not in seen:
                seen.add(key)
                clean.append(pp)
        self.file_paths = clean
        if not clean:
            self.file_label.configure(text="Файлы не выбраны")
            self.run_btn.configure(state="disabled")
            return
        label = str(clean[0]) if len(clean) == 1 else f"Выбрано файлов: {len(clean)} • {clean[0].parent}"
        self.file_label.configure(text=label)
        self.run_btn.configure(state="normal")
        self.log_msg(f"Выбрано файлов: {len(clean)}")

    def choose_files(self):
        paths = filedialog.askopenfilenames(
            title="Выберите файлы МТР",
            filetypes=[
                ("Поддерживаемые файлы", "*.xlsx *.xlsm *.csv *.pdf"),
                ("Excel", "*.xlsx *.xlsm"), ("CSV", "*.csv"), ("PDF", "*.pdf"),
            ],
        )
        if paths:
            self._set_files(paths)

    def choose_folder(self):
        folder = filedialog.askdirectory(title="Выберите папку с файлами МТР")
        if not folder:
            return
        root = Path(folder)
        paths = sorted(p for p in root.iterdir() if p.is_file())
        self._set_files(paths)

    def _set_busy(self, busy: bool):
        def _do():
            self.choose_btn.configure(state="disabled" if busy else "normal")
            self.folder_btn.configure(state="disabled" if busy else "normal")
            self.run_btn.configure(state="disabled" if busy or not self.file_paths else "normal")
            if busy:
                self.progress.start(12)
            else:
                self.progress.stop()
        self.after(0, _do)

    def start_processing(self):
        if not self.file_paths:
            return
        self._set_busy(True)
        threading.Thread(target=self._process_worker, daemon=True).start()

    def _process_one(self, src: Path):
        suffix = src.suffix.lower()
        if suffix in {".xlsx", ".xlsm"}:
            dst = src.with_name(src.stem + "_обезличено" + suffix)
            report = process_excel(src, dst, self.az, self.log_msg)
            summary = f"{src.name}: {report['rows']} строк; зелёных {report['green']}; жёлтых {report['yellow']}; изменено {report['changed']}."
        elif suffix == ".csv":
            dst = src.with_name(src.stem + "_обезличено.csv")
            report = process_csv(src, dst, self.az, self.log_msg)
            summary = f"{src.name}: {report['rows']} строк; зелёных {report['green']}; жёлтых {report['yellow']}; изменено {report['changed']}."
        elif suffix == ".pdf":
            dst = src.with_name(src.stem + "_обезличено.pdf")
            report = process_pdf(src, dst, self.az, self.log_msg)
            tag = "REVIEW" if report.get("review") else "ГОТОВО"
            summary = (f"{src.name}: {tag}; {report['pages']} стр.; удалений {report['redactions']}; "
                       f"OCR-страниц {report.get('ocr_pages', 0)}; OCR-ошибок {report.get('ocr_failed_pages', 0)}.")
        else:
            raise ValueError("Формат не поддерживается.")
        return dst, report, summary

    def _process_worker(self):
        ok = review = errors = 0
        outputs = []
        details = []
        total = len(self.file_paths)
        for idx, src in enumerate(list(self.file_paths), 1):
            try:
                self.log_msg(f"[{idx}/{total}] {src.name}")
                dst, report, summary = self._process_one(src)
                outputs.append(dst)
                if report.get("review"):
                    review += 1
                else:
                    ok += 1
                details.append(summary)
                self.log_msg(summary)
            except Exception as e:
                errors += 1
                details.append(f"{src.name}: ОШИБКА — {e}")
                self.log_msg(f"{src.name}: ОШИБКА — {e}")
                # Critical v15 behavior: one file never stops the remaining queue.
                continue

        self.last_outputs = outputs
        if outputs:
            self.after(0, lambda: self.open_btn.configure(state="normal"))
        summary = f"Пакет завершён: успешно {ok}; REVIEW {review}; ошибок {errors}; всего {total}."
        self.log_msg(summary)
        self.after(0, lambda: messagebox.showinfo(APP_TITLE, summary))
        self._set_busy(False)

    def open_output_folder(self):
        if not self.last_outputs:
            return
        folder = str(self.last_outputs[-1].parent)
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                import subprocess
                subprocess.Popen(["open", folder])
            else:
                import subprocess
                subprocess.Popen(["xdg-open", folder])
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Не удалось открыть папку:\n{e}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
