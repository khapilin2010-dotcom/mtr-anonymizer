"""Capture the real Windows Tk screens for release review (not a UI mockup)."""
from pathlib import Path
from contextlib import ExitStack
import logging
import gc
import sys
import tempfile
import time
import tkinter as tk
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from PIL import ImageGrab
from openpyxl import Workbook
from excel.simple_ui import run
from excel.simple_store import SimpleKnowledgeStore
from excel.history_import_ui import SelectionDialog
from excel.theme import apply_theme
from excel.MTR_Excel import APP_VERSION
from excel.file_io import process_file
from excel.operator_engine import OperatorExpertAnonymizer
from excel.semantics import features


def capture(widget, destination):
    widget.lift(); widget.update()
    time.sleep(0.3); widget.update()
    assert widget.winfo_ismapped() and widget.winfo_width() > 800 and widget.winfo_height() > 600
    def check_buttons(parent):
        for child in parent.winfo_children():
            if child.winfo_ismapped() and child.winfo_class() in ('TButton', 'TCheckbutton'):
                assert child.winfo_height() >= child.winfo_reqheight(), str(child)
                assert child.winfo_rooty() + child.winfo_height() <= widget.winfo_rooty() + widget.winfo_height(), str(child)
                assert child.winfo_rootx() + child.winfo_width() <= widget.winfo_rootx() + widget.winfo_width(), str(child)
            check_buttons(child)
    check_buttons(widget)
    x, y = widget.winfo_rootx(), widget.winfo_rooty()
    image = ImageGrab.grab(bbox=(x, y, x + widget.winfo_width(), y + widget.winfo_height()))
    image.save(destination)


def main():
    output = Path('ui-preview'); output.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='MTR_UI_') as tmp, ExitStack() as cleanup:
        cleanup.callback(logging.shutdown)
        folder = Path(tmp)
        # Populate a real session so the review capture uses production widgets,
        # queue restoration and learning, not a separate demonstration layout.
        store = SimpleKnowledgeStore(folder / 'local', folder / 'База решений', 'Инженер')
        known = 'Панель ПКМ-ТСТ-1'
        store.decide(dict(id='demo', session='demo', code='DEMO-1', source=known,
                          automatic='Панель', factory='', classification=features(known)), known,
                     'Оставить как в исходном')
        path = folder / 'Пример проверки.xlsx'; wb = Workbook(); ws = wb.active
        ws.append(['Код Автодокс', 'Наименование', 'Завод'])
        ws.append(['DEMO-3', 'Панель ПКМ-ТСТ-3 ООО «Тестовый завод» IP66', ''])
        wb.save(path); wb.close()
        process_file(path, folder, OperatorExpertAnonymizer(store.snapshot(), auto_apply_confirmed=True), knowledge_store=store)
        def descendants(widget):
            for child in widget.winfo_children():
                yield child
                yield from descendants(child)
        def capture_start(root):
            capture(root, output / 'start.png')
            resume = next(w for w in descendants(root) if w.winfo_class() == 'TButton' and w.cget('text') == 'Продолжить последнюю проверку')
            resume.invoke(); root.update()
            texts = [w for w in descendants(root) if w.winfo_class() == 'Text' and w.winfo_ismapped()]
            original = next(w for w in texts if w.cget('state') == 'disabled')
            final = next(w for w in texts if w.cget('state') == 'normal')
            assert original.tag_ranges('deleted') and final.tag_ranges('restored')
            final.insert('end-1c', ' в комплекте')
            root.update(); time.sleep(.15); root.update()
            assert final.tag_ranges('added')
            capture(root, output / 'review.png')
            # Exercise the real selection buttons; preserve the manual suffix.
            a = original.get('1.0', 'end-1c').index('Тестовый')
            original.tag_add('sel', f'1.0+{a}c', f'1.0+{a+len('Тестовый завод')}c')
            keep = next(w for w in descendants(root) if w.winfo_class() == 'TButton' and w.cget('text') == 'Оставить выделенное')
            keep.invoke(); root.update()
            restored = final.get('1.0', 'end-1c')
            assert 'Тестовый завод' in restored and 'ООО' not in restored and restored.endswith('в комплекте'), restored
            keep_all = next(w for w in descendants(root) if w.winfo_class() == 'TButton' and w.cget('text') == 'Оставить всё удалённое')
            keep_all.invoke(); root.update()
            assert final.get('1.0', 'end-1c') == original.get('1.0', 'end-1c') + ' в комплекте'
            assert not original.tag_ranges('deleted')
            capture(root, output / 'review-restored.png')
            for timer in root.tk.call('after', 'info'):
                root.after_cancel(timer)
            root.destroy()
        with patch.object(tk.Tk, 'mainloop', capture_start):
            run(folder / 'local', folder / 'База решений', APP_VERSION)
        # Release Tk cycles on the UI thread before the next dialog starts its worker.
        gc.collect()
        store = SimpleKnowledgeStore(folder / 'import-local', folder / 'База решений', 'Инженер')
        path = folder / 'Проверенная выборка.xlsx'; wb = Workbook(); ws = wb.active
        ws.append(['Код Автодокс', 'Исходное наименование', 'Обезличенное наименование', 'Завод'])
        source = 'Модуль автоматизированной технологической обвязки скважин МОС-3/1 Чертеж МОС-29.М1'
        ws.append(['631-336536', source, source, ''])
        ws.append(['000002', 'Клапан DN50 PN16', 'Клапан DN50 PN16', ''])
        wb.save(path); wb.close()
        root = tk.Tk(); root.geometry('1x1+0+0'); apply_theme(root); root.update()
        dialog = SelectionDialog(root, store, path)
        def idle():
            deadline = time.monotonic() + 20
            while dialog.busy and time.monotonic() < deadline:
                root.update(); time.sleep(0.02)
            assert not dialog.busy
            root.update()
        idle(); dialog.preview(); idle()
        capture(dialog.window, output / 'selection-import.png')
        dialog.close(); root.destroy()


if __name__ == '__main__':
    main()
