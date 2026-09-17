"""Capture the real Windows Tk screens for release review (not a UI mockup)."""
from pathlib import Path
from contextlib import ExitStack
import logging
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


def capture(widget, destination):
    widget.lift(); widget.update()
    widget.after(250, lambda: None)
    time.sleep(0.3); widget.update()
    x, y = widget.winfo_rootx(), widget.winfo_rooty()
    image = ImageGrab.grab(bbox=(x, y, x + widget.winfo_width(), y + widget.winfo_height()))
    image.save(destination)


def main():
    output = Path('ui-preview'); output.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='MTR_UI_') as tmp, ExitStack() as cleanup:
        cleanup.callback(logging.shutdown)
        folder = Path(tmp)
        def capture_start(root):
            capture(root, output / 'start.png')
            for timer in root.tk.call('after', 'info'):
                root.after_cancel(timer)
            root.destroy()
        with patch.object(tk.Tk, 'mainloop', capture_start):
            run(folder / 'local', folder / 'База решений', APP_VERSION)
        store = SimpleKnowledgeStore(folder / 'import-local', folder / 'База решений', 'Инженер')
        path = folder / 'Проверенная выборка.xlsx'; wb = Workbook(); ws = wb.active
        ws.append(['Код Автодокс', 'Исходное наименование', 'Обезличенное наименование', 'Завод'])
        source = 'Модуль автоматизированной технологической обвязки скважин МОС-3/1 Чертеж МОС-29.М1'
        ws.append(['631-336536', source, source, ''])
        ws.append(['000002', 'Клапан DN50 PN16', 'Клапан DN50 PN16', ''])
        wb.save(path); wb.close()
        root = tk.Tk(); root.withdraw(); apply_theme(root)
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
