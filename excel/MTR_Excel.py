"""MTR Excel desktop application and packaged self-test entry point."""
import argparse
import json
import logging
import os
from pathlib import Path
import queue
import sys
import threading

if __package__ in (None, ''):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from excel.excel_engine import Anonymizer
from excel.file_io import SUPPORTED, process_file

APP_VERSION = '1.3 RC1'
APP_DIR = Path(os.environ.get('LOCALAPPDATA', Path.home())) / 'MTR_Excel'


def self_test(marker):
    from excel.self_test import run
    report = run()
    Path(marker).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--self-test', action='store_true')
    parser.add_argument('--marker', default='SELF_TEST_OK.json')
    parser.add_argument('--output')
    parser.add_argument('--shared', help='Единая папка базы знаний')
    parser.add_argument('--import-review', metavar='FILE', help='Вернуть проверенный Excel')
    parser.add_argument('files', nargs='*')
    args = parser.parse_args()
    if args.self_test:
        self_test(args.marker)
        return
    from excel.knowledge import KnowledgeStore
    from excel.expert_engine import ExpertAnonymizer
    store = KnowledgeStore(APP_DIR, args.shared)
    if args.import_review:
        from excel.review_io import import_review
        print(json.dumps(import_review(args.import_review, store), ensure_ascii=False))
        return
    if args.files:
        if not args.output:
            parser.error('Для командной строки требуется --output ПАПКА')
        store.sync()
        az = ExpertAnonymizer(store.snapshot())
        errors = []
        for file in args.files:
            try:
                process_file(file, args.output, az, knowledge_store=store if store.shared else None)
            except Exception as exc:
                errors.append(f'{file}: {exc}')
        store.sync()
        if errors:
            raise RuntimeError('\n'.join(errors))
        return
    gui()


def gui():
    from excel.expert_ui import gui as expert_gui
    expert_gui(APP_DIR)


if __name__ == '__main__':
    try:
        main()
    except Exception:
        APP_DIR.mkdir(parents=True, exist_ok=True)
        import traceback
        (APP_DIR / 'startup_error.log').write_text(traceback.format_exc(), encoding='utf-8')
        if '--self-test' not in sys.argv and not any(x in sys.argv for x in ('--output',)):
            try:
                from tkinter import messagebox
                messagebox.showerror('Ошибка запуска MTR Excel', f'Не удалось запустить программу. Журнал: {APP_DIR / "startup_error.log"}')
            except Exception:
                pass
        raise
