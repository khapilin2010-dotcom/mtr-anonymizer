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

APP_VERSION = '1.2 RC3'
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
    parser.add_argument('files', nargs='*')
    args = parser.parse_args()
    if args.self_test:
        self_test(args.marker)
        return
    if args.files:
        if not args.output:
            parser.error('Для командной строки требуется --output ПАПКА')
        az = Anonymizer()
        errors = []
        for file in args.files:
            try:
                process_file(file, args.output, az)
            except Exception as exc:
                errors.append(f'{file}: {exc}')
        if errors:
            raise RuntimeError('\n'.join(errors))
        return
    gui()


def gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    APP_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(filename=APP_DIR / 'mtr_excel.log', level=logging.INFO,
                        encoding='utf-8', format='%(asctime)s %(levelname)s %(message)s')
    root = tk.Tk()
    root.title('Обезличивание МТР — MTR Excel')
    root.geometry('880x640')
    root.minsize(720, 520)
    root.configure(bg='#F1F6FC')
    style = ttk.Style(root)
    style.theme_use('clam')
    style.configure('TButton', padding=9, font=('Segoe UI', 10))
    style.configure('TLabel', background='#F1F6FC', font=('Segoe UI', 10))
    style.configure('TFrame', background='#F1F6FC')
    frame = ttk.Frame(root, padding=22)
    frame.pack(fill='both', expand=True)
    ttk.Label(frame, text='MTR Excel', font=('Segoe UI', 25, 'bold'), foreground='#1767A6').pack(anchor='w')
    ttk.Label(frame, text='Обезличивание МТР • XLSX / XLSM / XLS / CSV').pack(anchor='w', pady=(0, 12))
    ttk.Label(frame, text='Результат — отдельный файл с четырьмя столбцами проверки.').pack(anchor='w')
    files = []
    events = queue.Queue()
    busy = False
    listbox = tk.Listbox(frame, height=9, font=('Segoe UI', 10), selectmode='extended')
    listbox.pack(fill='both', expand=True, pady=12)
    output = tk.StringVar(value=str(Path.home() / 'Documents' / 'MTR_Excel_Результаты'))
    status = tk.StringVar(value='Добавьте файлы или папку.')
    buttons = []

    def add(paths):
        for path in paths:
            p = Path(path)
            if p.suffix.lower() in SUPPORTED and not p.name.startswith('~$') and p not in files:
                files.append(p)
                listbox.insert('end', str(p))

    def add_folder():
        folder = filedialog.askdirectory()
        if folder:
            add(sorted(p for p in Path(folder).rglob('*') if p.is_file()))

    def choose_output():
        folder = filedialog.askdirectory()
        if folder:
            output.set(folder)

    def remove():
        for i in reversed(listbox.curselection()):
            del files[i]
            listbox.delete(i)

    bar = ttk.Frame(frame)
    bar.pack(fill='x')
    for title, command in (
        ('Добавить файлы', lambda: add(filedialog.askopenfilenames(filetypes=[('Excel и CSV', '*.xlsx *.xlsm *.xls *.csv')]))),
        ('Добавить папку', add_folder), ('Убрать выбранные', remove)):
        b = ttk.Button(bar, text=title, command=command)
        b.pack(side='left', padx=(0, 8))
        buttons.append(b)
    location = ttk.Frame(frame)
    location.pack(fill='x', pady=12)
    ttk.Label(location, text='Папка результата:').pack(anchor='w')
    entry = ttk.Entry(location, textvariable=output)
    entry.pack(side='left', fill='x', expand=True)
    b = ttk.Button(location, text='Выбрать', command=choose_output)
    b.pack(side='right')
    buttons.append(b)
    progress = ttk.Progressbar(frame, mode='determinate')
    progress.pack(fill='x', pady=6)
    ttk.Label(frame, textvariable=status, wraplength=800).pack(anchor='w')

    def start():
        nonlocal busy
        if not files or not output.get().strip():
            messagebox.showerror('Нужны файлы', 'Добавьте файлы и укажите папку результата.')
            return
        targets, folder = list(files), Path(output.get())
        busy = True
        for button in buttons:
            button.configure(state='disabled')
        entry.configure(state='disabled')
        progress.configure(maximum=len(targets), value=0)
        status.set('Загрузка базы производителей…')

        def worker():
            success, failures, review = [], [], 0
            try:
                az = Anonymizer()
                for i, src in enumerate(targets):
                    try:
                        dst, report = process_file(src, folder, az, lambda msg: events.put(('status', msg)))
                        success.append(str(dst))
                        review += report['yellow'] + report['skipped_formulas']
                        logging.info('Processed %s: %s', src.name, report)
                    except Exception as exc:
                        failures.append(f'{src.name}: {exc}')
                        logging.exception('File failed: %s', src.name)
                    events.put(('progress', i + 1))
            except Exception as exc:
                failures.append(str(exc))
                logging.exception('Batch failed')
            events.put(('done', (success, failures, review)))
        threading.Thread(target=worker, daemon=True).start()

    def poll():
        nonlocal busy
        try:
            while True:
                kind, value = events.get_nowait()
                if kind == 'status':
                    status.set(value)
                elif kind == 'progress':
                    progress.configure(value=value)
                elif kind == 'done':
                    busy = False
                    for button in buttons:
                        button.configure(state='normal')
                    entry.configure(state='normal')
                    success, failures, review = value
                    text = f'Обработано файлов: {len(success)}. Ошибок: {len(failures)}. Строк для проверки: {review}.'
                    status.set(text)
                    if failures:
                        messagebox.showwarning('Обработка завершена с ошибками', text + '\n\n' + '\n'.join(failures))
                    else:
                        messagebox.showinfo('Обработка завершена', text)
        except queue.Empty:
            pass
        root.after(100, poll)

    actions = ttk.Frame(frame)
    actions.pack(fill='x', pady=12)
    b = ttk.Button(actions, text='Обезличить Excel', command=start)
    b.pack(side='left')
    buttons.append(b)

    def open_results():
        path = Path(output.get())
        if path.is_dir():
            os.startfile(str(path))

    ttk.Button(actions, text='Открыть результаты', command=open_results).pack(side='left', padx=8)
    ttk.Label(frame, text=f'разработал Хапилин Виктор • {APP_VERSION}', foreground='#52677F').pack(anchor='e')

    def close():
        if busy:
            messagebox.showinfo('Идёт обработка', 'Дождитесь окончания обработки файлов.')
        else:
            root.destroy()
    root.protocol('WM_DELETE_WINDOW', close)
    poll()
    root.mainloop()


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
