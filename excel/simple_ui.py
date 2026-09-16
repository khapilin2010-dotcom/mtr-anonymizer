"""Simple field UI: file -> automatic pass -> one row -> one engineer decision.

The knowledge database stays under the hood.  Every explicit decision is saved
immediately, while the generated workbook is updated in one batch when the
review is finished (or when the user explicitly saves current progress).
"""
import json
import logging
import os
from pathlib import Path
import queue
import threading

from excel.excel_engine import Anonymizer
from excel.expert_engine import protection_losses
from excel.file_io import SUPPORTED, process_file
from excel.knowledge import case_key, normalize
from excel.operator_engine import OperatorExpertAnonymizer
from excel.output_update import apply_decisions
from excel.simple_store import SimpleKnowledgeStore


def run(app_dir, default_knowledge, version=''):
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    app_dir = Path(app_dir)
    default_knowledge = Path(default_knowledge)
    app_dir.mkdir(parents=True, exist_ok=True)
    default_knowledge.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(filename=app_dir / 'mtr_excel.log', level=logging.INFO, encoding='utf-8',
                        format='%(asctime)s %(levelname)s %(message)s')

    configured = default_knowledge
    config_path = app_dir / 'knowledge_config.json'
    if config_path.exists():
        try:
            candidate = Path(json.loads(config_path.read_text('utf-8')).get('folder', ''))
            if candidate.exists():
                configured = candidate
        except Exception:
            pass

    root = tk.Tk()
    root.title('MTR Excel — обезличивание')
    root.geometry('1120x780')
    root.minsize(940, 690)
    style = ttk.Style(root)
    style.theme_use('clam')
    style.configure('Big.TButton', padding=(14, 11), font=('Segoe UI', 10, 'bold'))
    style.configure('Primary.TButton', padding=(16, 12), font=('Segoe UI', 11, 'bold'))

    state = {
        'store': SimpleKnowledgeStore(app_dir, configured),
        'files': [], 'outputs': [], 'queue': [], 'index': 0,
        'decisions': {}, 'busy': False, 'current': None,
    }
    events = queue.Queue()
    cancel = threading.Event()
    base = Anonymizer()

    status = tk.StringVar(value='Добавьте Excel-файл и нажмите «Обезличить и проверить».')
    knowledge_label = tk.StringVar(value=str(state['store'].shared or configured))
    output_dir = tk.StringVar(value=str(Path.home() / 'Documents' / 'MTR_Excel_Результаты'))
    counter = tk.StringVar(value='')
    row_info = tk.StringVar(value='')
    removed_info = tk.StringVar(value='')

    header = ttk.Frame(root, padding=(18, 14, 18, 8)); header.pack(fill='x')
    ttk.Label(header, text='MTR Excel', font=('Segoe UI', 23, 'bold')).pack(side='left')
    ttk.Label(header, text=(version + ' • разработал Хапилин Виктор').strip(' •'),
              font=('Segoe UI', 9)).pack(side='right')

    body = ttk.Frame(root, padding=(18, 6, 18, 10)); body.pack(fill='both', expand=True)
    process_frame = ttk.Frame(body)
    review_frame = ttk.Frame(body)
    done_frame = ttk.Frame(body)

    # ---- common helpers -------------------------------------------------
    progress = ttk.Progressbar(root, mode='indeterminate')
    progress.pack(fill='x', padx=18)
    ttk.Label(root, textvariable=status, wraplength=1060).pack(fill='x', padx=18, pady=(4, 10))

    def show(frame):
        for item in (process_frame, review_frame, done_frame):
            item.pack_forget()
        frame.pack(fill='both', expand=True)

    def run_job(work, done=None):
        if state['busy']:
            return
        state['busy'] = True
        cancel.clear()
        progress.configure(mode='indeterminate')
        progress.start(12)
        def worker():
            try:
                result = work()
                events.put(('done', result, done))
            except Exception as exc:
                logging.exception('Background operation failed')
                events.put(('error', str(exc), None))
        threading.Thread(target=worker, daemon=True).start()

    def poll():
        try:
            while True:
                kind, value, callback = events.get_nowait()
                state['busy'] = False
                progress.stop()
                if kind == 'error':
                    status.set(value)
                    messagebox.showerror('Ошибка', value, parent=root)
                elif callback:
                    callback(value)
        except queue.Empty:
            pass
        root.after(100, poll)

    def choose_knowledge():
        folder = filedialog.askdirectory(title='Папка базы знаний', initialdir=knowledge_label.get())
        if not folder:
            return
        try:
            new_store = SimpleKnowledgeStore(app_dir, folder)
            new_store.sync(auto_backup=False)
            state['store'] = new_store
            knowledge_label.set(str(new_store.shared))
            status.set('База знаний подключена. Все новые решения будут сохраняться сюда.')
        except Exception as exc:
            messagebox.showerror('Не удалось подключить базу', str(exc), parent=root)

    # ---- process screen -------------------------------------------------
    ttk.Label(process_frame, text='1. Выберите файл', font=('Segoe UI', 17, 'bold')).pack(anchor='w')
    ttk.Label(process_frame,
              text='Программа создаст отдельный обезличенный Excel. Исходный файл останется без изменений.',
              font=('Segoe UI', 10)).pack(anchor='w', pady=(2, 10))

    listbox = tk.Listbox(process_frame, height=9, selectmode='extended', font=('Segoe UI', 10))
    listbox.pack(fill='both', expand=True)

    filebar = ttk.Frame(process_frame); filebar.pack(fill='x', pady=7)

    def add_files(paths):
        for raw in paths:
            path = Path(raw)
            if path.suffix.lower() in SUPPORTED and path not in state['files'] and not path.name.startswith('~$'):
                state['files'].append(path)
                listbox.insert('end', str(path))

    ttk.Button(filebar, text='Добавить Excel…', command=lambda: add_files(
        filedialog.askopenfilenames(filetypes=[('Excel / CSV', '*.xlsx *.xlsm *.xls *.csv')]))).pack(side='left')

    def remove_selected():
        for index in reversed(listbox.curselection()):
            state['files'].pop(index)
            listbox.delete(index)
    ttk.Button(filebar, text='Убрать выбранное', command=remove_selected).pack(side='left', padx=6)

    location = ttk.LabelFrame(process_frame, text='Куда сохранить результат', padding=9)
    location.pack(fill='x', pady=(8, 4))
    ttk.Entry(location, textvariable=output_dir).pack(side='left', fill='x', expand=True)
    ttk.Button(location, text='Выбрать…', command=lambda: output_dir.set(
        filedialog.askdirectory(title='Папка результата') or output_dir.get())).pack(side='left', padx=(6, 0))

    kb = ttk.Frame(process_frame); kb.pack(fill='x', pady=(8, 3))
    ttk.Label(kb, text='База решений:').pack(side='left')
    ttk.Label(kb, textvariable=knowledge_label, foreground='#555555').pack(side='left', padx=5)
    ttk.Button(kb, text='Сменить…', command=choose_knowledge).pack(side='right')

    def build_queue(session_ids):
        store = state['store']
        store.sync(auto_backup=False)
        snapshot = store.snapshot()
        rows = []
        for sid in session_ids:
            page = store.session_rows(sid, '', 0, 1000000)
            for row in page:
                entry = snapshot.entries.get(case_key(row.get('code', ''), row.get('source', ''), row.get('factory', '')))
                # Exact engineer decisions are already final knowledge and do not
                # need to annoy the engineer again on every run.
                if entry and entry.get('kind') == 'case' and entry.get('status') in ('ACTIVE', 'TRUSTED'):
                    continue
                changed = normalize(row.get('source', '')) != normalize(row.get('automatic', ''))
                if changed or row.get('status') != 'ЗЕЛЁНЫЙ':
                    rows.append(row)
        return rows

    def process_all():
        if not state['files']:
            messagebox.showinfo('Нет файлов', 'Добавьте хотя бы один Excel-файл.', parent=root)
            return
        destination = Path(output_dir.get().strip())
        if not str(destination):
            messagebox.showinfo('Папка результата', 'Укажите папку для результата.', parent=root)
            return
        targets = list(state['files'])
        store = state['store']
        def work():
            store.sync(auto_backup=False)
            outputs, sessions, errors = [], [], []
            for path in targets:
                try:
                    az = OperatorExpertAnonymizer(store.snapshot(), auto_apply_confirmed=False)
                    dst, report = process_file(path, destination, az, knowledge_store=store, cancel=cancel)
                    outputs.append(str(dst))
                    if report.get('session'):
                        sessions.append(report['session'])
                except Exception as exc:
                    errors.append(path.name + ': ' + str(exc))
            store.sync(auto_backup=False)
            return outputs, sessions, errors, build_queue(sessions)
        def done(result):
            outputs, sessions, errors, rows = result
            state['outputs'] = outputs
            state['queue'] = rows
            state['index'] = 0
            state['decisions'] = {}
            if errors:
                messagebox.showwarning('Часть файлов не обработана', '\n'.join(errors[:15]), parent=root)
            if not outputs:
                status.set('Не удалось обработать файлы.')
                return
            if rows:
                status.set(f'Обработка завершена. На проверку: {len(rows)} строк.')
                show(review_frame)
                render_current()
            else:
                status.set('Готово. Строк, требующих проверки, нет.')
                show(done_frame)
                done_text.set('Проверка не потребовалась. Итоговые файлы готовы.')
        status.set('Обезличиваю файл…')
        run_job(work, done)

    ttk.Button(process_frame, text='Обезличить и проверить', style='Primary.TButton',
               command=process_all).pack(fill='x', pady=(12, 4))

    def latest_session():
        store = state['store']
        with store.db() as db:
            row = db.execute('SELECT session FROM rows GROUP BY session ORDER BY max(rowid) DESC LIMIT 1').fetchone()
        if not row:
            messagebox.showinfo('Проверка', 'Предыдущих обработок пока нет.', parent=root)
            return
        sid = row[0]
        queue_rows = build_queue([sid])
        state['queue'] = queue_rows
        state['index'] = 0
        state['decisions'] = {}
        meta = store.metadata(sid)
        output = meta.get('output_file')
        state['outputs'] = [output] if output else []
        if not queue_rows:
            messagebox.showinfo('Проверка', 'В последней обработке нет непринятых решений.', parent=root)
            return
        show(review_frame)
        render_current()
    ttk.Button(process_frame, text='Продолжить последнюю проверку', command=latest_session).pack(fill='x')

    # ---- review screen --------------------------------------------------
    top_review = ttk.Frame(review_frame); top_review.pack(fill='x')
    ttk.Label(top_review, text='2. Проверьте результат', font=('Segoe UI', 17, 'bold')).pack(side='left')
    ttk.Label(top_review, textvariable=counter, font=('Segoe UI', 11, 'bold')).pack(side='right')
    ttk.Label(review_frame, textvariable=row_info, foreground='#555555').pack(anchor='w', pady=(3, 8))

    source_box = ttk.LabelFrame(review_frame, text='Исходное наименование', padding=7)
    source_box.pack(fill='both', expand=True, pady=4)
    source_text = tk.Text(source_box, height=7, wrap='word', font=('Consolas', 10))
    source_text.pack(fill='both', expand=True)
    source_text.configure(state='disabled')

    final_box = ttk.LabelFrame(review_frame, text='Результат — его можно исправить прямо здесь', padding=7)
    final_box.pack(fill='both', expand=True, pady=4)
    final_text = tk.Text(final_box, height=7, wrap='word', font=('Consolas', 10))
    final_text.pack(fill='both', expand=True)

    ttk.Label(review_frame, textvariable=removed_info, wraplength=1040, foreground='#8a3b12').pack(anchor='w', pady=(3, 8))

    buttons = ttk.Frame(review_frame); buttons.pack(fill='x', pady=(4, 4))
    btn_program = ttk.Button(buttons, text='✓ Оставить результат программы', style='Big.TButton')
    btn_program.pack(side='left', fill='x', expand=True, padx=(0, 4))
    btn_fix = ttk.Button(buttons, text='✎ Сохранить моё исправление', style='Big.TButton')
    btn_fix.pack(side='left', fill='x', expand=True, padx=4)
    btn_original = ttk.Button(buttons, text='↩ Оставить как в исходном', style='Big.TButton')
    btn_original.pack(side='left', fill='x', expand=True, padx=(4, 0))

    smallbar = ttk.Frame(review_frame); smallbar.pack(fill='x', pady=(4, 0))
    ttk.Button(smallbar, text='Пропустить пока', command=lambda: next_row()).pack(side='left')

    def set_text(widget, value, disabled=False):
        widget.configure(state='normal')
        widget.delete('1.0', 'end')
        widget.insert('1.0', value)
        if disabled:
            widget.configure(state='disabled')

    def render_current():
        if not state['queue'] or state['index'] >= len(state['queue']):
            finish_review()
            return
        row = state['queue'][state['index']]
        state['current'] = row
        counter.set(f"{state['index'] + 1} из {len(state['queue'])}")
        code = row.get('code', '') or 'без кода'
        row_info.set(f"Код Автодокс: {code}   •   Лист: {row.get('sheet','')}   •   Строка: {row.get('row','')}   •   Статус программы: {row.get('status','')}")
        set_text(source_text, row.get('source', ''), True)
        set_text(final_text, row.get('automatic', ''))
        removed = row.get('removed') or []
        removed_info.set('Программа удалила: ' + ('; '.join(map(str, removed)) if removed else 'ничего'))
        final_text.focus_set()

    def validate(row, final):
        final = str(final).strip()
        if not final:
            raise ValueError('Итоговое наименование не может быть пустым.')
        losses = protection_losses(row.get('source', ''), final, base,
                                   row.get('code', ''), row.get('factory', ''))
        if losses:
            raise ValueError('Нельзя сохранить: потеряны защищённые технические признаки: ' + ', '.join(losses))
        return final

    def save_decision(action, final):
        row = state['current']
        try:
            final = validate(row, final)
            event = state['store'].decide(row, final, action)
            state['store'].sync(auto_backup=False)
            state['decisions'].setdefault(row['session'], []).append(
                {'row': row, 'final': final, 'action': action})
            status.set('Решение сохранено в базе.' if event else 'Такое решение уже было в базе.')
            next_row()
        except Exception as exc:
            messagebox.showerror('Решение не сохранено', str(exc), parent=root)

    def current_final():
        return final_text.get('1.0', 'end-1c').strip()

    btn_program.configure(command=lambda: save_decision('Правильно', state['current']['automatic']))
    btn_fix.configure(command=lambda: save_decision('Исправить', current_final()))
    btn_original.configure(command=lambda: save_decision('Оставить как в исходном', state['current']['source']))

    def next_row():
        state['index'] += 1
        render_current()

    savebar = ttk.Frame(review_frame); savebar.pack(fill='x', pady=(8, 0))
    ttk.Button(savebar, text='Сохранить принятые решения в Excel сейчас', command=lambda: flush_outputs(False)).pack(side='right')

    def flush_outputs(show_message=True, callback=None):
        decisions = {sid: list(items) for sid, items in state['decisions'].items() if items}
        if not decisions:
            if callback:
                callback([])
            elif show_message:
                messagebox.showinfo('Сохранение', 'Новых решений для записи в Excel нет.', parent=root)
            return
        store = state['store']
        def work():
            updated = []
            for sid, items in decisions.items():
                path = store.metadata(sid).get('output_file')
                if not path:
                    raise ValueError('Не найден итоговый файл для сеанса ' + sid)
                apply_decisions(path, items)
                updated.append(path)
            return updated
        def done(paths):
            for sid in decisions:
                state['decisions'][sid] = []
            if show_message:
                messagebox.showinfo('Готово', 'Исправления записаны в итоговый Excel.', parent=root)
            if callback:
                callback(paths)
        status.set('Записываю решения в итоговый Excel…')
        run_job(work, done)

    def finish_review():
        def after(_paths):
            try:
                state['store'].sync()
            except Exception:
                pass
            done_text.set('Проверка завершена. Решения сохранены в базе, итоговый Excel обновлён.')
            status.set('Готово. Итоговый Excel можно использовать.')
            show(done_frame)
        flush_outputs(False, after)
        if not any(state['decisions'].values()):
            after([])

    # ---- done screen ----------------------------------------------------
    done_text = tk.StringVar(value='Готово.')
    ttk.Label(done_frame, text='3. Готово', font=('Segoe UI', 19, 'bold')).pack(anchor='w', pady=(10, 8))
    ttk.Label(done_frame, textvariable=done_text, wraplength=900, font=('Segoe UI', 11)).pack(anchor='w', pady=(0, 18))

    def open_results():
        folder = Path(output_dir.get())
        if folder.exists() and hasattr(os, 'startfile'):
            os.startfile(str(folder))
    ttk.Button(done_frame, text='Открыть папку с результатами', style='Primary.TButton', command=open_results).pack(fill='x', pady=4)

    def new_run():
        state['files'].clear(); listbox.delete(0, 'end')
        state['queue'] = []; state['index'] = 0; state['decisions'] = {}; state['current'] = None
        status.set('Добавьте следующий Excel-файл.')
        show(process_frame)
    ttk.Button(done_frame, text='Обработать ещё один файл', command=new_run).pack(fill='x', pady=4)

    # Closing with unsaved workbook decisions should not silently lose the
    # visible result, although the knowledge events are already durable.
    def on_close():
        if any(state['decisions'].values()) and not state['busy']:
            if messagebox.askyesno('Сохранить результат?',
                                   'Есть решения, ещё не записанные в итоговый Excel. Сохранить их перед выходом?', parent=root):
                def finish_close(_): root.destroy()
                flush_outputs(False, finish_close)
                return
        root.destroy()

    root.protocol('WM_DELETE_WINDOW', on_close)
    show(process_frame)
    root.after(100, poll)
    root.mainloop()
