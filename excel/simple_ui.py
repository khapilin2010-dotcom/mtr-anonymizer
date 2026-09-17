"""Simple field UI: file -> automatic pass -> one row -> one engineer decision.

The normal engineer workflow deliberately stays small.  Knowledge storage,
audit and synchronization remain under the hood, but the user can export the
knowledge base to Excel and load controlled corrections back when necessary.
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
from excel.knowledge import normalize
from excel.knowledge_edit_io import export_editable, preview_editable, commit_editable
from excel.operator_engine import OperatorExpertAnonymizer
from excel.output_update import apply_decisions
from excel.simple_store import SimpleKnowledgeStore
from excel.simple_review import build_review_queue
from excel.theme import apply_theme, text_colors, MUTED
from excel.history_import_ui import show_import_dialog


def run(app_dir, default_knowledge, version='', smoke=False):
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
            if str(candidate) != '.':
                configured = candidate
        except Exception:
            pass

    root = tk.Tk()
    root.title('MTR Excel — обезличивание')
    root.geometry('1120x790')
    root.minsize(940, 700)
    apply_theme(root)

    state = {
        'store': SimpleKnowledgeStore(app_dir) if config_path.exists() else SimpleKnowledgeStore(app_dir, configured),
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

    header = ttk.Frame(root, padding=(18, 14, 18, 8), style='Header.TFrame'); header.pack(fill='x')
    ttk.Label(header, style='Header.TLabel', text='MTR Excel', font=('Segoe UI', 23, 'bold')).pack(side='left')
    ttk.Label(header, style='Header.TLabel', text=(version + ' • разработал Хапилин Виктор').strip(' •'),
              font=('Segoe UI', 9)).pack(side='right')

    body = ttk.Frame(root, padding=(18, 6, 18, 10)); body.pack(fill='both', expand=True)
    process_frame = ttk.Frame(body)
    review_frame = ttk.Frame(body)
    done_frame = ttk.Frame(body)

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
                events.put(('done', work(), done))
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
        if state['busy'] or any(state['decisions'].values()):
            messagebox.showinfo('База решений', 'Сначала завершите текущую проверку.', parent=root)
            return
        folder = filedialog.askdirectory(title='Папка базы решений', initialdir=knowledge_label.get())
        if not folder:
            return
        try:
            new_store = SimpleKnowledgeStore(app_dir, folder)
            new_store.sync(auto_backup=False)
            state['store'] = new_store
            knowledge_label.set(str(new_store.shared))
            status.set('База решений подключена. Новые решения будут сохраняться сюда.')
        except Exception as exc:
            messagebox.showerror('Не удалось подключить базу', str(exc), parent=root)

    def export_base():
        if state['busy']:
            return
        path = filedialog.asksaveasfilename(
            title='Выгрузить базу решений в Excel', defaultextension='.xlsx',
            filetypes=[('Excel', '*.xlsx')], initialfile='MTR_База_решений.xlsx')
        if not path:
            return

        def work():
            return export_editable(path, state['store'])

        def done(saved):
            status.set('База выгружена: ' + str(saved))
            messagebox.showinfo(
                'База выгружена',
                'В Excel меняйте только «Новое решение», «Действие» и «Причина изменения».\n'
                'После правки загрузите этот файл кнопкой «Загрузить исправленную базу».',
                parent=root)

        status.set('Выгружаю базу решений…')
        run_job(work, done)

    def import_base():
        if state['busy'] or any(state['decisions'].values()):
            messagebox.showinfo('База решений', 'Сначала завершите текущую проверку.', parent=root)
            return
        path = filedialog.askopenfilename(
            title='Загрузить исправленную базу', filetypes=[('Excel', '*.xlsx')])
        if not path:
            return

        def preview_work():
            return preview_editable(path, state['store'])

        def preview_done(plan):
            text = f"Изменений к применению: {plan['ready']}. Проблемных строк: {len(plan['issues'])}."
            if plan['issues']:
                text += '\nПроблемные строки будут пропущены.'
            if not plan['ready']:
                messagebox.showinfo('База решений', text, parent=root)
                return
            if not messagebox.askyesno('Применить изменения базы?', text + '\n\nПродолжить?', parent=root):
                return

            def commit_work():
                return commit_editable(path, state['store'])

            def commit_done(report):
                status.set(f"Изменений базы применено: {report['applied']}. Проблем: {len(report['issues'])}.")
                message = status.get()
                if report['issues']:
                    message += '\n\n' + '\n'.join(f'Строка {n}: {reason}' for n, reason in report['issues'][:12])
                messagebox.showinfo('База обновлена', message, parent=root)

            status.set('Применяю исправления базы…')
            run_job(commit_work, commit_done)

        status.set('Проверяю исправленную базу…')
        run_job(preview_work, preview_done)

    def import_selection():
        if state['busy'] or any(state['decisions'].values()):
            messagebox.showinfo('База решений', 'Сначала завершите текущую проверку.', parent=root)
            return
        path = filedialog.askopenfilename(
            title='Загрузить ранее обезличенную выборку',
            filetypes=[('Excel / CSV', '*.xlsx *.xlsm *.xls *.csv')])
        if path:
            report = show_import_dialog(root, state['store'], path)
            if report:
                status.set(f"База пополнена: {report['accepted']} решений. Можно обрабатывать новые файлы.")

    def open_knowledge_folder():
        folder = Path(knowledge_label.get())
        if folder.exists() and hasattr(os, 'startfile'):
            os.startfile(str(folder))

    # ---- 1. file selection --------------------------------------------
    ttk.Label(process_frame, text='1. Выберите файл', font=('Segoe UI', 17, 'bold')).pack(anchor='w')
    ttk.Label(process_frame,
              text='Программа создаст отдельный обезличенный Excel. Исходный файл останется без изменений.',
              font=('Segoe UI', 10)).pack(anchor='w', pady=(2, 10))

    listbox = tk.Listbox(process_frame, **{k: v for k, v in text_colors().items() if k != 'insertbackground'}, height=7, selectmode='extended', font=('Segoe UI', 10))
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

    kb = ttk.LabelFrame(process_frame, text='База решений инженеров', padding=9)
    kb.pack(fill='x', pady=(8, 3))
    kb_top = ttk.Frame(kb); kb_top.pack(fill='x')
    ttk.Label(kb_top, textvariable=knowledge_label, foreground=MUTED).pack(side='left', fill='x', expand=True)
    ttk.Button(kb_top, text='Сменить папку…', command=choose_knowledge).pack(side='right')
    kb_buttons = ttk.Frame(kb); kb_buttons.pack(fill='x', pady=(6, 0))
    ttk.Button(kb_buttons, text='Выгрузить базу в Excel…', command=export_base).pack(side='left')
    ttk.Button(kb_buttons, text='Загрузить исправленную базу…', command=import_base).pack(side='left', padx=6)
    ttk.Button(kb_buttons, text='Открыть папку базы', command=open_knowledge_folder).pack(side='left')

    ttk.Button(kb, text='Загрузить ранее обезличенную выборку…', command=import_selection).pack(anchor='w', pady=(8, 0))

    def build_queue(session_ids):
        return build_review_queue(state['store'], session_ids, base)

    def process_all():
        if state['busy']:
            return
        if any(state['decisions'].values()):
            flush_outputs(False, lambda _: process_all())
            return
        if not state['files']:
            messagebox.showinfo('Нет файлов', 'Добавьте хотя бы один Excel-файл.', parent=root)
            return
        destination_text = output_dir.get().strip()
        if not destination_text:
            messagebox.showinfo('Папка результата', 'Укажите папку для результата.', parent=root)
            return
        destination = Path(destination_text)
        targets = list(state['files'])
        store = state['store']

        def work():
            store.sync(auto_backup=False)
            outputs, sessions, errors = [], [], []
            for path in targets:
                try:
                    # Exact decisions take precedence, then confirmed scoped rules.
                    # Unconfirmed and disputed rules still require review.
                    az = OperatorExpertAnonymizer(store.snapshot(), auto_apply_confirmed=True)
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
                done_text.set('Итоговый файл готов. Строк, требующих проверки, нет.')
                show(done_frame)

        status.set('Обезличиваю файл…')
        run_job(work, done)

    ttk.Button(process_frame, text='Обезличить и проверить', style='Primary.TButton',
               command=process_all).pack(fill='x', pady=(12, 4))

    def latest_session():
        if state['busy']:
            return
        store = state['store']
        with store.db() as db:
            hit = db.execute('SELECT session FROM rows GROUP BY session ORDER BY max(rowid) DESC LIMIT 1').fetchone()
        if not hit:
            messagebox.showinfo('Проверка', 'Предыдущих обработок пока нет.', parent=root)
            return
        sid = hit[0]
        try:
            queue_rows = build_queue([sid])
        except Exception as exc:
            messagebox.showerror('Не удалось восстановить проверку', str(exc), parent=root)
            return
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

    # ---- 2. row review --------------------------------------------------
    top_review = ttk.Frame(review_frame); top_review.pack(fill='x')
    ttk.Label(top_review, text='2. Проверьте результат', font=('Segoe UI', 17, 'bold')).pack(side='left')
    ttk.Label(top_review, textvariable=counter, font=('Segoe UI', 11, 'bold')).pack(side='right')
    ttk.Label(review_frame, textvariable=row_info, foreground=MUTED).pack(anchor='w', pady=(3, 8))

    source_box = ttk.LabelFrame(review_frame, text='Исходное наименование', padding=7)
    source_box.pack(fill='both', expand=True, pady=4)
    source_text = tk.Text(source_box, **text_colors(), height=7, wrap='word', font=('Consolas', 10))
    source_text.pack(fill='both', expand=True)
    source_text.configure(state='disabled')

    final_box = ttk.LabelFrame(review_frame, text='Результат — его можно исправить прямо здесь', padding=7)
    final_box.pack(fill='both', expand=True, pady=4)
    final_text = tk.Text(final_box, **text_colors(), height=7, wrap='word', font=('Consolas', 10))
    final_text.pack(fill='both', expand=True)

    ttk.Label(review_frame, textvariable=removed_info, wraplength=1040,
              foreground='#8a3b12').pack(anchor='w', pady=(3, 8))

    buttons = ttk.Frame(review_frame); buttons.pack(fill='x', pady=(4, 4))
    btn_program = ttk.Button(buttons, text='✓ Оставить результат программы', style='Big.TButton')
    btn_program.pack(side='left', fill='x', expand=True, padx=(0, 4))
    btn_fix = ttk.Button(buttons, text='✎ Сохранить моё исправление', style='Big.TButton')
    btn_fix.pack(side='left', fill='x', expand=True, padx=4)
    btn_original = ttk.Button(buttons, text='↩ Оставить как в исходном', style='Big.TButton')
    btn_original.pack(side='left', fill='x', expand=True, padx=(4, 0))

    smallbar = ttk.Frame(review_frame); smallbar.pack(fill='x', pady=(4, 0))
    ttk.Button(smallbar, text='Пропустить пока', command=lambda: next_row()).pack(side='left')
    ttk.Button(smallbar, text='← Вернуться к выбору файла', command=lambda: show(process_frame)).pack(side='right')

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
        row_info.set(
            f"Код Автодокс: {code}   •   Лист: {row.get('sheet','')}   •   "
            f"Строка: {row.get('row','')}   •   Статус программы: {row.get('status','')}")
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
        if state['busy'] or not state['current']:
            return
        row = state['current']
        try:
            final = validate(row, final)
            event = state['store'].decide(row, final, action)
            sync = state['store'].sync(auto_backup=False)
            if not sync['online'] or sync.get('pending') or sync.get('errors'):
                messagebox.showwarning('Общая база недоступна', 'Решение сохранено на этом компьютере. После восстановления связи оно будет отправлено в общую базу.', parent=root)
            row['final'] = final
            state['decisions'].setdefault(row['session'], []).append(
                {'row': row, 'final': final, 'action': action})
            status.set('Решение сохранено в базе и будет приоритетным для этой же строки.'
                       if event else 'Такое решение уже есть в базе.')
            next_row()
        except Exception as exc:
            messagebox.showerror('Решение не сохранено', str(exc), parent=root)

    def current_final():
        return final_text.get('1.0', 'end-1c').strip()

    btn_program.configure(command=lambda: save_decision('Правильно', state['current']['automatic']))
    btn_fix.configure(command=lambda: save_decision('Исправить', current_final()))
    btn_original.configure(command=lambda: save_decision('Оставить как в исходном', state['current']['source']))

    def next_row():
        if state['busy']:
            return
        state['index'] += 1
        render_current()

    savebar = ttk.Frame(review_frame); savebar.pack(fill='x', pady=(8, 0))
    ttk.Button(savebar, text='Сохранить принятые решения в Excel сейчас',
               command=lambda: flush_outputs(True)).pack(side='right')

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
            for sid, saved in decisions.items():
                state['decisions'][sid] = [item for item in state['decisions'][sid] if item not in saved]
            if show_message:
                messagebox.showinfo('Готово', 'Исправления записаны в итоговый Excel.', parent=root)
            if callback:
                callback(paths)

        status.set('Записываю решения в итоговый Excel…')
        run_job(work, done)

    def finish_review():
        def after(_paths):
            try:
                state['store'].sync(auto_backup=False)
            except Exception:
                pass
            remaining = sum('final' not in row for row in state['queue'])
            if remaining:
                done_text.set(f'Итоговый Excel сохранён. Осталось проверить пропущенных строк: {remaining}. Нажмите «Продолжить проверку».')
                status.set('Проверка ещё не завершена.')
            else:
                done_text.set('Проверка завершена. Решения сохранены в базе, итоговый Excel обновлён.')
                status.set('Готово. Итоговый Excel можно использовать.')
            show(done_frame)

        flush_outputs(False, after)

    # ---- 3. done --------------------------------------------------------
    done_text = tk.StringVar(value='Готово.')
    ttk.Label(done_frame, text='3. Готово', font=('Segoe UI', 19, 'bold')).pack(anchor='w', pady=(10, 8))
    ttk.Label(done_frame, textvariable=done_text, wraplength=900, font=('Segoe UI', 11)).pack(anchor='w', pady=(0, 18))

    def open_results():
        folder = Path(output_dir.get())
        if folder.exists() and hasattr(os, 'startfile'):
            os.startfile(str(folder))

    ttk.Button(done_frame, text='Открыть папку с результатами', style='Primary.TButton',
               command=open_results).pack(fill='x', pady=4)

    def new_run():
        state['files'].clear()
        listbox.delete(0, 'end')
        state['queue'] = []
        state['index'] = 0
        state['decisions'] = {}
        state['current'] = None
        status.set('Добавьте следующий Excel-файл.')
        show(process_frame)

    ttk.Button(done_frame, text='Продолжить проверку', command=latest_session).pack(fill='x', pady=4)
    ttk.Button(done_frame, text='Обработать ещё один файл', command=new_run).pack(fill='x', pady=4)
    ttk.Button(done_frame, text='Выгрузить базу решений в Excel…', command=export_base).pack(fill='x', pady=4)

    def on_close():
        if state['busy']:
            messagebox.showinfo('Операция выполняется', 'Дождитесь окончания записи или обработки.', parent=root)
            return
        if any(state['decisions'].values()) and not state['busy']:
            if messagebox.askyesno('Сохранить результат?',
                                   'Есть решения, ещё не записанные в итоговый Excel. Сохранить их перед выходом?',
                                   parent=root):
                flush_outputs(False, lambda _: root.destroy())
                return
        root.destroy()

    root.protocol('WM_DELETE_WINDOW', on_close)
    show(process_frame)
    root.after(100, poll)
    if smoke:
        for frame in (process_frame, review_frame, done_frame):
            show(frame)
            root.update()
        for timer in root.tk.call('after', 'info'):
            root.after_cancel(timer)
        root.destroy()
        return True
    root.mainloop()
