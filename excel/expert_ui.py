"""Desktop review and shared-knowledge workflow; all IO jobs run off Tk thread."""
import json
import logging
import os
from pathlib import Path
import queue
import threading

from excel.knowledge import KnowledgeStore, VERSION, case_key, canonical, restore_backup
from excel.expert_engine import ExpertAnonymizer, protection_losses
from excel.file_io import SUPPORTED, process_file
from excel.review_io import import_review, export_knowledge


def gui(app_dir, smoke=False):
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, simpledialog
    app_dir.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(filename=app_dir / 'mtr_excel.log', level=logging.INFO, encoding='utf-8',
                        format='%(asctime)s %(levelname)s %(message)s')
    logging.info('Startup %s', VERSION)
    root = tk.Tk(); root.title('MTR Excel — общая база знаний'); root.geometry('1200x850'); root.minsize(980, 720)
    style = ttk.Style(root); style.theme_use('clam'); style.configure('TButton', padding=6)
    store = KnowledgeStore(app_dir)
    state = {'busy': False, 'store': store, 'rows': {}, 'selected': None, 'offset': 0, 'knowledge': None}
    events = queue.Queue(); cancel = threading.Event(); buttons = []
    status = tk.StringVar(value='Укажите общую папку подразделения, затем добавьте файлы.')
    shared = tk.StringVar(value=str(store.shared or 'Общая папка не настроена'))
    sync_status = tk.StringVar(value='База ещё не синхронизирована')
    output = tk.StringVar(value=str(Path.home() / 'Documents' / 'MTR_Excel_Результаты'))
    top = ttk.Frame(root, padding=12); top.pack(fill='x')
    ttk.Label(top, text='MTR Excel', font=('Segoe UI', 22, 'bold')).pack(side='left')
    ttk.Label(top, text=VERSION + ' • разработал Хапилин Виктор • ' + store.user).pack(side='right')
    settings = ttk.Frame(root, padding=(12, 0)); settings.pack(fill='x')
    ttk.Label(settings, textvariable=shared).pack(anchor='w')
    ttk.Label(settings, textvariable=sync_status).pack(anchor='w')
    setting_buttons = ttk.Frame(settings); setting_buttons.pack(fill='x')
    tabs = ttk.Notebook(root); tabs.pack(fill='both', expand=True, padx=12, pady=10)
    processing = ttk.Frame(tabs, padding=10); reviewing = ttk.Frame(tabs, padding=10); knowledge = ttk.Frame(tabs, padding=10)
    tabs.add(processing, text='Обезличить'); tabs.add(reviewing, text='Проверка строк'); tabs.add(knowledge, text='База знаний')
    footer = ttk.Frame(root, padding=12); footer.pack(fill='x')
    progress = ttk.Progressbar(footer, mode='indeterminate'); progress.pack(fill='x')
    ttk.Label(footer, textvariable=status, wraplength=1100).pack(anchor='w')

    def button(parent, label, command):
        b = ttk.Button(parent, text=label, command=command); b.pack(side='left', padx=(0, 5), pady=4); buttons.append(b); return b

    def sync_label():
        s = state['store']; info = s.last_sync
        shared.set(str(s.shared or 'Общая папка не настроена'))
        sync_status.set(info['message'] + ' • В очереди: ' + str(info['pending']) + ' • ' + info.get('time', '')[:19].replace('T', ' '))

    def run_job(work, done=None):
        if state['busy']:
            return
        state['busy'] = True; cancel.clear(); progress.configure(mode='indeterminate'); progress.start(12)
        for b in buttons: b.configure(state='disabled')
        status.set('Выполняется…')
        def worker():
            try:
                value = work(); events.put(('done', (value, done)))
            except Exception as exc:
                logging.exception('Operation failed'); events.put(('error', str(exc)))
        threading.Thread(target=worker, daemon=True).start()

    def notify(message): events.put(('status', message))

    def finish():
        state['busy'] = False; progress.stop()
        for b in buttons: b.configure(state='normal')
        sync_label()

    def poll():
        try:
            while True:
                kind, value = events.get_nowait()
                if kind == 'status':
                    import re
                    hit = re.search(r'(\d+)\s*/\s*(\d+)', value)
                    if hit and int(hit[2]):
                        percent = min(100, int(hit[1]) * 100 / int(hit[2])); progress.stop(); progress.configure(mode='determinate', maximum=100, value=percent)
                        value += f' ({percent:.0f}%)'
                    status.set(value)
                elif kind == 'error':
                    finish(); status.set(value); messagebox.showerror('Не удалось выполнить', value)
                else:
                    result, callback = value; finish()
                    if callback:
                        try:
                            callback(result)
                        except Exception as exc:
                            logging.exception('UI update failed')
                            status.set(str(exc)); messagebox.showerror('Ошибка отображения', str(exc))
        except queue.Empty:
            pass
        root.after(100, poll)

    def select_share():
        folder = filedialog.askdirectory(title='Единая папка базы знаний подразделения')
        if folder:
            def work():
                new = KnowledgeStore(app_dir, folder); new.sync(); return new
            def done(new):
                state['store'] = new; sync_label(); refresh_sessions(); refresh_knowledge()
                status.set('Общая база подключена. Эта папка должна быть одинаковой у всех инженеров.')
            run_job(work, done)

    def sync_now():
        run_job(lambda: state['store'].sync(), lambda _: (refresh_knowledge(), status.set('Синхронизация завершена.')))

    button(setting_buttons, 'Общая папка…', select_share)
    button(setting_buttons, 'Синхронизировать', sync_now)

    def backup():
        path = filedialog.asksaveasfilename(title='Резервная копия общей базы', defaultextension='.zip', filetypes=[('ZIP', '*.zip')])
        if path: run_job(lambda: state['store'].backup(path), lambda _: status.set('Резервная копия сохранена и проверена.'))

    def restore():
        path = filedialog.askopenfilename(title='Резервная копия', filetypes=[('ZIP', '*.zip')])
        if not path: return
        folder = filedialog.askdirectory(title='Пустая папка для восстановления')
        if folder: run_job(lambda: restore_backup(path, folder), lambda _: status.set('Копия восстановлена. Для подключения выберите её через «Общая папка».'))

    button(setting_buttons, 'Резервная копия…', backup); button(setting_buttons, 'Восстановить…', restore)
    ttk.Button(setting_buttons, text='Отменить обработку', command=cancel.set).pack(side='right')

    files = []
    ttk.Label(processing, text='Исходные файлы сохраняются. В результате можно исправить наименование и отметить «Правильно».').pack(anchor='w')
    listbox = tk.Listbox(processing, height=11, selectmode='extended'); listbox.pack(fill='both', expand=True, pady=8)
    file_buttons = ttk.Frame(processing); file_buttons.pack(fill='x')
    def add(paths):
        for p in map(Path, paths):
            if p.suffix.lower() in SUPPORTED and p not in files and not p.name.startswith('~$'):
                files.append(p); listbox.insert('end', str(p))
    def folder():
        name = filedialog.askdirectory()
        if name:
            run_job(lambda: [str(p) for p in Path(name).rglob('*') if p.is_file() and p.suffix.lower() in SUPPORTED], add)
    def remove():
        for i in reversed(listbox.curselection()): files.pop(i); listbox.delete(i)
    button(file_buttons, 'Добавить файлы', lambda: add(filedialog.askopenfilenames(filetypes=[('Excel / CSV', '*.xlsx *.xlsm *.xls *.csv')])))
    button(file_buttons, 'Добавить папку', folder); button(file_buttons, 'Убрать выбранные', remove)
    ttk.Label(processing, text='Папка результата').pack(anchor='w', pady=(12, 0))
    outbar = ttk.Frame(processing); outbar.pack(fill='x')
    ttk.Entry(outbar, textvariable=output).pack(side='left', fill='x', expand=True)
    def choose_output():
        name = filedialog.askdirectory()
        if name: output.set(name)
    button(outbar, 'Выбрать…', choose_output)
    def start():
        if not state['store'].shared:
            messagebox.showinfo('Общая база', 'Сначала выберите общую папку подразделения. Все инженеры должны указать одну и ту же папку.'); return
        if not files:
            messagebox.showinfo('Нет файлов', 'Добавьте исходные файлы.'); return
        targets = list(files); destination = output.get().strip(); active_store = state['store']
        if not destination:
            messagebox.showerror('Нет папки', 'Укажите папку результата.'); return
        def work():
            active_store.sync(); snapshot = active_store.snapshot(); az = ExpertAnonymizer(snapshot)
            good, failed, review = [], [], 0
            for src in targets:
                if cancel.is_set(): break
                try:
                    dst, report = process_file(src, destination, az, notify, active_store if active_store.shared else None, cancel)
                    good.append((str(dst), report)); review += report['yellow'] + report['red'] + report['skipped_formulas']
                except InterruptedError: break
                except Exception as exc: failed.append(src.name + ': ' + str(exc)); logging.exception('Processing failed')
            active_store.sync()
            return good, failed, review, cancel.is_set()
        def done(value):
            good, failed, review, stopped = value
            message = f'Готово файлов: {len(good)}. Ошибок: {len(failed)}. Для проверки: {review}.'
            if stopped: message += ' Обработка отменена; завершённые файлы сохранены.'
            status.set(message); refresh_sessions()
            if failed: show_text('Ошибки обработки', '\n'.join(failed))
        run_job(work, done)
    actionbar = ttk.Frame(processing); actionbar.pack(fill='x', pady=12)
    button(actionbar, 'Обезличить', start)
    def open_output():
        if Path(output.get()).is_dir() and hasattr(os, 'startfile'): os.startfile(output.get())
    button(actionbar, 'Открыть результаты', open_output)

    def import_file():
        path = filedialog.askopenfilename(title='Вернуть проверенный Excel', filetypes=[('Excel / CSV', '*.xlsx *.xlsm *.xls *.csv')])
        if not path: return
        def done(report):
            message = f"Принято решений: {report['accepted']}. Повторных: {report['repeated']}. Без отметки и изменений: {report['untouched']}. Проблем: {len(report['issues'])}."
            status.set(message); refresh_knowledge(); refresh_sessions()
            if report['issues']:
                show_text('Строки, которые не внесены в базу', message + '\n\n' + '\n'.join(f"{x['sheet']} / {x['row']}: {x['reason']}" for x in report['issues']))
        run_job(lambda: import_review(path, state['store'], notify, cancel), done)
    button(actionbar, 'Загрузить проверенный Excel…', import_file)

    def show_text(title, content):
        win = tk.Toplevel(root); win.title(title); win.geometry('940x640')
        text = tk.Text(win, wrap='word'); text.pack(fill='both', expand=True); text.insert('1.0', content); text.configure(state='disabled')
        def save():
            path = filedialog.asksaveasfilename(defaultextension='.txt')
            if path: Path(path).write_text(content, encoding='utf-8')
        ttk.Button(win, text='Сохранить…', command=save).pack(anchor='e')

    reviewbar = ttk.Frame(reviewing); reviewbar.pack(fill='x')
    sessions = ttk.Combobox(reviewbar, state='readonly', width=47); sessions.pack(side='left')
    search = tk.StringVar(); ttk.Entry(reviewbar, textvariable=search, width=27).pack(side='left', padx=8)
    filter_status = tk.StringVar(value='Все')
    ttk.Combobox(reviewbar, textvariable=filter_status, state='readonly', values=['Все', 'ЖЁЛТЫЙ', 'ЗЕЛЁНЫЙ', 'КРАСНЫЙ'], width=13).pack(side='left')
    session_ids = []
    columns = ('row', 'code', 'source', 'automatic', 'status')
    tree = ttk.Treeview(reviewing, columns=columns, show='headings', height=9, selectmode='extended')
    for col, title, width in zip(columns, ['Строка', 'Код', 'Исходное', 'Автоматическое', 'Статус'], [60, 100, 340, 340, 90]):
        tree.heading(col, text=title); tree.column(col, width=width, stretch=col in ('source', 'automatic'))
    tree.pack(fill='both', expand=True, pady=6)
    for tag, color in [('ЖЁЛТЫЙ', '#fff2cc'), ('ЗЕЛЁНЫЙ', '#d9ead3'), ('КРАСНЫЙ', '#f4cccc')]: tree.tag_configure(tag, background=color)
    pagination = ttk.Frame(reviewing); pagination.pack(fill='x')
    page_label = tk.StringVar(); ttk.Label(pagination, textvariable=page_label).pack(side='right')

    def refresh_sessions():
        if state['busy']: return
        s = state['store']
        with s.db() as db:
            data = db.execute("SELECT session,count(*),min(json_extract(body,'$.sheet')) FROM rows GROUP BY session ORDER BY max(rowid) DESC LIMIT 50").fetchall()
        session_ids[:] = [x[0] for x in data]
        sessions['values'] = [f'{len(data)-i}. {sheet} • {count} строк' for i, (_, count, sheet) in enumerate(data)]
        if data: sessions.current(0); state['offset'] = 0; load_rows()
        else: tree.delete(*tree.get_children()); state['rows'] = {}

    def load_rows(reset=False):
        if state['busy'] or sessions.current() < 0: return
        if reset: state['offset'] = 0
        sid = session_ids[sessions.current()]
        query = search.get(); selected_filter = filter_status.get(); offset = state['offset']
        def work():
            with state['store'].db() as db:
                sql = 'SELECT body FROM rows WHERE session=? AND (source LIKE ? OR code LIKE ?)'
                args = [sid, '%' + query + '%', '%' + query + '%']
                if selected_filter != 'Все': sql += " AND json_extract(body,'$.status')=?"; args.append(selected_filter)
                sql += ' ORDER BY rowid LIMIT 200 OFFSET ?'; args.append(offset)
                return [json.loads(x[0]) for x in db.execute(sql, args)]
        def done(rows):
            tree.delete(*tree.get_children()); state['rows'] = {r['id']: r for r in rows}; state['selected'] = None
            for r in rows: tree.insert('', 'end', iid=r['id'], values=(r['row'], r['code'], r['source'], r['automatic'], r['status']), tags=(r['status'],))
            page_label.set(f'{offset + 1 if rows else 0}–{offset + len(rows)} • до 200 строк на странице')
        run_job(work, done)
    button(reviewbar, 'Найти', lambda: load_rows(True))
    def page(delta): state['offset'] = max(0, state['offset'] + delta); load_rows()
    button(pagination, '← 200', lambda: page(-200)); button(pagination, '200 →', lambda: page(200))
    sessions.bind('<<ComboboxSelected>>', lambda _: load_rows(True))
    ttk.Label(reviewing, text='Итоговое наименование выбранной строки (можно исправлять и зелёные)').pack(anchor='w')
    final_text = tk.Text(reviewing, height=4, wrap='word'); final_text.pack(fill='x')
    details = tk.StringVar(); ttk.Label(reviewing, textvariable=details, wraplength=1100).pack(anchor='w', pady=5)
    def selected(_=None):
        selection = tree.selection()
        if not selection: return
        row = state['rows'][selection[0]]; state['selected'] = row
        final_text.delete('1.0', 'end'); final_text.insert('1.0', row['automatic'])
        details.set(row['reason'] + '\nУдалено: ' + '; '.join(row['removed']) + '\nЗавод: ' + row['detected_factory'])
    tree.bind('<<TreeviewSelect>>', selected)
    decisions = ttk.Frame(reviewing); decisions.pack(fill='x')
    def decide(action):
        chosen = [dict(state['rows'][key]) for key in tree.selection()]
        if not chosen: return
        if action == 'Исправить' and len(chosen) != 1:
            messagebox.showinfo('Одна строка', 'Редактирование наименования применяется к одной строке. Для группы используйте контекстное правило.'); return
        edited = final_text.get('1.0', 'end-1c')
        if action == 'Исправить' and not edited.strip():
            messagebox.showerror('Пустое наименование', 'Введите итоговое наименование.'); return
        def work():
            active_store = state['store']; count = 0
            for row in chosen:
                value = row['source'] if action == 'Оставить как в исходном' else edited if action == 'Исправить' else row['automatic']
                count += active_store.decide(row, value, action) is not None
            active_store.sync(); return count
        run_job(work, lambda count: (status.set(f'Сохранено решений: {count}. Они применятся при следующей обработке.'), refresh_knowledge()))
    for label in ('Правильно', 'Исправить', 'Оставить как в исходном'):
        button(decisions, label, lambda action=label: decide(action))

    def propose():
        row = state['selected']
        if not row: return
        win = tk.Toplevel(root); win.title('Правило для похожих ресурсов'); win.geometry('680x450')
        ttk.Label(win, text='Правило действует только для указанного завода, категории и роли.\nПервое решение — кандидат; автоматическое применение после второго инженера.', wraplength=640).pack(padx=12, pady=12)
        variables = {}
        defaults = [('fragment', 'Точный фрагмент', ''), ('factory', 'Завод изделия (точно как в справочнике)', row['factory'] or row['detected_factory']), ('category', 'Категория (слово/фраза в исходном)', ''), ('role', 'Роль', 'изделие'), ('action', 'Решение', 'KEEP')]
        for key, label, default in defaults:
            ttk.Label(win, text=label).pack(anchor='w', padx=12); var = tk.StringVar(value=default); variables[key] = var
            values = ['KEEP', 'DELETE'] if key == 'action' else ['изделие', 'комплектующее'] if key == 'role' else None
            widget = ttk.Combobox(win, textvariable=var, values=values, state='readonly') if values else ttk.Entry(win, textvariable=var)
            widget.pack(fill='x', padx=12)
        ttk.Label(win, text='KEEP — сохранять; DELETE — удалять. Для комплектующего нужны явный контекст и его завод в той же части текста.', wraplength=640).pack(padx=12, pady=8)
        def save():
            values = {k: v.get() for k, v in variables.items()}; sample = dict(row, factory=values['factory'])
            def work():
                event = state['store'].propose_rule(sample, values['fragment'], values['action'], values['category'], values['role']); state['store'].sync(); return event
            win.destroy(); run_job(work, lambda _: (refresh_knowledge(), status.set('Контекстное правило сохранено.')))
        ttk.Button(win, text='Сохранить правило', command=save).pack()
    button(decisions, 'Применить к похожим…', propose)
    button(decisions, 'Подробности', lambda: show_text('Исходное, автоматическое и основания', canonical(state['selected'])) if state['selected'] else None)

    knowledge_search = tk.StringVar(); knowledge_filter = tk.StringVar(value='Все')
    kbbar = ttk.Frame(knowledge); kbbar.pack(fill='x')
    ttk.Entry(kbbar, textvariable=knowledge_search, width=40).pack(side='left')
    ttk.Combobox(kbbar, textvariable=knowledge_filter, state='readonly', values=['Все', 'CANDIDATE', 'ACTIVE', 'TRUSTED', 'DISPUTED', 'DISABLED'], width=17).pack(side='left', padx=8)
    ktree = ttk.Treeview(knowledge, columns=('kind', 'state', 'vote', 'value', 'example'), show='headings', selectmode='browse')
    for c, title, width in [('kind', 'Вид', 60), ('state', 'Доверие', 90), ('vote', 'Инженеров', 80), ('value', 'Решение', 250), ('example', 'Пример / область', 400)]:
        ktree.heading(c, text=title); ktree.column(c, width=width)
    ktree.pack(fill='both', expand=True, pady=8)
    kbsummary = tk.StringVar(); ttk.Label(knowledge, textvariable=kbsummary).pack(anchor='w')
    ttk.Label(knowledge, text='Доверие — состояние знания, не вероятность правильности. Статическая база сохраняется как исходная, без экспертных голосов.').pack(anchor='w')
    def refresh_knowledge():
        if state['busy']: return
        query = knowledge_search.get().casefold(); selected_filter = knowledge_filter.get()
        def work(): return state['store'].snapshot()
        def done(snapshot):
            state['knowledge'] = snapshot; ktree.delete(*ktree.get_children()); shown = 0
            for key, e in snapshot.entries.items():
                if selected_filter != 'Все' and e['status'] != selected_filter: continue
                if query and query not in canonical(e).casefold(): continue
                sample = e['sample']; example = sample.get('source', sample.get('example', ''))
                ktree.insert('', 'end', iid=key, values=(e['kind'], e['status'], e['confirmations'], e['value'], example)); shown += 1
                if shown >= 500: break
            kbsummary.set(f'Версия знаний: {snapshot.version}. Всего: {len(snapshot.entries)}. Показано: {shown} (до 500; уточните поиск).')
        run_job(work, done)
    button(kbbar, 'Найти / обновить', refresh_knowledge)
    kbactions = ttk.Frame(knowledge); kbactions.pack(fill='x')
    def selected_entry():
        chosen = ktree.selection()
        return state['knowledge'].entries.get(chosen[0]) if chosen and state['knowledge'] else None
    def history():
        e = selected_entry()
        if e: show_text('История решений', json.dumps(e, ensure_ascii=False, indent=2))
    button(kbactions, 'История и примеры', history)
    def vote_rule():
        e = selected_entry()
        if not e or e['kind'] != 'rule':
            messagebox.showinfo('Контекстное правило', 'Для ресурса подтвердите конкретную строку на вкладке проверки.'); return
        sample = e['sample']
        choice = simpledialog.askstring('Независимое решение', 'Введите KEEP или DELETE для фрагмента «' + sample['fragment'] + '»:\n' + sample['example'])
        if not choice: return
        action = choice.strip().upper()
        if action not in ('KEEP', 'DELETE'): messagebox.showerror('Решение', 'Допустимы KEEP и DELETE.'); return
        row = dict(id=sample['row_id'], session=sample['session'], source=sample['example'], factory=sample['scope']['factory'])
        def work():
            state['store'].propose_rule(row, sample['fragment'], action, sample['scope']['category'], sample['scope']['role']); state['store'].sync()
        run_job(work, lambda _: refresh_knowledge())
    button(kbactions, 'Подтвердить / оспорить правило', vote_rule)
    def control(disabled):
        e = selected_entry()
        if not e: return
        reason = simpledialog.askstring('Причина', 'Укажите причину ' + ('отключения' if disabled else 'возврата') + ' знания:')
        if not reason: return
        def work(): state['store'].control(e['key'], disabled, reason); state['store'].sync()
        run_job(work, lambda _: refresh_knowledge())
    button(kbactions, 'Отключить', lambda: control(True)); button(kbactions, 'Включить', lambda: control(False))
    def export():
        path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel', '*.xlsx')])
        if path: run_job(lambda: export_knowledge(path, state['store'].snapshot()), lambda _: status.set('Знания и история экспортированы в Excel.'))
    button(kbactions, 'Экспорт Excel…', export)

    def heartbeat():
        if not state['busy'] and state['store'].shared:
            run_job(lambda: state['store'].sync(), lambda _: status.set('Автоматическая синхронизация завершена.'))
        root.after(60000, heartbeat)
    def close():
        if state['busy']: messagebox.showinfo('Выполняется операция', 'Дождитесь завершения. Обработку можно отменить кнопкой сверху.')
        else: root.destroy()
    root.protocol('WM_DELETE_WINDOW', close)
    if smoke:
        root.update(); root.destroy(); return True
    poll()
    if store.shared: root.after(200, lambda: run_job(lambda: state['store'].sync(), lambda _: refresh_sessions()))
    root.after(1000, refresh_sessions); root.after(60000, heartbeat)
    root.mainloop()
