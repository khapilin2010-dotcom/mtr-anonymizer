"""Engineer quality center for review, training and accumulated-knowledge audit."""
import json
from pathlib import Path

from excel.expert_engine import protection_losses
from excel.excel_engine import Anonymizer
from excel.knowledge import case_key, canonical, normalize
from excel.quality_tools import (
    export_manager_report, is_unknown_row, knowledge_audit, previous_decision,
    row_risks, same_case,
)


def open_quality_center(parent, store):
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, simpledialog

    store.sync()
    snapshot = store.snapshot()
    base = Anonymizer()
    win = tk.Toplevel(parent)
    win.title('MTR Excel — центр проверки и качества')
    win.geometry('1320x860')
    win.minsize(1080, 720)
    win.transient(parent)

    style = ttk.Style(win)
    style.configure('Danger.TLabel', foreground='#a00000')
    style.configure('Info.TLabel', foreground='#205493')

    state = {'snapshot': snapshot, 'rows': [], 'visible': [], 'selected': None,
             'sessions': [], 'session_ids': []}
    status = tk.StringVar(value='Центр качества готов.')

    top = ttk.Frame(win, padding=(12, 10)); top.pack(fill='x')
    ttk.Label(top, text='Центр проверки и качества', font=('Segoe UI', 18, 'bold')).pack(side='left')
    ttk.Label(top, textvariable=status, wraplength=720).pack(side='right')

    tabs = ttk.Notebook(win); tabs.pack(fill='both', expand=True, padx=12, pady=(0, 10))
    review = ttk.Frame(tabs, padding=8); audit_tab = ttk.Frame(tabs, padding=8)
    tabs.add(review, text='Проверка строк')
    tabs.add(audit_tab, text='Контроль базы знаний')

    # ---- review tab -----------------------------------------------------
    controls = ttk.Frame(review); controls.pack(fill='x')
    session_var = tk.StringVar(); session_box = ttk.Combobox(controls, state='readonly', textvariable=session_var, width=46)
    session_box.pack(side='left')
    query_var = tk.StringVar(); ttk.Entry(controls, textvariable=query_var, width=28).pack(side='left', padx=6)
    filter_var = tk.StringVar(value='Сомнительные')
    filters = ['Сомнительные', 'Только жёлтые', 'Только красные', 'Все изменённые',
               'Опасные', 'Новые / режим обучения', 'Все']
    ttk.Combobox(controls, state='readonly', textvariable=filter_var, values=filters, width=25).pack(side='left')
    training_var = tk.BooleanVar(value=False)
    ttk.Checkbutton(controls, text='Новые сначала', variable=training_var).pack(side='left', padx=8)

    columns = ('row', 'code', 'source', 'automatic', 'status', 'risk', 'previous')
    tree = ttk.Treeview(review, columns=columns, show='headings', selectmode='extended', height=14)
    specs = [
        ('row', 'Строка', 65), ('code', 'Код', 105), ('source', 'Исходное', 310),
        ('automatic', 'Автоматическое', 310), ('status', 'Статус', 90),
        ('risk', 'Риск', 180), ('previous', 'Ранее', 110),
    ]
    for name, title, width in specs:
        tree.heading(name, text=title); tree.column(name, width=width, stretch=name in ('source', 'automatic', 'risk'))
    tree.pack(fill='both', expand=True, pady=(8, 4))
    tree.tag_configure('КРАСНЫЙ', background='#f4cccc')
    tree.tag_configure('ЖЁЛТЫЙ', background='#fff2cc')
    tree.tag_configure('ЗЕЛЁНЫЙ', background='#d9ead3')
    tree.tag_configure('ОПАСНО', background='#f8d7da')

    detail = ttk.Panedwindow(review, orient='horizontal'); detail.pack(fill='x', pady=4)
    left = ttk.LabelFrame(detail, text='Исходное'); middle = ttk.LabelFrame(detail, text='Итог инженера / автоматический результат')
    detail.add(left, weight=1); detail.add(middle, weight=1)
    source_text = tk.Text(left, height=7, wrap='word'); source_text.pack(fill='both', expand=True, padx=4, pady=4)
    source_text.configure(state='disabled')
    final_text = tk.Text(middle, height=7, wrap='word'); final_text.pack(fill='both', expand=True, padx=4, pady=4)

    diff_box = ttk.LabelFrame(review, text='Предпросмотр изменений'); diff_box.pack(fill='x', pady=4)
    diff_text = tk.Text(diff_box, height=5, wrap='word'); diff_text.pack(fill='x', padx=4, pady=4)
    diff_text.tag_configure('removed', background='#f4cccc', overstrike=True)
    diff_text.tag_configure('added', background='#d9ead3')
    diff_text.tag_configure('same', foreground='#555555')
    diff_text.configure(state='disabled')

    detail_var = tk.StringVar()
    ttk.Label(review, textvariable=detail_var, wraplength=1240).pack(anchor='w', pady=(2, 4))

    def session_list():
        with store.db() as db:
            data = db.execute("SELECT session,count(*),max(rowid) FROM rows GROUP BY session ORDER BY max(rowid) DESC LIMIT 100").fetchall()
        labels, ids = [], []
        for i, (sid, count, _) in enumerate(data, 1):
            meta = store.metadata(sid)
            name = Path(meta.get('file', '')).name or sid[:8]
            labels.append(f'{i}. {name} • {count} строк')
            ids.append(sid)
        state['sessions'], state['session_ids'] = labels, ids
        session_box['values'] = labels
        if labels and not session_var.get():
            session_box.current(0)

    def current_session():
        try:
            return state['session_ids'][session_box.current()]
        except Exception:
            return None

    def refresh_snapshot():
        store.sync()
        state['snapshot'] = store.snapshot()

    def match_filter(row):
        selected = filter_var.get()
        risks = row_risks(row, state['snapshot'], base)
        if selected == 'Только жёлтые': return row.get('status') == 'ЖЁЛТЫЙ'
        if selected == 'Только красные': return row.get('status') == 'КРАСНЫЙ'
        if selected == 'Все изменённые': return normalize(row.get('source', '')) != normalize(row.get('automatic', ''))
        if selected == 'Опасные': return bool(risks)
        if selected == 'Новые / режим обучения': return is_unknown_row(row, state['snapshot'])
        if selected == 'Сомнительные': return row.get('status') in ('ЖЁЛТЫЙ', 'КРАСНЫЙ') or bool(risks)
        return True

    def display_rows():
        tree.delete(*tree.get_children())
        query = query_var.get().casefold().strip()
        rows = [r for r in state['rows'] if match_filter(r)]
        if query:
            rows = [r for r in rows if query in (str(r.get('code', '')) + ' ' + r.get('source', '') + ' ' + r.get('automatic', '')).casefold()]
        if training_var.get():
            rows.sort(key=lambda r: (not is_unknown_row(r, state['snapshot']), not bool(row_risks(r, state['snapshot'], base)), r.get('row', 0)))
        state['visible'] = rows
        for row in rows:
            risks = row_risks(row, state['snapshot'], base)
            previous = previous_decision(row, state['snapshot'])
            previous_text = '' if not previous else f"{previous['status']} / {previous['confirmations']}"
            tag = 'ОПАСНО' if risks else row.get('status', '')
            tree.insert('', 'end', iid=row['id'], values=(row.get('row', ''), row.get('code', ''),
                        row.get('source', ''), row.get('automatic', ''), row.get('status', ''),
                        '; '.join(risks[:2]), previous_text), tags=(tag,))
        status.set(f'Показано {len(rows)} из {len(state["rows"])} строк текущего сеанса.')

    def load_session(*_):
        sid = current_session()
        if not sid:
            return
        refresh_snapshot()
        state['rows'] = store.session_rows(sid, '', 0, 10000)
        display_rows()

    def render_diff(source, final):
        from difflib import SequenceMatcher
        diff_text.configure(state='normal'); diff_text.delete('1.0', 'end')
        matcher = SequenceMatcher(None, source, final, autojunk=False)
        for op, a, b, c, d in matcher.get_opcodes():
            if op == 'equal':
                diff_text.insert('end', source[a:b], 'same')
            elif op == 'delete':
                diff_text.insert('end', source[a:b], 'removed')
            elif op == 'insert':
                diff_text.insert('end', final[c:d], 'added')
            else:
                diff_text.insert('end', source[a:b], 'removed')
                diff_text.insert('end', final[c:d], 'added')
        diff_text.configure(state='disabled')

    def selected_row(*_):
        ids = tree.selection()
        if not ids:
            state['selected'] = None; return
        by_id = {r['id']: r for r in state['rows']}
        row = by_id.get(ids[0]); state['selected'] = row
        if not row: return
        source_text.configure(state='normal'); source_text.delete('1.0', 'end'); source_text.insert('1.0', row.get('source', '')); source_text.configure(state='disabled')
        final_text.delete('1.0', 'end'); final_text.insert('1.0', row.get('automatic', ''))
        render_diff(row.get('source', ''), row.get('automatic', ''))
        risks = row_risks(row, state['snapshot'], base)
        previous = previous_decision(row, state['snapshot'])
        previous_line = 'Ранее подтверждённого точного решения нет.' if not previous else (
            f"Ранее: {previous['status']}, подтверждений {previous['confirmations']}: {previous['value']}")
        detail_var.set((row.get('reason', '') or 'Без отдельного пояснения') + '\n' + previous_line +
                       ('\nРиски: ' + '; '.join(risks) if risks else ''))

    tree.bind('<<TreeviewSelect>>', selected_row)
    session_box.bind('<<ComboboxSelected>>', load_session)
    filter_var.trace_add('write', lambda *_: display_rows())
    training_var.trace_add('write', lambda *_: display_rows())

    buttons = ttk.Frame(review); buttons.pack(fill='x', pady=5)

    def current_final():
        return final_text.get('1.0', 'end-1c').strip()

    def validate_final(row, value):
        if not value:
            raise ValueError('Итоговое наименование не может быть пустым.')
        losses = protection_losses(row.get('source', ''), value, base, row.get('code', ''), row.get('factory', ''))
        if losses:
            raise ValueError('Решение теряет защищённые признаки: ' + ', '.join(losses))

    def save_decision(action):
        row = state['selected']
        if not row:
            messagebox.showinfo('Строка', 'Выберите строку.', parent=win); return
        value = row['source'] if action == 'Оставить как в исходном' else row['automatic'] if action == 'Правильно' else current_final()
        try:
            validate_final(row, value)
            event = store.decide(row, value, action)
            store.sync(); refresh_snapshot(); display_rows(); selected_row()
            status.set('Решение сохранено.' if event else 'Такое решение этого инженера уже записано.')
        except Exception as exc:
            messagebox.showerror('Решение не сохранено', str(exc), parent=win)

    def apply_identical():
        row = state['selected']
        if not row:
            messagebox.showinfo('Строка', 'Выберите строку.', parent=win); return
        value = current_final()
        candidates = [r for r in state['rows'] if same_case(row, r)]
        if len(candidates) < 2:
            messagebox.showinfo('Одинаковые случаи', 'Других полностью одинаковых случаев в этом сеансе нет.', parent=win); return
        try:
            for item in candidates: validate_final(item, value)
        except Exception as exc:
            messagebox.showerror('Массовое решение заблокировано', str(exc), parent=win); return
        if not messagebox.askyesno('Массовое решение', f'Применить этот итог к {len(candidates)} одинаковым строкам?\n\n{value}', parent=win):
            return
        count = 0
        for item in candidates:
            count += int(store.decide(item, value, 'Массовое подтверждение одинаковых строк') is not None)
        store.sync(); refresh_snapshot(); display_rows(); status.set(f'Сохранено массовых решений: {count}.')

    def why():
        row = state['selected']
        if not row: return
        previous = previous_decision(row, state['snapshot'])
        payload = {
            'Причина статуса': row.get('reason', ''),
            'Что удалено': row.get('removed', []),
            'Защищено': row.get('provenance', {}).get('protected', []),
            'Источник знания': row.get('provenance', {}).get('source', ''),
            'Доверие/статус знания': row.get('provenance', {}).get('status', ''),
            'Ранее подтверждённое решение': previous,
            'Риски контроля качества': row_risks(row, state['snapshot'], base),
            'Полное основание': row.get('provenance', {}),
        }
        show_text('Почему программа предлагает этот результат', json.dumps(payload, ensure_ascii=False, indent=2, default=str))

    def history():
        row = state['selected']
        if not row: return
        key = case_key(row.get('code', ''), row.get('source', ''), row.get('factory', ''))
        entry = state['snapshot'].entries.get(key)
        if not entry:
            show_text('История строки', 'Для этой строки ещё нет накопленных экспертных решений.'); return
        lines = []
        for event in sorted(entry.get('history', []), key=lambda x: x.get('timestamp', '')):
            if event.get('kind') != 'case': continue
            lines.append(f"{event.get('timestamp','')} • {event.get('user','')} • {event.get('action','')}\n{event.get('value','')}\n")
        show_text('История решений строки', '\n'.join(lines) or 'История пуста.')

    def undo():
        row = state['selected']
        if not row: return
        key = case_key(row.get('code', ''), row.get('source', ''), row.get('factory', ''))
        entry = state['snapshot'].entries.get(key)
        events = [] if not entry else [e for e in entry.get('history', []) if e.get('kind') == 'case' and e.get('user') == store.user]
        events.sort(key=lambda x: x.get('timestamp', ''))
        if len(events) < 2:
            messagebox.showinfo('Отмена', 'У этого инженера нет предыдущего решения для данной строки.', parent=win); return
        previous = events[-2].get('value', '')
        if not messagebox.askyesno('Вернуть предыдущее решение', 'Записать как текущее предыдущее решение?\n\n' + previous, parent=win): return
        try:
            validate_final(row, previous)
            store.decide(row, previous, 'Возврат предыдущего решения'); store.sync(); refresh_snapshot(); display_rows(); selected_row()
            status.set('Предыдущее решение восстановлено новым аудируемым событием.')
        except Exception as exc:
            messagebox.showerror('Не удалось вернуть', str(exc), parent=win)

    def show_text(title, content):
        dialog = tk.Toplevel(win); dialog.title(title); dialog.geometry('900x620'); dialog.transient(win)
        text = tk.Text(dialog, wrap='word'); text.pack(fill='both', expand=True, padx=8, pady=8)
        text.insert('1.0', content); text.configure(state='disabled')
        ttk.Button(dialog, text='Закрыть', command=dialog.destroy).pack(pady=(0, 8))

    ttk.Button(controls, text='Найти / обновить', command=load_session).pack(side='left', padx=6)
    ttk.Button(buttons, text='Правильно', command=lambda: save_decision('Правильно')).pack(side='left', padx=2)
    ttk.Button(buttons, text='Сохранить исправление', command=lambda: save_decision('Исправить')).pack(side='left', padx=2)
    ttk.Button(buttons, text='Оставить исходное', command=lambda: save_decision('Оставить как в исходном')).pack(side='left', padx=2)
    ttk.Button(buttons, text='Применить к одинаковым…', command=apply_identical).pack(side='left', padx=8)
    ttk.Button(buttons, text='Почему?', command=why).pack(side='left', padx=2)
    ttk.Button(buttons, text='История строки', command=history).pack(side='left', padx=2)
    ttk.Button(buttons, text='Вернуть предыдущее', command=undo).pack(side='left', padx=2)

    # ---- knowledge audit tab -------------------------------------------
    audit_summary = tk.StringVar()
    ttk.Label(audit_tab, textvariable=audit_summary, wraplength=1240).pack(anchor='w')
    audit_cols = ('severity', 'issue', 'kind', 'state', 'value', 'votes', 'opp', 'uses', 'example')
    audit_tree = ttk.Treeview(audit_tab, columns=audit_cols, show='headings', height=20)
    for name, title, width in [
        ('severity', 'Уровень', 80), ('issue', 'Проблема', 280), ('kind', 'Вид', 60),
        ('state', 'Статус', 90), ('value', 'Решение', 140), ('votes', 'За', 55),
        ('opp', 'Против', 55), ('uses', 'Применений', 80), ('example', 'Пример', 430),
    ]:
        audit_tree.heading(name, text=title); audit_tree.column(name, width=width, stretch=name in ('issue', 'example'))
    audit_tree.pack(fill='both', expand=True, pady=8)
    audit_tree.tag_configure('КРАСНЫЙ', background='#f4cccc')
    audit_tree.tag_configure('ЖЁЛТЫЙ', background='#fff2cc')

    def refresh_audit():
        refresh_snapshot()
        issues = knowledge_audit(state['snapshot'])
        audit_tree.delete(*audit_tree.get_children())
        for index, item in enumerate(issues):
            audit_tree.insert('', 'end', iid='audit-' + str(index), values=(item['severity'], item['issue'], item['kind'], item['status'],
                              item['value'], item['confirmations'], item['opposition'], item['applications'], item['example']),
                              tags=(item['severity'],))
        red = sum(x['severity'] == 'КРАСНЫЙ' for x in issues); yellow = sum(x['severity'] == 'ЖЁЛТЫЙ' for x in issues)
        audit_summary.set(f'Всего знаний: {len(state["snapshot"].entries)}. Замечаний: {len(issues)}. Красных: {red}. Жёлтых: {yellow}. '
                          'Неиспользованные правила и конфликты здесь видны до того, как они станут массовой ошибкой.')

    audit_buttons = ttk.Frame(audit_tab); audit_buttons.pack(fill='x')
    ttk.Button(audit_buttons, text='Обновить аудит', command=refresh_audit).pack(side='left')

    def export_report():
        path = filedialog.asksaveasfilename(parent=win, defaultextension='.xlsx', filetypes=[('Excel', '*.xlsx')],
                                            initialfile='MTR_Отчёт_руководителю.xlsx')
        if not path: return
        try:
            export_manager_report(path, store); status.set('Отчёт руководителю сохранён: ' + path)
            messagebox.showinfo('Готово', 'Отчёт сохранён.', parent=win)
        except Exception as exc:
            messagebox.showerror('Не удалось создать отчёт', str(exc), parent=win)
    ttk.Button(audit_buttons, text='Отчёт руководителю…', command=export_report).pack(side='left', padx=6)

    ttk.Label(audit_tab, text='Режим обучения находится на вкладке «Проверка строк»: выберите «Новые / режим обучения» '
                              'или включите «Новые сначала». Программа приоритетно покажет случаи без точного накопленного решения.',
              wraplength=1240).pack(anchor='w', pady=8)

    session_list(); refresh_audit()
    if state['sessions']:
        load_session()
    win.focus_set()
    return win
