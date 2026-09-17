"""Small, asynchronous wizard for loading previously reviewed selections."""
from pathlib import Path
import queue
import threading
import tkinter as tk
from tkinter import ttk, messagebox

from excel.history_import import best_candidate, preview_selection, commit_selection
from excel.mapped_import import NONE, header_candidates, guess_mapping
from excel.theme import text_colors, MUTED


class SelectionDialog:
    def __init__(self, parent, store, path):
        self.store, self.path = store, Path(path)
        self.plan = self.report = None
        self.approved = set()
        self.busy = False
        self.events = queue.Queue()
        self.window = w = tk.Toplevel(parent)
        w.title('Загрузить ранее обезличенную выборку')
        w.geometry('1020x790'); w.minsize(850, 690)
        w.transient(parent)
        w.protocol('WM_DELETE_WINDOW', self.close)
        body = ttk.Frame(w, padding=16); body.pack(fill='both', expand=True)
        ttk.Label(body, text='Пополнить базу проверенными решениями',
                  font=('Segoe UI', 17, 'bold')).pack(anchor='w')
        ttk.Label(body, text='Нужны исходное и обезличенное наименования. Код и завод укажите, если они есть в файле.',
                  wraplength=930).pack(anchor='w', pady=(5, 2))
        ttk.Label(body, text=self.path.name, foreground=MUTED).pack(anchor='w', pady=(0, 8))
        self.candidate = ttk.Combobox(body, state='readonly')
        self.candidate.pack(fill='x')
        self.candidate.bind('<<ComboboxSelected>>', self.select_candidate)
        fields = ttk.Frame(body); fields.pack(fill='x', pady=8)
        self.fields = {}
        for i, (key, label) in enumerate((('source', 'Исходное наименование *'), ('final', 'Обезличенное наименование *'),
                                        ('code', 'Код Автодокс'), ('factory', 'Завод / производитель'))):
            c, r = i % 2, (i // 2) * 2
            fields.columnconfigure(c, weight=1)
            ttk.Label(fields, text=label).grid(row=r, column=c, sticky='w', padx=(0, 12))
            box = ttk.Combobox(fields, state='readonly', width=42)
            box.grid(row=r + 1, column=c, sticky='ew', padx=(0, 12), pady=(2, 6))
            box.bind('<<ComboboxSelected>>', self.invalidate)
            self.fields[key] = box
        bar = ttk.Frame(body); bar.pack(fill='x')
        self.preview_button = ttk.Button(bar, text='Проверить выборку', command=self.preview)
        self.preview_button.pack(side='left')
        self.progress = ttk.Progressbar(bar, mode='indeterminate'); self.progress.pack(side='left', fill='x', expand=True, padx=12)
        self.summary = tk.StringVar(value='Читаю листы и заголовки…')
        ttk.Label(body, textvariable=self.summary, wraplength=940).pack(anchor='w', pady=8)
        tabs = ttk.Notebook(body); tabs.pack(fill='both', expand=True)
        review = ttk.Frame(tabs, padding=6); issues = ttk.Frame(tabs, padding=6)
        tabs.add(review, text='Решения и конфликты'); tabs.add(issues, text='Пропущенные строки')
        self.rows = ttk.Treeview(review, columns=('row', 'code', 'action'), show='headings', height=5, selectmode='browse')
        for key, label, width in (('row', 'Строка', 70), ('code', 'Код Автодокс', 170), ('action', 'Действие при загрузке', 510)):
            self.rows.heading(key, text=label); self.rows.column(key, width=width, minwidth=55, stretch=key == 'action')
        scroll = ttk.Scrollbar(review, orient='vertical', command=self.rows.yview)
        self.rows.configure(yscrollcommand=scroll.set)
        self.rows.grid(row=0, column=0, sticky='nsew'); scroll.grid(row=0, column=1, sticky='ns')
        review.columnconfigure(0, weight=1); review.rowconfigure(0, weight=1)
        self.details = tk.Text(review, height=6, wrap='word', **text_colors())
        self.details.grid(row=1, column=0, columnspan=2, sticky='ew', pady=6); self.details.configure(state='disabled')
        self.take = tk.BooleanVar(value=False)
        self.take_button = ttk.Checkbutton(review, text='Принять вариант выборки для этой позиции', variable=self.take, command=self.toggle)
        self.take_button.grid(row=2, column=0, columnspan=2, sticky='w'); self.take_button.configure(state='disabled')
        self.rows.bind('<<TreeviewSelect>>', self.show_selected)
        self.issues = tk.Text(issues, wrap='word', **text_colors()); self.issues.pack(fill='both', expand=True)
        self.issues.configure(state='disabled')
        self.reviewed = tk.BooleanVar(value=False)
        self.reviewed_button = ttk.Checkbutton(body, text='Выборка проверена инженером', variable=self.reviewed, command=self.update_commit)
        self.reviewed_button.pack(anchor='w', pady=(10, 6))
        footer = ttk.Frame(body); footer.pack(fill='x')
        self.commit_button = ttk.Button(footer, text='Добавить решения в базу', style='Primary.TButton', command=self.commit, state='disabled')
        self.commit_button.pack(side='left')
        ttk.Button(footer, text='Закрыть', command=self.close).pack(side='right')
        self.after_id = w.after(70, self.poll)
        self.job(lambda: header_candidates(self.path), self.loaded)

    def close(self):
        if self.busy:
            messagebox.showinfo('Загрузка выборки', 'Дождитесь завершения текущей операции.', parent=self.window)
            return
        self.window.after_cancel(self.after_id)
        self.window.destroy()

    def job(self, work, done):
        if self.busy:
            return
        self.busy = True
        self.set_controls()
        self.progress.start(12)
        def worker():
            try:
                self.events.put((True, work(), done))
            except Exception as exc:
                self.events.put((False, str(exc), None))
        threading.Thread(target=worker, daemon=True).start()

    def poll(self):
        try:
            ok, value, done = self.events.get_nowait()
            self.busy = False; self.progress.stop(); self.set_controls()
            if ok:
                done(value)
            else:
                self.plan = None; self.update_commit(); self.summary.set(value)
                messagebox.showerror('Загрузка выборки', value, parent=self.window)
        except queue.Empty:
            pass
        self.after_id = self.window.after(70, self.poll)

    def set_controls(self):
        for box in (self.candidate, *self.fields.values()):
            box.configure(state='disabled' if self.busy else 'readonly')
        self.preview_button.configure(state='disabled' if self.busy else 'normal')
        self.reviewed_button.configure(state='disabled' if self.busy else 'normal')
        self.take_button.configure(state='disabled')
        self.update_commit()
        if not self.busy:
            self.show_selected()

    def loaded(self, candidates):
        self.candidates = candidates
        self.candidate.configure(values=[f"{c['sheet']} — заголовки в строке {c['row']}" for c in candidates])
        self.candidate.current(best_candidate(candidates))
        self.select_candidate()

    def select_candidate(self, _event=None):
        candidate = self.candidates[self.candidate.current()]
        values = [NONE] + [f'{i + 1}. {name or "Без заголовка"}' for i, name in enumerate(candidate['headers'])]
        mapping = guess_mapping(candidate['headers'])
        for key, box in self.fields.items():
            box.configure(values=values)
            box.current(0 if mapping[key] is None else mapping[key] + 1)
        self.invalidate()

    def invalidate(self, _event=None):
        self.plan = None; self.approved.clear(); self.reviewed.set(False)
        self.rows.delete(*self.rows.get_children())
        self.put_text(self.details, ''); self.put_text(self.issues, '')
        self.take_button.configure(state='disabled'); self.take.set(False)
        self.summary.set('Проверьте выбранные столбцы и нажмите «Проверить выборку».')
        self.update_commit()

    @staticmethod
    def put_text(widget, value):
        widget.configure(state='normal'); widget.delete('1.0', 'end'); widget.insert('1.0', value); widget.configure(state='disabled')

    def preview(self):
        if self.busy or not hasattr(self, 'candidates'):
            return
        candidate = self.candidates[self.candidate.current()]
        mapping = {key: box.current() - 1 if box.current() > 0 else None for key, box in self.fields.items()}
        self.invalidate(); self.summary.set('Проверяю выборку и сравниваю с базой…')
        self.job(lambda: preview_selection(self.path, self.store, candidate, mapping), self.show_plan)

    def action(self, item):
        if not item['conflict']:
            return 'Добавить новое решение'
        if item['key'] in self.approved:
            return 'Принять вариант выборки'
        return 'Конфликт — сохранить решение базы'

    def show_plan(self, plan):
        self.plan = plan
        self.summary.set(f"Новых: {plan['new']} • Повторов: {plan['repeated']} • Конфликтов: {plan['conflicts']} • Пропущенных строк: {len(plan['issues'])}")
        for item in sorted(plan['plan'], key=lambda x: not x['conflict']):
            self.rows.insert('', 'end', iid=item['key'], values=(item['row'], item['code'], self.action(item)))
        self.put_text(self.issues, '\n'.join(f'{sheet}, строка {row}: {reason}' for sheet, row, reason in plan['issues']) or 'Проблемных строк нет.')
        if self.rows.get_children():
            self.rows.selection_set(self.rows.get_children()[0])
        self.update_commit()

    def selected(self):
        selected = self.rows.selection()
        return next((item for item in self.plan['plan'] if selected and item['key'] == selected[0]), None) if self.plan else None

    def show_selected(self, _event=None):
        item = self.selected()
        if not item:
            return
        statuses = {'ACTIVE': 'ДЕЙСТВУЕТ', 'TRUSTED': 'ДЕЙСТВУЕТ', 'DISABLED': 'ОТКЛЮЧЕНО', 'DISPUTED': 'ТРЕБУЕТ ПРОВЕРКИ'}
        self.put_text(self.details, f"Исходное: {item['source']}\nЗавод: {item['factory'] or 'не указан'}\n"
                      f"В базе ({statuses.get(item['current_status'], 'нет решения')}): {item['current'] or '—'}\n"
                      f"В выборке: {item['final']}")
        self.take.set(item['key'] in self.approved)
        self.take_button.configure(state='normal' if item['conflict'] and not self.busy else 'disabled')

    def toggle(self):
        item = self.selected()
        if not item or not item['conflict'] or self.busy:
            return
        if self.take.get():
            self.approved.add(item['key'])
        else:
            self.approved.discard(item['key'])
        self.rows.item(item['key'], values=(item['row'], item['code'], self.action(item)))
        self.update_commit()

    def update_commit(self):
        enabled = not self.busy and self.plan and self.reviewed.get() and (self.plan['new'] or self.approved)
        self.commit_button.configure(state='normal' if enabled else 'disabled')

    def commit(self):
        if self.busy or not self.plan or not self.reviewed.get():
            return
        plan, approved = self.plan, set(self.approved)
        self.summary.set('Сохраняю решения в базе…')
        self.job(lambda: commit_selection(self.path, self.store, plan, approved), self.completed)

    def completed(self, report):
        self.report = report
        self.invalidate()
        message = f"Добавлено решений: {report['accepted']}. Повторов: {report['repeated']}.\nКонфликтов оставлено без изменений: {report['skipped_conflicts']}. Пропущенных строк: {len(report['issues'])}."
        if not report['sync'].get('online') or report['sync'].get('errors'):
            message += '\nРешения сохранены локально. Синхронизация с общей папкой пока не завершена.'
        self.summary.set(message)
        self.put_text(self.issues, '\n'.join(f'{sheet}, строка {row}: {reason}' for sheet, row, reason in report['issues']))
        messagebox.showinfo('База пополнена', message, parent=self.window)


def show_import_dialog(parent, store, path):
    dialog = SelectionDialog(parent, store, path)
    dialog.window.grab_set()
    parent.wait_window(dialog.window)
    return dialog.report
