"""Recoverable row review shared by the desktop UI and regression tests."""
from collections import defaultdict
from excel.knowledge import case_key, code_key, normalize
from excel.operator_engine import OperatorExpertAnonymizer
from excel.output_update import apply_decisions
from excel.auto_review import REVIEW_POLICY


def skips_review(row):
    knowledge = row.get('provenance', {})
    return knowledge.get('exact_applied') is True or knowledge.get('auto_reviewed') is True


def persist_automatic(store, rows):
    """Write first, hide second. Automatic decisions never train the base."""
    groups = defaultdict(list)
    for row in rows:
        action = 'Точное решение инженера' if row['provenance'].get('exact_applied') else 'Проверено автоматически по базе'
        groups[row['session']].append(dict(row=row, final=row['automatic'], action=action, automatic_review=bool(row['provenance'].get('auto_reviewed'))))
    for sid, updates in groups.items():
        output = store.metadata(sid).get('output_file')
        if not output:
            raise ValueError('Не найден итоговый файл предыдущей проверки. Обработайте исходный файл заново.')
        apply_decisions(output, updates)
    # Do not mark any row final until every write succeeded.
    for row in rows:
        row['final'] = row['automatic']


def refresh_review_row(store, row, base=None, engine=None):
    engine = engine or OperatorExpertAnonymizer(store.snapshot(), base=base, auto_apply_confirmed=True)
    snapshot = engine.snapshot
    exact = snapshot.entries.get(case_key(row['code'], row['source'], row['factory']))
    history = snapshot.cases_by_code.get(code_key(row['code']))
    if (not exact and not history and row.get('provenance', {}).get('version') == snapshot.version
            and row.get('provenance', {}).get('review_policy') == REVIEW_POLICY):
        return row
    result = engine.anonymize(row['source'], row['code'], row['factory'])
    return dict(row, initial_automatic=row.get('initial_automatic', row['automatic']),
                automatic=result['text'], status=result['status'],
                reason=result.get('reason', ''), removed=result.get('removed', []),
                classification=result.get('classification'), provenance=result.get('knowledge', {}))


def build_review_queue(store, session_ids, base=None):
    store.sync(auto_backup=False)
    engine = OperatorExpertAnonymizer(store.snapshot(), base=base, auto_apply_confirmed=True)
    queue = []
    for sid in session_ids:
        automatic = []
        for saved in store.session_rows(sid, '', 0, 1000000):
            row = refresh_review_row(store, saved, base, engine)
            if skips_review(row):
                automatic.append(row)
            elif (row.get('provenance', {}).get('series_rules') or normalize(row['source']) != normalize(row['automatic']) or row['status'] != 'ЗЕЛЁНЫЙ'):
                queue.append(row)
        persist_automatic(store, automatic)
    return queue


def advance_review_queue(store, rows, index, base=None):
    """Refresh only up to the next manual row; batch-write a trusted run once."""
    engine = OperatorExpertAnonymizer(store.snapshot(), base=base, auto_apply_confirmed=True)
    automatic = []
    current = index
    while current < len(rows):
        row = refresh_review_row(store, rows[current], base, engine)
        rows[current] = row
        if not skips_review(row):
            break
        automatic.append(row)
        current += 1
    persist_automatic(store, automatic)
    return current
