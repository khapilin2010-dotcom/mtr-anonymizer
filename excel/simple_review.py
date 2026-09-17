"""Recoverable row review shared by the desktop UI and regression tests."""
from excel.knowledge import case_key, code_key, normalize
from excel.operator_engine import OperatorExpertAnonymizer
from excel.output_update import apply_decisions


def build_review_queue(store, session_ids, base=None):
    """Persist a saved result into the output before omitting its review row.

    This also recovers a decision saved immediately before an application crash.
    Presence of a database entry alone is never a reason to omit a row.
    """
    store.sync(auto_backup=False)
    snapshot = store.snapshot()
    engine = OperatorExpertAnonymizer(snapshot, base=base, auto_apply_confirmed=True)
    queue = []
    for sid in session_ids:
        updates = []
        for row in store.session_rows(sid, '', 0, 1000000):
            entry = snapshot.entries.get(case_key(row['code'], row['source'], row['factory']))
            has_history = snapshot.cases_by_code.get(code_key(row['code']))
            if entry or has_history or row.get('provenance', {}).get('exact_applied'):
                result = engine.anonymize(row['source'], row['code'], row['factory'])
                exact = result.get('knowledge', {}).get('exact_applied') is True
                if exact:
                    updates.append(dict(row=row, final=result['text'], action='Точное решение инженера'))
                    continue
                row = dict(row, automatic=result['text'], status=result['status'],
                           reason=result['reason'], removed=result['removed'],
                           provenance=result.get('knowledge', {}))
            if normalize(row['source']) != normalize(row['automatic']) or row['status'] != 'ЗЕЛЁНЫЙ':
                queue.append(row)
        if updates:
            output = store.metadata(sid).get('output_file')
            if not output:
                raise ValueError('Не найден итоговый файл предыдущей проверки. Обработайте исходный файл заново.')
            # Failure propagates: the UI must not announce completion or hide
            # rows while the workbook still contains the obsolete automatic text.
            apply_decisions(output, updates)
    return queue


def refresh_review_row(store, row, base=None):
    """Recompute a queued row against decisions accepted during this session."""
    snapshot = store.snapshot()
    if row.get('provenance', {}).get('version') == snapshot.version:
        return row
    result = OperatorExpertAnonymizer(snapshot, base=base, auto_apply_confirmed=True).anonymize(
        row['source'], row['code'], row['factory'])
    return dict(row, initial_automatic=row.get('initial_automatic', row['automatic']),
                automatic=result['text'], status=result['status'],
                reason=result.get('reason', ''), removed=result.get('removed', []),
                classification=result.get('classification'), provenance=result.get('knowledge', {}))
