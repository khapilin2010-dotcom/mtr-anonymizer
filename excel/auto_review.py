"""Automatic review only after three distinct, explicit engineer decisions."""
from excel.knowledge import case_key, code_key
from excel.review_display import change_spans

REVIEW_POLICY = 5

SAFE_REMOVALS = {'ТУ', 'реквизиты письма', 'ИНН/КПП', 'контакт'}
BENIGN_REASONS = {'Производитель не определён', 'Нет подтверждённых правил для производителя'}


def can_auto_review(snapshot, source, code, factory, result, applied):
    if not applied or any(e.get('case_confirmations', 0) < 3 for e in applied):
        return False
    if result['status'] == 'КРАСНЫЙ' or snapshot.settings_conflict:
        return False
    # Changed source with an old code, or a disabled/disputed exact decision,
    # must still be reviewed even when a broader series rule is trusted.
    if snapshot.entries.get(case_key(code, source, factory)) or snapshot.cases_by_code.get(code_key(code)):
        return False
    reasons = {s.strip() for s in result.get('reason', '').split(';') if s.strip()}
    if reasons - BENIGN_REASONS:
        return False
    deleted, added = change_spans(source, result['text'])
    if added or not result['text'].strip():
        return False
    approved = [(t['start'], t['end']) for t in result.get('trace', []) if t.get('rule') in SAFE_REMOVALS]
    for a, b in deleted:
        for index in range(a, b):
            if source[index] not in ' \t\r\n,;().«»"' and not any(x <= index < y for x, y in approved):
                return False
    return True
