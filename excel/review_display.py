"""Literal character spans shared by the review screen and UI tests."""
from difflib import SequenceMatcher


def change_spans(source, final):
    deleted, added = [], []
    for action, a, b, c, d in SequenceMatcher(None, source, final, autojunk=False).get_opcodes():
        if action in ('delete', 'replace') and source[a:b].strip():
            deleted.append((a, b))
        if action in ('insert', 'replace') and final[c:d].strip():
            added.append((c, d))
    return deleted, added


def restored_spans(source, automatic, final):
    """Mark source text restored into final, not arbitrary inserted wording."""
    removed, _ = change_spans(source, automatic)
    result = []
    for action, a, b, c, d in SequenceMatcher(None, source, final, autojunk=False).get_opcodes():
        if action == 'equal':
            for start, end in removed:
                left, right = max(a, start), min(b, end)
                if left < right and source[left:right].strip():
                    result.append((c + left - a, c + right - a))
    return result
