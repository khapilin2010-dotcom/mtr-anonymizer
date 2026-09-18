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


def restore_deleted(source, final, selection=None):
    """Restore selected missing source characters at their original anchors.

    Keep user insertions in the result. A selection may span both retained and
    deleted text; only the missing part is restored, never duplicate retained text.
    """
    start, end = selection if selection is not None else (0, len(source))
    insertions = []
    for action, a, b, c, d in SequenceMatcher(None, source, final, autojunk=False).get_opcodes():
        if action not in ('delete', 'replace'):
            continue
        left, right = max(a, start), min(b, end)
        if left >= right or not source[left:right].strip():
            continue
        piece = source[left:right]
        # Preserve separators when only part of a larger red span is selected.
        if c and not final[c-1].isspace() and not piece[0].isspace() and any(x.isspace() for x in source[a:left]):
            piece = ' ' + piece
        if c < len(final) and not final[c].isspace() and not piece[-1].isspace() and any(x.isspace() for x in source[right:b]):
            piece += ' '
        insertions.append((c, piece))
    for position, piece in reversed(insertions):
        final = final[:position] + piece + final[position:]
    return final
