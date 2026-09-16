"""Operator-facing expert engine.

Exact engineer decisions are authoritative for the same unchanged resource.
They are applied directly before static anonymization, so an explicitly saved
engineer result cannot be removed again on the next pass.  Broader learned
rules stay conservative unless automatic application is enabled.
"""
import os

from excel.expert_engine import ExpertAnonymizer as _ExpertAnonymizer, protection_losses
from excel.knowledge import case_key, fragments, normalize
from excel.semantics import features


def auto_apply_enabled():
    return os.environ.get('MTR_AUTO_APPLY_CONFIRMED', '').strip().lower() in {
        '1', 'true', 'yes', 'on', 'да'
    }


def _same_exact_context(entry, source, factory):
    """A code match alone is not enough for an exact engineer decision.

    ``case_key`` intentionally uses the Autodocs code when it exists.  That is
    useful for grouping history, but it means a changed source string may still
    resolve to the same entry.  We only call a decision *exact* when the source
    and factory from the saved event still match the current row.
    """
    sample = entry.get('sample', {}) if entry else {}
    saved_source = sample.get('source', '')
    saved_factory = sample.get('factory', '')
    return (
        bool(saved_source)
        and normalize(saved_source) == normalize(source)
        and normalize(saved_factory) == normalize(factory)
    )


class OperatorExpertAnonymizer(_ExpertAnonymizer):
    def __init__(self, snapshot, base=None, auto_apply_confirmed=None):
        super().__init__(snapshot, base)
        self.auto_apply_confirmed = (auto_apply_enabled() if auto_apply_confirmed is None
                                     else bool(auto_apply_confirmed))

    def _apply_exact_case(self, entry, source, code, factory):
        """Return the engineer's exact result without re-running static rules."""
        final = str(entry.get('value', ''))
        if not final.strip():
            return None
        # Old/imported databases may contain unsafe edits.  Never silently use
        # an exact result that loses protected technical information.
        losses = protection_losses(source, final, self.base, code, factory)
        if losses:
            return None
        info = features(source, factory, code, self.base)
        removed, _restored = fragments(source, final)
        result = {
            'text': final,
            'factory': info.get('factory', factory),
            'status': 'ЗЕЛЁНЫЙ',
            'changed': normalize(final) != normalize(source),
            'reason': 'Точное решение инженера',
            'removed': ['Решение инженера: ' + x for x in removed if str(x).strip()],
            'trace': [],
            'protected': list(info.get('technical', [])),
            'classification': info,
            'knowledge': {
                'version': self.snapshot.version,
                'source': 'Точное решение инженера',
                'status': entry.get('status', 'ACTIVE'),
                'score': entry.get('score'),
                'confirmations': entry.get('confirmations', 0),
                'opposition': entry.get('opposition', 0),
                'key': entry.get('key'),
                'events': [e.get('id') for e in entry.get('live', []) if e.get('id')],
                'exact_applied': True,
                'auto_apply': True,
            },
        }
        return result

    def anonymize(self, name, code='', factory=''):
        source = str(name or '')
        entry = self.snapshot.entries.get(case_key(code, source, factory))

        # Critical product rule: the engineer's decision for the same unchanged
        # row outranks every static/general rule.  Do not call the parent engine
        # here; doing so can repeat the very deletion the engineer corrected.
        if (entry and entry.get('kind') == 'case'
                and entry.get('status') in ('ACTIVE', 'TRUSTED')
                and _same_exact_context(entry, source, factory)):
            exact = self._apply_exact_case(entry, source, code, factory)
            if exact is not None:
                return exact

        # A code may point to a historical case whose source has since changed.
        # Never treat that stale full-row decision as exact.  In that situation
        # continue with the conservative operator path and show it for review.
        stale_case = bool(entry and entry.get('kind') == 'case'
                          and not _same_exact_context(entry, source, factory))
        if self.auto_apply_confirmed and not stale_case:
            return super().anonymize(source, code, factory)

        info = features(source, factory, code, self.base)
        matches = self.matching_rules(source, info)

        # General KEEP knowledge is safe to apply automatically: it can only
        # prevent an existing deletion.  General DELETE knowledge remains
        # advisory while automatic broader rules are disabled.
        keep_spans = [span for e, spans in matches
                      if e['value'] == 'KEEP' and e['status'] in ('ACTIVE', 'TRUSTED')
                      for span in spans]
        disabled = self.snapshot.static_disabled
        result = self.base.anonymize(source, code, factory,
                                     expert_keeps=keep_spans,
                                     disabled_rules=disabled)
        result['classification'] = info
        result['knowledge'] = dict(version=self.snapshot.version,
                                   source='Статическая база + рекомендации инженеров',
                                   status='ADVISORY', events=[], score=None,
                                   auto_apply=False, exact_applied=False)

        red = []
        advice = []
        if entry and entry['status'] != 'DISABLED':
            self.describe(result, entry,
                          'Код Автодокс' if str(code).strip()
                          else 'Точное наименование и завод')
            result['knowledge']['auto_apply'] = False
            result['knowledge']['exact_applied'] = False
            result['knowledge']['proposed'] = entry['value']
            if entry['status'] == 'DISPUTED':
                red.append('Инженеры сохранили разные решения')
            elif stale_case:
                advice.append('Для этого кода есть решение, но исходное наименование изменилось; требуется проверка')
            elif entry['status'] in ('ACTIVE', 'TRUSTED'):
                advice.append('Есть ранее подтверждённое решение')

        active_delete = []
        disputed = []
        for e, _spans in matches:
            if e['status'] == 'DISPUTED':
                disputed.append(e['sample'].get('fragment', e['key']))
            elif e['status'] in ('ACTIVE', 'TRUSTED') and e['value'] == 'DELETE':
                active_delete.append(e['sample'].get('fragment', e['key']))
        if disputed:
            red.append('Спорные накопленные правила: ' + ', '.join(dict.fromkeys(disputed)))
        if active_delete:
            advice.append('Есть ранее подтверждённые удаления: ' + ', '.join(dict.fromkeys(active_delete)))
            result['knowledge']['proposed_delete'] = list(dict.fromkeys(active_delete))

        if self.snapshot.settings_conflict:
            red.append('Конфликт настроек доверия')

        if red:
            return self.conflict(result, '; '.join(red))
        if advice:
            old_reason = result.get('reason', '').strip()
            result['status'] = 'ЖЁЛТЫЙ'
            result['reason'] = '; '.join(advice + ([old_reason] if old_reason else []))
        return result
