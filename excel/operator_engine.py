"""Operator-facing expert engine.

Exact engineer decisions are authoritative for the same resource immediately
after one explicit confirmation.  They are therefore applied on every future
matching run without an extra checkbox.  The optional auto-apply flag now only
controls broader learned DELETE rules; conservative KEEP rules may still protect
text automatically because keeping extra text cannot cause over-anonymization.
"""
import os

from excel.expert_engine import ExpertAnonymizer as _ExpertAnonymizer
from excel.knowledge import case_key
from excel.semantics import features


def auto_apply_enabled():
    return os.environ.get('MTR_AUTO_APPLY_CONFIRMED', '').strip().lower() in {
        '1', 'true', 'yes', 'on', 'да'
    }


class OperatorExpertAnonymizer(_ExpertAnonymizer):
    def __init__(self, snapshot, base=None, auto_apply_confirmed=None):
        super().__init__(snapshot, base)
        self.auto_apply_confirmed = (auto_apply_enabled() if auto_apply_confirmed is None
                                     else bool(auto_apply_confirmed))

    def anonymize(self, name, code='', factory=''):
        source = str(name or '')
        entry = self.snapshot.entries.get(case_key(code, source, factory))

        # One explicit engineer decision for this exact resource becomes ACTIVE
        # immediately in Snapshot.  Exact decisions always outrank the static
        # anonymizer, including «Оставить как в исходном».  The checkbox is not
        # required for this path; otherwise an engineer could correct the same
        # row and watch the program repeat the old mistake on the next run.
        if entry and entry.get('kind') == 'case' and entry.get('status') in ('ACTIVE', 'TRUSTED'):
            return super().anonymize(source, code, factory)

        if self.auto_apply_confirmed:
            return super().anonymize(source, code, factory)

        info = features(source, factory, code, self.base)
        matches = self.matching_rules(source, info)

        # General KEEP knowledge is safe to apply automatically: it can only
        # prevent an existing deletion. General DELETE knowledge remains
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
                                   auto_apply=False)

        red = []
        advice = []
        if entry and entry['status'] != 'DISABLED':
            self.describe(result, entry,
                          'Код Автодокс' if str(code).strip()
                          else 'Точное наименование и завод')
            result['knowledge']['auto_apply'] = False
            result['knowledge']['proposed'] = entry['value']
            if entry['status'] == 'DISPUTED':
                red.append('Инженеры сохранили разные решения')
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
