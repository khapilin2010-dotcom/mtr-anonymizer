"""Operator-facing expert engine.

Static anonymization rules keep working as before.  Accumulated expert knowledge
is conservative by default: previous DELETE/case corrections are shown as a
recommendation but are not applied automatically until the engineer explicitly
enables the option in the launcher.  Confirmed KEEP rules may still protect text
because keeping extra text cannot cause over-anonymization.
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
        if self.auto_apply_confirmed:
            return super().anonymize(name, code, factory)

        source = str(name or '')
        info = features(source, factory, code, self.base)
        entry = self.snapshot.entries.get(case_key(code, source, factory))
        matches = self.matching_rules(source, info)

        # KEEP knowledge is safe to apply automatically: it can only prevent an
        # existing deletion.  DELETE knowledge remains advisory in this mode.
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
                advice.append('Есть ранее подтверждённое решение; автоматическое применение выключено')

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
