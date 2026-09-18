"""Simplified knowledge-store policy for the field Excel application.

The durable event store remains the implementation from ``knowledge.py``.  This
wrapper only changes two product decisions:

* when no folder is supplied, use ``MTR_KNOWLEDGE_DIR`` (normally the folder
  next to the EXE);
* knowledge administration is open to every engineer who has filesystem access
  to that folder.  Windows ACLs remain the trust boundary.

SQLite is still LOCAL cache only; it is never placed on a shared/network disk.
"""
from pathlib import Path
from collections import defaultdict
import os

from excel.knowledge import KnowledgeStore as _KnowledgeStore, Snapshot, contextual_events


class SimpleSnapshot(Snapshot):
    def __init__(self, events, admins=()):
        from excel.series_learning import derived_events
        super().__init__(events + derived_events(events, admins), admins)
        self.series_index = defaultdict(list)
        from excel.semantics import folded
        for entry in self.entries.values():
            if entry['sample'].get('series_rule'):
                entry['case_confirmations'] = len({e['origin_case'] for e in entry['live'] if e['value'] == 'KEEP' and e.get('origin_case')})
                self.series_index[folded(entry['sample']['fragment'])].append(entry)
            if (entry['sample'].get('series_rule') and entry['value'] == 'KEEP'
                    and entry['status'] == 'CANDIDATE'
                    and not any(e.get('requested_status') == 'CANDIDATE' for e in entry['history'])):
                entry['status'] = 'ACTIVE'

    @staticmethod
    def live(events, cross_user=False):
        # A deliberate correction supersedes every observed exact decision.
        # Concurrent, unseen corrections remain disputed instead of being lost.
        return Snapshot.live(events, cross_user=cross_user or bool(events and events[0]['kind'] == 'case'))


class SimpleKnowledgeStore(_KnowledgeStore):
    replace_observed_cases = True

    def decide(self, row, final, action='Исправить'):
        from excel.knowledge import case_key
        key = case_key(row.get('code', ''), row['source'], row.get('factory', ''))
        disabled = self.snapshot().entries.get(key, {}).get('status') == 'DISABLED'
        event = super().decide(row, final, action)
        if disabled:
            self.control(key, False, 'Инженер заново проверил и подтвердил точное решение')
        return event
    def __init__(self, local_dir, shared_dir=None, user=None):
        if shared_dir is None:
            configured = os.environ.get('MTR_KNOWLEDGE_DIR', '').strip()
            if configured:
                shared_dir = configured
        if shared_dir is not None:
            Path(shared_dir).mkdir(parents=True, exist_ok=True)
        super().__init__(local_dir, shared_dir, user)

    def snapshot(self):
        # In this product mode filesystem ACLs define who may edit the knowledge
        # base.  Therefore control events from every engineer who could write an
        # event are effective, not only from the first manifest administrator.
        events = self.events()
        editors = {self.user}
        editors.update(e.get('user') for e in events if e.get('user'))
        snapshot = SimpleSnapshot(events, sorted(editors))
        snapshot.observed_events = events
        return snapshot

    def administrate(self, key, reason, **changes):
        """Allow knowledge edits to every engineer with write access to the folder.

        The event still records the Windows identity and complete history, so an
        incorrect edit is auditable and can be superseded later.
        """
        self.check_shared()
        if not str(reason or '').strip():
            raise ValueError('Укажите основание изменения.')
        events = [e for e in contextual_events(self.events())
                  if e['kind'] == 'control' and e['key'] == key and not e.get('usage')]
        live = Snapshot.live(events, cross_user=True)
        inherited = ({k: live[0][k]
                      for k in ('classification', 'scope', 'protected',
                                'requested_status', 'settings', 'decision_override')
                      if k in live[0]} if len(live) == 1 else {})
        changes = dict(inherited, **changes)
        previous = self.snapshot().entries.get(key, {}).get('status', '')
        return self.emit(dict(kind='control', key=key,
                              value=changes.pop('value', 'ENABLED'),
                              reason=str(reason).strip(), previous_status=previous,
                              supersedes=[e['id'] for e in events], **changes))
