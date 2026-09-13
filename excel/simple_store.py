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
import os

from excel.knowledge import KnowledgeStore as _KnowledgeStore, Snapshot


class SimpleKnowledgeStore(_KnowledgeStore):
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
        return Snapshot(events, sorted(editors))

    def administrate(self, key, reason, **changes):
        """Allow knowledge edits to every engineer with write access to the folder.

        The event still records the Windows identity and complete history, so an
        incorrect edit is auditable and can be superseded later.
        """
        self.check_shared()
        if not str(reason or '').strip():
            raise ValueError('Укажите основание изменения.')
        events = [e for e in self.events()
                  if e['kind'] == 'control' and e['key'] == key and not e.get('usage')]
        live = Snapshot.live(events, cross_user=True)
        inherited = ({k: live[0][k]
                      for k in ('classification', 'scope', 'protected',
                                'requested_status', 'settings')
                      if k in live[0]} if len(live) == 1 else {})
        changes = dict(inherited, **changes)
        previous = self.snapshot().entries.get(key, {}).get('status', '')
        return self.emit(dict(kind='control', key=key,
                              value=changes.pop('value', 'ENABLED'),
                              reason=str(reason).strip(), previous_status=previous,
                              supersedes=[e['id'] for e in events], **changes))
