"""Shared immutable decisions; SQLite is a disposable LOCAL index, never on SMB.

The department share (including its ACLs) is the trust boundary. No Internet IO.
Events have causal supersession, independent-user votes and content checksums.
"""
from collections import defaultdict
from contextlib import contextmanager
from datetime import datetime, timezone
import difflib
import getpass
import gzip
import hashlib
import json
import logging
import os
from pathlib import Path
import shutil
import sqlite3
import tempfile
import uuid
import zipfile

SCHEMA = 1
VERSION = '1.4 RC1'
DEFAULT_SETTINGS = dict(rule_active_votes=2, rule_trusted_votes=3, case_trusted_votes=3, similarity_min=50, analog_limit=20)

def canonical(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(',', ':'))

def digest(value):
    return hashlib.sha256(canonical(value).encode('utf-8')).hexdigest()

def normalize(value):
    # Do not equate Cyrillic/Latin characters, units, punctuation or model numbers.
    return ' '.join(str(value or '').split())

def code_key(value):
    from excel.rules import code_candidates
    candidates = code_candidates(str(value or '').strip())
    return str(candidates[-1]) if candidates else ''

def identity():
    domain = os.environ.get('USERDOMAIN', '')
    return ((domain + '\\') if domain else '') + getpass.getuser()

def case_key(code, source, factory=''):
    code = code_key(code)
    return 'code:' + code if code else 'text:' + digest([normalize(source), normalize(factory)])

def fragments(before, after):
    removed, restored = [], []
    # Exact fragments, not inferred votes for all unchanged tokens.
    for op, a, b, c, d in difflib.SequenceMatcher(None, before, after, autojunk=False).get_opcodes():
        if op in ('replace', 'delete') and before[a:b].strip():
            removed.append(before[a:b])
        if op in ('replace', 'insert') and after[c:d].strip():
            restored.append(after[c:d])
    return removed, restored

def atomic_copy(source, target):
    target.parent.mkdir(parents=True, exist_ok=True)
    temp = target.with_name('.' + target.name + '.' + uuid.uuid4().hex + '.tmp')
    try:
        with open(source, 'rb') as src, open(temp, 'xb') as dst:
            shutil.copyfileobj(src, dst)
            dst.flush(); os.fsync(dst.fileno())
        # Unique event/session IDs: a different existing payload is corruption.
        if target.exists():
            if file_hash(source) != file_hash(target):
                raise ValueError('Коллизия идентификатора: ' + target.name)
        else:
            os.replace(temp, target)
    finally:
        temp.unlink(missing_ok=True)

def file_hash(path):
    h = hashlib.sha256()
    with open(path, 'rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()

def atomic_json(path, value):
    path = Path(path); path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix='.mtr-', dir=path.parent)
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            f.write(canonical(value)); f.flush(); os.fsync(f.fileno())
        os.replace(tmp, path)
    finally:
        Path(tmp).unlink(missing_ok=True)

def uid(value):
    return str(uuid.UUID(str(value)))

class KnowledgeStore:
    def __init__(self, local_dir, shared_dir=None, user=None):
        self.local = Path(local_dir)
        self.local.mkdir(parents=True, exist_ok=True)
        self.config_path = self.local / 'knowledge_config.json'
        self.config = json.loads(self.config_path.read_text('utf-8')) if self.config_path.exists() else {}
        self.user = user or identity()
        if shared_dir is not None:
            self.configure(shared_dir)
        self.shared = Path(self.config['folder']) if self.config.get('folder') else None
        self.store_id = self.config.get('store_id', 'unconfigured')
        self.cache = self.local / self.store_id
        self.cache.mkdir(parents=True, exist_ok=True)
        self.db_path = self.cache / 'index.sqlite3'
        with self.db() as db:
            db.execute('PRAGMA journal_mode=WAL')
            db.executescript('''
                CREATE TABLE IF NOT EXISTS events(id TEXT PRIMARY KEY, body TEXT NOT NULL, pending INTEGER NOT NULL);
                CREATE INDEX IF NOT EXISTS event_owner_key ON events(json_extract(body,'$.key'),json_extract(body,'$.user'));
                CREATE TABLE IF NOT EXISTS rows(id TEXT PRIMARY KEY, session TEXT NOT NULL, code TEXT, source TEXT, factory TEXT, body TEXT NOT NULL);
                CREATE INDEX IF NOT EXISTS row_match ON rows(code,source,factory);
                CREATE INDEX IF NOT EXISTS row_session ON rows(session);
                CREATE TABLE IF NOT EXISTS session_meta(id TEXT PRIMARY KEY, body TEXT NOT NULL);
                CREATE TABLE IF NOT EXISTS features(id TEXT PRIMARY KEY, session TEXT, factory TEXT, category TEXT, body TEXT);
                CREATE INDEX IF NOT EXISTS feature_scope ON features(session,factory,category);
                CREATE TABLE IF NOT EXISTS tokens(session TEXT, token TEXT, id TEXT, PRIMARY KEY(session,token,id));
                CREATE INDEX IF NOT EXISTS token_lookup ON tokens(token,session,id);
                CREATE INDEX IF NOT EXISTS token_id ON tokens(id);
                CREATE TABLE IF NOT EXISTS catalog_ready(session TEXT PRIMARY KEY, version TEXT);
                CREATE TABLE IF NOT EXISTS sessions(id TEXT PRIMARY KEY, pending INTEGER NOT NULL);
            ''')
        self.last_sync = {'online': False, 'pending': self.pending(), 'errors': [], 'message': 'Синхронизация ещё не выполнялась'}

    @contextmanager
    def db(self):
        db = sqlite3.connect(self.db_path, timeout=30)
        try:
            with db:
                yield db
        finally:
            db.close()

    def configure(self, folder):
        folder = Path(folder)
        if not folder.is_dir():
            raise ValueError('Выберите существующую общую папку подразделения.')
        manifest = folder / 'mtr-knowledge.json'
        if not manifest.exists():
            value = {'schema': SCHEMA, 'store_id': str(uuid.uuid4()), 'admins': [self.user]}
            # O_EXCL prevents two initializers from creating different identities.
            try:
                fd = os.open(manifest, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                with os.fdopen(fd, 'w', encoding='utf-8') as f:
                    f.write(canonical(value)); f.flush(); os.fsync(f.fileno())
            except FileExistsError:
                pass
        value = json.loads(manifest.read_text('utf-8'))
        if value.get('schema') != SCHEMA:
            raise ValueError('Неподдерживаемая версия общей базы.')
        uid(value['store_id'])
        self.config = dict(value, folder=str(folder.absolute()))
        atomic_json(self.config_path, self.config)

    def check_shared(self):
        if self.shared is None:
            raise OSError('Общая папка не настроена')
        # Do not recreate a disconnected share or silently start a second database.
        value = json.loads((self.shared / 'mtr-knowledge.json').read_text('utf-8'))
        if value.get('schema') != SCHEMA or value.get('store_id') != self.store_id:
            raise ValueError('В выбранной папке другая база или версия формата.')
        self.config['admins'] = value.get('admins', [])

    def pending(self):
        with self.db() as db:
            return db.execute('SELECT count(*) FROM events WHERE pending=1').fetchone()[0] + db.execute('SELECT count(*) FROM sessions WHERE pending=1').fetchone()[0]

    def events(self):
        with self.db() as db:
            return [json.loads(x[0]) for x in db.execute('SELECT body FROM events ORDER BY id')]

    def own_events(self, key, kind):
        with self.db() as db:
            return [json.loads(x[0]) for x in db.execute(
                "SELECT body FROM events WHERE json_extract(body,'$.key')=? AND json_extract(body,'$.user')=? AND json_extract(body,'$.kind')=?", (key, self.user, kind))]

    def snapshot(self):
        return Snapshot(self.events(), self.config.get('admins', []))

    def emit(self, payload):
        if not self.shared:
            raise ValueError('Сначала укажите единую общую папку базы знаний.')
        event = dict(payload, schema=SCHEMA, store_id=self.store_id, id=str(uuid.uuid4()),
                     user=self.user, timestamp=datetime.now(timezone.utc).isoformat(), version=VERSION)
        self.validate(event)
        with self.db() as db:
            db.execute('INSERT INTO events VALUES(?,?,1)', (event['id'], canonical(event)))
        logging.info('Knowledge event %s %s by %s', event['id'], event['kind'], self.user)
        return event

    def validate(self, event):
        uid(event['id'])
        if event.get('schema') != SCHEMA or event.get('store_id') != self.store_id:
            raise ValueError('Несовместимое событие')
        if not isinstance(event.get('user'), str) or not event['user']:
            raise ValueError('Нет автора решения')
        if event.get('kind') not in ('case', 'rule', 'control'):
            raise ValueError('Неизвестный тип решения')
        if not isinstance(event.get('key'), str) or not event['key']:
            raise ValueError('Нет ключа решения')
        if event['kind'] in ('case', 'rule'):
            if not isinstance(event.get('value'), str):
                raise ValueError('Неверный результат решения')
            for old in event.get('supersedes', []):
                uid(old)
        if event['kind']=='control':
            from excel.semantics import CLASSES
            if event.get('settings'):validate_settings(event['settings'])
            if event.get('classification') and event['classification'] not in CLASSES:raise ValueError('Неизвестная классификация')
            if event.get('requested_status') not in (None,'CANDIDATE','ACTIVE','TRUSTED','DISABLED'):raise ValueError('Неизвестный статус')
            scope=event.get('scope')
            if scope is not None and (not scope.get('factory') or not scope.get('category') or scope.get('role') not in ('изделие','комплектующее')):raise ValueError('Неполная область применения')
        if event['kind'] == 'case':
            if event['key'] != case_key(event['code'], event['source'], event['factory']):
                raise ValueError('Неверный ключ ресурса')
        elif event['kind'] == 'rule':
            scope = event['scope']
            if event['value'] not in ('KEEP', 'DELETE') or not scope.get('factory') or not scope.get('category') or scope.get('role') not in ('изделие', 'комплектующее') or not event.get('fragment', '').strip():
                raise ValueError('Укажите фрагмент, завод, категорию и роль правила.')
            if event['key'] != 'rule:' + digest([event['fragment'], scope]):
                raise ValueError('Неверный ключ правила')
        elif event.get('value') not in ('DISABLED', 'ENABLED'):
            raise ValueError('Неверное состояние правила')

    def decide(self, row, final, action='Исправить'):
        final = str(final)
        key = case_key(row.get('code', ''), row['source'], row.get('factory', ''))
        own = self.own_events(key, 'case')
        live = Snapshot.live(own)
        if len(live) == 1 and live[0]['value'] == final and not row.get('explicit_feedback'):
            return None  # Same user + same decision, even across sessions/imports.
        removed, restored = fragments(row.get('automatic', ''), final)
        original_deleted, _ = fragments(row['source'], final)
        from excel.semantics import feedback_facts, features
        classification = row.get('classification') or features(row['source'],row.get('factory',''))
        return self.emit(dict(classification=classification,facts=feedback_facts(row,final,action,classification),batch_id=row.get('batch_id',''),kind='case', key=key, value=final, source=row['source'], code=row.get('code', ''),
                              factory=row.get('factory', ''), automatic=row.get('automatic', ''),
                              row_id=row['id'], session=row['session'], action=action,
                              removed=removed, restored=restored, original_deleted=original_deleted,
                              provenance=row.get('provenance', {}), supersedes=[e['id'] for e in own]))

    def propose_rule(self, row, fragment, action, category, role='изделие'):
        scope = {'factory': normalize(row.get('factory', '')), 'category': normalize(category), 'role': role}
        fragment = str(fragment).strip()
        if fragment not in row['source']:
            raise ValueError('Фрагмент должен буквально присутствовать в исходной строке.')
        key = 'rule:' + digest([fragment, scope])
        own = self.own_events(key, 'rule')
        if len(Snapshot.live(own)) == 1 and Snapshot.live(own)[0]['value'] == action:
            return None
        return self.emit(dict(kind='rule', key=key, value=action, fragment=fragment, scope=scope,
                              example=row['source'], row_id=row['id'], session=row['session'],
                              supersedes=[e['id'] for e in own]))

    def control(self, key, disabled, reason):
        return self.administrate(key,reason,value='DISABLED' if disabled else 'ENABLED',requested_status='DISABLED' if disabled else None)

    def sync(self, auto_backup=True):
        errors = []
        try:
            self.check_shared()
            with self.db() as db:
                pending_sessions = [x[0] for x in db.execute('SELECT id FROM sessions WHERE pending=1')]
            for sid in pending_sessions:
                atomic_copy(self.cache / 'sessions' / (sid + '.jsonl.gz'), self.shared / 'sessions' / (sid + '.jsonl.gz'))
                with self.db() as db:
                    db.execute('UPDATE sessions SET pending=0 WHERE id=?', (sid,))
            with self.db() as db:
                pending = [json.loads(x[0]) for x in db.execute('SELECT body FROM events WHERE pending=1')]
            for event in pending:
                path = self.cache / 'outbox' / (event['id'] + '.json')
                atomic_json(path, {'event': event, 'sha256': digest(event)})
                atomic_copy(path, self.shared / 'events' / event['id'][:2] / path.name)
                with self.db() as db:
                    db.execute('UPDATE events SET pending=0 WHERE id=?', (event['id'],))
                path.unlink(missing_ok=True)
            known = {e['id'] for e in self.events()}
            folder = self.shared / 'events'
            if folder.exists():
                for path in folder.glob('*/*.json'):
                    if path.stem in known:
                        continue
                    try:
                        envelope = json.loads(path.read_text('utf-8'))
                        event = envelope['event']
                        self.validate(event)
                        if event['id'] != path.stem or digest(event) != envelope['sha256']:
                            raise ValueError('Контрольная сумма не совпала')
                        with self.db() as db:
                            db.execute('INSERT OR IGNORE INTO events VALUES(?,?,0)', (event['id'], canonical(event)))
                    except Exception as exc:
                        errors.append(path.name + ': ' + str(exc))
                        logging.exception('Invalid shared knowledge event %s', path.name)
            status = {'online': True, 'errors': errors, 'message': 'Синхронизировано' if not errors else 'Есть повреждённые события; проверьте журнал'}
        except Exception as exc:
            logging.warning('Knowledge offline: %s', exc)
            status = {'online': False, 'errors': [str(exc)], 'message': 'Нет доступа к общей базе. Используется последний кэш; решения ожидают отправки.'}
        status.update(pending=self.pending(), time=datetime.now(timezone.utc).isoformat())
        self.last_sync = status
        if auto_backup and status['online'] and not status['errors'] and not status['pending']:
            try:
                signature = digest([self.snapshot().version, sorted(p.name for p in (self.shared / 'sessions').glob('*.jsonl.gz'))])[:20]
                backup_path = self.shared / 'backups' / (datetime.now(timezone.utc).strftime('%Y-%m-%d') + '-' + signature + '.zip')
                if not backup_path.exists():
                    self._backup_snapshot(backup_path)
            except Exception as exc:
                status['errors'].append('Автоматическая копия: ' + str(exc))
                status['message'] = 'Знания синхронизированы; ошибка резервной копии'
                logging.exception('Automatic backup failed')
        atomic_json(self.cache / 'sync_status.json', status)
        return status

    def administrate(self,key,reason,**changes):
        self.check_shared()
        if self.user not in self.config.get('admins',[]):raise PermissionError('Требуются права администратора общей базы.')
        if not reason.strip():raise ValueError('Укажите основание изменения.')
        events=[e for e in self.events() if e['kind']=='control' and e['key']==key and not e.get('usage')]
        live=Snapshot.live([e for e in events if e['user'] in self.config.get('admins',[])],cross_user=True)
        inherited={k:live[0][k] for k in ('classification','scope','protected','requested_status','settings') if k in live[0]} if len(live)==1 else {}
        changes=dict(inherited,**changes);previous=self.snapshot().entries.get(key,{}).get('status','')
        return self.emit(dict(kind='control',key=key,value=changes.pop('value','ENABLED'),reason=reason,previous_status=previous,supersedes=[e['id'] for e in events],**changes))

    def configure_trust(self,settings,reason):
        settings=dict(DEFAULT_SETTINGS,**settings);validate_settings(settings)
        return self.administrate('settings',reason,settings=settings)

    def metadata(self,sid):
        with self.db() as db:row=db.execute('SELECT body FROM session_meta WHERE id=?',(sid,)).fetchone()
        return json.loads(row[0]) if row else {'session':sid}

    def register_output(self,sid,path):
        meta=dict(self.metadata(sid),output_file=str(Path(path).absolute()))
        with self.db() as db:db.execute('INSERT OR REPLACE INTO session_meta VALUES(?,?)',(sid,canonical(meta)))

    def new_session(self, source_file, knowledge_version):
        return SessionWriter(self, source_file, knowledge_version)

    def load_session(self, sid):
        sid = uid(sid)
        with self.db() as db:
            if db.execute('SELECT 1 FROM sessions WHERE id=?', (sid,)).fetchone():
                return
        self.check_shared()
        source = self.shared / 'sessions' / (sid + '.jsonl.gz')
        local = self.cache / 'sessions' / source.name
        local.unlink(missing_ok=True)
        atomic_copy(source, local)
        count, h = 0, hashlib.sha256()
        # All rows and footer verified before the transaction is committed.
        with self.db() as db, gzip.open(local, 'rt', encoding='utf-8') as f:
            head = json.loads(next(f))
            if head.get('session') != sid or head.get('store_id') != self.store_id or head.get('schema') != SCHEMA:
                raise ValueError('Несовместимый исходный сеанс')
            footer = None
            for line in f:
                obj = json.loads(line)
                if 'footer' in obj:
                    footer = obj
                    if f.read():
                        raise ValueError('Данные после конца сеанса')
                    break
                uid(obj['id'])
                if obj['session'] != sid:
                    raise ValueError('Неверный идентификатор сеанса строки')
                h.update(line.encode('utf-8')); count += 1
                db.execute('INSERT INTO rows VALUES(?,?,?,?,?,?)', (obj['id'], sid, code_key(obj['code']), normalize(obj['source']), normalize(obj['factory']), canonical(obj)))
            if footer != {'footer': count, 'sha256': h.hexdigest()}:
                raise ValueError('Неполный или повреждённый сеанс')
            db.execute('INSERT INTO sessions VALUES(?,0)', (sid,))
            db.execute('INSERT OR REPLACE INTO session_meta VALUES(?,?)',(sid,canonical(head)))

    def row(self, row_id):
        with self.db() as db:
            hit = db.execute('SELECT body FROM rows WHERE id=?', (row_id,)).fetchone()
        return json.loads(hit[0]) if hit else None

    def match_rows(self, code, source, factory='', sid=None):
        with self.db() as db:
            sql = 'SELECT body FROM rows WHERE code=? AND source=?'
            params = [code_key(code), normalize(source)]
            if sid:
                sql += ' AND session=?'; params.append(sid)
            if factory:
                sql += ' AND factory=?'; params.append(normalize(factory))
            return [json.loads(x[0]) for x in db.execute(sql, params)]

    def session_rows(self, sid, query='', offset=0, limit=200):
        with self.db() as db:
            rows = db.execute('SELECT body FROM rows WHERE session=? AND (source LIKE ? OR code LIKE ?) ORDER BY rowid LIMIT ? OFFSET ?',
                              (sid, '%' + query + '%', '%' + query + '%', limit, offset)).fetchall()
        return [json.loads(x[0]) for x in rows]

    def backup(self, target):
        self.sync(auto_backup=False)
        self.check_shared()
        if self.pending() or self.last_sync['errors']:
            raise ValueError('Резервная копия требует успешной синхронизации без ошибок.')
        return self._backup_snapshot(target)

    def _backup_snapshot(self, target):
        target = Path(target)
        target.parent.mkdir(parents=True, exist_ok=True)
        fd, temporary = tempfile.mkstemp(prefix='.mtr-backup-', dir=target.parent); os.close(fd)
        try:
            paths = [self.shared / 'mtr-knowledge.json'] + sorted((self.shared / 'events').glob('*/*.json')) + sorted((self.shared / 'sessions').glob('*.jsonl.gz'))
            hashes = {}
            with zipfile.ZipFile(temporary, 'w', zipfile.ZIP_DEFLATED) as z:
                for path in paths:
                    rel = path.relative_to(self.shared).as_posix()
                    data = path.read_bytes(); hashes[rel] = hashlib.sha256(data).hexdigest(); z.writestr(rel, data)
                z.writestr('backup-manifest.json', canonical(hashes))
            verify_backup(temporary)
            os.replace(temporary, target)
        finally:
            Path(temporary).unlink(missing_ok=True)
        return target

class SessionWriter:
    def __init__(self, store, source_file, knowledge_version):
        self.store = store; self.id = str(uuid.uuid4()); self.count = 0
        self.path = store.cache / 'sessions' / (self.id + '.jsonl.gz')
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.temp = self.path.with_suffix('.tmp')
        self.f = gzip.open(self.temp, 'wt', encoding='utf-8', newline='\n')
        self.meta=dict(schema=SCHEMA,store_id=store.store_id,session=self.id,file=str(Path(source_file).absolute()),knowledge=knowledge_version,version=VERSION,timestamp=datetime.now(timezone.utc).isoformat())
        self.usage=defaultdict(int);self.statistics=defaultdict(int)
        self.f.write(canonical(self.meta)+'\n')
        self.h = hashlib.sha256(); self.db = sqlite3.connect(store.db_path, timeout=30)
        self.db.execute('BEGIN'); self.closed = False

    def add(self, sheet, row_number, code, source, factory, result):
        info=result.get('knowledge',{})
        for key in set(info.get('rules',[])+([info['key']] if info.get('key') else [])):self.usage[key]+=1
        self.statistics[result['status']]+=1
        self.statistics['expert' if info.get('status','LEGACY')!='LEGACY' else 'legacy']+=1
        row = dict(classification=result.get('classification'),source_file=self.meta['file'],id=str(uuid.uuid4()), session=self.id, sheet=sheet, row=row_number, code=code, source=source,
                   factory=factory, automatic=result['text'], status=result['status'], reason=result['reason'],
                   detected_factory=result['factory'], removed=result['removed'], provenance=dict(result.get('knowledge', {}), trace=result.get('trace', []), protected=result.get('protected', [])))
        line = canonical(row) + '\n'; self.f.write(line); self.h.update(line.encode('utf-8')); self.count += 1
        self.db.execute('INSERT INTO rows VALUES(?,?,?,?,?,?)', (row['id'], self.id, code_key(code), normalize(source), normalize(factory), canonical(row)))
        return row

    def finish(self):
        self.f.write(canonical({'footer': self.count, 'sha256': self.h.hexdigest()}) + '\n'); self.f.close()
        os.replace(self.temp, self.path)
        self.db.execute('INSERT INTO session_meta VALUES(?,?)',(self.id,canonical(dict(self.meta,statistics=dict(self.statistics)))))
        if self.usage:
            event=dict(schema=SCHEMA,store_id=self.store.store_id,id=str(uuid.uuid4()),user=self.store.user,timestamp=datetime.now(timezone.utc).isoformat(),version=VERSION,kind='control',key='usage:'+self.id,value='ENABLED',usage=dict(self.usage),supersedes=[])
            self.store.validate(event);self.db.execute('INSERT INTO events VALUES(?,?,1)',(event['id'],canonical(event)))
        self.db.execute('INSERT INTO sessions VALUES(?,1)', (self.id,)); self.db.commit(); self.db.close(); self.closed = True

    def abort(self):
        if not self.closed:
            self.f.close(); self.db.rollback(); self.db.close(); self.temp.unlink(missing_ok=True); self.path.unlink(missing_ok=True); self.closed = True

def validate_settings(settings):
    if set(settings)!=set(DEFAULT_SETTINGS):raise ValueError('Неверный набор настроек')
    if any(type(v) is not int for v in settings.values()):raise ValueError('Пороги должны быть целыми числами')
    if not 1<=settings['rule_active_votes']<=settings['rule_trusted_votes']<=100:raise ValueError('Активный порог должен быть не выше доверенного')
    if not 1<=settings['case_trusted_votes']<=100 or not 0<=settings['similarity_min']<=100 or not 1<=settings['analog_limit']<=100:raise ValueError('Недопустимый порог')

class Snapshot:
    def __init__(self,events,admins=()):
        from excel.semantics import folded,factory_key,feedback_facts
        self.version=digest([sorted(e['id'] for e in events),sorted(admins)])[:16]
        self.settings=dict(DEFAULT_SETTINGS);self.entries={};self.rules=defaultdict(list);self.protected_rules=[]
        groups=defaultdict(list);controls=defaultdict(list);usage=defaultdict(int);dates={}
        for e in events:
            (controls if e['kind']=='control' else groups)[e['key']].append(e)
            for key,count in e.get('usage',{}).items():usage[key]+=int(count);dates[key]=max(dates.get(key,''),e['timestamp'])
        def actions(key):return self.live([e for e in controls[key] if e['user'] in admins],cross_user=True)
        config=actions('settings');self.settings_conflict=len({canonical(e.get('settings',{})) for e in config})>1
        if config and not self.settings_conflict:self.settings.update(config[0].get('settings',{}))
        self.static_disabled={k for k in controls if k.startswith('static:') and any(e['value']=='DISABLED' for e in actions(k))}
        for key,history in list(groups.items()):
            if history[0]['kind']!='case':continue
            effective={e['id'] for e in self.live(history)}
            for event in history:
                facts=event.get('facts')
                if facts is None:facts=feedback_facts(event,event['value'],event.get('action','Исправить'))
                for fact in facts:
                    scope=fact['scope'];derived_key='learn:'+digest([folded(fact['fragment']),factory_key(scope['factory']),folded(scope['category']),scope['role']])
                    groups[derived_key].append(dict(event,id=event['id']+':'+digest([derived_key,fact['action']])[:12],kind='rule',key=derived_key,value=fact['action'],fragment=fact['fragment'],scope=scope,example=event['source'],classification=fact.get('classification','не определено'),eligible=fact.get('eligible',False),origin_event=event['id'],retired=event['id'] not in effective,supersedes=[]))
        for key,history in groups.items():
            live=[e for e in self.live(history) if not e.get('retired')];variants=defaultdict(set)
            for e in live:variants[e['value']].add(e['user'])
            ranked=sorted(variants,key=lambda v:(-len(variants[v]),v));value=ranked[0] if ranked else ''
            count=len(variants.get(value,set()));opposition=sum(len(v) for v in variants.values())-count;kind=history[0]['kind']
            threshold=self.settings['case_trusted_votes' if kind=='case' else 'rule_trusted_votes']
            state='DISABLED' if not live else 'DISPUTED' if len(variants)>1 else 'TRUSTED' if count>=threshold else 'ACTIVE' if count>=(1 if kind=='case' else self.settings['rule_active_votes']) else 'CANDIDATE'
            sample=dict(sorted(live or history,key=lambda e:e['id'])[0]);admin=actions(key)
            edits={(e.get('classification'),canonical(e.get('scope')),e.get('requested_status'),e.get('protected')) for e in admin}
            if len(edits)>1:state='DISPUTED'
            if any(e['value']=='DISABLED' for e in admin):state='DISABLED'
            elif len(edits)<=1:
                for e in admin:
                    for field in ('classification','scope','protected'):
                        if field in e:sample[field]=e[field]
                    if e.get('requested_status') and len(variants)==1:state=e['requested_status']
            if kind=='rule' and (not sample['scope'].get('factory') or sample['scope'].get('category')=='не определено') and state not in ('DISABLED','DISPUTED'):state='CANDIDATE'
            entry=dict(key=key,kind=kind,status=state,value=value,confirmations=count,opposition=opposition,score=round(min(100,100*count/max(1,threshold))*count/max(1,count+opposition)),variants={v:sorted(u) for v,u in variants.items()},history=history+controls[key],sample=sample,live=live,applications=usage[key],last_used=dates.get(key,''),first_seen=min(e['timestamp'] for e in history),affected_rows=len({e.get('row_id') for e in history if e.get('row_id')}))
            self.entries[key]=entry
            if kind=='rule':
                self.rules[sample['scope']['factory']].append(entry)
                if sample.get('protected') and value=='KEEP' and state in ('ACTIVE','TRUSTED'):self.protected_rules.append(entry)

    @staticmethod
    def live(events, cross_user=False):
        by_id = {e['id']: e for e in events}
        superseded = set()
        for e in events:
            for old in e.get('supersedes', []):
                previous = by_id.get(old)
                if previous and previous['key'] == e['key'] and (cross_user or previous['user'] == e['user']):
                    superseded.add(old)
        return [e for e in events if e['id'] not in superseded]


def verify_backup(path):
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if len(set(names)) != len(names):
            raise ValueError('Повторяющиеся имена в архиве')
        manifest = json.loads(z.read('backup-manifest.json'))
        if set(names) != set(manifest) | {'backup-manifest.json'}:
            raise ValueError('Неверный состав резервной копии')
        for name, checksum in manifest.items():
            parts = name.split('/')
            allowed = name == 'mtr-knowledge.json' or (len(parts) == 3 and parts[0] == 'events' and parts[-1].endswith('.json')) or (len(parts) == 2 and parts[0] == 'sessions' and parts[-1].endswith('.jsonl.gz'))
            if not allowed or '..' in parts or '\\' in name or ':' in name or name.startswith('/'):
                raise ValueError('Недопустимый путь в резервной копии')
            if hashlib.sha256(z.read(name)).hexdigest() != checksum:
                raise ValueError('Повреждена резервная копия: ' + name)
    return True

def restore_backup(path, empty_folder):
    verify_backup(path)
    folder = Path(empty_folder)
    folder.mkdir(parents=True, exist_ok=True)
    if any(folder.iterdir()):
        raise ValueError('Восстановление разрешено только в пустую папку.')
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if name != 'backup-manifest.json':
                target = folder / name; target.parent.mkdir(parents=True, exist_ok=True); target.write_bytes(z.read(name))
    return folder
