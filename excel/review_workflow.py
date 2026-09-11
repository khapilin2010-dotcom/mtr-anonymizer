"""Explicit feedback, contextual search and batch review with preflight."""
from collections import defaultdict
import json
import re
import uuid
from excel.knowledge import digest,canonical,case_key,code_key
from excel.excel_engine import Anonymizer,review_tokens
from excel.expert_engine import ExpertAnonymizer
from excel.semantics import features,feedback_facts,folded,literal_spans,similarity,fragment_class
class StaleReview(ValueError):pass
class ReviewWorkflow:
    def __init__(self,store,base=None):self.store=store;self.base=base or Anonymizer()
    def prepare(self,proposals,snapshot=None):
        snapshot=snapshot or self.store.snapshot();engine=ExpertAnonymizer(snapshot,self.base);batch=str(uuid.uuid4());items=[];issues=[];by_key=defaultdict(set)
        for proposal in proposals:
            row=dict(proposal['row']);final=str(proposal['final']);action=proposal.get('action','Исправить')
            if not final.strip():issues.append(dict(row_id=row['id'],reason='Пустое наименование'));continue
            row['classification']=row.get('classification') or features(row['source'],row.get('factory',''),row.get('code',''),self.base)
            row['batch_id']=batch
            if proposal.get('explicit_feedback'):row['explicit_feedback']=proposal['explicit_feedback']
            key=case_key(row.get('code',''),row['source'],row.get('factory',''));entry=snapshot.entries.get(key)
            warnings,_,_=engine.assess(row['source'],final,row.get('code',''),row.get('factory',''),row['classification'])
            if entry and entry['status']!='DISABLED' and any(e['user']!=self.store.user and e['value']!=final for e in entry['live']):warnings.append('Другой инженер подтвердил иной результат')
            facts=feedback_facts(row,final,action,row['classification']);items.append(dict(row=row,final=final,action=action,facts=facts,key=key,warnings=warnings));by_key[key].add(final)
        ambiguous={k for k,v in by_key.items() if len(v)>1};accepted=[];conflicts=[]
        for item in items:
            if item['key'] in ambiguous:issues.append(dict(row_id=item['row']['id'],reason='Разные результаты для одного ресурса в одном пакете'));continue
            accepted.append(item)
            if item['warnings']:conflicts.append(dict(row_id=item['row']['id'],code=item['row'].get('code',''),warnings=item['warnings']))
        return dict(version=snapshot.version,items=accepted,issues=issues,conflicts=conflicts,fingerprint=digest(accepted),statistics=dict(rows=len(accepted),conflicts=len(conflicts),candidates=sum(len(x['facts']) for x in accepted)))
    def commit(self,plan,accept_conflicts=False,cancel=None):
        if plan['conflicts'] and not accept_conflicts:return dict(accepted=0,repeated=0,issues=plan['issues'],conflicts=plan['conflicts'],needs_confirmation=True)
        if self.store.snapshot().version!=plan['version']:raise StaleReview('База изменилась. Повторите предварительную проверку.')
        if digest(plan['items'])!=plan['fingerprint']:raise ValueError('Состав пакета изменился')
        accepted=repeated=0
        for item in plan['items']:
            if cancel and cancel.is_set():break
            row=dict(item['row'],provenance=dict(item['row'].get('provenance',{}),conflicts_acknowledged=bool(item['warnings'] and accept_conflicts),review_warnings=item['warnings'],prior_knowledge=plan['version']))
            event=self.store.decide(row,item['final'],item['action']);accepted+=event is not None;repeated+=event is None
        return dict(accepted=accepted,repeated=repeated,issues=plan['issues'],conflicts=plan['conflicts'],needs_confirmation=False,sync=self.store.sync(),cancelled=bool(cancel and cancel.is_set()))
    def current(self,row,snapshot=None):
        snapshot=snapshot or self.store.snapshot()
        if getattr(self,'_version',None)!=snapshot.version:self._engine=ExpertAnonymizer(snapshot,self.base);self._version=snapshot.version
        result=self._engine.anonymize(row['source'],row.get('code',''),row.get('factory',''));entry=snapshot.entries.get(case_key(row.get('code',''),row['source'],row.get('factory','')))
        own=[e for e in entry['live'] if e['user']==self.store.user] if entry else []
        return dict(row,result=result,final=result['text'],proposal=own[0]['value'] if len(own)==1 else result.get('knowledge',{}).get('proposed',''),entry=entry,kept=result.get('protected',[]),classification=result['classification'])
    def index_session(self,sid,progress=None,cancel=None):
        with self.store.db() as db:
            if db.execute('SELECT 1 FROM catalog_ready WHERE session=? AND version=?',(sid,'semantics-1')).fetchone():return
            for count,(rid,body) in enumerate(db.execute('SELECT id,body FROM rows WHERE session=?',(sid,)),1):
                if cancel and cancel.is_set():raise InterruptedError('Индексация отменена')
                row=json.loads(body);info=row.get('classification') or features(row['source'],row.get('factory',''),row.get('code',''),self.base);info['review_fragments']=review_tokens(row['automatic'])
                light={k:row.get(k,'') for k in ('id','session','sheet','row','code','source','factory','automatic','status')}
                db.execute('INSERT OR REPLACE INTO features VALUES(?,?,?,?,?)',(rid,sid,info['factory_key'],info['category'],canonical(dict(row=light,features=info))))
                db.execute('DELETE FROM tokens WHERE id=?',(rid,));db.executemany('INSERT OR IGNORE INTO tokens VALUES(?,?,?)',[(sid,t,rid) for t in info['tokens']])
                if progress and count%500==0:progress('Индекс аналогов: '+str(count)+' строк')
            db.execute('INSERT OR REPLACE INTO catalog_ready VALUES(?,?)',(sid,'semantics-1'))
    def similar(self,source,code='',factory='',sid=None,exclude='',limit=None):
        snapshot=self.store.snapshot();query=features(source,factory,code,self.base);ids=set();result=[]
        with self.store.db() as db:
            restriction=' AND session=?' if sid else '';suffix=[sid] if sid else []
            if code:ids.update(x[0] for x in db.execute('SELECT id FROM rows WHERE code=?'+restriction+' LIMIT 500',[code_key(code)]+suffix))
            counts=[(db.execute('SELECT count(*) FROM tokens WHERE token=?'+restriction,[t]+suffix).fetchone()[0],t) for t in query['tokens']]
            for count,t in sorted((n,t) for n,t in counts if n)[:6]:ids.update(x[0] for x in db.execute('SELECT id FROM tokens WHERE token=?'+restriction+' LIMIT 1500',[t]+suffix))
            for rid in ids-{exclude}:
                hit=db.execute('SELECT body FROM features WHERE id=?',(rid,)).fetchone()
                if not hit:continue
                value=json.loads(hit[0]);score=similarity(query,value['features'])
                if code and code_key(code)==code_key(value['row']['code']):score=100
                if score<snapshot.settings['similarity_min']:continue
                entry=snapshot.entries.get(case_key(value['row']['code'],value['row']['source'],value['row']['factory']))
                result.append(dict(row=value['row'],score=score,knowledge=entry['status'] if entry else 'не проверено',final=entry['value'] if entry and entry['status'] in ('ACTIVE','TRUSTED') else value['row']['automatic'],same_factory=bool(query['factory_key'] and query['factory_key']==value['features']['factory_key'])))
        return sorted(result,key=lambda x:(-x['score'],x['row']['id']))[:limit or snapshot.settings['analog_limit']]
    def groups(self,sid,progress=None,cancel=None,include_green=False):
        self.index_session(sid,progress,cancel);snapshot=self.store.snapshot();engine=ExpertAnonymizer(snapshot,self.base);groups={}
        with self.store.db() as db:
            for count,(body,) in enumerate(db.execute('SELECT body FROM features WHERE session=?',(sid,)),1):
                if cancel and cancel.is_set():raise InterruptedError('Группировка отменена')
                value=json.loads(body);row=value['row'];info=value['features'];exact=snapshot.entries.get(case_key(row['code'],row['source'],row['factory']));matching=engine.matching_rules(row['source'],info)
                result=engine.anonymize(row['source'],row['code'],row['factory']) if exact or matching or snapshot.settings_conflict else None
                status=result['status'] if result else row['status']
                if not include_green and status=='ЗЕЛЁНЫЙ':continue
                resolved={folded(e['sample']['fragment']) for e,_ in matching if e['status'] in ('ACTIVE','TRUSTED')} if status!='КРАСНЫЙ' else set()
                fragments=list(info['review_fragments'])
                if include_green:fragments=list(dict.fromkeys(fragments+review_tokens(row['source'])))
                if not fragments:fragments=['[проверить строку целиком]']
                for fragment in fragments:
                    if folded(fragment) in resolved:continue
                    spans=literal_spans(row['source'],fragment);components=[c for c in info['components'] if spans and all(c['start']<=a and b<=c['end'] for a,b in spans)]
                    context=components[0] if len(components)==1 else info;role='комплектующее' if components else 'изделие';key=digest([context['factory_key'],context['category'],role,folded(fragment)])
                    if key not in groups:groups[key]=dict(key=key,fragment=fragment,factory=context['factory'],category=context['category'],role=role,classification=fragment_class(fragment,row['source'],info),row_ids=[],examples=[])
                    group=groups[key];group['row_ids'].append(row['id'])
                    if len(group['examples'])<4 and row['source'] not in group['examples']:group['examples'].append(row['source'])
                if progress and count%1000==0:progress('Группировка: '+str(count)+' строк')
        return sorted(groups.values(),key=lambda x:(-len(x['row_ids']),x['key']))
    def group_plan(self,group,action,row_ids=None):
        if group['fragment'].startswith('['):raise ValueError('Эти строки требуют индивидуального исправления.')
        proposals=[];snapshot=self.store.snapshot()
        for rid in group['row_ids'] if row_ids is None else row_ids:
            row=self.store.row(rid)
            if row:proposals.append(dict(row=row,final=edit_fragment(row['source'],self.current(row,snapshot)['final'],group['fragment'],action),action='Групповое решение',explicit_feedback=[dict(fragment=group['fragment'],action=action)]))
        return self.prepare(proposals,snapshot)

def edit_fragment(source,current,fragment,action):
    spans=literal_spans(current,fragment)
    if action=='DELETE':
        for a,b in reversed(spans):current=current[:a]+' '+current[b:]
    elif action=='KEEP':
        import difflib
        original=literal_spans(source,fragment)
        for start,end in original[len(spans):]:
            matcher=difflib.SequenceMatcher(None,source,current,autojunk=False)
            insert=next((b for a,b,size in matcher.get_matching_blocks() if a>=end and size),len(current));current=current[:insert].rstrip()+' '+source[start:end]+' '+current[insert:].lstrip()
    else:raise ValueError('Неверное действие')
    return re.sub(r'\s+',' ',current).strip(' ,;')
