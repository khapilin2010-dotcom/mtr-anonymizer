"""Contextual expert decisions over the existing manufacturer catalog."""
from collections import Counter,defaultdict
from excel.excel_engine import Anonymizer,protected_ranges,merge_ranges
from excel.knowledge import case_key,fragments
from excel.manufacturer_policy import AERO_INN,AERO_NAME_RE,aero_designation_ranges
from excel.semantics import features,factory_key,category_matches,literal_spans
occurrences=literal_spans

def protection_losses(source,final,base,code='',factory='',extra=()):
    spans=protected_ranges(source)+list(extra);_,identity=base.resolve_factory(code,factory)
    if identity==AERO_INN or AERO_NAME_RE.search(source):spans+=aero_designation_ranges(source)
    return [v for v,n in Counter(source[a:b] for a,b in merge_ranges(spans)).items() if final.count(v)<n]

class ExpertAnonymizer:
    def __init__(self,snapshot,base=None):
        self.snapshot=snapshot;self.base=base or Anonymizer();self.version=self.base.version
        self.product_rules=defaultdict(list);self.component_rules=defaultdict(list);self.global_protected=[]
        for e in snapshot.entries.values():
            if e['kind']!='rule':continue
            scope=e['sample']['scope'];factory=scope['factory']
            if factory=='*' and e['sample'].get('protected'):self.global_protected.append(e);continue
            keys={factory_key(factory)};_,inn=self.base.resolve_factory('',factory)
            if inn and not inn.startswith('name:'):keys.add('inn:'+inn)
            for key in keys-{''}:(self.component_rules if scope['role']=='комплектующее' else self.product_rules)[key].append(e)
    def matching_rules(self,source,info):
        matches={}
        def add(e,text,offset=0):
            if not category_matches(text,e['sample']['scope']['category']):return
            spans=[(a+offset,b+offset) for a,b in literal_spans(text,e['sample']['fragment'])]
            if not spans:return
            if e['key'] not in matches:matches[e['key']]=(e,[])
            matches[e['key']][1].extend(s for s in spans if s not in matches[e['key']][1])
        end=min((c['start'] for c in info['components']),default=len(source))
        for e in self.product_rules.get(info['factory_key'],[])+self.product_rules.get(factory_key(info['factory']),[]):add(e,source[:end])
        for e in self.global_protected:add(e,source)
        for c in info['components']:
            text=source[c['start']:c['end']];candidates=list(self.component_rules.get(c['factory_key'],[]))
            if not c['factory_key']:
                key=factory_key(text);candidates += [e for k,values in self.component_rules.items() if k and k in key for e in values]
            for e in candidates:add(e,text,c['start'])
        return list(matches.values())
    def assess(self,source,final,code='',factory='',info=None):
        info=info or features(source,factory,code,self.base);matches=self.matching_rules(source,info)
        extra=[span for e,spans in matches if e['sample'].get('protected') and e['value']=='KEEP' and e['status'] in ('ACTIVE','TRUSTED') for span in spans]
        conflicts=['Удаляется защищённый признак: '+v for v in protection_losses(source,final,self.base,code,factory,extra)]
        for e,spans in matches:
            if e['status']=='DISPUTED':conflicts.append('Противоположные решения: '+e['sample']['fragment']);continue
            if e['status'] not in ('ACTIVE','TRUSTED'):continue
            present=len(literal_spans(final,e['sample']['fragment']))
            if (e['value']=='KEEP' and present<len(spans)) or (e['value']=='DELETE' and present):conflicts.append(e['value']+' / '+e['sample']['fragment'])
        if self.snapshot.settings_conflict:conflicts.append('Конфликт настроек доверия')
        return list(dict.fromkeys(conflicts)),matches,extra
    def anonymize(self,name,code='',factory=''):
        source=str(name or '');info=features(source,factory,code,self.base);entry=self.snapshot.entries.get(case_key(code,source,factory))
        matches=self.matching_rules(source,info)
        extra=[span for e,spans in matches if e['sample'].get('protected') and e['value']=='KEEP' and e['status'] in ('ACTIVE','TRUSTED') for span in spans]
        if entry and entry['status'] in ('ACTIVE','TRUSTED') and entry['value'].strip() and not protection_losses(source,entry['value'],self.base,code,factory,extra):
            removed,_=fragments(source,entry['value'])
            result=dict(text=entry['value'],factory=info['factory'],status='ЗЕЛЁНЫЙ',changed=entry['value']!=source,reason='Проверенное решение; независимых инженеров: '+str(entry['confirmations']),removed=['Экспертное решение: '+x for x in removed],trace=[],protected=info['technical']+[source[a:b] for a,b in extra],classification=info)
            self.describe(result,entry,'Код Автодокс' if str(code).strip() else 'Точное наименование и завод')
            conflicts,_,_=self.assess(source,entry['value'],code,factory,info)
            return self.conflict(result,'; '.join(conflicts)) if conflicts else result
        disabled=self.snapshot.static_disabled
        result=self.base.anonymize(source,code,factory,disabled_rules=disabled);result['classification']=info
        result['knowledge']=dict(version=self.snapshot.version,source='Статическая база',status='LEGACY',events=[],score=None)
        if entry and entry['status']!='DISABLED':
            self.describe(result,entry,'Код Автодокс' if str(code).strip() else 'Точное наименование и завод')
            if entry['status']=='DISPUTED':return self.conflict(result,'Инженеры сохранили разные решения')
            if entry['status'] in ('ACTIVE','TRUSTED'):
                result['knowledge']['proposed']=entry['value'];losses=protection_losses(source,entry['value'],self.base,code,factory,extra)
                return self.conflict(result,'Решение удаляет защищённые признаки: '+', '.join(losses) if losses else 'Экспертное наименование пустое')
        keeps=list(extra);deletes=[];applied=[];conflicts=[]
        protect=protected_ranges(source)+extra;_,identity=self.base.resolve_factory(code,factory)
        if identity==AERO_INN or AERO_NAME_RE.search(source):protect+=aero_designation_ranges(source)
        for e,spans in matches:
            if e['status']=='DISPUTED':conflicts.append('Спорное правило: '+e['sample']['fragment']);continue
            if e['status'] not in ('ACTIVE','TRUSTED'):continue
            if e['value']=='DELETE':
                if any(a<y and x<b for a,b in spans for x,y in protect):conflicts.append('DELETE пересекается с защитой: '+e['sample']['fragment']);continue
                deletes += [(a,b,'Экспертное DELETE: '+e['key']) for a,b in spans]
            else:keeps+=spans
            applied.append(e)
        if any(a<y and x<b for a,b,_ in deletes for x,y in keeps):conflicts.append('Одновременно применимы KEEP и DELETE');deletes=[]
        if applied or keeps:
            result=self.base.anonymize(source,code,factory,expert_keeps=keeps,expert_deletes=deletes,disabled_rules=disabled);result['classification']=info
            result['knowledge']=dict(version=self.snapshot.version,source='Контекстные решения инженеров',status='ACTIVE',events=[h['id'] for e in applied for h in e['live']],rules=[e['key'] for e in applied],confirmations=min((e['confirmations'] for e in applied),default=0),opposition=sum(e['opposition'] for e in applied),score=min((e['score'] for e in applied),default=0))
        if self.snapshot.settings_conflict:conflicts.append('Конфликт настроек доверия')
        return self.conflict(result,'; '.join(conflicts)) if conflicts else result
    def describe(self,result,entry,source):
        result['knowledge']=dict(version=self.snapshot.version,source=source,status=entry['status'],score=entry['score'],confirmations=entry['confirmations'],opposition=entry['opposition'],key=entry['key'],events=[e['id'] for e in entry['live']])
    @staticmethod
    def conflict(result,reason):
        result.update(status='КРАСНЫЙ',reason=reason+'; требуется проверка');result.get('knowledge',{})['score']=0
        return result
