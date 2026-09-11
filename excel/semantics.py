"""Local deterministic classification; inferred labels remain hypotheses."""
import difflib
import re
import unicodedata
CATEGORIES=[('вентиляционная установка',r'приточ|вытяж|вентиляц|кондиционер'),('клапан',r'клапан'),('кран',r'\bкран'),('задвижка',r'задвиж'),('электродвигатель',r'электродвигател|\bдвигател'),('привод',r'привод'),('насос',r'насос'),('трансформатор',r'трансформатор'),('кабель',r'кабел|\bпровод\b'),('труба',r'\bтруб[аы]|трубопровод'),('арматура трубопровода',r'фланец|фланц|отвод|тройник|переход|заглуш'),('шкаф управления',r'шкаф|\bщит\b|панел|шукгс'),('автоматический выключатель',r'автомат|выключател'),('датчик',r'датчик|преобразовател|термометр'),('манометр',r'манометр'),('расходомер',r'расходомер|счетчик|счётчик'),('светильник',r'светильник|прожектор|лампа'),('блок',r'\bблок'),('муфта',r'муфт'),('фильтр',r'фильтр'),('опора',r'опор'),('ёмкость',r'резервуар|емкост|ёмкост|\bбак\b')]
PATTERNS=[(name,re.compile(pattern,re.I)) for name,pattern in CATEGORIES]
TOKEN=re.compile(r'[\w]+(?:[-./+][\w]+)*')
COMPONENT=re.compile(r'(?i)\b(?:с\s+)?(?:электропривод\w*|привод\w*|электродвигател\w*|двигател\w*|комплектующ\w*|в\s+комплекте)\b')
STOP=set('и в с со по на для из к от до или при не тип типа комплект комплекте шт мм м кг ооо ао пао оао зао нпо'.split())
CLASSES=('производитель','бренд','серия','модель','тип оборудования','техническое обозначение','не определено')
def folded(value):
    value=unicodedata.normalize('NFKC',str(value or '')).casefold().replace('ё','е')
    return ' '.join(re.sub('[«»“”„"]',' ',re.sub('[‐‑‒–—−]','-',value)).split())
def words(value):return [m.group() for m in TOKEN.finditer(folded(value))]
def factory_key(value):
    from excel.rules import extract_inn,normalize_name
    inn=extract_inn(value)
    return 'inn:'+str(inn) if inn else folded(normalize_name(value))
def category(text):
    hits=[(m.start(),i,name) for i,(name,rx) in enumerate(PATTERNS) if (m:=rx.search(text))]
    if hits:return min(hits)[2]
    first=next((t for t in words(text) if t not in STOP and t.isalpha()),'')
    return 'прочее: '+first if first else 'не определено'
def category_matches(text,label):
    if label=='*':return True
    if label.startswith('прочее: ') or label in dict(CATEGORIES):return category(text)==label
    return bool(literal_spans(folded(text),folded(label)))
def literal_spans(text,fragment):
    parts=re.split(r'(\s+|[‐‑‒–—−-])',fragment.strip())
    pattern=''.join(r'\s+' if p.isspace() else r'[‐‑‒–—−-]' if re.fullmatch('[‐‑‒–—−-]',p) else re.escape(p) for p in parts if p)
    return [m.span() for m in re.finditer(r'(?<!\w)'+pattern+r'(?!\w)',text,re.I)] if pattern else []
def features(source,factory='',code='',base=None):
    from excel.excel_engine import protected_ranges
    resolved,inn=base.resolve_factory(code,factory) if base else (factory,'')
    main=resolved or factory;origin='код/поле производителя' if main else 'не определён';components=[]
    for clause in re.finditer(r'[^;\n]+',source):
        marker=COMPONENT.search(clause.group())
        if marker and clause.start()+marker.start()>0:
            start=clause.start()+marker.start();components.append(dict(start=start,end=clause.end(),category=category(source[start:clause.end()])))
    if not main and base:
        ids={i for a,b,i in base._aliases(source) if not any(c['start']<=a<c['end'] for c in components)}
        if len(ids)==1:inn=next(iter(ids));main=base.names.get(inn,'');origin='название в тексте изделия'
    for c in components:
        ids={i for a,b,i in base._aliases(source[c['start']:c['end']])} if base else set()
        c['factory']=base.names.get(next(iter(ids)),'') if len(ids)==1 else '';c['factory_key']=factory_key(c['factory'])
    return dict(factory=main,factory_key='inn:'+inn if inn and not inn.startswith('name:') else factory_key(main),factory_source=origin,category=category(source),role='изделие',components=components,technical=[source[a:b] for a,b in protected_ranges(source)],tokens=list(dict.fromkeys(t for t in words(source) if t not in STOP and len(t)>1)),normalized=folded(source))
def fragment_class(fragment,source,info):
    if any(folded(fragment)==folded(t) for t in info.get('technical',[])):return 'техническое обозначение'
    if factory_key(fragment) and factory_key(fragment) in factory_key(info.get('factory','')):return 'производитель'
    if re.search(r'(?i)\b(?:ооо|ао|пао|оао|зао|нпо)\b',fragment):return 'производитель'
    if re.search(r'\d',fragment):return 'модель'
    if category(fragment) in dict(CATEGORIES):return 'тип оборудования'
    return 'не определено'
def changed_tokens(before,after):
    left=list(TOKEN.finditer(before));right=list(TOKEN.finditer(after));deleted=[];restored=[]
    for op,a,b,c,d in difflib.SequenceMatcher(None,[folded(m.group()) for m in left],[folded(m.group()) for m in right],autojunk=False).get_opcodes():
        if op in ('delete','replace') and b>a:deleted.append(before[left[a].start():left[b-1].end()])
        if op in ('insert','replace') and d>c:restored.append(after[right[c].start():right[d-1].end()])
    return deleted,restored
def feedback_facts(row,final,action,info=None):
    from excel.excel_engine import review_tokens
    info=info or row.get('classification') or features(row['source'],row.get('factory',''))
    source=row['source'];automatic=row.get('automatic',source)
    removed,_=changed_tokens(source if action=='Правильно' else automatic,final)
    _,restored=changed_tokens(automatic,final)
    facts=[('DELETE',x) for x in removed]+[('KEEP',x) for x in restored]
    if action=='Правильно':facts += [('KEEP',x) for x in review_tokens(final)]
    facts += [(x['action'],x['fragment']) for x in row.get('explicit_feedback',[])]
    output={}
    for direction,part in facts:
        positions=literal_spans(source,part)
        if len(part)<2 or not re.search(r'[A-Za-zА-Яа-яЁё]',part) or not positions:continue
        components=[c for c in info['components'] if all(c['start']<=a and b<=c['end'] for a,b in positions)]
        context=components[0] if len(components)==1 else info;role='комплектующее' if components else 'изделие'
        scope=dict(factory=context.get('factory',''),category=context['category'],role=role)
        output[(direction,folded(part),str(scope))]=dict(fragment=part,action=direction,scope=scope,classification=fragment_class(part,source,info),origin='явная проверка' if action=='Правильно' else 'исправление инженера',eligible=bool(scope['factory'] and scope['category']!='не определено'))
    return list(output.values())
def similarity(a,b):
    x,y=set(a['tokens']),set(b['tokens']);tx,ty=set(map(folded,a['technical'])),set(map(folded,b['technical']))
    return round(100*(.45*len(x&y)/max(1,len(x|y))+.25*bool(a['factory_key'] and a['factory_key']==b['factory_key'])+.20*(a['category']==b['category'])+.10*len(tx&ty)/max(1,len(tx|ty))),1)
