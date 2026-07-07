# -*- coding: utf-8 -*-
"""Агент P&L для филиала CSS (морские проекты / DEM GROUP).
Читает выгрузку 1С ОСВ (структура: статья на уровне 3, проект/подразделение на уровне 4)
и заполняет единый лист P&L «Actual numbers 2026» по блокам проектов.

Ключевые правила:
- доход 6xxx -> Кредит (кол.6), расход 7xxx -> Дебет (кол.5);
- налоги по ЗП внутри проекта = ОПВ + ОСМС + соц.отчисления + соц.налог -> одна строка;
- «Аренда судна» -> строка конкретного судна в блоке проекта;
- новая статья, которой нет в блоке: >300к -> отдельной строкой в блоке; <=300к -> «Прочие».
"""
import re, openpyxl
from collections import defaultdict

SHEET='Actual numbers 2026'
THRESHOLD=300_000

# ----- границы блоков проектов (прямые затраты) -----
SECTIONS={
    'Зайсан':(13,27),'TULPAR':(28,42),'USDT':(43,48),'Meric':(49,53),
    'FlatTop':(54,57),'LQB':(58,72),'IBEEV':(73,89),'JanDeNul':(90,94),
    'IPS':(95,96),'Прочие':(97,98),
}
DIRECT_TOTAL=99          # строка «Итого» прямых затрат
ADMIN=(101,119); ADMIN_TOTAL=120
R_OTHER_INCOME=122; R_OTHER_EXP=123; R_FIN_INCOME=125; R_RESTORE=126
R_TAXABLE=129; R_KPN=130
INCOME=(3,10); INCOME_TOTAL=11

import copy as _copy
_ORIG=dict(SECTIONS=_copy.deepcopy(SECTIONS),DIRECT_TOTAL=DIRECT_TOTAL,ADMIN=ADMIN,
    ADMIN_TOTAL=ADMIN_TOTAL,R_OTHER_INCOME=R_OTHER_INCOME,R_OTHER_EXP=R_OTHER_EXP,
    R_FIN_INCOME=R_FIN_INCOME,R_RESTORE=R_RESTORE,R_TAXABLE=R_TAXABLE,R_KPN=R_KPN,
    INCOME_TOTAL=INCOME_TOTAL)
def _reset_layout():
    global SECTIONS,DIRECT_TOTAL,ADMIN,ADMIN_TOTAL,R_OTHER_INCOME,R_OTHER_EXP
    global R_FIN_INCOME,R_RESTORE,R_TAXABLE,R_KPN,INCOME_TOTAL
    import copy
    SECTIONS=copy.deepcopy(_ORIG['SECTIONS']); DIRECT_TOTAL=_ORIG['DIRECT_TOTAL']
    ADMIN=_ORIG['ADMIN']; ADMIN_TOTAL=_ORIG['ADMIN_TOTAL']; R_OTHER_INCOME=_ORIG['R_OTHER_INCOME']
    R_OTHER_EXP=_ORIG['R_OTHER_EXP']; R_FIN_INCOME=_ORIG['R_FIN_INCOME']; R_RESTORE=_ORIG['R_RESTORE']
    R_TAXABLE=_ORIG['R_TAXABLE']; R_KPN=_ORIG['R_KPN']; INCOME_TOTAL=_ORIG['INCOME_TOTAL']

# ----- проект ОСВ -> блок P&L -----
def project_to_section(proj):
    p=proj.lower()
    if 'survey vessels' in p or 'zaisan' in p or 'зайсан' in p: return 'Зайсан'
    if 'tulpar' in p or 'тулпар' in p: return 'TULPAR'
    if 'flat top' in p or 'сом 6' in p or 'сом6' in p: return 'FlatTop'
    if 'usdt eva' in p or 'eva' in p: return 'USDT'
    if 'meric' in p: return 'Meric'
    if 'usdt' in p: return 'USDT'
    if 'lqb' in p: return 'LQB'
    if 'ibeev' in p: return 'IBEEV'
    if 'jan de nul' in p: return 'JanDeNul'
    if 'ui60102' in p or 'ips' in p: return 'IPS'
    return None

# ----- «Аренда судна»: проект -> строка судна -----
VESSEL_ROW={
    'survey vessels':'north caspian',
    'usdt eva':'caspian marine  eva',
    'usdt':'osg  bue chu',
}

def norm(s):
    return re.sub(r'[^a-zа-я0-9 ]',' ',str(s).lower()).replace('ё','е')

# ----- концепты статей прямых затрат: (ключевые слова в статье ОСВ) -> ключевое слово строки блока -----
def concept_of(item):
    n=norm(item)
    if 'аренда судна' in n: return 'VESSEL'
    if 'аренда помещения' in n or 'аренда контейнер' in n or 'контейнер' in n: return 'аренда помещения'
    if 'заработн' in n and 'налог' not in n: return 'заработная плата'
    if any(k in n for k in ('пенсионные взносы','осмс','социальные отчисления','социальный налог','обязательное соц','отчисления на соц')): return 'TAX'
    if 'проживан' in n or 'найм жил' in n: return 'проживание'
    if 'медицин' in n: return 'медицин'
    if 'обучение' in n: return 'обучение'
    if 'питани' in n: return 'питание'
    if 'проезд' in n: return 'проезд'
    if 'суточные' in n: return 'суточные'
    if 'связь' in n or 'связи' in n: return 'связь'
    if 'причальн' in n: return 'причальн'
    if 'транспортн' in n: return 'транспортн'
    if 'страхован' in n: return 'страхован'
    if 'перевыставл' in n: return 'перевыставл'
    if 'сиз' in n: return 'сиз'
    if 'материал' in n: return 'материал'
    if 'командировочн' in n: return 'командировочн'
    if 'ремонт' in n: return 'ремонт'
    if 'утилизац' in n: return 'утилизац'
    if 'инспекц' in n: return 'инспекц'
    if 'экспертиз' in n: return 'экспертиз'
    if 'себестоимость реализованного' in n: return 'перевыставл'  # товар -> перевыставляемые
    return 'OTHER'

def find_row(ws, rng, keyword, anti=None):
    a,b=rng
    for r in range(a,b+1):
        lbl=norm(ws.cell(row=r,column=2).value or '')
        if keyword in lbl and (not anti or anti not in lbl):
            return r
    return None

def section_other_row(ws, sec):
    """строка «Прочие проф услуги»/«Прочие расходы» блока (для мелочи <=300к)."""
    rng=SECTIONS[sec]
    r=find_row(ws,rng,'прочие проф')
    if r: return r
    r=find_row(ws,rng,'прочие')
    return r if r else rng[1]


def parse_1c(path):
    """Возвращает листья: (account, item, project, amount, is_income)."""
    wb=openpyxl.load_workbook(path,data_only=True); ws=wb[wb.sheetnames[0]]
    rows=[]
    for r in range(1,ws.max_row+1):
        name=ws.cell(row=r,column=1).value
        if name is None: continue
        name=str(name).strip()
        if name in ('<...>','Итого',''): continue
        od=ws.row_dimensions[r].outline_level if r in ws.row_dimensions else 0
        deb=ws.cell(row=r,column=5).value or 0
        cred=ws.cell(row=r,column=6).value or 0
        rows.append((r,od,name,float(deb),float(cred)))
    leaves=[]; stack={}; acct=None
    for i,(r,od,name,deb,cred) in enumerate(rows):
        stack[od]=name
        for k in list(stack):
            if k>od: del stack[k]
        m=re.match(r'^(\d{4}),',name)
        if m and od<=2: acct=m.group(1)
        nxt = rows[i+1][1] if i+1<len(rows) else 0
        is_leaf = nxt<=od
        if not is_leaf: continue
        if not (deb or cred): continue
        a=acct or ''
        if a.startswith('6'):
            item=stack.get(od-1,name); leaves.append((a,item,name,cred,True))
        elif a.startswith('7'):
            # 7400 группа (курсовые 7430 / прочие 7480) — статья это сам счёт
            if a in ('7430','7480') or a.startswith('74'):
                leaves.append((a,name,'',deb,False))
            else:
                item=stack.get(od-1,name); leaves.append((a,item,name,deb,False))
    return leaves


def fill(template, leaves, out, month_col='C'):
    wb=openpyxl.load_workbook(template); ws=wb[SHEET]
    _relocate(ws)
    plan=_plan_insertions(ws,leaves)
    if any(plan.values()): _apply_insertions(ws,plan)
    acc=defaultdict(float); flags=[]; log=[]; taxbuf=defaultdict(float)
    MC=month_col

    def add(row,amt): acc[row]+=amt

    for (account,item,project,amount,is_income) in leaves:
        if is_income:
            _route_income(ws,account,item,project,amount,add,log); continue
        # --- расходы ---
        if account.startswith('74'):
            add(R_OTHER_EXP,amount); log.append(('ПрРасх',item,project,amount)); continue
        if account in ('7210','7212') or 'основное подразделение' in norm(project):
            _route_admin(ws,item,amount,add,taxbuf,log); continue
        # прямые затраты по проектам (7010/7011)
        sec=project_to_section(project)
        if not sec:
            add(R_OTHER_EXP,amount); flags.append((item,project,amount,'проект не распознан → Прочие расходы')); continue
        _route_direct(ws,sec,item,project,amount,add,taxbuf,log,flags)

    # налоги по ЗП: суммы по (section-row) уже накоплены в taxbuf -> в acc
    for row,amt in taxbuf.items():
        add(row,amt)

    # запись значений
    for row,val in acc.items():
        ws.cell(row=row,column=_col(MC)).value=round(val,2)

    wb.save(out)
    return flags,log


def _col(letter): return openpyxl.utils.column_index_from_string(letter)


# ---------- доходы ----------
INCOME_ROW={  # проект дохода -> строка блока ДОХОД ОТ РЕАЛИЗАЦИИ
    'ips':3,'ui60102':3,'survey vessels':4,'zaisan':4,
    'eva':5,'usdt':5,'bue':5,'flat top':6,'сом 6':6,'barges':6,
    'lqb':7,'ibeev':8,'jan de nul':9,
}
def _route_income(ws,account,item,project,amount,add,log):
    if account.startswith('61'):  # доходы по вознаграждениям -> фин.деятельность
        add(R_FIN_INCOME,amount); log.append(('ФинДоход',item,project,amount)); return
    if account in ('6200','6250'):  # курсовая разница доход -> прочие доходы
        add(R_OTHER_INCOME,amount); log.append(('ПрДоход',item,project,amount)); return
    p=norm(project)+' '+norm(item)
    for k,row in INCOME_ROW.items():
        if k in p:
            add(row,amount); log.append(('Доход',item,project,amount)); return
    add(10,amount); log.append(('ДоходПрочий',item,project,amount))  # r10 Прочие доходы (в блоке реализации)


# ---------- админ ----------
ADMIN_MAP=[  # (ключевое слово статьи ОСВ) -> ключевое слово строки админ-блока
    ('амортизац','амортизац'),('аренда офис','аренда помещения'),('спонсор','спонсор'),
    ('страхован','страхован'),('медицин','медицин'),('питани','питание'),
    ('обслуживанию банк','обслуж'),('вознаграждение за риски','обслуж'),
    ('транспортн','транспортн'),('связь','связ'),('связи','связ'),
    ('амортиз','амортиз'),('обучение','обучение'),
]
def _route_admin(ws,item,amount,add,taxbuf,log):
    n=norm(item)
    # налоги по ЗП админ
    if any(k in n for k in ('пенсионные взносы','осмс','социальные отчисления','социальный налог')):
        row=find_row(ws,ADMIN,'налоги'); taxbuf[row]+=amount; log.append(('Админ.налогЗП',item,'',amount)); return
    # ЗП + предоставление персонала
    if ('заработн' in n and 'налог' not in n) or 'предоставление персонала' in n:
        row=find_row(ws,ADMIN,'заработная плата','налог'); add(row,amount); log.append(('Админ.ЗП',item,'',amount)); return
    # командировочные = проезд + суточные + найм жилья
    if 'проезд' in n or 'суточные' in n or 'найм' in n:
        row=find_row(ws,ADMIN,'командировочн'); add(row,amount); log.append(('Админ.команд',item,'',amount)); return
    # обслуживание банка
    if 'обслуживанию банк' in n or 'вознаграждение за риски' in n:
        row=find_row(ws,ADMIN,'обслуж'); add(row,amount); log.append(('Админ.банк',item,'',amount)); return
    # прочие: почта, реклама, сервисные, IT, за счёт чистой прибыли, прочие расходы
    if any(k in n for k in ('почте','реклам','сервисн','it-сервис','ит-сервис','чистой прибыли','прочие расход','прочие услуги')):
        row=find_row(ws,ADMIN,'прочие услуги'); add(row,amount); log.append(('Админ.прочие',item,'',amount)); return
    if 'связ' in n:
        row=find_row(ws,ADMIN,'связ'); add(row,amount); log.append(('Админ.связь',item,'',amount)); return
    # прямое совпадение по ключевым словам
    for kw,lblkw in ADMIN_MAP:
        if kw in n:
            row=find_row(ws,ADMIN,lblkw)
            if row: add(row,amount); log.append(('Админ',item,'',amount)); return
    # не нашли -> прочие услуги
    row=find_row(ws,ADMIN,'прочие услуги'); add(row,amount); log.append(('Админ.проч?',item,'',amount))


# ---------- прямые затраты ----------
def _route_direct(ws,sec,item,project,amount,add,taxbuf,log,flags):
    rng=SECTIONS[sec]; con=concept_of(item); ni=norm(item)
    if sec=='LQB' and 'перевыставление расходов по контракту' in ni:
        row=find_row(ws,rng,'проживание')
        if row: add(row,amount); log.append((sec,'проживание(контракт)',item,amount)); return
    if con=='TAX':
        row=find_row(ws,rng,'налог') or find_row(ws,rng,'налоги')
        if row: taxbuf[row]+=amount; log.append((sec,'налогЗП',item,amount)); return
    if con=='VESSEL':
        vr=None
        for k,lbl in VESSEL_ROW.items():
            if k in norm(project): vr=find_row(ws,rng,lbl); break
        if not vr:  # первая строка блока с именем судна (двоеточие)
            for r in range(rng[0],rng[1]+1):
                if ':' in str(ws.cell(row=r,column=2).value or ''): vr=r; break
        if vr: add(vr,amount); log.append((sec,'аренда судна',item,amount)); return
    if con=='заработная плата':
        row=find_row(ws,rng,'заработная плата','налог')
        if row: add(row,amount); log.append((sec,'ЗП',item,amount)); return
    if con!='OTHER':
        row=find_row(ws,rng,con)
        if row: add(row,amount); log.append((sec,con,item,amount)); return
    # не распознали по концепту: ищем точную метку (в т.ч. вставленную новую строку)
    exact=_find_label(ws,item,rng[0],rng[1])
    if exact:
        add(exact,amount); log.append((sec,'новая статья',item,amount)); return
    if amount>THRESHOLD:
        flags.append((item,project,amount,f'НОВАЯ статья >300к в блоке {sec}'))
        log.append((sec,'NEW>300k(не вставлена?)',item,amount))
    else:
        row=section_other_row(ws,sec); add(row,amount); log.append((sec,'прочие<300к',item,amount))


# ================= динамическая вставка новых статей =================
def _find_label(ws, keyword, r_from=1, r_to=140, anti=None):
    kw=norm(keyword)
    for r in range(r_from,r_to+1):
        lbl=norm(ws.cell(row=r,column=2).value or '')
        if kw in lbl and (not anti or anti not in lbl): return r
    return None

def _plan_insertions(ws, leaves):
    """Новые статьи >300к, которых нет в блоке проекта -> план вставки строк."""
    plan=defaultdict(list)
    for (account,item,project,amount,is_income) in leaves:
        if is_income or account.startswith('74'): continue
        if account in ('7210','7212') or 'основное подразделение' in norm(project): continue
        sec=project_to_section(project)
        if not sec: continue
        con=concept_of(item); ni=norm(item)
        if con!='OTHER': continue
        if sec=='LQB' and 'перевыставление расходов по контракту' in ni: continue
        # точное совпадение с существующей меткой блока?
        rng=SECTIONS[sec]
        if _find_label(ws,item,rng[0],rng[1]): continue
        if amount>THRESHOLD and item not in plan[sec]:
            plan[sec].append(item)
    return plan

def _apply_insertions(ws, plan):
    """Вставляет строки под новые статьи в конец блока проекта, затем пересобирает разметку и формулы."""
    order=sorted(plan.keys(), key=lambda s: SECTIONS[s][0], reverse=True)
    for sec in order:
        items=plan[sec]
        if not items: continue
        a,b=SECTIONS[sec]
        pos=_find_label(ws,'прочие проф',a,b) or (b+1)
        for it in items:
            ws.insert_rows(pos); ws.cell(row=pos,column=2).value=it; pos+=1
    _relocate(ws)
    _rebuild_totals(ws)

def _shift_sections(at_row, delta):
    global SECTIONS, DIRECT_TOTAL, ADMIN, ADMIN_TOTAL
    global R_OTHER_INCOME,R_OTHER_EXP,R_FIN_INCOME,R_RESTORE,R_TAXABLE,R_KPN,INCOME_TOTAL
    ns={}
    for k,(a,b) in SECTIONS.items():
        na=a+delta if a>=at_row else a; nb=b+delta if b>=at_row else b
        ns[k]=(na,nb)
    SECTIONS=ns
    def sh(v): return v+delta if v>=at_row else v
    DIRECT_TOTAL=sh(DIRECT_TOTAL); ADMIN=(sh(ADMIN[0]),sh(ADMIN[1])); ADMIN_TOTAL=sh(ADMIN_TOTAL)
    R_OTHER_INCOME=sh(R_OTHER_INCOME); R_OTHER_EXP=sh(R_OTHER_EXP); R_FIN_INCOME=sh(R_FIN_INCOME)
    R_RESTORE=sh(R_RESTORE); R_TAXABLE=sh(R_TAXABLE); R_KPN=sh(R_KPN); INCOME_TOTAL=sh(INCOME_TOTAL)


A_HEADER_MAP=[('зайсан','Зайсан'),('tulpar','TULPAR'),('bue chu','USDT'),('usdt meric','Meric'),
    ('сом 6','FlatTop'),('lqb','LQB'),('ibeev','IBEEV'),('jan de nul','JanDeNul'),('ips','IPS'),('прочие','Прочие')]

def _relocate(ws):
    """Определяет границы блоков и якорные строки по заголовкам (устойчиво к вставкам)."""
    global SECTIONS,DIRECT_TOTAL,ADMIN,ADMIN_TOTAL,R_OTHER_INCOME,R_OTHER_EXP
    global R_FIN_INCOME,R_RESTORE,R_TAXABLE,R_KPN,INCOME_TOTAL
    inc_hdr=_find_label(ws,'доход от реализации')
    INCOME_TOTAL=_find_label(ws,'итого',inc_hdr+1,inc_hdr+30)
    dir_hdr=_find_label(ws,'прямые затраты')
    DIRECT_TOTAL=_find_label(ws,'итого',dir_hdr+1,dir_hdr+120)
    heads=[]
    for r in range(dir_hdr+1,DIRECT_TOTAL):
        a=norm(ws.cell(row=r,column=1).value or '')
        for kw,sec in A_HEADER_MAP:
            if kw in a and sec not in [h[1] for h in heads]: heads.append((r,sec)); break
    SECTIONS={}
    for i,(r,sec) in enumerate(heads):
        end=(heads[i+1][0]-1) if i+1<len(heads) else (DIRECT_TOTAL-1)
        SECTIONS[sec]=(r,end)
    adm_hdr=_find_label(ws,'административные расходы')
    ADMIN_TOTAL=_find_label(ws,'итого',adm_hdr+1,adm_hdr+40)
    ADMIN=(adm_hdr+1,ADMIN_TOTAL-1)
    R_OTHER_INCOME=_find_label(ws,'прочие доходы',ADMIN_TOTAL+1)
    R_OTHER_EXP=_find_label(ws,'прочие расходы',ADMIN_TOTAL+1)
    R_FIN_INCOME=_find_label(ws,'доходы от финансов',ADMIN_TOTAL+1)
    R_RESTORE=_find_label(ws,'восстанов',ADMIN_TOTAL+1) or (R_FIN_INCOME+1 if R_FIN_INCOME else None)
    R_TAXABLE=_find_label(ws,'налогооблагаемый',ADMIN_TOTAL+1)
    R_KPN=_find_label(ws,'кпн',ADMIN_TOTAL+1)

def _rebuild_totals(ws):
    """Пересобирает формулы итогов и колонку O=SUM(C:N) для всех строк данных."""
    cols=[openpyxl.utils.get_column_letter(c) for c in range(3,15)]  # C..N
    O='O'
    inc=_find_label(ws,'доход от реализации'); inc_tot=INCOME_TOTAL
    dir_hdr=_find_label(ws,'прямые затраты'); dir_tot=DIRECT_TOTAL
    adm_hdr=_find_label(ws,'административные расходы'); adm_tot=ADMIN_TOTAL
    # итог доход = SUM(строки между шапкой и ИТОГО)
    for tot,first in [(inc_tot,inc+1),(dir_tot,dir_hdr+1),(adm_tot,adm_hdr+1)]:
        for c in cols+[O]:
            ws[f'{c}{tot}']=f'=SUM({c}{first}:{c}{tot-1})'
    # налогооблагаемый доход
    tax=R_TAXABLE
    for c in cols+[O]:
        ws[f'{c}{tax}']=f'={c}{inc_tot}-{c}{dir_tot}-{c}{adm_tot}+{c}{R_OTHER_INCOME}-{c}{R_OTHER_EXP}+{c}{R_FIN_INCOME}-{c}{R_RESTORE}'
    # колонка O для строк данных
    for r in range(3,R_KPN+1):
        f=ws.cell(row=r,column=15).value
        cur=ws.cell(row=r,column=2).value
        if cur and not (isinstance(f,str) and f.startswith('=')):
            ws[f'O{r}']=f'=SUM(C{r}:N{r})'
