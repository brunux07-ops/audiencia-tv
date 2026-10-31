"""Atualizador conservador: nunca inventa ratings e preserva registros existentes.
Se o site mudar sua estrutura, registre a URL como pendente em dados.json.
"""
import json, re, calendar, html, time, os
from pathlib import Path
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

ROOT=Path(__file__).parent
DB=ROOT/'dados.json'
TODAY=datetime.now(ZoneInfo('America/Sao_Paulo')).date()
START=date(2026,9,21)
NETS={'ABC','CBS','NBC','FOX','CW'}
DAYS=['Segunda','Terça','Quarta','Quinta','Sexta','Domingo']
SPORT=re.compile(r'\b(NFL|NBA|MLB|NHL|WNBA|NASCAR|football|baseball|basketball|hockey|soccer|volleyball|world series|the OT|pre.?game|post.?game|playoffs?|olympics?|sports|racing|college football)\b',re.I)
TYPES=[('(Series Premiere)','Estreia da série'),('(Series Finale)','Final da série'),('(P)','Estreia'),('(R)','Reprise'),('(F)','Final')]
S=requests.Session();S.headers.update({'User-Agent':'Mozilla/5.0 (compatible; ratings-archive/1.0)'})
REQUEST_LIMIT=int(os.getenv('RATINGS_REQUEST_LIMIT','10'))
REQUEST_DELAY=float(os.getenv('RATINGS_REQUEST_DELAY','3'))
requests_made=0
blocked=False
class Blocked(Exception):pass

def get(url):
 global requests_made, blocked
 if blocked or requests_made>=REQUEST_LIMIT:
  return None
 if requests_made:time.sleep(REQUEST_DELAY)
 requests_made+=1
 try:
  r=S.get(url,timeout=25,allow_redirects=True)
  if r.status_code==429 or 'google.com/sorry/' in r.url:
   blocked=True
   print('AVISO: bloqueio 429; interrompendo consultas nesta execucao. Sem tentativas adicionais.')
   return None
  r.raise_for_status()
  return BeautifulSoup(r.text,'html.parser')
 except requests.RequestException as e:
  print('AVISO: acesso indisponivel:',url,str(e))
  return None

def number(s):
 s=s.strip().replace(',','').replace('−','-')
 if not s or s.upper() in ('TBA','N/A','--','-'):return None
 m=re.search(r'(?<!\w)\d+(?:\.\d+)?',s)
 return m.group() if m else None

def parse_post(url,night):
 soup=get(url)
 if not soup:return []
 records=[]
 for table in soup.find_all('table'):
  lines=[]
  for tr in table.find_all('tr'):
   cells=[c.get_text(' ',strip=True) for c in tr.find_all(['th','td'],recursive=False)]
   if cells:lines.append(cells)
  if not lines:continue
  head=' '.join(' '.join(x) for x in lines[:3]).lower()
  if not any(x in head for x in ['18-49','18–49','18-49 rating','adults 18','demo']):continue
  # Infer columns from a header, reject ambiguous tables rather than guessing.
  headers=[c.lower() for c in lines[0]]
  def find_col(words):
   return next((i for i,h in enumerate(headers) if any(w in h for w in words)),None)
  ti=find_col(['time','hour']);pi=find_col(['show','program']);ni=find_col(['network','net','channel']);ri=find_col(['18-49','18–49','demo']);vi=find_col(['viewers','million','audience'])
  if None in (pi,ni,ri):continue
  current_time=''
  for cells in lines[1:]:
   if len(cells)<=max(pi,ni,ri):continue
   network=cells[ni].strip().upper()
   if network not in NETS:continue
   title=cells[pi].strip()
   if SPORT.search(title):continue
   typ=''
   for marker,meaning in TYPES:
    if marker in title:
     title=title.replace(marker,'').strip();typ=meaning;break
   if ti is not None and ti<len(cells) and cells[ti].strip():current_time=cells[ti].strip()
   rating=number(cells[ri]);viewers=number(cells[vi]) if vi is not None and vi<len(cells) else None
   if not title or not current_time:continue
   records.append({'data':night.isoformat(),'horario':current_time,'programa':title,'rede':network,'rating':rating,'publico':viewers,'tipo':typ,'obs':'parcial' if rating is None or viewers is None else '', 'fonte':url})
 return records

def discover():
 links={}
 # Consultar somente o mes atual e o anterior, evitando revarrer a temporada inteira.
 months={(TODAY.year,TODAY.month)}
 previous=TODAY.replace(day=1)-timedelta(days=1)
 months.add((previous.year,previous.month))
 for yy,mm in sorted(months,reverse=True):
  if blocked:break
  soup=get(f'https://www.thetvratingsguide.com/{yy}/{mm:02d}/')
  if not soup:continue
  for a in soup.find_all('a',href=True):
   title=a.get_text(' ',strip=True)
   m=re.search(r'\b(Monday|Tuesday|Wednesday|Thursday|Friday|Sunday)\s+TV Ratings\s+(\d{1,2})/(\d{1,2})/(\d{2,4})',title,re.I)
   if not m:continue
   month,day,year=map(int,m.groups()[1:]);year=year+2000 if year<100 else year
   try:d=date(year,month,day)
   except ValueError:continue
   if START<=d<=TODAY:links[d.isoformat()]=a['href']
 return links

def sorttime(s):
 m=re.search(r'(\d{1,2})(?::(\d\d))?\s*(AM|PM)',s,re.I)
 if not m:return 9999
 return (int(m[1])%12+(12 if m[3].upper()=='PM' else 0))*60+int(m[2] or 0)

existing=json.loads(DB.read_text()) if DB.exists() else {'registros':[],'pendentes':[]}
old=existing.get('registros',[])
bykey={(x['data'],x['horario'],x['programa'],x['rede']):x for x in old}
links=discover()
# Preservar pendencias anteriores, inclusive quando o site bloquear a listagem.
pending_by_date={x['data']:x for x in existing.get('pendentes',[]) if x.get('data')}
for item in pending_by_date.values():
 if item.get('url'):links.setdefault(item['data'],item['url'])
# Verificar apenas datas recentes ou com pendencias. Nunca refazer dezenas de posts antigos.
recent_cutoff=TODAY-timedelta(days=14)
checked=0
for ds,url in sorted(links.items(),reverse=True):
 night=date.fromisoformat(ds)
 if night.weekday()==5:continue
 if night<recent_cutoff and ds not in pending_by_date:continue
 if blocked or requests_made>=REQUEST_LIMIT:break
 rows=parse_post(url,night)
 checked+=1
 if not rows:
  pending_by_date[ds]={'data':ds,'url':url,'motivo':'Acesso indisponivel ou tabela nao interpretada; tentar novamente'}
  continue
 pending_by_date.pop(ds,None)
 for row in rows:
  key=(row['data'],row['horario'],row['programa'],row['rede'])
  prior=bykey.get(key,{})
  for field in ('rating','publico'):
   if row[field] is None and prior.get(field) is not None:row[field]=prior[field]
  row['obs']='parcial' if row['rating'] is None or row['publico'] is None else ''
  bykey[key]=row
pending=sorted(pending_by_date.values(),key=lambda x:x['data'],reverse=True)
print(f'Consultas HTTP: {requests_made}; posts analisados: {checked}; bloqueio 429: {blocked}; pendencias: {len(pending)}')

changed=any(k not in {(x['data'],x['horario'],x['programa'],x['rede']) for x in old} or any((bykey[k].get(f) is not None and next((z.get(f) for z in old if (z['data'],z['horario'],z['programa'],z['rede'])==k),None) is None) for f in ('rating','publico')) for k in bykey)
rows=list(bykey.values());rows.sort(key=lambda r:(-date.fromisoformat(r['data']).toordinal(),sorttime(r['horario']),-float(r['rating'] or 0)))
# Nao reescrever dados.json nem publicar uma pagina quando nao ha novas audiencias.
# Isso evita alterar o site por causa de bloqueios ou consultas sem resultados.
if changed:
 DB.write_text(json.dumps({'registros':rows,'pendentes':pending,'ultima_verificacao':TODAY.isoformat()},ensure_ascii=False,indent=2))
if not changed:
 print('Nenhum dado novo confirmado. Preservando index.html e audiencia.xlsx.');raise SystemExit(0)
wb=Workbook();wb.remove(wb.active)
for day in DAYS:
 ws=wb.create_sheet(day);ws.append(['Data','Horário','Programa','Rede','18-49','Público (milhões)','Tipo','Obs','Fonte'])
 for x in rows:
  d=date.fromisoformat(x['data']);dayname=DAYS[d.weekday()] if d.weekday()<5 else ('Domingo' if d.weekday()==6 else '')
  if dayname!=day:continue
  ws.append([d.strftime('%d/%m/%Y'),x['horario'],x['programa'],x['rede'],float(x['rating']) if x['rating'] is not None else None,float(x['publico']) if x['publico'] is not None else None,x['tipo'],x['obs'],x['fonte']])
 for c in ws[1]:c.font=Font(color='FFFFFF',bold=True);c.fill=PatternFill('solid',fgColor='182033')
 ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
 for col,w in {'A':15,'B':12,'C':40,'D':12,'E':12,'F':20,'G':20,'H':22,'I':65}.items():ws.column_dimensions[col].width=w
wb.save(ROOT/'audiencia.xlsx')
# A static self-contained page: no external JS/data dependency.
blob=json.dumps(rows,ensure_ascii=False).replace('</','<\\/')
page='''<!doctype html><html lang="pt-BR"><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Audiência TV aberta EUA</title><style>body{font:16px system-ui;background:#f4f6fa;color:#19202e;margin:0}.wrap{max-width:1150px;margin:auto;padding:22px}h1 span{color:#b72b40}button,a.download{padding:12px;border:1px solid #cad1dd;border-radius:9px;background:white;color:#182032;font-weight:700;margin:4px;cursor:pointer}button.active{background:#19202e;color:white}.toolbar{display:flex;flex-wrap:wrap;margin:20px 0}.panel{background:white;border:1px solid #d5dce6;border-radius:12px;margin:16px 0;overflow:auto}table{border-collapse:collapse;width:100%;min-width:680px}th,td{padding:12px;text-align:left;border-bottom:1px solid #e2e5eb}th{background:#e8edf5}small{color:#647084}.warn{background:#fff4d8;padding:12px;border-radius:9px}input{padding:12px;border:1px solid #ccd3df;border-radius:9px;width:90%;max-width:400px}</style><div class="wrap"><h1>Audiência <span>TV aberta</span> EUA</h1><p>ABC, CBS, NBC, FOX e CW · Nielsen Live+SD · Sem esportes e sábados</p><p><a class="download" href="audiencia.xlsx" download>↓ Baixar planilha (.xlsx)</a></p><div id="stats"></div><div class="toolbar" id="days"></div><div class="toolbar"><button id="bydate" class="active" onclick="mode='data';draw()">Por data</button><button id="byweek" onclick="mode='semanas';draw()">Comparar semanas</button></div><div class="toolbar" id="nets"></div><label><input type="checkbox" id="reprises" onchange="draw()"> Esconder reprises</label><p><input id="search" placeholder="Buscar programa..." oninput="draw()"></p><div id="content"></div><p><small>Fonte: The TV Ratings Guide. Números ausentes ficam vazios. Atualização: '''+TODAY.isoformat()+'''</small></p></div><script>const rows='''+blob+''';let day=0,net='Todas',mode='data';const names=['Segunda','Terça','Quarta','Quinta','Sexta','Domingo'];function wd(s){let n=new Date(s+'T12:00:00').getDay();return n===0?5:n-1}function esc(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}function draw(){document.getElementById('days').innerHTML=names.map((x,i)=>`<button class="${day===i?'active':''}" onclick="day=${i};draw()">${x}</button>`).join('');document.getElementById('nets').innerHTML=['Todas','ABC','CBS','NBC','FOX','CW'].map(x=>`<button class="${net===x?'active':''}" onclick="net='${x}';draw()">${x}</button>`).join('');document.getElementById('bydate').className=mode==='data'?'active':'';document.getElementById('byweek').className=mode==='semanas'?'active':'';let list=rows.filter(x=>wd(x.data)===day&&(net==='Todas'||x.rede===net)&&(!document.getElementById('reprises').checked||x.tipo!=='Reprise')&&x.programa.toLowerCase().includes(document.getElementById('search').value.toLowerCase()));let dates=[...new Set(list.map(x=>x.data))].sort().reverse();document.getElementById('stats').textContent=`${dates.length} noites · ${list.length} registros`;if(mode==='data'){document.getElementById('content').innerHTML=dates.map(d=>`<div class="panel"><h2 style="padding:12px">${d.split('-').reverse().join('/')}</h2><table><thead><tr><th>Horário</th><th>Programa</th><th>Rede</th><th>18–49</th><th>Público (milhões)</th><th>Tipo</th></tr></thead><tbody>${list.filter(x=>x.data===d).map(x=>`<tr><td>${esc(x.horario)}</td><td>${esc(x.programa)}</td><td>${esc(x.rede)}</td><td>${esc(x.rating)}</td><td>${esc(x.publico)}</td><td>${esc(x.tipo)}</td></tr>`).join('')}</tbody></table></div>`).join('')||'<p class="warn">Nenhum dado verificado para esta seleção.</p>'}else{let shows=[...new Set(list.map(x=>x.programa+'|'+x.rede))].sort();document.getElementById('content').innerHTML=`<div class="panel"><table><thead><tr><th>Programa</th><th>Rede</th>${dates.map(d=>`<th>${d.slice(8)}/${d.slice(5,7)}</th>`).join('')}</tr></thead><tbody>${shows.map(s=>{let [p,n]=s.split('|');return `<tr><td>${esc(p)}</td><td>${esc(n)}</td>${dates.map(d=>`<td>${esc(list.find(x=>x.programa===p&&x.rede===n&&x.data===d)?.rating??'—')}</td>`).join('')}</tr>`}).join('')}</tbody></table></div>`}}draw();</script></html>'''
(ROOT/'index.html').write_text(page)
print('Registros:',len(rows),'URLs encontradas:',len(links),'Pendencias:',len(pending))
