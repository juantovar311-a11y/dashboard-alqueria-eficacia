#!/usr/bin/env python3
"""
Servidor universal para Dashboard Alquería · Eficacia.

Modo 1 (local): lee Base_Datos_Dashboard.xlsx en la misma carpeta.
Modo 2 (público/cloud): si defines DASHBOARD_EXCEL_URL, descarga el Excel
público desde esa URL y lo vuelve a consultar periódicamente.

No requiere librerías externas: solo Python 3.
"""
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse, urlsplit, urlunsplit, parse_qsl, urlencode
from urllib.request import Request, urlopen
from datetime import datetime, timedelta
import io, json, zipfile, xml.etree.ElementTree as ET, re, os, threading, time, html

BASE = Path(__file__).resolve().parent
LOCAL_EXCEL = BASE / os.environ.get('DASHBOARD_EXCEL_FILE', 'Base_Datos_Dashboard.xlsx')
POP_EXCEL = BASE / os.environ.get('DASHBOARD_POP_FILE', 'Material_POP_Campanas.xlsx')
EXCEL_URL = os.environ.get('DASHBOARD_EXCEL_URL', '').strip()
HOST = '0.0.0.0'
PORT = int(os.environ.get('PORT') or os.environ.get('DASHBOARD_PORT', '8765'))
REMOTE_SECONDS = max(30, int(os.environ.get('DASHBOARD_REMOTE_SECONDS', '60')))

NS = {'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
RELNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
CACHE = {
    'local_mtime': None,
    'remote_checked': 0.0,
    'remote_etag': '',
    'remote_last_modified': '',
    'xlsx_bytes': None,
    'payload': b'[]',
    'rows': 0,
    'source': 'sin fuente',
    'error': '',
    'updated': '',
}
LOCK = threading.Lock()
POP_LOCK = threading.Lock()
POP_CACHE = {
    'local_mtime': None,
    'payload': b'{"rows":[],"summary":[]}',
    'rows': 0,
    'summary_rows': 0,
    'source': 'sin fuente POP',
    'error': '',
    'updated': '',
}


def cloud_url_candidates(url):
    """Build likely direct-download variants for OneDrive/SharePoint shared links."""
    raw = (url or '').strip()
    if not raw:
        return []
    out = []

    def add(u):
        if u and u not in out:
            out.append(u)

    def add_download_param(u):
        try:
            p = urlsplit(u)
            host = p.netloc.lower()
            q = dict(parse_qsl(p.query, keep_blank_values=True))
            if ('1drv.ms' in host or 'onedrive.live.com' in host or 'sharepoint.com' in host):
                q.pop('web', None)
                q['download'] = '1'
                return urlunsplit((p.scheme, p.netloc, p.path, urlencode(q), p.fragment))
        except Exception:
            pass
        return u

    # Prefer the force-download form first, then fall back to the original share link.
    add(add_download_param(raw))
    add(raw)
    return out


def onedrive_final_download_urls(final_url):
    """Build direct-download candidates from a OneDrive preview/editor URL."""
    out=[]
    def add(u):
        if u and u not in out: out.append(u)
    try:
        p=urlsplit(final_url)
        host=p.netloc.lower()
        if 'onedrive.live.com' in host:
            q=dict(parse_qsl(p.query, keep_blank_values=True))
            # Common OneDrive web endpoints accept /download with the same query.
            if q:
                add(urlunsplit((p.scheme,p.netloc,'/download',urlencode(q),'')))
            # Also try forcing download=1 on the final preview/editor URL.
            q2=dict(q); q2['download']='1'
            add(urlunsplit((p.scheme,p.netloc,p.path,urlencode(q2),'')))
    except Exception:
        pass
    return out


def extract_download_candidates(body, base_url=''):
    """Extract likely file-download URLs embedded in OneDrive HTML/JSON responses."""
    try:
        text=body.decode('utf-8','ignore') if isinstance(body,(bytes,bytearray)) else str(body)
    except Exception:
        return []
    # Decode the most common escaping forms used inside script JSON blobs.
    text=html.unescape(text)
    text=text.replace('\\u0026','&').replace('\\u003d','=').replace('\\/','/')
    found=[]
    def add(u):
        u=u.strip().strip('"\'()[]{};,')
        if u.startswith('http') and u not in found:
            found.append(u)
    # Named fields seen in OneDrive/Office payloads.
    for pat in (
        r'(?i)["\'](?:downloadUrl|download_url|@microsoft\.graph\.downloadUrl)["\']\s*:\s*["\']([^"\']+)',
        r'(?i)["\'](?:downloadUrl|download_url)["\']\s*=\s*["\']([^"\']+)',
    ):
        for m in re.finditer(pat,text): add(m.group(1))
    # Generic absolute URLs, restricted to hosts/paths that look file-oriented.
    for m in re.finditer(r'https?://[^"\'<>\s]+',text):
        u=m.group(0)
        lu=u.lower()
        if ('download' in lu or '1drv' in lu or 'onedrive' in lu or 'files.' in lu or 'storage.live.com' in lu):
            add(u)
    return found[:40]


def fetch_cloud_xlsx(url, headers):
    """Fetch XLSX bytes from a public OneDrive/SharePoint URL, following redirects and embedded download URLs."""
    errors=[]
    tried=set()
    candidates=cloud_url_candidates(url)
    idx=0
    while idx < len(candidates) and idx < 60:
        candidate=candidates[idx]; idx+=1
        if candidate in tried: continue
        tried.add(candidate)
        try:
            req=Request(candidate,headers=headers)
            with urlopen(req,timeout=35) as resp:
                data=resp.read()
                final_url=resp.geturl()
                ctype=resp.headers.get('Content-Type','')
                # XLSX is a ZIP package and always begins with PK.
                if data.startswith(b'PK'):
                    return data,resp.headers,final_url
                for direct in onedrive_final_download_urls(final_url):
                    if direct not in tried and direct not in candidates: candidates.append(direct)
                # New OneDrive pages sometimes embed a temporary file URL in HTML/JSON.
                if 'html' in ctype.lower() or data[:1] in (b'<',b'{',b'['):
                    for embedded in extract_download_candidates(data,final_url):
                        if embedded not in tried and embedded not in candidates: candidates.append(embedded)
                errors.append(f'{final_url} devolvió {ctype or "contenido no XLSX"}')
        except Exception as e:
            errors.append(f'{candidate}: {e}')
    detail=' | '.join(errors[-4:]) if errors else 'sin detalle'
    raise ValueError('OneDrive no devolvió el archivo .xlsx. El vínculo debe permitir descarga sin iniciar sesión. Detalle: '+detail)

def norm(s):
    s = '' if s is None else str(s).strip()
    import unicodedata
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    return re.sub(r'\s+', ' ', s).upper()

def clean(v):
    if v is None: return ''
    s = str(v).strip()
    return '' if s.upper() in {'0','N/A','#N/A','NA','#NA','NULL','NONE','NAN','-'} else s

def nval(v):
    if v in (None,''): return 0.0
    try: return float(v)
    except Exception:
        s=str(v).strip().replace(' ','')
        if re.match(r'^\d{1,3}(\.\d{3})+(,\d+)?$',s): s=s.replace('.','').replace(',','.')
        elif ',' in s and '.' not in s: s=s.replace(',','.')
        try:return float(s)
        except Exception:return 0.0

def excel_date(v):
    try:
        x=float(v)
        return (datetime(1899,12,30)+timedelta(days=x)).strftime('%Y-%m-%d')
    except Exception:
        s=clean(v)
        if re.match(r'^\d{4}-\d{2}-\d{2}',s): return s[:10]
        for fmt in ('%d/%m/%Y','%m/%d/%Y','%Y/%m/%d'):
            try:return datetime.strptime(s,fmt).strftime('%Y-%m-%d')
            except Exception:pass
        return None

def cell_col(ref):
    m=re.match(r'([A-Z]+)',ref or '')
    if not m:return -1
    n=0
    for ch in m.group(1):n=n*26+ord(ch)-64
    return n-1

def first_sheet_rows(source):
    # source can be Path, bytes, or BytesIO
    if isinstance(source, (bytes, bytearray)):
        source = io.BytesIO(source)
    with zipfile.ZipFile(source) as z:
        shared=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            root=ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root.findall('m:si',NS):
                shared.append(''.join(t.text or '' for t in si.findall('.//m:t',NS)))
        wb=ET.fromstring(z.read('xl/workbook.xml'))
        relroot=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        relmap={r.attrib['Id']:r.attrib['Target'] for r in relroot}
        sh=wb.find('m:sheets/m:sheet',NS)
        if sh is None:
            return []
        rid=sh.attrib[f'{{{RELNS}}}id']
        target=relmap[rid].lstrip('/')
        if not target.startswith('xl/'):target='xl/'+target
        root=ET.fromstring(z.read(target))
        rows=[]
        for row in root.findall('.//m:sheetData/m:row',NS):
            vals={}
            for c in row.findall('m:c',NS):
                idx=cell_col(c.attrib.get('r',''))
                typ=c.attrib.get('t','')
                if typ=='inlineStr':
                    val=''.join(t.text or '' for t in c.findall('.//m:t',NS))
                else:
                    v=c.find('m:v',NS); val='' if v is None else (v.text or '')
                    if typ=='s' and val!='':
                        try:val=shared[int(val)]
                        except Exception:pass
                    elif typ=='b': val='TRUE' if val=='1' else 'FALSE'
                vals[idx]=val
            if vals:
                mx=max(vals); rows.append([vals.get(i,'') for i in range(mx+1)])
        return rows

def pick(d,*names):
    for n in names:
        k=norm(n)
        if k in d:return d[k]
    return ''

def normalize_rows(rows):
    if not rows:return []
    headers=[norm(x) for x in rows[0]]
    out=[]
    for vals in rows[1:]:
        d={headers[i]:vals[i] if i<len(vals) else '' for i in range(len(headers)) if headers[i]}
        cod=clean(pick(d,'CÓDIGO CLIENTE','CODIGO CLIENTE'))
        pdv=clean(pick(d,'PDV'))
        ruta=clean(pick(d,'RUTA'))
        camp=clean(pick(d,'NOMBRE PLAN COMERCIAL','CAMPAÑA','CAMPANA'))
        tipo=clean(pick(d,'TIPO EXHIBICIÓN OBJETIVO','TIPO EXHIBICION OBJETIVO','TIPO EXHIBICIÓN','TIPO EXHIBICION','TIPO DE EXHIBICIÓN','TIPO DE EXHIBICION'))
        if not (cod or pdv or ruta or camp or tipo):continue
        c=norm(pick(d,'CUMPLE IMPLEMENTACIÓN','CUMPLE IMPLEMENTACION','CUMPLE ESTRATEGIA','CUMPLIMIENTO'))
        cumple='SI' if c in {'SI','YES','TRUE','1'} else ('NO' if c in {'NO','FALSE','0'} else (c or 'SIN DATO'))
        out.append({
            'region':clean(pick(d,'REGIONAL','REGION','REGIÓN')),
            'ciudad':clean(pick(d,'CIUDAD','CIUDAD EFICACIA','MUNICIPIO')),
            'ruta':ruta,
            'merc':clean(pick(d,'MERCADERISTA')),
            'sup':clean(pick(d,'SUPERVISOR')),
            'lider':clean(pick(d,'LIDER DE EJECUCION','LÍDER DE EJECUCIÓN','LIDER','LÍDER')),
            'canal':clean(pick(d,'CANAL')),
            'razon':clean(pick(d,'RAZÓN SOCIAL','RAZON SOCIAL')),
            'cod':cod,
            'pdv':pdv,
            'formato':clean(pick(d,'FORMATO')),
            'fecha':excel_date(pick(d,'FECHA')),
            'camp':camp,
            'catObj':clean(pick(d,'CATEGORÍA OBJETIVO','CATEGORIA OBJETIVO')),
            'marcaObj':clean(pick(d,'MARCA OBJETIVO')),
            'tipo':tipo,
            'tipoImpl':clean(pick(d,'TIPO EXHIBICIÓN IMPLEMENTADA','TIPO EXHIBICION IMPLEMENTADA')),
            'obj':nval(pick(d,'CANTIDAD IMPLEMENTACIÓN OBJETIVO','CANTIDAD IMPLEMENTACION OBJETIVO','CANTIDAD OBJETIVO','OBJETIVO')),
            'cumple':cumple,
            'catImpl':clean(pick(d,'CATEGORÍA IMPLEMENTADA','CATEGORIA IMPLEMENTADA')),
            'marcaImpl':clean(pick(d,'MARCA IMPLEMENTADA')),
            'impl':nval(pick(d,'CANTIDAD IMPLEMENTADA','CANTIDAD IMPLEMENTADA ESTRATEGIA','IMPLEMENTADO')),
            'causal':clean(pick(d,'CAUSAL','CAUSAL ESTRATEGIA')),
            'pct':nval(pick(d,'% IMPLEMENTACIÓN','% IMPLEMENTACION','% IMPLEMENTACIÓN ESTRATEGIA','% IMPLEMENTACION ESTRATEGIA')),
            'foto':clean(pick(d,'FOTO')),
            'gestion':clean(pick(d,'GESTIÓN','GESTION','GESTIÓN ESTRATEGIA','GESTION ESTRATEGIA')),
        })
    return out


def sheet_rows(source, wanted_name=None):
    """Read an XLSX sheet by name with stdlib only. Falls back to first sheet."""
    if isinstance(source, (bytes, bytearray)):
        source = io.BytesIO(source)
    with zipfile.ZipFile(source) as z:
        shared=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            root=ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root.findall('m:si',NS):
                shared.append(''.join(t.text or '' for t in si.findall('.//m:t',NS)))
        wbxml=ET.fromstring(z.read('xl/workbook.xml'))
        relroot=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        relmap={r.attrib['Id']:r.attrib['Target'] for r in relroot}
        sheets=wbxml.findall('m:sheets/m:sheet',NS)
        sh=None
        if wanted_name:
            wn=norm(wanted_name)
            for candidate in sheets:
                if norm(candidate.attrib.get('name',''))==wn:
                    sh=candidate;break
        if sh is None and sheets:
            sh=sheets[0]
        if sh is None:return []
        rid=sh.attrib[f'{{{RELNS}}}id']
        target=relmap[rid].lstrip('/')
        if not target.startswith('xl/'): target='xl/'+target
        root=ET.fromstring(z.read(target))
        rows=[]
        for row in root.findall('.//m:sheetData/m:row',NS):
            vals={}
            for c in row.findall('m:c',NS):
                idx=cell_col(c.attrib.get('r',''))
                typ=c.attrib.get('t','')
                if typ=='inlineStr':
                    val=''.join(t.text or '' for t in c.findall('.//m:t',NS))
                else:
                    v=c.find('m:v',NS); val='' if v is None else (v.text or '')
                    if typ=='s' and val!='':
                        try:val=shared[int(val)]
                        except Exception:pass
                    elif typ=='b': val='TRUE' if val=='1' else 'FALSE'
                vals[idx]=val
            if vals:
                mx=max(vals);rows.append([vals.get(i,'') for i in range(mx+1)])
        return rows

SPANISH_MONTHS = {
    'ENERO':1,'FEBRERO':2,'MARZO':3,'ABRIL':4,'MAYO':5,'JUNIO':6,
    'JULIO':7,'AGOSTO':8,'SEPTIEMBRE':9,'SETIEMBRE':9,'OCTUBRE':10,
    'NOVIEMBRE':11,'DICIEMBRE':12
}

def parse_execution_range(raw, year=2026):
    """Best-effort parser for strings such as '15 de agosto a 31 de agosto'."""
    s=clean(raw)
    if not s:return (None,None,False)
    ns=norm(s)
    open_end=('FIN EXISTENCIA' in ns or 'HASTA AGOTAR' in ns)
    # Day + Spanish month occurrences, allowing "de".
    matches=re.findall(r'(\d{1,2})\s*(?:DE\s*)?([A-ZÁÉÍÓÚÑ]+)', ns)
    dates=[]
    for day,month_name in matches:
        m=SPANISH_MONTHS.get(norm(month_name))
        if not m:continue
        try:dates.append(datetime(year,m,int(day)).strftime('%Y-%m-%d'))
        except Exception:pass
    start=dates[0] if dates else None
    end=None if open_end else (dates[1] if len(dates)>1 else None)
    return (start,end,open_end)

def raw_text(v):
    if v is None:return ''
    s=str(v).strip()
    return '' if s.upper() in {'N/A','#N/A','NA','#NA','NULL','NONE','NAN'} else s

def normalize_pop_rows(rows):
    if not rows:return []
    headers=[norm(x) for x in rows[0]]
    out=[]
    for vals in rows[1:]:
        d={headers[i]:vals[i] if i<len(vals) else '' for i in range(len(headers)) if headers[i]}
        month=clean(pick(d,'EJECUCION','EJECUCIÓN'))
        campaign=clean(pick(d,'CAMPAÑA','CAMPANA'))
        article=clean(pick(d,'ARTICULO','ARTÍCULO'))
        pop=clean(pick(d,'POP'))
        regional=clean(pick(d,'REGIONAL'))
        if not (month or campaign or article or pop or regional):continue
        exec_raw=clean(pick(d,'FECHA EJECUCION','FECHA EJECUCIÓN'))
        es,ee,eo=parse_execution_range(exec_raw)
        negotiated_raw=raw_text(pick(d,'TIEMPO NEGOCIADO'))
        negotiated_date=excel_date(negotiated_raw)
        delta=nval(pick(d,'UNIDADES FALTANTES'))
        assigned=nval(pick(d,'CANTIDAD ASIGNADA'))
        arrived=nval(pick(d,'CANTIDAD LLEGA'))
        output=nval(pick(d,'CANTIDAD SALIDA'))
        balance=nval(pick(d,'SALDO'))
        out.append({
            'month':month,
            'campaign':campaign,
            'article':article,
            'pop':pop,
            'product':clean(pick(d,'PRODUCTO FOCO')),
            'fecha':excel_date(pick(d,'FECHA')),
            'negotiatedRaw':negotiated_raw,
            'negotiatedDate':negotiated_date,
            'regional':regional,
            'assigned':assigned,
            'perPdv':nval(pick(d,'CANT X PDV','CANT X PDV.','CANTIDAD X PDV')),
            'pdvObjective':nval(pick(d,'PDV OBJETIVO','PDV OBJETIVOS')),
            'entryDate':excel_date(pick(d,'FECHA INGRESO')),
            'arrived':arrived,
            'arrivalDelta':delta,
            'shortage':max(-delta,0),
            'surplus':max(delta,0),
            'output':output,
            'lastOutputDate':excel_date(pick(d,'FECHA ÚLTIMA SALIDA','FECHA ULTIMA SALIDA')),
            'balance':balance,
            'observations':raw_text(pick(d,'OBSERVACIONES')),
            'receiptDate':excel_date(pick(d,'FECHA RECEPCION POP','FECHA RECEPCIÓN POP')),
            'executionRaw':exec_raw,
            'execStart':es,
            'execEnd':ee,
            'execOpen':eo,
        })
    return out

def normalize_pop_summary(rows):
    if not rows:return []
    headers=[norm(x) for x in rows[0]]
    out=[]
    for vals in rows[1:]:
        d={headers[i]:vals[i] if i<len(vals) else '' for i in range(len(headers)) if headers[i]}
        campaign=clean(pick(d,'CAMPAÑA','CAMPANA'))
        article=clean(pick(d,'ARTICULO','ARTÍCULO'))
        name=clean(pick(d,'NOMBRE'))
        if not (campaign or article or name):continue
        expected_raw=raw_text(pick(d,'ESPERADO'))
        out.append({
            'campaign':campaign,
            'article':article,
            'name':name,
            'expected':nval(expected_raw) if re.match(r'^\s*-?[\d.,]+\s*$',expected_raw or '') else 0,
            'expectedRaw':expected_raw,
            'arrived':nval(pick(d,'LLEGADO')),
            'output':nval(pick(d,'SALIDO')),
            'balance':nval(pick(d,'SALDO')),
            'arrivalDelta':nval(pick(d,'FALTANTE')),
            'arrivalText':raw_text(pick(d,'FECHA LLEGADA POP A CEDI')),
            'expectedDelivery':excel_date(pick(d,'FECHA ESPERADA ENTREGA POP')),
            'delay':raw_text(pick(d,'RETRASO')),
        })
    return out

def refresh_pop_if_needed(force=False):
    if not POP_EXCEL.exists():
        raise FileNotFoundError(f'No existe {POP_EXCEL.name}')
    mtime=POP_EXCEL.stat().st_mtime_ns
    if force or POP_CACHE['local_mtime'] != mtime or POP_CACHE['rows']==0:
        pop_rows=normalize_pop_rows(sheet_rows(POP_EXCEL,'SabanaPO'))
        summary=normalize_pop_summary(sheet_rows(POP_EXCEL,'Resumen'))
        if not pop_rows:
            raise ValueError('El Excel de Material POP no contiene registros válidos o los encabezados cambiaron.')
        POP_CACHE['payload']=json.dumps({'rows':pop_rows,'summary':summary},ensure_ascii=False,separators=(',',':')).encode('utf-8')
        POP_CACHE['rows']=len(pop_rows)
        POP_CACHE['summary_rows']=len(summary)
        POP_CACHE['source']=POP_EXCEL.name
        POP_CACHE['error']=''
        POP_CACHE['updated']=datetime.now().isoformat(timespec='seconds')
        POP_CACHE['local_mtime']=mtime
        print(f'[{datetime.now():%H:%M:%S}] Material POP actualizado: {len(pop_rows):,} registros · {len(summary):,} resumen')

def pop_payload(force=False):
    with POP_LOCK:
        try:
            refresh_pop_if_needed(force=force)
        except Exception as e:
            POP_CACHE['error']=str(e)
            if POP_CACHE['rows']==0:raise
        return POP_CACHE['payload']

def parse_to_payload(source, source_label):
    data=normalize_rows(first_sheet_rows(source))
    if not data:
        raise ValueError('El Excel no contiene registros válidos o los encabezados cambiaron.')
    CACHE['payload']=json.dumps(data,ensure_ascii=False,separators=(',',':')).encode('utf-8')
    CACHE['rows']=len(data)
    CACHE['source']=source_label
    CACHE['error']=''
    CACHE['updated']=datetime.now().isoformat(timespec='seconds')
    print(f'[{datetime.now():%H:%M:%S}] Fuente actualizada: {source_label} · {len(data):,} registros')

def refresh_local_if_needed(force=False):
    if not LOCAL_EXCEL.exists():
        raise FileNotFoundError(f'No existe {LOCAL_EXCEL.name}')
    mtime=LOCAL_EXCEL.stat().st_mtime_ns
    if force or CACHE['local_mtime'] != mtime or CACHE['rows']==0:
        parse_to_payload(LOCAL_EXCEL, LOCAL_EXCEL.name)
        CACHE['local_mtime']=mtime

def refresh_remote_if_needed(force=False):
    now=time.time()
    if not force and CACHE['rows'] and (now-CACHE['remote_checked'] < REMOTE_SECONDS):
        return
    CACHE['remote_checked']=now
    headers={'User-Agent':'Mozilla/5.0 Dashboard-Alqueria-Eficacia/2.0','Cache-Control':'no-cache'}
    if CACHE['remote_etag']: headers['If-None-Match']=CACHE['remote_etag']
    if CACHE['remote_last_modified']: headers['If-Modified-Since']=CACHE['remote_last_modified']
    try:
        data, response_headers, final_url = fetch_cloud_xlsx(EXCEL_URL, headers)
        CACHE['remote_etag']=response_headers.get('ETag','')
        CACHE['remote_last_modified']=response_headers.get('Last-Modified','')
        if force or CACHE['xlsx_bytes'] != data or CACHE['rows']==0:
            CACHE['xlsx_bytes']=data
            parse_to_payload(data, 'Excel OneDrive')
            print(f'[{datetime.now():%H:%M:%S}] URL final OneDrive: {final_url}')
    except Exception as e:
        CACHE['error']=str(e)
        if CACHE['rows']==0:
            raise
        print(f'[{datetime.now():%H:%M:%S}] Aviso cloud: {e} · se conserva última base válida')


def payload(force=False):
    with LOCK:
        if EXCEL_URL:
            refresh_remote_if_needed(force=force)
        else:
            refresh_local_if_needed(force=force)
        return CACHE['payload']

class Handler(SimpleHTTPRequestHandler):
    def __init__(self,*args,**kwargs): super().__init__(*args,directory=str(BASE),**kwargs)
    def end_headers(self):
        self.send_header('Cache-Control','no-store, no-cache, must-revalidate, max-age=0')
        super().end_headers()
    def send_json(self, obj, status=200):
        body=json.dumps(obj,ensure_ascii=False,separators=(',',':')).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type','application/json; charset=utf-8')
        self.send_header('Content-Length',str(len(body)))
        self.end_headers(); self.wfile.write(body)
    def do_GET(self):
        path=urlparse(self.path).path
        if path in ('/data.json','/api/data'):
            try:
                body=payload();self.send_response(200);self.send_header('Content-Type','application/json; charset=utf-8');self.send_header('Content-Length',str(len(body)));self.end_headers();self.wfile.write(body)
            except Exception as e:
                self.send_json({'error':str(e)},500)
            return
        if path in ('/api/pop','/pop.json'):
            try:
                body=pop_payload();self.send_response(200);self.send_header('Content-Type','application/json; charset=utf-8');self.send_header('Content-Length',str(len(body)));self.end_headers();self.wfile.write(body)
            except Exception as e:
                self.send_json({'error':str(e)},500)
            return
        if path=='/api/pop_status':
            try: pop_payload()
            except Exception as e: POP_CACHE['error']=str(e)
            self.send_json({'rows':POP_CACHE['rows'],'summary_rows':POP_CACHE['summary_rows'],'source':POP_CACHE['source'],'updated':POP_CACHE['updated'],'error':POP_CACHE['error']})
            return
        if path=='/api/pop_refresh':
            try:
                pop_payload(force=True)
                self.send_json({'ok':True,'rows':POP_CACHE['rows'],'summary_rows':POP_CACHE['summary_rows'],'source':POP_CACHE['source'],'updated':POP_CACHE['updated']})
            except Exception as e:
                self.send_json({'ok':False,'error':str(e),'rows':POP_CACHE['rows']},500)
            return
        if path=='/api/status':
            try: payload()
            except Exception as e: CACHE['error']=str(e)
            self.send_json({'rows':CACHE['rows'],'source':CACHE['source'],'updated':CACHE['updated'],'error':CACHE['error'],'cloud':bool(EXCEL_URL),'refresh_seconds':REMOTE_SECONDS})
            return
        if path=='/api/refresh':
            try:
                body=payload(force=True)
                self.send_json({'ok':True,'rows':CACHE['rows'],'source':CACHE['source'],'updated':CACHE['updated']})
            except Exception as e:
                self.send_json({'ok':False,'error':str(e),'rows':CACHE['rows']},500)
            return
        if path=='/health':
            self.send_json({'ok':True,'rows':CACHE['rows'],'pop_rows':POP_CACHE['rows']})
            return
        if path=='/': self.path='/index.html'
        super().do_GET()

if __name__=='__main__':
    print('\nDASHBOARD ALQUERÍA · EFICACIA — AUTOMÁTICO')
    print(f'Puerto: {PORT}')
    if EXCEL_URL:
        print('Fuente: Excel OneDrive/SharePoint configurado en DASHBOARD_EXCEL_URL')
        print(f'Revisión cloud: cada {REMOTE_SECONDS} segundos')
    else:
        print(f'Fuente local: {LOCAL_EXCEL}')
    print(f'Material POP: {POP_EXCEL}')
    print(f'Abre: http://localhost:{PORT}')
    print('El dashboard consulta /api/data automáticamente y el servidor refresca la fuente cloud.\n')
    # Abrir el puerto PRIMERO para que Render detecte el servicio de inmediato.
    # La lectura de los Excel se hace en segundo plano; así un archivo grande o
    # una demora de parsing no bloquea el arranque del Web Service.
    server = ThreadingHTTPServer((HOST,PORT),Handler)
    print(f'Servidor escuchando en {HOST}:{PORT} · health disponible en /health')

    def preload_data():
        try:
            payload(force=True)
        except Exception as e:
            print('Aviso inicial base ejecución:',e)
        try:
            pop_payload(force=True)
        except Exception as e:
            print('Aviso inicial Material POP:',e)

    threading.Thread(target=preload_data, daemon=True, name='preload-dashboard-data').start()
    server.serve_forever()
