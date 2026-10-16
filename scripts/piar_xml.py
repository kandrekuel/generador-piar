"""Operaciones OOXML conservadoras. No genera contenido pedagógico."""
import copy
import hashlib
import json
import re
import unicodedata
import zipfile
from pathlib import Path
from lxml import etree as E

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}
def q(s): return '{' + W + '}' + s
def norm(s):
    return re.sub(r'[^A-Z0-9]', '', ''.join(c for c in unicodedata.normalize('NFD', s.upper()) if not unicodedata.combining(c)))
def text(e): return ''.join(e.xpath('.//w:t/text()', namespaces=NS))
def paragraphs(e): return e.xpath('.//w:p', namespaces=NS)
def celltext(e): return '\n'.join(text(p) for p in paragraphs(e))

AREAS = ['Ciencias Naturales y Educación Ambiental', 'Ciencias Sociales',
         'Educación Artística y Cultural', 'Educación Ética y en Valores Humanos',
         'Educación Física, Recreación y Deportes', 'Educación Religiosa', 'Emprendimiento',
         'Humanidades y Lengua Castellana', 'Idioma Extranjero Inglés', 'Matemáticas',
         'Tecnología e Informática']
ALIASES = dict(zip(map(norm, AREAS), AREAS))
for alias, i in [('CIENCIAS NATURALES',0),('ARTISTICA',2),('EDUCACION ARTISTICA',2),('ETICA',3),
                 ('EDUCACION ETICA',3),('EDUCACION FISICA',4),('RELIGION',5),
                 ('LENGUA CASTELLANA',7),('INGLES',8)]:
    ALIASES[norm(alias)] = AREAS[i]
PERIODS = ['PRIMER PERIODO','SEGUNDO PERIODO','TERCER PERIODO','CUARTO PERIODO']
NEUTRAL = 'Información insuficiente para seleccionar una opción específica'

def area_name(s): return ALIASES.get(norm(s))
def periods(s):
    return [i+1 for i,p in enumerate(PERIODS) if norm(p) in norm(s)]
def row_cells(tr):
    return [tc for tc in tr.xpath('.//w:tc',namespaces=NS)
            if next(tc.iterancestors(q('tr')),None) is tr]
def area_rows(root):
    result=[]
    for tr in root.xpath('.//w:tr',namespaces=NS):
        cells=row_cells(tr)
        if len(cells)==5:
            area=area_name(celltext(cells[0]))
            if area: result.append((area,tr,cells))
    return result
def sha(path): return hashlib.sha256(Path(path).read_bytes()).hexdigest()
def load(path):
    with zipfile.ZipFile(path) as z:
        return E.fromstring(z.read('word/document.xml'),E.XMLParser(resolve_entities=False,no_network=True))
def options(s):
    return [x.get(q('displayText')) or x.get(q('value')) or '' for x in s.xpath('./w:sdtPr/w:dropDownList/w:listItem | ./w:sdtPr/w:comboBox/w:listItem',namespaces=NS)]
def path(root,e): return root.getroottree().getpath(e)
def fingerprint(e): return hashlib.sha256(E.tostring(e,method='c14n')).hexdigest()
def target(root,op):
    es=root.xpath(op['xpath'],namespaces=NS)
    if len(es)!=1: raise ValueError('El destino XPath no es único: '+op['xpath'])
    e=es[0]
    if 'expected_text' in op and text(e)!=op['expected_text']:
        raise ValueError('El texto cambió; volver a inspeccionar: '+op['xpath'])
    if 'expected_hash' in op and fingerprint(e)!=op['expected_hash']:
        raise ValueError('La estructura cambió; volver a inspeccionar: '+op['xpath'])
    if not ('expected_text' in op or 'expected_hash' in op):
        raise ValueError('Cada cambio requiere expected_text o expected_hash')
    return e

def paragraph_like(container,value):
    proto=next(iter(container.xpath('.//w:p',namespaces=NS)),None)
    p=E.Element(q('p'))
    if proto is not None and proto.find(q('pPr')) is not None:
        p.append(copy.deepcopy(proto.find(q('pPr'))))
    r=E.SubElement(p,q('r'))
    rp=container.find('.//'+q('rPr'))
    if rp is not None:r.append(copy.deepcopy(rp))
    t=E.SubElement(r,q('t'));t.text=value
    t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
    return p
def append_unique(container,value):
    if value not in [text(p) for p in paragraphs(container)]:
        container.append(paragraph_like(container,value))
def set_paragraph(p,value):
    if p.tag!=q('p'):raise ValueError('Se esperaba párrafo')
    if p.xpath('.//w:sdt | .//w:drawing | .//w:fldChar | .//w:fldSimple',namespaces=NS):
        raise ValueError('Párrafo con controles, campos o imágenes: editar elementos específicos')
    ts=p.xpath('.//w:t',namespaces=NS)
    if ts:
        ts[0].text=value
        ts[0].set('{http://www.w3.org/XML/1998/namespace}space','preserve')
        for t in ts[1:]:t.text=''
    else:
        r=E.SubElement(p,q('r'));E.SubElement(r,q('t')).text=value
def select(s,value,add_neutral=False):
    if s.tag!=q('sdt'):raise ValueError('Se esperaba control de contenido')
    lists=s.xpath('./w:sdtPr/w:dropDownList | ./w:sdtPr/w:comboBox',namespaces=NS)
    if len(lists)!=1:raise ValueError('Control sin lista seleccionable')
    if value not in options(s):
        if not add_neutral:raise ValueError('Opción inexistente; solo se permite agregar una neutral')
        if not value.startswith(('Información insuficiente','No se cuenta con información','No aplica')):
            raise ValueError('La opción añadida debe ser neutral y justificada')
        item=E.SubElement(lists[0],q('listItem'));item.set(q('displayText'),value);item.set(q('value'),value)
    content=s.find(q('sdtContent'))
    if content is None:raise ValueError('Control sin contenido')
    ts=content.xpath('.//w:t',namespaces=NS)
    if not ts:raise ValueError('Control sin texto visible; requiere revisión específica')
    ts[0].text=value
    for t in ts[1:]:t.text=''
    for e in s.xpath('./w:sdtPr/w:showingPlcHdr',namespaces=NS):e.getparent().remove(e)
    selected=next(x for x in lists[0] if (x.get(q('displayText')) or x.get(q('value')))==value)
    lists[0].set(q('lastValue'),selected.get(q('value')) or value)

def clear_plain_cell(tc):
    if tc.xpath('.//w:sdt | .//w:drawing',namespaces=NS):
        raise ValueError('No se borran celdas con controles o imágenes')
    p=paragraph_like(tc,'')
    for ch in list(tc):
        if ch.tag!=q('tcPr'):tc.remove(ch)
    tc.append(p)

def ensure_area_period(root,area,period,template_root):
    canonical=area_name(area)
    if canonical is None:raise ValueError('Área no configurada: '+area)
    if period not in range(1,5):raise ValueError('Período fuera de 1 a 4')
    found=[(tr,cs) for a,tr,cs in area_rows(root) if a==canonical and period in periods(celltext(tr))]
    if len(found)>1:raise ValueError('Área/período duplicado: revisar antes de escribir')
    if found:return
    existing=[(tr,cs) for a,tr,cs in area_rows(root) if a==canonical]
    if existing and any(not periods(celltext(tr)) for tr,cs in existing):
        raise ValueError('Hay una fila sin período; identificarla antes de agregar otra')
    current=area_rows(root)
    if not current:raise ValueError('No se encontró tabla académica compatible de cinco columnas')
    # Siempre clonar la plantilla vacía de confianza, no datos de otro estudiante.
    proto=area_rows(template_root)[-1][1]
    row=copy.deepcopy(proto);cells=row_cells(row)
    clear_plain_cell(cells[0])
    cells[0].remove(cells[0].find(q('p')))
    for ch in canonical.upper():cells[0].append(paragraph_like(cells[0],ch if ch!=' ' else ''))
    clear_plain_cell(cells[1]);append_unique(cells[1],PERIODS[period-1])
    clear_plain_cell(cells[4])
    for tc in cells[2:4]:
        for t in tc.xpath('.//w:t',namespaces=NS):
            if t.text:
                for label in PERIODS:t.text=t.text.replace(label,PERIODS[period-1])
        for s in tc.xpath('.//w:sdt',namespaces=NS):
            if options(s):select(s,NEUTRAL,True)
    # IDs de controles únicos; no copiar bookmarks ni metadatos de enlace.
    ids=[int(e.get(q('val'))) for e in root.xpath('.//w:sdtPr/w:id',namespaces=NS) if (e.get(q('val')) or '').lstrip('-').isdigit()]
    nextid=max([0]+ids)+1
    for pr in row.xpath('.//w:sdtPr',namespaces=NS):
        ident=pr.find(q('id'))
        if ident is None:ident=E.SubElement(pr,q('id'))
        ident.set(q('val'),str(nextid));nextid+=1
        for db in pr.findall(q('dataBinding')):pr.remove(db)
    for e in row.xpath('.//w:bookmarkStart | .//w:bookmarkEnd',namespaces=NS):e.getparent().remove(e)
    anchor=existing[-1][0] if existing else current[-1][1]
    anchor.addnext(row)

def inspect(pathname):
    root=load(pathname)
    result={'sha256':sha(pathname),'areas':[], 'controles':[], 'parrafos':[], 'celdas':[]}
    for a,tr,cs in area_rows(root):
        result['areas'].append({'area':a,'periodos':periods(celltext(tr)), 'celdas':[path(root,c) for c in cs]})
    for s in root.xpath('.//w:sdt',namespaces=NS):
        result['controles'].append({'xpath':path(root,s),'expected_hash':fingerprint(s),'texto':text(s),'opciones':options(s)})
    for p in root.xpath('.//w:p',namespaces=NS):
        if text(p).strip():result['parrafos'].append({'xpath':path(root,p),'expected_text':text(p)})
    for c in root.xpath('.//w:tc',namespaces=NS):
        result['celdas'].append({'xpath':path(root,c),'expected_hash':fingerprint(c),'texto':celltext(c)})
    return result

def write_package(src,dst,root):
    src,dst=Path(src),Path(dst)
    if src.resolve()==dst.resolve():raise ValueError('La salida debe ser una copia distinta del original')
    if dst.exists():raise ValueError('La salida ya existe; elegir un nombre nuevo')
    dst.parent.mkdir(parents=True,exist_ok=True)
    tmp=dst.with_suffix(dst.suffix+'.tmp')
    if tmp.exists():raise ValueError('Ya existe un archivo temporal de salida')
    try:
        with zipfile.ZipFile(src) as zin,zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                zout.writestr(info,E.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True) if info.filename=='word/document.xml' else zin.read(info.filename))
        load(tmp)
        tmp.replace(dst)
    finally:
        if tmp.exists():tmp.unlink()

def emit(data,out=None):
    s=json.dumps(data,ensure_ascii=False,indent=2)
    if out:Path(out).write_text(s,encoding='utf-8')
    else:print(s)
