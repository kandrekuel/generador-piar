"""Auditoría estructural; no certifica corrección pedagógica ni identidad."""
import argparse
import re
from piar_xml import *

def audit(src,required_periods=(),required_areas=()):
    root=load(src);out={'areas':[],'alertas':[],'controles':[],'identidad':'Revisar datos del titular y documentos aportados; no inferir nombres por mayúsculas.'}
    seen=set()
    for area,tr,cs in area_rows(root):
        per=periods(celltext(tr));bar=celltext(cs[2]);adj=celltext(cs[3]);support={}
        for key in ['Metodología','Comunicación','Recursos','Espacio','Tiempo']:
            clean=unicodedata.normalize('NFD',adj)
            clean=''.join(c for c in clean if not unicodedata.combining(c))
            match=re.search(re.escape(unicodedata.normalize('NFD',key).encode('ascii','ignore').decode())+r'\s*:\s*([^\n]+)',clean,re.I)
            value=match.group(1).strip() if match else ''
            support[key]=bool(value) and not any(s in norm(value) for s in ['ELIJAUNELEMENTO','FALTADATO','INFORMACIONINSUFICIENTE'])
        out['areas'].append({'area':area,'periodos':per,'apoyos_con_contenido':support,'evaluacion_vacia':not celltext(cs[4]).strip()})
        if not per:out['alertas'].append(area+': período sin identificar')
        for p in per:
            if (area,p) in seen:out['alertas'].append(f'{area}: período {p} duplicado')
            seen.add((area,p))
        if re.search(r'(?im)^\s*DBA\b',bar):out['alertas'].append(area+': revisar posible DBA en barreras')
        if not all(support.values()):out['alertas'].append(area+': revisar contenido de apoyos')
    for s in root.xpath('.//w:sdt',namespaces=NS):
        opts=options(s);visible=text(s)
        status='sin_lista' if not opts else ('pendiente' if 'Elija un elemento' in visible else ('valido' if visible in opts else 'fuera_de_lista'))
        out['controles'].append({'xpath':path(root,s),'estado':status,'texto':visible,'opciones':len(opts)})
        if status in ['pendiente','fuera_de_lista']:out['alertas'].append('Control '+status+': '+path(root,s))
    out['cobertura_faltante']=[{'area':a,'periodo':p} for a in required_areas for p in required_periods if (a,p) not in seen]
    out['placeholders_total']=text(root).count('Elija un elemento')
    out['falta_dato_total']=text(root).count('FALTA DATO')
    return out

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('docx');p.add_argument('--json');p.add_argument('--periodos',nargs='*',type=int,choices=range(1,5),default=[]);p.add_argument('--grado',type=int)
    a=p.parse_args();emit(audit(a.docx,a.periodos,AREAS if a.grado in [8,9] else []),a.json)
