"""Edición protegida por hash y destinos explícitos. Ver references/edicion.md."""
import argparse
from piar_xml import *

def apply(src,changes,dst):
    if changes.get('source_sha256')!=sha(src):raise ValueError('Hash del documento incorrecto; volver a inspeccionar')
    root=load(src)
    original_controls=[(s,options(s)) for s in root.xpath('.//w:sdt',namespaces=NS)]
    operations=changes.get('operations',[])
    # Resolver destinos sobre la entrada original para que inserciones no desplacen XPath.
    resolved=[(op,target(root,op)) for op in operations]
    for op,e in resolved:
        kind=op['op']
        if kind=='select':select(e,op['value'],op.get('add_neutral',False))
        elif kind=='paragraph':
            if next(e.iterancestors(q('sdt')),None) is not None:raise ValueError('Usar select para texto dentro de desplegable')
            set_paragraph(e,op['value'])
        elif kind=='text_node':
            if e.tag!=q('t') or next(e.iterancestors(q('sdt')),None) is not None:raise ValueError('Nodo de texto no elegible')
            e.text=op['value']
        elif kind=='append':
            if e.tag!=q('tc'):raise ValueError('append requiere una celda')
            for line in op['lines']:append_unique(e,line)
        elif kind=='reset_evaluation':
            if not changes.get('new_academic_year'):raise ValueError('Solo limpiar seguimiento al actualizar año')
            eligible=[cs[4] for _,_,cs in area_rows(root)]
            for tr in root.xpath('.//w:tr',namespaces=NS):
                cs=row_cells(tr)
                if len(cs)==5 and norm(celltext(cs[1])) in ['CONVIVENCIA','SOCIALIZACION','PARTICIPACION','AUTONOMIA','AUTOCONTROL']:
                    eligible.append(cs[4])
            if not any(e is c for c in eligible):raise ValueError('No es una columna de evaluación reconocida')
            clear_plain_cell(e)
        else:raise ValueError('Operación desconocida: '+kind)
    template=load(Path(__file__).resolve().parent.parent/'assets'/'plantilla-piar.docx')
    for item in changes.get('ensure_area_period',[]):ensure_area_period(root,item['area'],item['periodo'],template)
    for s,opts in original_controls:
        if s.getroottree().getroot() is not root:raise ValueError('Se perdió un control original')
        if not set(opts).issubset(options(s)):raise ValueError('Se perdió una opción original')
    write_package(src,dst,root)

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('entrada');p.add_argument('cambios');p.add_argument('salida');a=p.parse_args()
    apply(a.entrada,json.loads(Path(a.cambios).read_text(encoding='utf-8-sig')),a.salida)
    print('Guardado:',a.salida)
