"""Pruebas estructurales con datos ficticios; salidas solo en un directorio de pruebas."""
import argparse
import tempfile
from piar_xml import *
from fill_piar import apply
from audit_piar import audit

def run(out):
    out=Path(out);out.mkdir(parents=True,exist_ok=True)
    src=Path(__file__).resolve().parent.parent/'assets'/'plantilla-piar.docx'
    root=load(src);s=next(s for s in root.xpath('.//w:sdt',namespaces=NS) if options(s))
    original_opts=options(s)
    first=area_rows(root)[0][2]
    patch={'source_sha256':sha(src),'operations':[
        {'op':'select','xpath':path(root,s),'expected_hash':fingerprint(s),'value':NEUTRAL,'add_neutral':True},
        {'op':'append','xpath':path(root,first[4]),'expected_hash':fingerprint(first[4]),'lines':['SEGUIMIENTO FICTICIO PARA PRUEBA']}],
        'ensure_area_period':[{'area':'Emprendimiento','periodo':2}]}
    dst=out/'prueba-1.docx';apply(src,patch,dst)
    r=load(dst);s2=r.xpath(path(root,s),namespaces=NS)[0]
    assert set(original_opts).issubset(options(s2)) and text(s2)==NEUTRAL
    emp=[cs for a,tr,cs in area_rows(r) if a=='Emprendimiento']
    assert len(emp)==1 and not celltext(emp[0][4]).strip()
    assert len(emp[0][2].xpath('.//w:sdt',namespaces=NS))==5
    assert len(emp[0][3].xpath('.//w:sdt',namespaces=NS))==5
    with zipfile.ZipFile(src) as a,zipfile.ZipFile(dst) as b:
        assert a.namelist()==b.namelist()
        for n in a.namelist():
            if n!='word/document.xml':assert a.read(n)==b.read(n),n
    second=out/'prueba-2.docx'
    apply(dst,{'source_sha256':sha(dst),'ensure_area_period':[{'area':'Emprendimiento','periodo':2}]},second)
    assert E.tostring(load(dst))==E.tostring(load(second))
    assert 'SEGUIMIENTO FICTICIO PARA PRUEBA' in text(load(second))
    r=load(second);ev=area_rows(r)[0][2][4]
    third=out/'prueba-3.docx'
    op={'op':'reset_evaluation','xpath':path(r,ev),'expected_hash':fingerprint(ev)}
    try:apply(second,{'source_sha256':sha(second),'operations':[op]},third)
    except ValueError:pass
    else:raise AssertionError('Permitió borrar seguimiento sin cambio de año')
    apply(second,{'source_sha256':sha(second),'new_academic_year':True,'operations':[op]},third)
    assert 'SEGUIMIENTO FICTICIO PARA PRUEBA' not in text(load(third))
    assert not any('evaluación' in a for a in audit(src)['alertas'])
    assert not all(audit(src)['areas'][0]['apoyos_con_contenido'].values())
    try:apply(src,{'source_sha256':'incorrecto'},out/'invalido.docx')
    except ValueError:pass
    else:raise AssertionError('Aceptó inventario de otro documento')
    # Nombre dividido en runs: corrección dirigida, sin reemplazo global.
    p=E.fromstring(f'<w:p xmlns:w="{W}"><w:r><w:t>Ana </w:t></w:r><w:r><w:t>Prueba</w:t></w:r></w:p>')
    set_paragraph(p,'Nombre de prueba corregido');assert text(p)=='Nombre de prueba corregido'
    # Cobertura completa de 11 áreas y 4 períodos, sin eliminar controles originales.
    allout=out/'cobertura.docx'
    apply(src,{'source_sha256':sha(src),'ensure_area_period':[{'area':a,'periodo':p} for a in AREAS for p in range(1,5)]},allout)
    report=audit(allout,range(1,5),AREAS)
    assert not report['cobertura_faltante']
    assert len(report['areas'])==44
    assert all(a['evaluacion_vacia'] for a in report['areas'])
    ids=load(allout).xpath('.//w:sdtPr/w:id/@w:val',namespaces=NS)
    assert len(ids)==len(set(ids))
    print('OK: opciones, controles, paquete, idempotencia, evaluaciones, hash, nombre repartido y 44 filas académicas.')

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('directorio');a=p.parse_args();run(a.directorio)
