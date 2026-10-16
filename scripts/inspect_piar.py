"""Inventario sin modificar el archivo."""
import argparse
from piar_xml import inspect,emit
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('docx');p.add_argument('--json');a=p.parse_args()
    emit(inspect(a.docx),a.json)
