"""Rasteriza un PDF de revisión sin editarlo. Requiere pypdfium2 y Pillow."""
import argparse
from pathlib import Path
import pypdfium2 as pdfium

def render(source, destination, scale=1.5):
    destination=Path(destination)
    destination.mkdir(parents=True,exist_ok=True)
    with pdfium.PdfDocument(source) as document:
        for i in range(len(document)):
            target=destination/f'pagina-{i+1:03d}.png'
            if target.exists():raise FileExistsError(target)
            page=document[i]
            bitmap=page.render(scale=scale)
            bitmap.to_pil().save(target)
            bitmap.close();page.close()
        print(f'{len(document)} páginas renderizadas en {destination}')

if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('pdf');parser.add_argument('directorio');parser.add_argument('--scale',type=float,default=1.5)
    args=parser.parse_args()
    if args.scale<=0:parser.error('La escala debe ser positiva')
    render(args.pdf,args.directorio,args.scale)
