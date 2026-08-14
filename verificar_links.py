# -*- coding: utf-8 -*-
"""Chequea que cada link del indice tenga su marcador de destino en el Word,
y que el PDF haya conservado los enlaces internos."""
import sys, io, re, zipfile
from pathlib import Path
from docx import Document

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
docx = Path(sys.argv[1] if len(sys.argv) > 1
            else r"salida\Orden de servicio SEMANA.docx")
pdf = Path(sys.argv[2]) if len(sys.argv) > 2 else None

d = Document(docx)
body = d.element.body
anclas = [h.get(W + 'anchor') for h in body.findall(f'.//{W}hyperlink')
          if h.get(W + 'anchor')]
marcadores = {b.get(W + 'name') for b in body.findall(f'.//{W}bookmarkStart')}

print(f"{docx.name}")
print(f"  links en el indice : {len(anclas)}")
print(f"  marcadores destino : {len(marcadores & set(anclas))}")
huerfanos = [a for a in anclas if a not in marcadores]
if huerfanos:
    print(f"  !! LINKS ROTOS: {huerfanos}")
else:
    print("  OK: todos los links tienen destino")

if pdf and pdf.exists():
    datos = pdf.read_bytes()
    print(f"\n{pdf.name}")
    print(f"  anotaciones /Link  : {datos.count(b'/Link')}")
    print(f"  destinos internos  : {datos.count(b'/Dest')}")
