# -*- coding: utf-8 -*-
"""Utilidad: convierte un .docx a PDF y rasteriza las paginas pedidas a PNG,
para poder mirar como quedo sin abrir Word.

    python ver.py "salida\\Orden de servicio SEMANA.docx" 1,2,3

Usa el mismo conversor que el generador: Word en Windows, LibreOffice en Mac.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from generar_orden import exportar_pdf   # noqa: E402

if len(sys.argv) < 2:
    sys.exit(__doc__)

docx = Path(sys.argv[1]).resolve()
paginas = [int(x) for x in sys.argv[2].split(',')] if len(sys.argv) > 2 else [1]
out = Path(sys.argv[3]) if len(sys.argv) > 3 else docx.parent / "_preview"
out.mkdir(parents=True, exist_ok=True)

pdf = out / (docx.stem + ".pdf")
exportar_pdf(docx, pdf)
print(f"PDF: {pdf}")

try:
    import pypdfium2 as pdfium
except ImportError:
    sys.exit("  (para las imagenes hace falta: pip install pypdfium2)")

doc = pdfium.PdfDocument(str(pdf))
print(f"  {len(doc)} paginas")
for n in paginas:
    if not 1 <= n <= len(doc):
        print(f"  pagina {n} fuera de rango")
        continue
    p = out / f"{docx.stem}_p{n:02d}.png"
    doc[n - 1].render(scale=1.7).to_pil().save(p)
    print(f"  {p}")
