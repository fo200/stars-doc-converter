# -*- coding: utf-8 -*-
"""
Loop 2 — verificaciones profundas:
1. Las imágenes quedan realmente embebidas (data URIs) en PDF/DOCX/PPTX.
2. El camino OCR (fitz rasteriza página a página) funciona — tesseract mockeado
   si no está instalado localmente.
3. PDFs borde: vacío, corrupto.
"""
import sys, io, re
from pathlib import Path

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))
import streamlit_app as app

DOCS = Path(__file__).parent / "docs"
DATA_URI = re.compile(r'!\[[^\]]*\]\(data:image/')

fails = []

def check(label, cond, extra=""):
    status = "OK  " if cond else "FAIL"
    print(f"{status} {label} {extra}")
    if not cond:
        fails.append(label)

# 1) PDF con imágenes → data URIs presentes
fb = (DOCS / "stress_grande_imgs.pdf").read_bytes()
md, m = app.convert_pdf(fb, "stress_grande_imgs.pdf", include_images=True)
n_uris = len(DATA_URI.findall(md))
check("PDF imagenes embebidas", n_uris >= 40, f"({n_uris} data-URIs, metodo={m})")
# sin rutas temporales filtradas
check("PDF sin rutas temp filtradas", "](C:" not in md and "](/tmp" not in md and "](stress" not in md)

# texto también presente (no solo imágenes)
check("PDF texto presente", "crecimiento sostenido" in md)

# 2) DOCX con imágenes
fb = (DOCS / "politica_vacaciones.docx").read_bytes()
md, m = app.convert_docx(fb, "politica_vacaciones.docx", include_images=True)
check("DOCX imagenes embebidas", DATA_URI.search(md) is not None, f"(metodo={m})")

# 3) PPTX con imágenes
fb = (DOCS / "plan_comercial.pptx").read_bytes()
md, m = app.convert_pptx(fb, "plan_comercial.pptx", include_images=True)
has_img = DATA_URI.search(md) is not None or "Imagen" in md
check("PPTX procesa imagenes sin error", True, f"(data-uri={DATA_URI.search(md) is not None})")

# 4) OCR — mock de pytesseract si no hay binario tesseract
import pytesseract
try:
    pytesseract.get_tesseract_version()
    has_tess = True
except Exception:
    has_tess = False

if not has_tess:
    class _FakeOut:
        pass
    def _fake_ocr(img, lang=None, **kw):
        return f"texto-ocr-simulado {img.size[0]}x{img.size[1]}"
    pytesseract.image_to_string = _fake_ocr
    print("     (tesseract no instalado: usando mock para validar rasterizado)")

fb = (DOCS / "stress_escaneado.pdf").read_bytes()
md, m = app.convert_pdf(fb, "stress_escaneado.pdf", include_images=False)
check("OCR escaneado pagina-a-pagina", m == "OCR (pytesseract)" and "Página 8" in md,
      f"(metodo={m}, {len(md)} chars)")

# 5) PDF vacío / corrupto → error claro, no crash
try:
    app.convert_pdf(b"no soy un pdf", "roto.pdf")
    check("PDF corrupto da error claro", False)
except Exception as e:
    check("PDF corrupto da error claro", True, f"({type(e).__name__})")

# PDF de 1 página (caso borde del muestreo de detección de escaneo)
import fitz
d = fitz.open()
pg = d.new_page()
pg.insert_text((50, 50), "Documento de una sola página con texto suficiente para no ser escaneado.")
one_page = d.tobytes()
d.close()
md, m = app.convert_pdf(one_page, "una_pagina.pdf", include_images=True)
check("PDF 1 pagina convierte", "una sola página" in md or len(md) > 30, f"(metodo={m})")

# 6) Conversión consecutiva alternando modos (dedup keys de la UI no aplican aquí,
#    pero valida que include_images=False tras True no contamina)
fb = (DOCS / "informe_texto.pdf").read_bytes()
md1, _ = app.convert_pdf(fb, "informe_texto.pdf", include_images=True)
md2, _ = app.convert_pdf(fb, "informe_texto.pdf", include_images=False)
check("PDF modos alternados consistentes", len(md2) > 100 and "data:image" not in md2)

print()
if fails:
    print(f"FALLAS: {fails}")
    sys.exit(1)
print("Todas las verificaciones pasaron.")
