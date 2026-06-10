# -*- coding: utf-8 -*-
"""Ejecuta los conversores de streamlit_app.py sobre los docs de prueba."""
import sys, io, traceback
from pathlib import Path

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

import streamlit_app as app

DOCS = Path(__file__).parent / "docs"
OUT  = Path(__file__).parent / "out"
OUT.mkdir(exist_ok=True)

CASES = [
    ("informe_texto.pdf",             dict(include_images=True)),
    ("informe_texto.pdf",             dict(include_images=False)),
    ("portada_imagen_resto_texto.pdf", dict(include_images=False)),
    ("politica_vacaciones.docx",      dict(include_images=True)),
    ("politica_vacaciones.docx",      dict(include_images=False)),
    ("plan_comercial.pptx",           dict(include_images=True)),
    ("dotacion.xlsx",                 dict()),
    ("acta_utf8.txt",                 dict()),
    ("acta_latin1.txt",               dict()),
    ("nomina_semicolon.csv",          dict()),
    ("roster_comma.csv",              dict()),
    ("legacy.doc",                    dict(include_images=False)),
]

def convert(name, kwargs):
    fb  = (DOCS / name).read_bytes()
    ext = Path(name).suffix.lower()
    fn  = {'.pdf': app.convert_pdf, '.docx': app.convert_docx,
           '.doc': app.convert_docx,
           '.txt': app.convert_txt, '.pptx': app.convert_pptx,
           '.xlsx': app.convert_xlsx, '.xls': app.convert_xlsx,
           '.csv': app.convert_csv}[ext]
    if ext in ('.txt', '.xlsx', '.xls', '.csv'):
        return fn(fb, name)
    return fn(fb, name, **kwargs)

for name, kwargs in CASES:
    tag = "_img" if kwargs.get("include_images") else ""
    out_name = Path(name).stem + tag + ".md"
    try:
        md, method = convert(name, kwargs)
        (OUT / out_name).write_text(md, encoding="utf-8")
        print(f"OK    {out_name:45s} metodo={method:30s} {len(md):,} chars")
    except Exception as e:
        print(f"ERROR {out_name:45s} {type(e).__name__}: {e}")
        traceback.print_exc(limit=3)
