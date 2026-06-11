# -*- coding: utf-8 -*-
"""Mide leak por llamada: convierte el mismo PDF N veces seguidas."""
import sys, gc, time
from pathlib import Path

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))
import streamlit_app as app
import psutil

DOCS = Path(__file__).parent / "docs"
PROC = psutil.Process()
def rss(): return PROC.memory_info().rss / 1024 / 1024

name = sys.argv[1] if len(sys.argv) > 1 else "informe_texto.pdf"
fb = (DOCS / name).read_bytes()

print(f"Doc: {name} ({len(fb)/1024:.0f} KB) — 6 conversiones seguidas con imágenes")
for i in range(6):
    gc.collect()
    b = rss(); t0 = time.perf_counter()
    md, _ = app.convert_pdf(fb, name, include_images=True)
    dt = time.perf_counter() - t0
    del md; gc.collect()
    print(f"  run {i+1}: {dt:5.1f}s  RAM {b:6.0f} -> {rss():6.0f} MB (delta {rss()-b:+6.1f})")
