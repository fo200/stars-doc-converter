# -*- coding: utf-8 -*-
"""¿Cuánta RAM cuesta importar pymupdf4llm y cuándo se carga el modelo layout?"""
import gc, psutil
proc = psutil.Process()
def rss(): return proc.memory_info().rss / 1048576

print(f"baseline:                  {rss():.0f} MB")
import fitz
print(f"tras import fitz:          {rss():.0f} MB")
import pymupdf4llm
print(f"tras import pymupdf4llm:   {rss():.0f} MB")
pymupdf4llm.use_layout(False)
gc.collect()
print(f"tras use_layout(False):    {rss():.0f} MB")

from pathlib import Path
fb = (Path(__file__).parent / "docs" / "informe_texto.pdf").read_bytes()
doc = fitz.open(stream=fb, filetype="pdf")
md = pymupdf4llm.to_markdown(doc, filename="x.pdf")
doc.close()
gc.collect()
print(f"tras to_markdown clásico:  {rss():.0f} MB  (md={len(md)} chars)")
