# -*- coding: utf-8 -*-
"""Compara modo clásico (use_layout(False)) vs layout en velocidad/RAM."""
import sys, time, gc, tempfile
from pathlib import Path
import psutil, fitz, pymupdf4llm

pymupdf4llm.use_layout(False)
proc = psutil.Process()
fb = (Path(__file__).parent / "docs" / "stress_grande_imgs.pdf").read_bytes()

for run in range(3):
    gc.collect()
    b = proc.memory_info().rss / 1048576
    t0 = time.perf_counter()
    doc = fitz.open(stream=fb, filetype="pdf")
    with tempfile.TemporaryDirectory() as tmp:
        md = pymupdf4llm.to_markdown(doc, write_images=True, image_path=tmp,
                                     image_format="png", filename="stress.pdf")
    doc.close()
    fitz.TOOLS.store_shrink(100)   # vaciar caché global de MuPDF
    dt = time.perf_counter() - t0
    n = len(md)
    del md
    gc.collect()
    a = proc.memory_info().rss / 1048576
    print(f"run {run+1}: {dt:5.1f}s  md={n/1024:.0f} KB  RAM {b:.0f}->{a:.0f} MB (delta {a-b:+.1f})")
