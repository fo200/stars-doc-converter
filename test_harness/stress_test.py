# -*- coding: utf-8 -*-
"""
Stress test: simula el escenario real del usuario —
documentos consecutivos (memoria acumulada) e imágenes embebidas.
Mide tiempo y RAM por conversión y detecta leaks.
"""
import sys, gc, time, traceback
from pathlib import Path

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

import streamlit_app as app

DOCS = Path(__file__).parent / "docs"

try:
    import psutil
    PROC = psutil.Process()
    def rss_mb(): return PROC.memory_info().rss / 1024 / 1024
except ImportError:
    PROC = None
    def rss_mb(): return 0.0


def convert(name, **kwargs):
    fb = (DOCS / name).read_bytes()
    ext = Path(name).suffix.lower()
    fn = {'.pdf': app.convert_pdf, '.docx': app.convert_docx,
          '.txt': app.convert_txt, '.pptx': app.convert_pptx,
          '.xlsx': app.convert_xlsx, '.csv': app.convert_csv}[ext]
    if ext in ('.txt', '.xlsx', '.csv'):
        return fn(fb, name)
    return fn(fb, name, **kwargs)


# Simula "documentos seguidos": el mismo flujo que la app cuando el usuario
# convierte varios archivos uno tras otro (con y sin imágenes).
SEQUENCE = [
    ("stress_grande_imgs.pdf", dict(include_images=False)),
    ("stress_grande_imgs.pdf", dict(include_images=True)),
    ("politica_vacaciones.docx", dict(include_images=True)),
    ("plan_comercial.pptx", dict(include_images=True)),
    ("stress_grande_imgs.pdf", dict(include_images=True)),
    ("informe_texto.pdf", dict(include_images=True)),
    ("stress_grande_imgs.pdf", dict(include_images=True)),
]

if __name__ == "__main__":
    results = []
    gc.collect()
    base = rss_mb()
    print(f"RAM inicial: {base:.0f} MB\n")

    for i, (name, kw) in enumerate(SEQUENCE, 1):
        gc.collect()
        before = rss_mb()
        t0 = time.perf_counter()
        try:
            md, method = convert(name, **kw)
            dt = time.perf_counter() - t0
            gc.collect()
            after = rss_mb()
            tag = "img" if kw.get("include_images") else "txt"
            print(f"[{i}] OK    {name:32s} ({tag}) {dt:6.1f}s  "
                  f"md={len(md)/1024:8.0f} KB  RAM {before:.0f}->{after:.0f} MB (delta {after-before:+.0f})")
            results.append((name, dt, len(md)))
            del md
        except Exception as e:
            dt = time.perf_counter() - t0
            print(f"[{i}] ERROR {name:32s} {dt:6.1f}s  {type(e).__name__}: {e}")
            traceback.print_exc(limit=3)

    gc.collect()
    print(f"\nRAM final: {rss_mb():.0f} MB (inicial {base:.0f} MB) — "
          f"residuo no liberado: {rss_mb()-base:+.0f} MB")

    # OCR path (escaneado) — solo si tesseract está disponible
    try:
        import pytesseract
        pytesseract.get_tesseract_version()
        has_tess = True
    except Exception:
        has_tess = False
    if has_tess:
        gc.collect(); before = rss_mb(); t0 = time.perf_counter()
        try:
            md, method = convert("stress_escaneado.pdf", include_images=False)
            print(f"\nOCR OK  stress_escaneado.pdf  {time.perf_counter()-t0:.1f}s  "
                  f"md={len(md)} chars  metodo={method}  RAM {before:.0f}->{rss_mb():.0f} MB")
        except Exception as e:
            print(f"\nOCR ERROR stress_escaneado.pdf: {type(e).__name__}: {e}")
            traceback.print_exc(limit=3)
    else:
        print("\n(OCR omitido: tesseract no instalado localmente)")
