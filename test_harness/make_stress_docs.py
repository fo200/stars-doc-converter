# -*- coding: utf-8 -*-
"""Genera documentos grandes con imágenes para stress test."""
import io, zipfile, random
from pathlib import Path

DOCS = Path(__file__).parent / "docs"
DOCS.mkdir(exist_ok=True)

import fitz  # PyMuPDF

random.seed(42)


def _png_bytes(w=900, h=600, seed=0):
    """PNG sintético con ruido (no comprime bien → tamaño realista)."""
    from PIL import Image
    import random as rnd
    r = rnd.Random(seed)
    img = Image.new("RGB", (w, h))
    px = img.load()
    for y in range(0, h, 4):
        for x in range(0, w, 4):
            c = (r.randint(0, 255), r.randint(0, 255), r.randint(0, 255))
            for dy in range(4):
                for dx in range(4):
                    if x + dx < w and y + dy < h:
                        px[x + dx, y + dy] = c
    buf = io.BytesIO()
    img.save(buf, "PNG")
    return buf.getvalue()


def make_big_pdf_with_images(path, pages=40, imgs_per_page=2):
    doc = fitz.open()
    img_cache = [_png_bytes(seed=i) for i in range(4)]
    lorem = ("Informe de gestión trimestral. Los resultados consolidados muestran "
             "un crecimiento sostenido en todas las unidades de negocio. ") * 6
    for p in range(pages):
        page = doc.new_page()
        page.insert_text((50, 50), f"Sección {p + 1}", fontsize=18)
        rect = fitz.Rect(50, 80, 545, 400)
        page.insert_textbox(rect, lorem, fontsize=10)
        y = 420
        for k in range(imgs_per_page):
            r = fitz.Rect(60 + k * 240, y, 280 + k * 240, y + 150)
            page.insert_image(r, stream=img_cache[(p + k) % len(img_cache)])
    doc.save(str(path))
    doc.close()
    print(f"{path.name}: {path.stat().st_size/1024/1024:.1f} MB, {pages} páginas")


def make_scanned_pdf(path, pages=8):
    """PDF 'escaneado': solo imágenes de texto renderizado, sin capa de texto."""
    from PIL import Image, ImageDraw
    doc = fitz.open()
    for p in range(pages):
        img = Image.new("RGB", (1240, 1754), "white")
        d = ImageDraw.Draw(img)
        for i in range(40):
            d.text((80, 60 + i * 40), f"Linea {i+1} de la pagina escaneada {p+1} del informe anual", fill="black")
        buf = io.BytesIO()
        img.save(buf, "JPEG", quality=70)
        page = doc.new_page(width=595, height=842)
        page.insert_image(fitz.Rect(0, 0, 595, 842), stream=buf.getvalue())
    doc.save(str(path))
    doc.close()
    print(f"{path.name}: {path.stat().st_size/1024/1024:.1f} MB, {pages} páginas (escaneado)")


if __name__ == "__main__":
    make_big_pdf_with_images(DOCS / "stress_grande_imgs.pdf", pages=40, imgs_per_page=2)
    make_scanned_pdf(DOCS / "stress_escaneado.pdf", pages=8)
