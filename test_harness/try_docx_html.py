# -*- coding: utf-8 -*-
"""Prueba: mammoth -> HTML -> markdownify, ¿conserva tablas?"""
import io
from pathlib import Path
import mammoth
from markdownify import markdownify as mdify

fb = (Path(__file__).parent / "docs" / "politica_vacaciones.docx").read_bytes()
result = mammoth.convert_to_html(io.BytesIO(fb), convert_image=mammoth.images.img_element(lambda img: {}))
html = result.value
md = mdify(html, heading_style="ATX", bullets="-")
print(md[:2500])
