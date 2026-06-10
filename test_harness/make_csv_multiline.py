# -*- coding: utf-8 -*-
"""Genera un CSV con saltos de línea dentro de celdas (como el que falló en Streamlit Cloud)."""
import csv, io
from pathlib import Path
OUT = Path(__file__).parent / "docs"

rows = [
    ["Nombre", "Área", "Comentarios"],
    ["María Soto", "Finanzas", "Revisó el informe.\nAprobó el plan."],
    ["Jorge Díaz", "TI | Datos", "Pendiente\nde respuesta"],
    ["Ana Núñez", "RRHH", "Sin observaciones"],
]

# Semicolon-separated, Windows endings (lo que exporta Excel en español)
buf = io.StringIO(newline='')
w = csv.writer(buf, delimiter=';', quoting=csv.QUOTE_MINIMAL)
for r in rows:
    w.writerow(r)
content = buf.getvalue().replace('\r\n', '\r\n')  # ya tiene CRLF en Windows
(OUT / "csv_multiline.csv").write_bytes(content.encode('cp1252'))
print("csv_multiline.csv OK")
