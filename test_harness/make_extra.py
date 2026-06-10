# -*- coding: utf-8 -*-
from pathlib import Path
OUT = Path(__file__).parent / "docs"
csv_es = "Nombre;Cargo;Sueldo\nMaría Soto;Analista Sr.;2.450.000\nJorge Díaz;Jefe TI | Infra;3.100.000\n"
(OUT / "nomina_semicolon.csv").write_bytes(csv_es.encode("cp1252"))
csv_en = 'name,role,salary\nAna,HR Lead,3200\nLuis,"Analyst, Jr",1800\n'
(OUT / "roster_comma.csv").write_bytes(csv_en.encode("utf-8"))
(OUT / "legacy.doc").write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 600)
print("extra docs OK")
