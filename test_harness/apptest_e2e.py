# -*- coding: utf-8 -*-
"""
Loop 3 — prueba end-to-end con streamlit.testing.AppTest:
1. La app arranca sin excepciones (estado vacío).
2. Renderiza resultados, incluido un markdown de 40+ MB con data URIs
   (ejercita truncado de preview, strip de base64 y botones de descarga).
3. Resultados con error se muestran sin romper la UI.
"""
import sys, base64
from pathlib import Path

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

from streamlit.testing.v1 import AppTest

APP = str(ROOT / "streamlit_app.py")
fails = []

def check(label, cond, extra=""):
    print(("OK  " if cond else "FAIL"), label, extra)
    if not cond:
        fails.append(label)

# ── 1. Arranque limpio ────────────────────────────────────────────────────────
at = AppTest.from_file(APP, default_timeout=60)
at.run()
check("arranque sin excepcion", not at.exception,
      f"({at.exception[0].message if at.exception else ''})")

# ── 2. Resultado gigante con imágenes (simula post-conversión) ───────────────
fake_uri = "data:image/png;base64," + base64.b64encode(b"x" * 400_000).decode()
big_md = ("# Documento grande\n\n" +
          ("Párrafo de texto con contenido relevante. " * 50 + "\n\n" +
           f"![figura]({fake_uri})\n\n") * 80)   # ~43 MB
print(f"     markdown simulado: {len(big_md)/1024/1024:.1f} MB")

at = AppTest.from_file(APP, default_timeout=120)
at.session_state["results"] = {
    "grande.pdf": {
        "ok": True, "filename": "grande.pdf", "markdown": big_md,
        "method": "pymupdf4llm (word PDF)", "digest": "abc", "include_images": True,
        "words": 100, "chars": len(big_md), "lines": big_md.count("\n") + 1,
    },
    "fallido.docx": {"ok": False, "filename": "fallido.docx", "error": "Es un .doc antiguo"},
}
at.session_state["active"] = "grande.pdf"
at.run()
check("render resultado grande sin excepcion", not at.exception,
      f"({at.exception[0].message if at.exception else ''})")

# El preview enviado al navegador NO debe contener data URIs gigantes
all_md = "\n".join(str(m.value) for m in at.markdown)
check("preview sin base64 crudo", "base64," + "e" * 50 not in all_md and fake_uri[:200] not in all_md)
check("preview truncado", any("primeros" in str(m.value) for m in at.caption))
check("error mostrado sin romper UI", any(".doc antiguo" in str(c.value) for c in at.caption))

# tamaño total del frame renderizado: el preview no debe pesar MBs
total_render = sum(len(str(m.value)) for m in at.markdown)
check("payload de preview acotado", total_render < 2_000_000, f"({total_render/1024:.0f} KB)")

print()
if fails:
    print(f"FALLAS: {fails}")
    sys.exit(1)
print("E2E AppTest: todo OK.")
