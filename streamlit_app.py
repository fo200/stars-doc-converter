import streamlit as st
import io, re, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="STARS · Doc → Markdown",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Library checks ────────────────────────────────────────────────────────────
@st.cache_resource(show_spinner=False)
def _check_libs():
    libs = {}
    for key, mod in [("fitz","fitz"),("pymupdf4llm","pymupdf4llm"),
                     ("mammoth","mammoth"),("pdf2image","pdf2image"),
                     ("pytesseract","pytesseract"),
                     ("openpyxl","openpyxl"),("xlrd","xlrd")]:
        try:    __import__(mod); libs[key] = True
        except: libs[key] = False
    return libs

LIBS = _check_libs()

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

/* ─ Variables ─────────────────────────────────────────────────── */
:root {
  --dl:      #005A70;
  --teal:    #10B5BF;
  --green:   #01969F;
  --navy:    #061D3D;
  --dark:    #302E2E;
  --bg:      #EDF2F4;
  --border:  #C9D8DC;
  --white:   #FFFFFF;
  --shadow:  0 2px 16px rgba(0,90,112,.10);
  --font:    'Inter', 'Segoe UI', sans-serif;
}

/* ─ Hide Streamlit chrome ─────────────────────────────────────── */
#MainMenu, footer                         { visibility: hidden; }
[data-testid="stHeader"]                  { display: none !important; }
[data-testid="stToolbar"]                 { display: none !important; }
[data-testid="stDecoration"]              { display: none !important; }
[data-testid="stDeployButton"]            { display: none !important; }
[data-testid="stStatusWidget"]            { display: none !important; }

/* ─ Base ──────────────────────────────────────────────────────── */
html, body, .stApp                        { font-family: var(--font); background: var(--bg); }
[data-testid="stAppViewBlockContainer"]   { padding-top: 0 !important; }
.main .block-container                    { padding-top: 0 !important; padding-left: 0 !important;
                                            padding-right: 0 !important; max-width: 100% !important; }

/* ─ Column cards ──────────────────────────────────────────────── */
[data-testid="column"] {
  background: var(--white);
  border-radius: 14px;
  box-shadow: var(--shadow);
  padding: 24px 26px 28px !important;
}

/* ─ Buttons — primary (deep-lake) ────────────────────────────── */
.stButton > button {
  background: var(--dl)     !important;
  color: var(--white)       !important;
  border: none              !important;
  border-radius: 8px        !important;
  font-family: var(--font)  !important;
  font-weight: 600          !important;
  font-size: 13.5px         !important;
  padding: 10px 20px        !important;
  transition: background .18s !important;
  box-shadow: none          !important;
}
.stButton > button:hover    { background: var(--navy) !important; }
.stButton > button:active   { background: var(--navy) !important; }

/* ─ Select-file outline buttons ──────────────────────────────── */
.sel-row .stButton > button {
  background: transparent       !important;
  color: var(--teal)            !important;
  border: 1.5px solid var(--teal) !important;
  border-radius: 6px            !important;
  font-size: 12px               !important;
  padding: 4px 12px             !important;
  min-height: 0                 !important;
  height: auto                  !important;
}
.sel-row .stButton > button:hover {
  background: var(--teal) !important;
  color: var(--white)     !important;
}

/* ─ Download buttons (teal) ──────────────────────────────────── */
.stDownloadButton > button {
  background: var(--teal)   !important;
  color: var(--white)       !important;
  border: none              !important;
  border-radius: 8px        !important;
  font-family: var(--font)  !important;
  font-weight: 600          !important;
  font-size: 13px           !important;
  transition: background .18s !important;
}
.stDownloadButton > button:hover { background: var(--green) !important; }

/* ─ File uploader ─────────────────────────────────────────────── */
[data-testid="stFileUploader"] section {
  background: var(--bg)               !important;
  border: 2px dashed var(--border)    !important;
  border-radius: 12px                 !important;
  transition: border-color .2s;
}
[data-testid="stFileUploader"] section:hover {
  border-color: var(--teal) !important;
}
[data-testid="stFileUploaderDropzone"] {
  background: var(--bg) !important;
}

/* ─ Progress bar ──────────────────────────────────────────────── */
[data-testid="stProgress"] > div                { border-radius: 4px; overflow: hidden; }
[data-testid="stProgress"] > div > div > div    { background: var(--teal) !important; }
.stProgress > div > div > div                   { background: var(--teal) !important; }

/* ─ Tabs ──────────────────────────────────────────────────────── */
.stTabs [role="tablist"]                        { border-bottom: 2px solid var(--bg) !important; }
.stTabs [role="tab"]                            { font-weight: 600 !important; color: #9CA3AF !important; font-size: 13px !important; }
.stTabs [role="tab"][aria-selected="true"]      { color: var(--teal) !important; border-bottom-color: var(--teal) !important; }

/* ─ Panel heading ─────────────────────────────────────────────── */
.panel-heading {
  font-size: 11.5px; font-weight: 700; color: var(--dl);
  text-transform: uppercase; letter-spacing: .6px;
  padding-bottom: 14px; border-bottom: 1.5px solid var(--bg);
  margin-bottom: 18px;
}

/* ─ Format badges row ─────────────────────────────────────────── */
.fmt-row { display:flex; gap:7px; flex-wrap:wrap; margin-bottom:16px; }
.badge { display:inline-block; padding:3px 9px; border-radius:5px; font-size:11px; font-weight:700; letter-spacing:.3px; }
.b-pdf  { background:#FEE2E2; color:#B91C1C; }
.b-docx,.b-doc { background:#DBEAFE; color:#1D4ED8; }
.b-txt  { background:#FEF9C3; color:#78350F; }
.b-pptx { background:#FCE7F3; color:#9D174D; }
.b-xlsx,.b-xls { background:#D1FAE5; color:#065F46; }
.b-ok   { background:#D1FAE5; color:#065F46; }
.b-err  { background:#FEE2E2; color:#B91C1C; }

/* ─ File queue rows ───────────────────────────────────────────── */
.file-row {
  display:flex; align-items:center; gap:10px;
  padding:10px 13px; border-radius:9px; margin-bottom:6px;
  background:var(--bg); border:1.5px solid transparent;
}
.file-row.active { border-color:var(--teal); background:#E3F4F6; }
.file-row .fname { flex:1; font-weight:600; font-size:13px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; color:var(--dark); }
.file-row .fmeta { font-size:11.5px; color:#9CA3AF; white-space:nowrap; }

/* ─ Stats cards ───────────────────────────────────────────────── */
.stat-row  { display:flex; gap:10px; margin:14px 0 18px; }
.stat-card { flex:1; background:var(--bg); border-radius:9px; padding:12px 8px; text-align:center; }
.stat-val  { font-weight:700; font-size:1.3rem; color:var(--teal); }
.stat-lbl  { font-size:10.5px; color:#9CA3AF; text-transform:uppercase; letter-spacing:.05em; margin-top:2px; }

/* ─ Method pill ───────────────────────────────────────────────── */
.method-pill { background:var(--dl); color:white; padding:3px 9px; border-radius:4px; font-size:11px; font-weight:700; margin-left:8px; }

/* ─ Empty state ───────────────────────────────────────────────── */
.empty-state { text-align:center; padding:64px 24px; color:#B0BFC5; }
.empty-icon  { font-size:3rem; margin-bottom:14px; opacity:.6; }
.empty-title { font-size:15px; font-weight:600; color:#8FA3AB; }
.empty-sub   { font-size:12.5px; margin-top:6px; line-height:1.6; color:#B0BFC5; }

/* ─ Markdown preview ──────────────────────────────────────────── */
.md-preview h1 { color:var(--navy); border-bottom:2px solid var(--teal); padding-bottom:6px; }
.md-preview h2 { color:var(--dl); }
.md-preview h3 { color:var(--green); }
.md-preview table  { border-collapse:collapse; width:100%; font-size:.88rem; }
.md-preview thead tr { background:var(--dl); color:white; }
.md-preview th, .md-preview td { padding:8px 12px; border:1px solid #D1D5DB; }
.md-preview tbody tr:nth-child(even) { background:#F4FAFA; }
.md-preview hr  { border:none; border-top:1.5px solid var(--border); }
.md-preview code { background:var(--bg); padding:2px 6px; border-radius:4px; font-size:.87em; }
.md-preview blockquote { border-left:3px solid var(--teal); padding-left:14px; color:#555; }
.md-preview a   { color:var(--teal); }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="
  background:#005A70; padding:0 40px; height:64px;
  display:flex; align-items:center; gap:14px;
  position:relative; margin-bottom:28px;
  font-family:'Inter','Segoe UI',sans-serif;
">
  <div style="
    width:38px;height:38px;border-radius:50%;background:#10B5BF;
    display:flex;align-items:center;justify-content:center;
    font-weight:700;font-size:15px;color:white;flex-shrink:0;
  ">S</div>
  <span style="color:white;font-weight:700;font-size:17px;">STARS Companies</span>
  <div style="width:1px;height:22px;background:rgba(255,255,255,.3);"></div>
  <span style="color:rgba(255,255,255,.72);font-size:13px;">Document → Markdown Converter</span>
  <div style="
    position:absolute;bottom:0;left:0;right:0;height:3px;
    background:linear-gradient(90deg,#10B5BF,#01969F);
  "></div>
</div>
""", unsafe_allow_html=True)

# ── Conversion functions ──────────────────────────────────────────────────────

def convert_pdf(file_bytes, filename):
    if not LIBS["fitz"]: raise ImportError("PyMuPDF no instalado.")
    import fitz, pymupdf4llm
    doc   = fitz.open(stream=file_bytes, filetype="pdf")
    page0 = doc[0]
    text0 = page0.get_text().strip()
    if len(text0) < 30:
        if not (LIBS["pdf2image"] and LIBS["pytesseract"]):
            raise ImportError("PDF escaneado: instala pdf2image y pytesseract.")
        from pdf2image import convert_from_bytes
        import pytesseract
        imgs = convert_from_bytes(file_bytes, dpi=300)
        pages = [f"## Página {i}\n\n{pytesseract.image_to_string(img, lang='spa+eng').strip()}"
                 for i, img in enumerate(imgs, 1)]
        return f"# {Path(filename).stem}\n\n" + "\n\n---\n\n".join(pages), "OCR (pytesseract)"
    CHUNK = 20
    total = len(doc)
    parts = [pymupdf4llm.to_markdown(doc, pages=list(range(s, min(s+CHUNK, total))))
             for s in range(0, total, CHUNK)]
    md = re.sub(r'\n{4,}', '\n\n\n', "\n\n".join(parts))
    label = "pymupdf4llm (word PDF)" if page0.get_fonts() else "pymupdf4llm (plain PDF)"
    return md, label


def convert_docx(file_bytes, filename):
    if not LIBS["mammoth"]: raise ImportError("mammoth no instalado.")
    import mammoth
    return mammoth.convert_to_markdown(io.BytesIO(file_bytes)).value, "mammoth (DOCX)"


def convert_txt(file_bytes, filename):
    return f"# {Path(filename).stem}\n\n{file_bytes.decode('utf-8', errors='replace')}", "texto plano"


def convert_pptx(file_bytes, filename):
    A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
    def pt(el):    return ''.join(t.text or '' for t in el.iter(f'{A}t')).strip()
    def is_title(sp): return any(ph.get('type','') in ('title','ctrTitle') or ph.get('idx','9')=='0'
                                  for ph in sp.iter(f'{P}ph'))
    slides = []
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
        fnames = sorted([f for f in z.namelist()
                         if f.startswith('ppt/slides/slide') and f.endswith('.xml')
                         and 'Layout' not in f and 'Master' not in f],
                        key=lambda x: int(''.join(c for c in x if c.isdigit()) or '0'))
        for i, sf in enumerate(fnames, 1):
            root = ET.parse(z.open(sf)).getroot()
            ttl, body = '', []
            for sp in root.iter(f'{P}sp'):
                paras = [pt(p) for p in sp.iter(f'{A}p') if pt(p)]
                if not paras: continue
                if not ttl and is_title(sp): ttl = paras[0]; body.extend(paras[1:])
                else: body.extend(paras)
            for gf in root.iter(f'{P}graphicFrame'):
                body.extend(pt(p) for p in gf.iter(f'{A}p') if pt(p))
            hdr = f"## Slide {i}" + (f": {ttl}" if ttl else "")
            slides.append(f"{hdr}\n\n" + '\n\n'.join(body) if body else hdr)
    return '\n\n---\n\n'.join(slides), "XML directo (PPTX)"


def _fmt_cell(value) -> str:
    import datetime
    if value is None: return ""
    if isinstance(value, datetime.datetime):
        if value.hour == value.minute == value.second == 0:
            return value.strftime("%Y-%m-%d")
        return value.strftime("%Y-%m-%d %H:%M")
    if isinstance(value, datetime.date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return str(value)


def _md_table(rows):
    if not rows: return "*(vacío)*"
    n = max(len(r) for r in rows)
    pad = lambda r: r + [''] * (n - len(r))
    lines = ['| ' + ' | '.join(pad(rows[0])) + ' |', '|' + '--- |' * n]
    for r in rows[1:]: lines.append('| ' + ' | '.join(pad(r)) + ' |')
    return '\n'.join(lines)


def convert_xlsx(file_bytes, filename):
    ext = Path(filename).suffix.lower()
    if ext == '.xls' and LIBS["xlrd"]:
        try:
            import xlrd
            wb = xlrd.open_workbook(file_contents=file_bytes)
            sheets = [f"## {s.name}\n\n" + _md_table(
                [[str(s.cell_value(r, c)) for c in range(s.ncols)] for r in range(s.nrows)]
            ) for s in wb.sheets()]
            return f"# {Path(filename).stem}\n\n" + '\n\n---\n\n'.join(sheets), "xlrd (XLS)"
        except: pass
    if not LIBS["openpyxl"]: raise ImportError("openpyxl no instalado.")
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    base = Path(filename).stem
    parts = [f"# {base}\n"]
    for ws in wb.worksheets:
        all_rows = list(ws.iter_rows(values_only=True))
        last = max((i for i, r in enumerate(all_rows)
                    if any(v is not None and str(v).strip() for v in r)), default=-1)
        if last < 0: continue
        fmt = [[_fmt_cell(c) for c in row] for row in all_rows[:last + 1]]
        fmt = [r for r in fmt if any(c.strip() for c in r)]
        if not fmt: continue
        parts.append(f"\n## {ws.title}\n\n" + _md_table(fmt))
    return '\n'.join(parts), "openpyxl (XLSX)"


@st.cache_data(show_spinner=False)
def run_conversion(file_bytes: bytes, filename: str) -> dict:
    ext = Path(filename).suffix.lower()
    fn = {'.pdf': convert_pdf, '.docx': convert_docx, '.doc': convert_docx,
          '.txt': convert_txt, '.pptx': convert_pptx,
          '.xlsx': convert_xlsx, '.xls': convert_xlsx}
    if ext not in fn: raise ValueError(f"Formato no soportado: {ext}")
    md, method = fn[ext](file_bytes, filename)
    return dict(filename=filename, markdown=md, method=method,
                words=len(md.split()), chars=len(md), lines=md.count('\n') + 1)


# ── Session state ─────────────────────────────────────────────────────────────
if "results" not in st.session_state: st.session_state.results = {}
if "active"  not in st.session_state: st.session_state.active  = None

# ── Layout: two-column ────────────────────────────────────────────────────────
col_left, col_right = st.columns([1, 1.65], gap="large")


# ════════════════════════════════ LEFT ═══════════════════════════════════════
with col_left:

    # Panel heading
    st.markdown('<div class="panel-heading">📁 &nbsp;Subir documentos</div>', unsafe_allow_html=True)

    # Format badges
    st.markdown("""
    <div class="fmt-row">
      <span class="badge b-pdf">PDF</span>
      <span class="badge b-docx">DOCX · DOC</span>
      <span class="badge b-pptx">PPTX</span>
      <span class="badge b-xlsx">XLSX · XLS</span>
      <span class="badge b-txt">TXT</span>
    </div>""", unsafe_allow_html=True)

    # File uploader
    uploaded = st.file_uploader(
        "Arrastra archivos aquí o haz clic para seleccionar",
        type=["pdf", "docx", "doc", "txt", "pptx", "xlsx", "xls"],
        accept_multiple_files=True,
    )

    # Convert button
    if uploaded:
        if st.button("⟳  Convertir todo", type="primary", use_container_width=True):
            st.session_state.results = {}
            st.session_state.active  = None
            bar = st.progress(0, text="Iniciando…")
            for i, uf in enumerate(uploaded):
                bar.progress(i / len(uploaded), text=f"Convirtiendo {uf.name}…")
                try:
                    fb = uf.getvalue()
                    if len(fb) == 0:
                        raise ValueError("Archivo vacío.")
                    r = run_conversion(fb, uf.name)
                    st.session_state.results[uf.name] = {"ok": True, **r}
                except Exception as e:
                    st.session_state.results[uf.name] = {"ok": False, "filename": uf.name, "error": str(e)}
            bar.progress(1.0, text="¡Listo!")
            for name, r in st.session_state.results.items():
                if r["ok"]: st.session_state.active = name; break
            st.rerun()

    # File queue
    results = st.session_state.results
    if results:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="panel-heading" style="font-size:10.5px;">Cola de conversión</div>',
                    unsafe_allow_html=True)

        for name, r in results.items():
            ext = Path(name).suffix.lower().lstrip('.')
            is_active  = name == st.session_state.active
            status_cls = "b-ok" if r["ok"] else "b-err"
            status_lbl = "✓ Listo" if r["ok"] else "✗ Error"
            words_str  = f"{r['words']:,} palabras" if r.get("words") else ""
            st.markdown(f"""
            <div class="file-row {'active' if is_active else ''}">
              <span class="badge b-{ext}">{ext.upper()}</span>
              <span class="fname">{name}</span>
              <span class="fmeta">{words_str}</span>
              <span class="badge {status_cls}">{status_lbl}</span>
            </div>""", unsafe_allow_html=True)

            with st.container():
                st.markdown('<div class="sel-row">', unsafe_allow_html=True)
                if r["ok"] and st.button(f"Ver resultado", key=f"sel_{name}"):
                    st.session_state.active = name
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)

            if not r["ok"]:
                st.caption(f"⚠️ {r.get('error', 'Error desconocido')}")

        # ZIP download when multiple files done
        done_ok = {n: r for n, r in results.items() if r.get("ok")}
        if len(done_ok) > 1:
            st.markdown("---")
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                for name, r in done_ok.items():
                    zf.writestr(Path(name).stem + ".md", r["markdown"])
            st.download_button(
                "⬇ Descargar todo (.zip)",
                data=buf.getvalue(),
                file_name="documentos_markdown.zip",
                mime="application/zip",
                use_container_width=True,
            )


# ════════════════════════════════ RIGHT ══════════════════════════════════════
with col_right:

    active  = st.session_state.active
    results = st.session_state.results

    # ── Empty state
    if not active or not results.get(active, {}).get("ok"):
        st.markdown('<div class="panel-heading">📝 &nbsp;Resultado Markdown</div>', unsafe_allow_html=True)
        st.markdown("""
        <div class="empty-state">
          <div class="empty-icon">📄</div>
          <div class="empty-title">Ningún archivo convertido aún</div>
          <div class="empty-sub">Sube uno o más archivos y haz clic en<br><b>Convertir todo</b> para ver el resultado aquí.</div>
        </div>""", unsafe_allow_html=True)

    # ── Result
    else:
        r  = results[active]
        md = r["markdown"]

        # Heading + download button in one row
        h_col, d_col = st.columns([3, 1])
        with h_col:
            st.markdown(
                f'<div class="panel-heading" style="border:none;padding-bottom:0;margin-bottom:0;">'
                f'📝 &nbsp;{r["filename"]}'
                f'<span class="method-pill">{r["method"]}</span>'
                f'</div>',
                unsafe_allow_html=True,
            )
        with d_col:
            st.download_button(
                "⬇ Descargar .md",
                data=md,
                file_name=Path(r["filename"]).stem + ".md",
                mime="text/markdown",
                use_container_width=True,
            )

        # Stats
        st.markdown(f"""
        <div class="stat-row">
          <div class="stat-card">
            <div class="stat-val">{r['words']:,}</div>
            <div class="stat-lbl">Palabras</div>
          </div>
          <div class="stat-card">
            <div class="stat-val">{r['chars']:,}</div>
            <div class="stat-lbl">Caracteres</div>
          </div>
          <div class="stat-card">
            <div class="stat-val">{r['lines']:,}</div>
            <div class="stat-lbl">Líneas</div>
          </div>
        </div>""", unsafe_allow_html=True)

        # Tabs: preview / raw
        tab_prev, tab_raw = st.tabs(["  Vista previa  ", "  Markdown raw  "])
        with tab_prev:
            st.markdown('<div class="md-preview">', unsafe_allow_html=True)
            st.markdown(md)
            st.markdown('</div>', unsafe_allow_html=True)
        with tab_raw:
            st.code(md, language="markdown", line_numbers=True)
