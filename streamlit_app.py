import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import html as html_lib
import json as _json
from docx import Document
from docx.shared import RGBColor, Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import io
import os

st.set_page_config(
    page_title="Gene Coverage Analyzer",
    page_icon="🧬",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ── Stitch Design System ───────────────────────────────────────────────────
# CSS is injected via JS (components.html) so Streamlit Cloud cannot strip
# the <style> tag the way it does with st.markdown.
_GCA_CSS = """
/* ── Tokens ──────────────────────────────────────────────────────────────── */
:root{
  --cp:        #ad2c00;
  --cp-dk:     #872000;
  --cs:        #565e74;
  --bg:        #f8f9ff;
  --surf:      #ffffff;
  --surf-lo:   #eff4ff;
  --surf-mid:  #e5eeff;
  --surf-hi:   #dce9ff;
  --surf-max:  #d3e4fe;
  --out:       #926f66;
  --out-v:     #e7bdb2;
  --on-surf:   #0b1c30;
  --err:       #ba1a1a;
  --err-c:     #ffdad6;
  --err-on:    #93000a;
  --r-card:    12px;
  --r-pill:    100px;
  --sh:        0 2px 4px rgba(0,0,0,.05);
  --sh-hover:  0 4px 12px rgba(0,0,0,.08);
}

/* ── Base ────────────────────────────────────────────────────────────────── */
html,body,[class*="css"],.stMarkdown{font-family:'Inter',sans-serif!important}
.stApp{background:var(--bg)!important}
.block-container{
  padding-top:0!important;
  padding-left:2rem!important;padding-right:2rem!important;
  max-width:1280px!important;
}

/* ── Hide chrome ─────────────────────────────────────────────────────────── */
#MainMenu,footer,[data-testid="stToolbar"],[data-testid="stDecoration"],.stDeployButton{display:none!important}
header[data-testid="stHeader"]{display:none!important}

/* ── Top navbar ──────────────────────────────────────────────────────────── */
.gca-nav{
  background:var(--surf);border-bottom:1px solid var(--out-v);
  height:60px;padding:0 1.5rem;
  display:flex;align-items:center;justify-content:space-between;
  margin:0 -2rem 1.75rem -2rem;
  box-shadow:0 1px 3px rgba(0,0,0,.06);
  position:sticky;top:0;z-index:999;
}
.gca-brand{display:flex;align-items:center;gap:10px}
.gca-brand-icon{
  font-family:'Material Symbols Outlined';font-size:26px;color:var(--cp);
  font-variation-settings:'FILL' 0,'wght' 400,'GRAD' 0,'opsz' 24;
}
.gca-brand-name{font-size:17px;font-weight:700;color:var(--cp);line-height:1.15;display:block}
.gca-brand-sub{font-size:11px;color:var(--cs);letter-spacing:.02em;display:block}
.gca-nav-links{display:flex;gap:1.25rem}
.gca-nav-links a{
  font-size:13px;color:var(--cs);text-decoration:none;
  padding-bottom:2px;border-bottom:2px solid transparent;transition:color .15s;
}
.gca-nav-links a:hover{color:var(--cp)}
.gca-nav-links a.active{color:var(--cp);border-bottom-color:var(--cp);font-weight:600}

/* ── Section label ───────────────────────────────────────────────────────── */
.gca-lbl{
  text-align:center;font-size:15px;font-weight:600;color:var(--on-surf);
  margin-bottom:.75rem;display:block;width:100%;
}

/* ── Step-1 centering wrapper ────────────────────────────────────────────── */
.gca-step1{
  display:flex;flex-direction:column;align-items:center;
  padding:1.5rem 0 1.25rem;
}

/* ── Radio → pill toggle ─────────────────────────────────────────────────── */
div[data-testid="stRadio"]{
  text-align:center!important;
}
div[data-testid="stRadio"]>div{
  background:var(--surf-hi);border-radius:var(--r-pill);
  border:1px solid var(--out-v);padding:3px;
  gap:0!important;display:inline-flex!important;
}
div[data-testid="stRadio"]>div>label{
  border-radius:var(--r-pill);padding:4px 20px!important;
  margin:0!important;font-size:13px;font-weight:400;color:var(--cs);
  transition:all .15s;cursor:pointer;
}
div[data-testid="stRadio"]>div>label:has(input:checked){
  background:var(--surf)!important;
  box-shadow:0 1px 3px rgba(0,0,0,.14);
  color:var(--cp)!important;font-weight:600;
}
div[data-testid="stRadio"]>div>label>div:first-child{display:none!important}

/* ── Upload column cards ─────────────────────────────────────────────────── */
/* Target only the direct stColumn children of the main upload horizontal block.
   We scope to the first stHorizontalBlock that has exactly 2 columns. */
[data-testid="stHorizontalBlock"]:first-of-type > [data-testid="stColumn"] > div:first-child{
  background:var(--surf);border:1px solid var(--out-v);
  border-radius:var(--r-card);box-shadow:var(--sh);
  padding:1.25rem 1.25rem 1rem !important;
  min-height:200px;
}

/* ── Buttons ─────────────────────────────────────────────────────────────── */
.stButton>button{
  background:var(--cp)!important;color:#fff!important;
  border:none!important;border-radius:8px!important;
  font-family:'Inter',sans-serif!important;font-size:14px!important;font-weight:500!important;
  padding:.5rem 1.25rem!important;letter-spacing:.01em;
  box-shadow:0 1px 3px rgba(173,44,0,.3)!important;
  transition:opacity .2s,box-shadow .2s!important;
}
.stButton>button:hover{opacity:.88!important;box-shadow:0 3px 8px rgba(173,44,0,.3)!important}
.stButton>button:active{opacity:.75!important}
.gca-sec .stButton>button{
  background:var(--surf)!important;color:var(--on-surf)!important;
  border:1px solid var(--out)!important;box-shadow:none!important;
}
.gca-sec .stButton>button:hover{border-color:var(--cp)!important;color:var(--cp)!important;opacity:1!important}

/* ── Download buttons ────────────────────────────────────────────────────── */
.stDownloadButton>button{
  background:var(--surf)!important;color:var(--on-surf)!important;
  border:1px solid var(--out)!important;border-radius:8px!important;
  font-family:'Inter',sans-serif!important;font-size:14px!important;font-weight:500!important;
  box-shadow:none!important;
}
.stDownloadButton>button:hover{border-color:var(--cp)!important;color:var(--cp)!important}

/* ── File uploader ───────────────────────────────────────────────────────── */
[data-testid="stFileUploader"] section{
  border:2px dashed var(--out)!important;border-radius:10px!important;
  background:white!important;transition:border-color .2s,background .2s!important;
}
[data-testid="stFileUploader"] section:hover{
  border-color:var(--cp)!important;background:rgba(173,44,0,.02)!important;
}

/* ── Tabs ────────────────────────────────────────────────────────────────── */
.stTabs [data-baseweb="tab-list"]{
  border-bottom:2px solid var(--out-v);gap:4px;background:transparent!important;
}
.stTabs [data-baseweb="tab"]{
  font-family:'Inter',sans-serif!important;font-size:13px!important;
  color:var(--cs)!important;padding:6px 12px!important;
  background:transparent!important;border-bottom:2px solid transparent;
}
.stTabs [data-baseweb="tab"][aria-selected="true"]{
  color:var(--cp)!important;border-bottom-color:var(--cp)!important;
  font-weight:600!important;background:transparent!important;
}
.stTabs [data-baseweb="tab-highlight"],.stTabs [data-baseweb="tab-border"]{display:none!important}

/* ── Expanders ───────────────────────────────────────────────────────────── */
[data-testid="stExpander"]{
  border:1px solid var(--out-v)!important;border-radius:var(--r-card)!important;
  background:var(--surf)!important;box-shadow:var(--sh)!important;overflow:hidden;
}
[data-testid="stExpander"] summary{
  font-family:'Inter',sans-serif!important;font-size:15px!important;
  font-weight:600!important;color:var(--on-surf)!important;
  background:var(--surf)!important;padding:1rem 1.25rem!important;
}
[data-testid="stExpander"] summary:hover{background:var(--surf-lo)!important}
[data-testid="stExpander"] summary svg{color:var(--cs)!important}

/* ── Text area ───────────────────────────────────────────────────────────── */
.stTextArea>div>div>textarea{
  font-family:'JetBrains Mono',monospace!important;font-size:13px!important;
  border-radius:8px!important;border-color:var(--out)!important;background:white!important;
}
.stTextArea>div>div>textarea:focus{
  border-color:var(--cp)!important;
  box-shadow:0 0 0 3px rgba(173,44,0,.12)!important;outline:none!important;
}

/* ── Headings ────────────────────────────────────────────────────────────── */
h3{font-size:15px!important;font-weight:600!important;color:var(--on-surf)!important}

/* ── Custom metric cards ─────────────────────────────────────────────────── */
.gca-metrics{display:grid;grid-template-columns:repeat(3,1fr);gap:1rem;margin:1.25rem 0}
.gca-metric{
  background:white;border:1px solid var(--out-v);border-radius:10px;
  padding:1.2rem 1.4rem;display:flex;flex-direction:column;
}
.gca-metric.err{
  background:rgba(186,26,26,.05);border-color:var(--err-c);
  position:relative;overflow:hidden;
}
.gca-metric.err::after{
  content:'';position:absolute;top:-20px;right:-20px;
  width:70px;height:70px;background:var(--err-c);border-radius:50%;opacity:.5;
}
.gca-metric-lbl{
  font-size:10px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;
  color:var(--cs);margin-bottom:.4rem;
}
.gca-metric.err .gca-metric-lbl{color:var(--err-on)}
.gca-metric-val{font-size:34px;font-weight:700;letter-spacing:-.02em;line-height:1.1;color:var(--on-surf)}
.gca-metric.err .gca-metric-val{color:var(--err)}

/* ── Results header ──────────────────────────────────────────────────────── */
.gca-rh{
  display:flex;justify-content:space-between;align-items:flex-end;
  padding-bottom:1rem;border-bottom:1px solid var(--out-v);margin-bottom:1.25rem;
}
.gca-rh-ttl{font-size:22px;font-weight:600;color:var(--on-surf);line-height:1.2}
.gca-rh-sub{font-size:12px;color:var(--cs);margin-top:2px}

/* ── Gene badges ─────────────────────────────────────────────────────────── */
.gca-gene-grid{display:flex;flex-wrap:wrap;gap:6px;max-height:320px;overflow-y:auto}
.gca-gene{
  display:inline-flex;align-items:center;padding:3px 12px;
  border-radius:var(--r-pill);font-family:'JetBrains Mono',monospace;font-size:12px;
  transition:transform .1s;cursor:default;
}
.gca-gene:hover{transform:translateY(-1px)}
.gca-gene.ok{background:var(--surf-hi);color:var(--on-surf);border:1px solid var(--out-v)}
.gca-gene.low{
  background:var(--err-c);color:var(--err-on);
  border:1px solid rgba(186,26,26,.35);font-weight:600;
}
.gca-gene .pct{font-size:10px;opacity:.7;margin-left:4px}

/* ── Card section wrapper ────────────────────────────────────────────────── */
.gca-card{
  background:var(--surf);border:1px solid var(--out-v);
  border-radius:var(--r-card);box-shadow:var(--sh);padding:1.5rem;
}
.gca-card-title{
  font-size:16px;font-weight:600;color:var(--on-surf);
  display:flex;align-items:center;justify-content:space-between;margin-bottom:1rem;
}
.gca-badge-ready{
  font-size:11px;font-weight:500;background:rgba(0,99,135,.08);color:#006387;
  border:1px solid rgba(0,99,135,.2);padding:2px 10px;border-radius:100px;
  display:inline-flex;align-items:center;gap:6px;
}
.gca-badge-ready::before{content:'';width:6px;height:6px;border-radius:50%;background:#006387}

/* ── Status / progress / divider ─────────────────────────────────────────── */
.stSuccess>div,.stError>div,.stInfo>div,.stWarning>div{border-radius:8px!important}
.stProgress>div>div{background:var(--cp)!important}
.stSpinner>div{border-top-color:var(--cp)!important}
hr{border-color:var(--out-v)!important;margin:1.5rem 0!important}

/* ── Footer bar ──────────────────────────────────────────────────────────── */
.gca-footer{
  text-align:center;padding:.75rem 1rem;background:var(--surf-max);
  border-top:1px solid var(--out-v);font-size:11px;color:var(--cs);
  display:flex;align-items:center;justify-content:center;gap:6px;
  margin:2rem -2rem -2rem -2rem;
}
.gca-footer .ms{
  font-family:'Material Symbols Outlined';font-size:14px;
  font-variation-settings:'FILL' 0,'wght' 400,'GRAD' 0,'opsz' 24;
}

@media(max-width:768px){
  .gca-metrics{grid-template-columns:1fr}
  .gca-nav-links{display:none}
  .block-container{padding-left:1rem!important;padding-right:1rem!important}
}
"""

_FONT_URLS = [
    "https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700"
    "&family=JetBrains+Mono:wght@400&display=swap",
    "https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined"
    ":opsz,wght,FILL,GRAD@24,400,0,0",
]

components.html(
    f"""<script>
(function(){{
  var p = window.parent.document;
  if (p.getElementById('gca-css')) return;
  {_json.dumps(_FONT_URLS)}.forEach(function(href){{
    var l = p.createElement('link');
    l.rel='stylesheet'; l.href=href;
    p.head.appendChild(l);
  }});
  var s = p.createElement('style');
  s.id='gca-css';
  s.textContent={_json.dumps(_GCA_CSS)};
  p.head.appendChild(s);
}})();
</script>""",
    height=0,
    scrolling=False,
)

# ── Top Nav Bar ────────────────────────────────────────────────────────────
st.markdown("""
<header class="gca-nav">
  <div class="gca-brand">
    <span class="gca-brand-icon">science</span>
    <div>
      <span class="gca-brand-name">Gene Coverage Analyzer</span>
      <span class="gca-brand-sub">Clinical Genomics Dashboard</span>
    </div>
  </div>
  <nav class="gca-nav-links">
    <a class="active" href="#">Dashboard</a>
    <a href="#">Reports</a>
    <a href="#">Batch</a>
  </nav>
</header>
""", unsafe_allow_html=True)

# ── Session State ──────────────────────────────────────────────────────────
if 'coverage_data' not in st.session_state:
    st.session_state.coverage_data = None
if 'panel_genes' not in st.session_state:
    st.session_state.panel_genes = []
if 'filtered_data' not in st.session_state:
    st.session_state.filtered_data = None
if 'file_basename' not in st.session_state:
    st.session_state.file_basename = "output"
if 'last_processed_file' not in st.session_state:
    st.session_state.last_processed_file = None
if 'mito_data' not in st.session_state:
    st.session_state.mito_data = None
if 'pipeline_html_bytes' not in st.session_state:
    st.session_state.pipeline_html_bytes = None
if 'pipeline_html_name' not in st.session_state:
    st.session_state.pipeline_html_name = None

# ── Helper functions ───────────────────────────────────────────────────────
@st.cache_data
def extract_gene_name(name):
    if isinstance(name, str):
        if ',' in name:
            return name.split(',')[0]
        elif ';' in name:
            return name.split(';')[0]
        return name
    return None

def parse_gene_list(text):
    """Returns (unique_genes, n_duplicates)."""
    if not text.strip():
        return [], 0
    text = text.replace('\n', ' ').replace(',', ' ').replace('\t', ' ')
    genes = [g.strip() for g in text.split() if g.strip()]
    seen, unique = set(), []
    for g in genes:
        if g not in seen:
            seen.add(g)
            unique.append(g)
    return unique, len(genes) - len(unique)

@st.cache_data(show_spinner=False)
def preprocess_excel_cached(file_bytes, file_name, skip_rows=1):
    try:
        data = pd.read_excel(io.BytesIO(file_bytes), skiprows=skip_rows)
        if 'Gene Name' not in data.columns:
            raise ValueError("Column 'Gene Name' not found")
        data['Gene_Name'] = data['Gene Name'].apply(extract_gene_name)
        data['Gene_ID'] = data['Gene_Name'].str.strip().str.upper()
        valid = (
            data['Gene_ID'].notna() & (data['Gene_ID'] != '') &
            ~data['Gene_ID'].str.startswith('ENSG') &
            ~data['Gene_ID'].str.startswith('INTRON')
        )
        data = data[valid].copy()
        data['Covered_Bases'] = (data['% 1x'] / 100.0) * data['Counted Bases']
        grouped = data.groupby('Gene_ID', as_index=False).agg(
            Total_Bases=('Counted Bases', 'sum'),
            Covered_Bases=('Covered_Bases', 'sum'),
        )
        grouped['% 1x'] = (grouped['Covered_Bases'] / grouped['Total_Bases'] * 100).round(2)
        grouped['Gene_Name'] = grouped['Gene_ID']
        grouped['Ref Name'] = grouped['Gene_ID']
        return grouped, os.path.splitext(file_name)[0]
    except Exception as e:
        raise Exception(f"Error: {str(e)}")

def create_word_document_with_mito(df_data):
    df_data = df_data.reset_index(drop=True)
    doc = Document()
    doc.add_heading('Appendix 1: Gene Coverage', 1)
    doc.add_heading('Indication Based Analysis:', 2)
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(0)

    chunk_size = 4
    chunks = [df_data[i:i+chunk_size] for i in range(0, len(df_data), chunk_size)]
    table = doc.add_table(rows=1, cols=chunk_size * 2)
    table.alignment = 1
    hdr_cells = table.rows[0].cells

    for i in range(chunk_size):
        for cell, text in [
            (hdr_cells[i*2],     "Gene Name"),
            (hdr_cells[i*2+1],   "Percentage of coding region covered"),
        ]:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            para = cell.paragraphs[0]
            para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            para.paragraph_format.space_after = Pt(0)
            run = para.add_run(text)
            run.font.name = 'Calibri'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')
            run.font.size = Pt(8)

    for chunk in chunks:
        chunk = chunk.reset_index(drop=True)
        row_cells = table.add_row().cells
        for i in range(chunk_size):
            if i < len(chunk):
                row = chunk.iloc[i]
                gene = str(row['Gene_ID'])
                pv = row['% 1x']
                percent = float(pv.iloc[0] if isinstance(pv, pd.Series) else pv) if pd.notna(pv) else 0.0

                for cell, content, italic in [
                    (row_cells[i*2],   gene,         True),
                    (row_cells[i*2+1], str(percent), False),
                ]:
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    p = cell.paragraphs[0]
                    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    p.paragraph_format.space_after = Pt(0)
                    r = p.add_run(content)
                    r.font.name = 'Calibri'
                    r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')
                    r.font.size = Pt(8)
                    r.italic = italic
                    if percent < 90:
                        r.font.color.rgb = RGBColor(255, 0, 0)
            else:
                for cell in [row_cells[i*2], row_cells[i*2+1]]:
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    p = cell.paragraphs[0]
                    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    p.paragraph_format.space_after = Pt(0)
                    r = p.add_run("–")
                    r.font.name = 'Calibri'
                    r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')
                    r.font.size = Pt(8)
                    r.font.color.rgb = RGBColor(255, 255, 255)

    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.append(tblPr)
    tblBorders = OxmlElement('w:tblBorders')
    for bn in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        b = OxmlElement(f'w:{bn}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), '000000')
        tblBorders.append(b)
    tblPr.append(tblBorders)
    return doc

def generate_html_report(df, patient_name):
    gene_col = next(
        (c for c in ['Gene_ID', 'Gene_Name', 'Gene', 'gene'] if c in df.columns),
        df.columns[0]
    )
    coverage_col = next(
        (c for c in ['% 1x', '% 1x ', 'Perc_1x', 'Coverage', 'Percentage'] if c in df.columns),
        df.columns[1]
    )

    df = df.copy()
    df[gene_col] = df[gene_col].astype(str).str.strip().str.upper()

    rows_html = []
    for sno, (_, row) in enumerate(df.iterrows(), start=1):
        gene = str(row.get(gene_col, '') or '')
        pct  = row.get(coverage_col, '')
        if pct == 0 or pct == '0' or pct == 0.0:
            pct_str = '0'
        elif pct != pct:
            pct_str = ''
        else:
            pct_str = str(pct)
        rows_html.append(
            f"<tr><td>{sno}</td><td>{html_lib.escape(gene)}</td><td>{pct_str}</td></tr>"
        )

    table_rows       = '\n'.join(rows_html)
    safe_patient_html = html_lib.escape(patient_name)
    safe_patient_js   = patient_name.replace('\\', '\\\\').replace("'", "\\'")

    LOGO_B64 = "iVBORw0KGgoAAAANSUhEUgAAAT0AAABtCAYAAADTYrbeAAAABHNCSVQICAgIfAhkiAAAABl0RVh0U29mdHdhcmUAZ25vbWUtc2NyZWVuc2hvdO8Dvz4AACAASURBVHic7Z13eFRV+sc/507KpPeQRhJIQggJJRA6EnoXsQEWWNsKYvspytpB14ZtRd1FVFyWsrgqIIJKkRqkiEBooQUJJCG992Rm7u+PIYHJTJKZSSHK/TwPT5hz7z3nzJ2Z733Pe97zHiHLsoyCgoLCDYJ0vTugoKCg0JYooqegoHBDoYiegoLCDYUiegoKCjcUiugpKCjcUCiip6CgcEOhiJ6CgsINhc317oDCjYOYvcKoTP50xnXoicKNjGLpKSgo3FAolp6CgkK7o7CwkISEBIqKgnBycmLIkCH4+Pi0SN2K6CkoKLQbysvLeffdd9m0aRMhISH4+fmRm5vL22+/zbBhw3j++edxd3dvVhuK6CkoKLQLSktLefLJJ8nNzWXx4sX07NkTIQQAp0+fZv78+cyePZvFixfj4eFhdTvXTfQ0Wh3pRRVczC/jUmE5VTVaVJJAJQnUNipCPZ3o4uuKm4Pt9eqigoJCG/LZZ59x8eJFli9fzpEjR1i7di1arRYXFxcmT57MP//5T2bOnMmiRYtYsGCB1e20meil5JWy/Vw2O5Kz+ep8NprMYjAnwYuzPT0C3Lk1yp9J0QH06ehZp/4KCgp/DsrLy1m/fj1Tp04lICCARYsWkZWVxYMPPsjKlSuZM2cO27dvZ+rUqSxZsoQnn3zSaquvVUUvr6yK5Qcv8I/9F0hNybWuktIqjp3N4tjZLF5dnwgu9tzbM5jHbwqnX4h3y3ZYQUHhunDy5EnKysqIj4+vK/P09KR3797s2rWLnJwcAOLj4/nkk084ejQyw4YNs6qtFhc9jVbH+uPpLN5/nm3H00Gra9kGSqpYueccK/ecIyTUm7dGd2NqbEdUkhJ9o6DwR6W6uhoArVpdV7Zv3z4mTpxIeXk5H330EZIk4eBgcL61ugAAAGHFJREFU3OdsKwMA"

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1.0">
  <title>Gene Coverage Report — {safe_patient_html}</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400&display=swap" rel="stylesheet">
  <style>
    *{{box-sizing:border-box;margin:0;padding:0}}
    body{{background:#f8f9ff;font-family:'Inter',sans-serif;color:#0b1c30}}
    .topbar{{
      background:#fff;border-bottom:1px solid #e7bdb2;height:60px;
      padding:0 1.5rem;display:flex;align-items:center;justify-content:space-between;
      position:sticky;top:0;z-index:99;box-shadow:0 1px 3px rgba(0,0,0,.06);
    }}
    .topbar-brand{{display:flex;align-items:center;gap:10px}}
    .topbar-logo{{width:180px;height:auto}}
    .topbar-title{{font-size:16px;font-weight:700;color:#ad2c00}}
    .patient-bar{{
      background:#fff;padding:10px 1.5rem;border-bottom:2px solid #ad2c00;
    }}
    .patient-bar h3{{font-size:16px;font-weight:600;color:#0b1c30}}
    .container{{padding:1.5rem;max-width:960px;margin:0 auto}}
    .controls{{
      display:flex;gap:10px;flex-wrap:wrap;justify-content:center;
      align-items:center;margin-bottom:1.25rem;
    }}
    input[type="text"]{{
      padding:7px 12px;font-size:13px;border:1px solid #926f66;
      border-radius:8px;min-width:260px;font-family:'Inter',sans-serif;
      background:#fff;outline:none;transition:border-color .2s,box-shadow .2s;
    }}
    input[type="text"]:focus{{border-color:#ad2c00;box-shadow:0 0 0 3px rgba(173,44,0,.12)}}
    button{{
      padding:7px 16px;background:#ad2c00;color:#fff;border:none;
      border-radius:8px;cursor:pointer;font-size:13px;font-weight:500;
      font-family:'Inter',sans-serif;transition:opacity .2s;
    }}
    button:hover{{opacity:.88}}
    button.active{{background:#5a1500}}
    .table-wrapper{{
      max-height:660px;overflow-y:auto;
      border:1px solid #e7bdb2;border-radius:10px;
    }}
    table{{border-collapse:collapse;width:100%}}
    th,td{{
      border-bottom:1px solid #e7bdb2;padding:8px 12px;
      font-size:13px;text-align:center;
    }}
    td:nth-child(2){{font-family:'JetBrains Mono',monospace;text-align:left}}
    th{{
      background:#ad2c00;color:#fff;position:sticky;top:0;
      cursor:pointer;font-weight:600;font-size:12px;letter-spacing:.03em;
    }}
    tr:last-child td{{border-bottom:none}}
    tr:hover td{{background:#f8f9ff}}
    .low-row td{{color:#93000a;background:rgba(255,218,214,.35)}}
    .low-row:hover td{{background:rgba(255,218,214,.55)}}
    @media(max-width:600px){{
      .controls{{flex-direction:column}}
      input[type="text"]{{min-width:auto;width:100%}}
      button{{width:100%}}
    }}
  </style>
</head>
<body>
  <div class="topbar">
    <div class="topbar-brand">
      <img class="topbar-logo" src="data:image/png;base64,{LOGO_B64}" alt="Logo">
      <span class="topbar-title">Gene Coverage Metrics</span>
    </div>
  </div>
  <div class="patient-bar"><h3>{safe_patient_html}</h3></div>
  <div class="container">
    <div class="controls">
      <input id="searchInput" type="text" placeholder="Search gene…">
      <button onclick="sortTable(true)">↑ Sort Ascending</button>
      <button onclick="sortTable(false)">↓ Sort Descending</button>
      <button id="toggleLow" onclick="toggleLow()">Show Low Coverage Only (&lt;90%)</button>
      <button id="dlBtn">⬇ Download CSV</button>
    </div>
    <div class="table-wrapper">
      <table id="dataTable">
        <thead>
          <tr>
            <th>S.No</th>
            <th>Gene</th>
            <th>% Coding Region Covered</th>
          </tr>
        </thead>
        <tbody>{table_rows}</tbody>
      </table>
    </div>
  </div>
  <script>
    const searchInput = document.getElementById('searchInput');
    const tbody = document.querySelector('tbody');
    let rows = Array.from(tbody.querySelectorAll('tr'));
    let showLowOnly = false;

    // Highlight low-coverage rows
    rows.forEach(r => {{
      const pct = parseFloat(r.cells[2].textContent);
      if (!isNaN(pct) && pct < 90) r.classList.add('low-row');
    }});

    function applyFilters() {{
      const q = searchInput.value.toLowerCase();
      let n = 0;
      rows.forEach(r => {{
        const cells = Array.from(r.getElementsByTagName('td'));
        const matchQ = !q || cells.some(c => c.textContent.toLowerCase().includes(q));
        const pct = parseFloat(r.cells[2].textContent);
        const isLow = !isNaN(pct) && pct < 90;
        const matchT = !showLowOnly || isLow;
        const vis = matchQ && matchT;
        r.style.display = vis ? '' : 'none';
        if (vis) {{ n++; r.cells[0].textContent = n; }}
      }});
    }}

    searchInput.addEventListener('keyup', applyFilters);

    function toggleLow() {{
      showLowOnly = !showLowOnly;
      const btn = document.getElementById('toggleLow');
      btn.classList.toggle('active', showLowOnly);
      btn.textContent = showLowOnly ? 'Show All Genes' : 'Show Low Coverage Only (<90%)';
      applyFilters();
    }}

    function sortTable(asc) {{
      rows.slice().sort((a, b) => {{
        const av = parseFloat(a.cells[2].textContent) || 0;
        const bv = parseFloat(b.cells[2].textContent) || 0;
        return asc ? av - bv : bv - av;
      }}).forEach(r => tbody.appendChild(r));
      rows = Array.from(tbody.querySelectorAll('tr'));
      applyFilters();
    }}

    document.getElementById('dlBtn').addEventListener('click', () => {{
      let csv = 'S.No,Gene,Percentage\\n';
      rows.filter(r => r.style.display !== 'none').forEach(r => {{
        csv += Array.from(r.cells).map(c => '"' + c.textContent.trim() + '"').join(',') + '\\n';
      }});
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([csv], {{type:'text/csv;charset=utf-8;'}}));
      a.download = 'gene_coverage_{safe_patient_js}.csv';
      a.click();
    }});
  </script>
</body>
</html>"""

# ──────────────────────────────────────────────────────────────────────────
# Main UI
# ──────────────────────────────────────────────────────────────────────────

# Step 1 — centred pill toggle
st.markdown('<div class="gca-step1"><span class="gca-lbl">Step 1 — Select Input Format</span></div>',
            unsafe_allow_html=True)
_, c_mid, _ = st.columns([1, 2, 1])
with c_mid:
    input_type = st.radio(
        "", ["📊 Pre-processed CSV", "📁 Raw Excel File"],
        horizontal=True, label_visibility="collapsed"
    )

st.markdown("<div style='margin-bottom:.75rem'></div>", unsafe_allow_html=True)

# Step 2 — data upload  |  panel genes
col1, col2 = st.columns(2, gap="large")

with col1:
    st.markdown(
        '<div class="gca-card-title">Coverage Data '
        '<span class="gca-badge-ready">Ready to upload</span></div>',
        unsafe_allow_html=True
    )

    if "Raw Excel" in input_type:
        # Clear mito when in Raw Excel mode
        st.session_state.mito_data = None

        raw_excel = st.file_uploader("Choose Excel file", type=['xlsx', 'xls'], key='raw_excel')

        if raw_excel:
            file_id = f"{raw_excel.name}_{raw_excel.size}"
            if st.session_state.last_processed_file != file_id:
                progress = st.progress(0, "Reading file…")
                try:
                    progress.progress(30, "Processing…")
                    file_bytes = raw_excel.read()
                    coverage_df, basename = preprocess_excel_cached(file_bytes, raw_excel.name)
                    progress.progress(80, "Generating CSV…")
                    st.session_state.coverage_data = coverage_df
                    st.session_state.file_basename  = basename
                    st.session_state.last_processed_file = file_id
                    st.session_state.pipeline_html_bytes = None
                    st.session_state.pipeline_html_name  = None
                    progress.progress(100, "Complete!")
                    progress.empty()
                except Exception as e:
                    progress.empty()
                    st.error(f"❌ {e}")

            if st.session_state.coverage_data is not None:
                st.success(f"✅ {len(st.session_state.coverage_data)} records loaded")
                st.info(
                    "💡 **No re-upload needed.** Load panel genes in Step 2B below — "
                    "results appear automatically. Download the CSV only if you need an offline copy."
                )
                st.download_button(
                    "⬇️ Download CSV (optional)",
                    data=st.session_state.coverage_data.to_csv(index=False),
                    file_name=f"{st.session_state.file_basename}_coverage.csv",
                    mime="text/csv",
                    use_container_width=True
                )
    else:
        coverage_file = st.file_uploader("Choose Coverage CSV", type=['csv'], key='coverage')
        if coverage_file:
            st.session_state.file_basename = os.path.splitext(coverage_file.name)[0]
            df = pd.read_csv(coverage_file)
            df.columns = df.columns.str.strip()
            coverage_col = next((c for c in ['% 1x', '%1x'] if c in df.columns), None)
            if coverage_col:
                st.session_state.coverage_data = df
                st.session_state.pipeline_html_bytes = None
                st.session_state.pipeline_html_name  = None
                st.success(f"✅ {len(df)} records loaded")
            else:
                st.error("❌ No coverage column found ('% 1x' or '%1x')")

        with st.expander("➕ Add Mitochondrial Data (optional)"):
            mito_file = st.file_uploader(
                "Mitochondrial Excel (header in row 2)",
                type=['xlsx', 'xls'], key='mito_file'
            )
            if mito_file:
                try:
                    mito_df = pd.read_excel(mito_file, header=1)
                    mito_df.columns = mito_df.columns.str.strip()
                    mito_cols = [c for c in mito_df.columns if '1x' in str(c).lower()]
                    if not mito_cols:
                        st.error("❌ No coverage column in mitochondrial file")
                        st.session_state.mito_data = None
                    elif 'Name' not in mito_df.columns:
                        st.error("❌ No 'Name' column in mitochondrial file")
                        st.session_state.mito_data = None
                    else:
                        mito_df = mito_df.rename(columns={mito_cols[0]: '% 1x'})
                        mito_df['Gene_ID'] = mito_df['Name'].astype(str).str.split('/').str[0].str.strip()
                        mito_df = mito_df[['Gene_ID', '% 1x']].copy()
                        mito_df['% 1x'] = pd.to_numeric(mito_df['% 1x'], errors='coerce')
                        before = len(mito_df)
                        mito_df = mito_df.dropna(subset=['% 1x'])
                        if before > len(mito_df):
                            st.warning(f"⚠️ Removed {before - len(mito_df)} rows with invalid coverage")
                        mito_df['% 1x'] = mito_df['% 1x'].round(2)
                        mito_df = mito_df[mito_df['Gene_ID'].str.strip() != '']
                        st.session_state.mito_data = mito_df
                        st.success(f"✅ {len(mito_df)} mitochondrial genes loaded")
                except Exception as e:
                    st.error(f"❌ {e}")
                    import traceback; st.code(traceback.format_exc())
                    st.session_state.mito_data = None
            else:
                st.session_state.mito_data = None

with col2:
    st.markdown('<div class="gca-card-title">Target Panel Genes</div>', unsafe_allow_html=True)
    tab1, tab2 = st.tabs(["📁 Upload Excel", "📝 Paste Genes"])

    with tab1:
        panel_file = st.file_uploader("Choose Excel", type=['xlsx', 'xls'], key='panel')
        if panel_file:
            panel_df = pd.read_excel(panel_file)
            if 'GENE' in panel_df.columns:
                genes = panel_df['GENE'].dropna().unique().tolist()
                st.session_state.panel_genes = genes
                st.success(f"✅ {len(genes)} genes loaded")
            else:
                st.error("❌ No 'GENE' column found")

    with tab2:
        gene_text = st.text_area(
            "Gene symbols",
            height=150,
            placeholder="BRCA1 BRCA2 TP53\nEGFR, KRAS\nTP53",
            label_visibility="collapsed"
        )
        c_left, c_right = st.columns([2, 1])
        with c_right:
            if st.button("Load Genes", use_container_width=True):
                genes, n_dup = parse_gene_list(gene_text)
                if genes:
                    st.session_state.panel_genes = genes
                    msg = f"✅ {len(genes)} genes"
                    if n_dup:
                        msg += f" ({n_dup} duplicates removed)"
                    st.success(msg)
                else:
                    st.warning("⚠️ No genes found")
        with c_left:
            if st.session_state.panel_genes:
                st.markdown(
                    f'<span style="font-size:12px;color:var(--cs)">'
                    f'● {len(st.session_state.panel_genes)} genes loaded</span>',
                    unsafe_allow_html=True
                )

# ── Results ────────────────────────────────────────────────────────────────
if st.session_state.coverage_data is not None and st.session_state.panel_genes:
    df = st.session_state.coverage_data.copy()
    panel_genes = [g.strip().upper() for g in st.session_state.panel_genes]

    # Ensure required columns exist (CSVs not generated by this app may lack them)
    if 'Gene_Name' not in df.columns:
        df['Gene_Name'] = df.get('Gene_ID', df.iloc[:, 0]).astype(str)
    if 'Ref Name' not in df.columns:
        df['Ref Name'] = df['Gene_Name']

    df['Gene_Name'] = df['Gene_Name'].astype(str).str.strip().str.upper()
    df['Ref Name']  = df['Ref Name'].astype(str).str.strip().str.upper()

    coverage_col = next((c for c in ['% 1x', '%1x'] if c in df.columns), None)
    if not coverage_col:
        st.error("Coverage column not found")
        st.stop()

    df_filtered = df[df['Gene_Name'].isin(panel_genes) | df['Ref Name'].isin(panel_genes)].copy()
    df_filtered['Gene_ID'] = df_filtered['Gene_Name'].fillna(df_filtered['Ref Name'])
    df_filtered['Gene_ID'] = df_filtered['Gene_ID'].str.strip().str.upper()
    df_filtered = df_filtered[~df_filtered['Gene_ID'].fillna('').str.startswith("Intron:")]
    df_filtered = df_filtered[~df_filtered['Gene_ID'].fillna('').str.startswith("ENSG")]
    df_filtered['Perc_1x'] = pd.to_numeric(df_filtered[coverage_col], errors='coerce')

    df_grouped = df_filtered.groupby('Gene_ID', as_index=False)['Perc_1x'].mean()
    df_grouped['Perc_1x'] = df_grouped['Perc_1x'].round(2)
    df_grouped = df_grouped.sort_values('Gene_ID')
    st.session_state.filtered_data = df_grouped

    st.divider()

    total = len(df_grouped)
    low   = len(df_grouped[df_grouped['Perc_1x'] < 90])
    avg   = df_grouped['Perc_1x'].mean()

    st.markdown(f"""
<div class="gca-rh">
  <div>
    <div class="gca-rh-ttl">Analysis Results</div>
    <div class="gca-rh-sub">Sample: {html_lib.escape(st.session_state.file_basename)}</div>
  </div>
</div>
<div class="gca-metrics">
  <div class="gca-metric">
    <div class="gca-metric-lbl">Total Genes Analyzed</div>
    <div class="gca-metric-val">{total}</div>
  </div>
  <div class="gca-metric err">
    <div class="gca-metric-lbl">Low Coverage (&lt; 90%)</div>
    <div class="gca-metric-val">{low}</div>
  </div>
  <div class="gca-metric">
    <div class="gca-metric-lbl">Average Coverage</div>
    <div class="gca-metric-val">{avg:.1f}%</div>
  </div>
</div>
""", unsafe_allow_html=True)

    btn_col1, btn_col2 = st.columns(2, gap="medium")

    with btn_col1:
        if st.button("📄 Generate Word Report", type="primary", use_container_width=True):
            with st.spinner("Creating…"):
                try:
                    has_mito = st.session_state.mito_data is not None

                    if has_mito:
                        df_panel = df_grouped.rename(columns={'Perc_1x': '% 1x'}).sort_values('Gene_ID')
                        mito_sorted = st.session_state.mito_data.sort_values('Gene_ID')
                        df_final = pd.concat([df_panel, mito_sorted], ignore_index=True)
                        df_final = df_final.drop_duplicates(subset='Gene_ID', keep='first').reset_index(drop=True)
                        st.info(
                            f"📊 Combined: {len(df_panel)} panel + "
                            f"{len(st.session_state.mito_data)} mito = {len(df_final)} total"
                        )
                    else:
                        df_final = df_grouped.rename(columns={'Perc_1x': '% 1x'})

                    word_filename = f"{st.session_state.file_basename}_report.docx"
                    doc = create_word_document_with_mito(df_final)
                    buf = io.BytesIO()
                    doc.save(buf); buf.seek(0)
                    st.session_state.doc_bytes    = buf
                    st.session_state.word_filename = word_filename
                    st.success("✅ Ready to download!")
                except Exception as e:
                    st.error(f"❌ {e}")
                    import traceback; st.code(traceback.format_exc())

        if 'doc_bytes' in st.session_state and st.session_state.doc_bytes:
            st.download_button(
                f"⬇️ {st.session_state.word_filename}",
                data=st.session_state.doc_bytes,
                file_name=st.session_state.word_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

    with btn_col2:
        csv_filename = f"{st.session_state.file_basename}_filtered.csv"
        with st.container():
            st.markdown('<div class="gca-sec">', unsafe_allow_html=True)
            st.download_button(
                f"📊 Download Filtered CSV",
                data=df_grouped.to_csv(index=False),
                file_name=csv_filename,
                mime="text/csv",
                use_container_width=True
            )
            st.markdown('</div>', unsafe_allow_html=True)

    with st.expander("🔬 View Gene Level Details"):
        st.caption("Red badges = coverage below 90%")
        badges = '<div class="gca-gene-grid">'
        for _, row in df_grouped.iterrows():
            pct  = row['Perc_1x']
            gene = html_lib.escape(str(row['Gene_ID']))
            cls  = 'low' if pct < 90 else 'ok'
            badges += (
                f'<span class="gca-gene {cls}" title="{pct}%">'
                f'{gene}<span class="pct">({pct}%)</span></span>'
            )
        badges += '</div>'
        st.markdown(badges, unsafe_allow_html=True)

    # ── Pipeline shortcut: direct HTML export ─────────────────────────────
    st.markdown("---")
    direct_html_on = st.checkbox(
        "🔗 **Direct HTML export** — generate the interactive HTML report from current results "
        "(skip manual CSV download → re-upload cycle)",
        key="pipeline_direct_html",
        value=False,
    )

    if direct_html_on:
        has_mito_p = st.session_state.mito_data is not None
        col_gen, col_info = st.columns([2, 3])
        with col_info:
            mito_note = " + mitochondrial genes appended at end" if has_mito_p else ""
            st.caption(f"Will export **{len(df_grouped)} panel genes**{mito_note}.")

        with col_gen:
            if st.button("🌐 Build HTML Report", use_container_width=True, type="primary",
                         key="pipeline_build_btn"):
                with st.spinner("Building report…"):
                    if has_mito_p:
                        df_for_html = pd.concat([
                            df_grouped.rename(columns={'Perc_1x': '% 1x'}).sort_values('Gene_ID'),
                            st.session_state.mito_data.sort_values('Gene_ID'),
                        ], ignore_index=True).drop_duplicates(subset='Gene_ID', keep='first')
                    else:
                        df_for_html = df_grouped.rename(columns={'Perc_1x': '% 1x'})

                    html_out  = generate_html_report(df_for_html, st.session_state.file_basename)
                    safe_stem = st.session_state.file_basename.replace(' ', '_').replace('/', '_')
                    st.session_state.pipeline_html_bytes = html_out.encode('utf-8')
                    st.session_state.pipeline_html_name  = safe_stem

        if st.session_state.pipeline_html_bytes:
            st.success("✅ HTML ready — download below.")
            st.download_button(
                f"⬇️ Download coverage_{st.session_state.pipeline_html_name}.html",
                data=st.session_state.pipeline_html_bytes,
                file_name=f"coverage_{st.session_state.pipeline_html_name}.html",
                mime="text/html",
                use_container_width=True,
                key="pipeline_html_dl",
            )

st.divider()

# ── Batch HTML Generator ───────────────────────────────────────────────────
with st.expander("📦 Batch HTML Report Generator"):
    st.markdown(
        "Upload multiple CSV files to get interactive HTML reports — "
        "each with search, sorting, and download features."
    )

    batch_files = st.file_uploader(
        "Upload CSV files",
        type=['csv'],
        accept_multiple_files=True,
        key='batch_csvs'
    )

    if batch_files:
        st.info(f"✅ {len(batch_files)} file(s) loaded")

        if st.button("🎨 Generate HTML Reports", type="primary", use_container_width=True):
            pb = st.progress(0, "Generating…")
            try:
                from datetime import datetime
                n = len(batch_files)

                if n == 1:
                    csv_file    = batch_files[0]
                    patient_name = os.path.splitext(csv_file.name)[0]
                    df           = pd.read_csv(csv_file)
                    html_content = generate_html_report(df, patient_name)
                    pb.progress(1.0, "Done!"); pb.empty()
                    st.success("✅ Report ready!")
                    safe_name = patient_name.replace(' ', '_').replace('/', '_')
                    st.download_button(
                        f"⬇️ Download {safe_name}.html",
                        data=html_content.encode('utf-8'),
                        file_name=f"coverage_{safe_name}.html",
                        mime="text/html",
                        use_container_width=True, type="primary"
                    )
                else:
                    import zipfile
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                        for idx, csv_file in enumerate(batch_files):
                            patient_name = os.path.splitext(csv_file.name)[0]
                            df           = pd.read_csv(csv_file)
                            html_content = generate_html_report(df, patient_name)
                            safe_name    = patient_name.replace(' ', '_').replace('/', '_')
                            zf.writestr(f"coverage_{safe_name}.html", html_content)
                            pb.progress((idx + 1) / n, f"Creating {idx+1}/{n}…")
                    zip_buf.seek(0); pb.empty()
                    st.success(f"✅ {n} reports generated!")
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    st.download_button(
                        "⬇️ Download All Reports (ZIP)",
                        data=zip_buf.getvalue(),
                        file_name=f"gene_coverage_reports_{ts}.zip",
                        mime="application/zip",
                        use_container_width=True, type="primary"
                    )
            except Exception as e:
                pb.empty()
                st.error(f"❌ {e}")
                import traceback; st.code(traceback.format_exc())

# ── Footer ─────────────────────────────────────────────────────────────────
st.markdown("""
<div class="gca-footer">
  <span class="ms">lock</span>
  All processing happens on the server. Data is not stored permanently.
</div>
""", unsafe_allow_html=True)
