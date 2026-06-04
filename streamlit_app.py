import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import RGBColor, Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import io
import os
import re

# Page config MUST be first
st.set_page_config(
    page_title="Gene Coverage Analyzer",
    page_icon="🧬",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Minimal CSS - only essential styles
st.markdown("""
    <style>
    .stButton>button {
        width: 100%;
    }
    </style>
""", unsafe_allow_html=True)

# Initialize session state ONCE
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

@st.cache_data
def extract_gene_name(name):
    """Extract the first gene name from comma or semicolon-separated names."""
    if isinstance(name, str):
        if ',' in name:
            return name.split(',')[0]
        elif ';' in name:
            return name.split(';')[0]
        return name
    return None

def parse_gene_list(text):
    """Parse gene list with support for newlines, commas, and spaces. Remove duplicates."""
    if not text.strip():
        return []
    
    # Replace newlines and commas with spaces
    text = text.replace('\n', ' ').replace(',', ' ').replace('\t', ' ')
    
    # Split by whitespace and filter empty strings
    genes = [gene.strip() for gene in text.split() if gene.strip()]
    
    # Remove duplicates while preserving order
    seen = set()
    unique_genes = []
    for gene in genes:
        if gene not in seen:
            seen.add(gene)
            unique_genes.append(gene)
    
    return unique_genes

@st.cache_data(show_spinner=False)
def preprocess_excel_cached(file_bytes, file_name, skip_rows=1):
    """Cached preprocessing — uses vectorised groupby for speed."""
    try:
        data = pd.read_excel(io.BytesIO(file_bytes), skiprows=skip_rows)

        if 'Gene Name' not in data.columns:
            raise ValueError("Column 'Gene Name' not found")

        data['Gene_Name'] = data['Gene Name'].apply(extract_gene_name)
        data['Gene_ID'] = data['Gene_Name'].str.strip().str.upper()

        valid = (
            data['Gene_ID'].notna() &
            (data['Gene_ID'] != '') &
            ~data['Gene_ID'].str.startswith('ENSG') &
            ~data['Gene_ID'].str.startswith('INTRON')
        )
        data = data[valid].copy()

        data['Covered_Bases'] = (data['% 1x'] / 100.0) * data['Counted Bases']

        grouped = (
            data.groupby('Gene_ID', as_index=False)
            .agg(
                Total_Bases=('Counted Bases', 'sum'),
                Covered_Bases=('Covered_Bases', 'sum'),
            )
        )
        grouped['% 1x '] = (grouped['Covered_Bases'] / grouped['Total_Bases'] * 100).round(2)
        grouped['Gene_Name'] = grouped['Gene_ID']
        grouped['Ref Name'] = grouped['Gene_ID']

        base_name = os.path.splitext(file_name)[0]
        return grouped, base_name
    except Exception as e:
        raise Exception(f"Error: {str(e)}")

def create_word_document(df_grouped, output_filename):
    """Create Word document with gene coverage table (legacy function - uses Perc_1x)"""
    # Rename for compatibility
    df_work = df_grouped.copy()
    if 'Perc_1x' in df_work.columns:
        df_work = df_work.rename(columns={'Perc_1x': '% 1x'})
    
    return create_word_document_with_mito(df_work, output_filename)

def create_word_document_with_mito(df_data, output_filename):
    """Create Word document with gene coverage table - works with mito or regular data"""
    
    # IMPORTANT: Reset the index to avoid iloc issues
    df_data = df_data.reset_index(drop=True)
    
    doc = Document()
    doc.add_heading('Appendix 1: Gene Coverage', 1)
    doc.add_heading('Indication Based Analysis:', 2)
    
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(0)
    
    chunk_size = 4
    chunks = [df_data[i:i+chunk_size] for i in range(0, len(df_data), chunk_size)]
    
    table = doc.add_table(rows=1, cols=chunk_size*2)
    table.alignment = 1
    hdr_cells = table.rows[0].cells
    
    for i in range(chunk_size):
        gene_hdr_cell = hdr_cells[i*2]
        gene_hdr_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        gene_hdr_para = gene_hdr_cell.paragraphs[0]
        gene_hdr_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        gene_hdr_para.paragraph_format.space_after = Pt(0)
        gene_hdr_run = gene_hdr_para.add_run("Gene Name")
        gene_hdr_run.font.name = 'Calibri'
        gene_hdr_run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')
        gene_hdr_run.font.size = Pt(8)
        
        perc_hdr_cell = hdr_cells[i*2+1]
        perc_hdr_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        perc_hdr_para = perc_hdr_cell.paragraphs[0]
        perc_hdr_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        perc_hdr_para.paragraph_format.space_after = Pt(0)
        perc_hdr_run = perc_hdr_para.add_run("Percentage of coding region covered")
        perc_hdr_run.font.name = 'Calibri'
        perc_hdr_run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')
        perc_hdr_run.font.size = Pt(8)
    
    for chunk in chunks:
        # IMPORTANT: Reset index for each chunk too
        chunk = chunk.reset_index(drop=True)
        
        row_cells = table.add_row().cells
        for i in range(chunk_size):
            if i < len(chunk):
                row = chunk.iloc[i]
                gene = str(row['Gene_ID'])
                
                # Get the percentage value - now using the correct index
                percent_value = row['% 1x']
                if isinstance(percent_value, pd.Series):
                    percent = float(percent_value.iloc[0])
                else:
                    percent = float(percent_value) if pd.notna(percent_value) else 0.0
                
                gene_cell = row_cells[i*2]
                gene_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                gene_para = gene_cell.paragraphs[0]
                gene_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                gene_para.paragraph_format.space_after = Pt(0)
                gene_run = gene_para.add_run(gene)
                gene_run.font.name = 'Calibri'
                gene_run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')
                gene_run.font.size = Pt(8)
                gene_run.italic = True
                if percent < 90:
                    gene_run.font.color.rgb = RGBColor(255, 0, 0)
                
                perc_cell = row_cells[i*2+1]
                perc_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                perc_para = perc_cell.paragraphs[0]
                perc_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                perc_para.paragraph_format.space_after = Pt(0)
                perc_run = perc_para.add_run(str(percent))
                perc_run.font.name = 'Calibri'
                perc_run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')
                perc_run.font.size = Pt(8)
                if percent < 90:
                    perc_run.font.color.rgb = RGBColor(255, 0, 0)
            else:
                for empty_cell in [row_cells[i*2], row_cells[i*2+1]]:
                    empty_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    para = empty_cell.paragraphs[0]
                    para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    para.paragraph_format.space_after = Pt(0)
                    run = para.add_run("–")
                    run.font.name = 'Calibri'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')
                    run.font.size = Pt(8)
                    run.font.color.rgb = RGBColor(255, 255, 255)
    
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.append(tblPr)
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        tblBorders.append(border)
    tblPr.append(tblBorders)
    
    return doc

# Main UI
st.title("🧬 Gene Coverage Analyzer")
st.caption("Process raw Excel data or upload pre-processed CSV to generate Word report")

# Step 1
input_type = st.radio(
    "**Step 1: Choose Input Type**",
    ["📊 Pre-processed CSV", "📁 Raw Excel File"],
    horizontal=True
)

# Step 2
col1, col2 = st.columns(2)

with col1:
    if "Raw Excel" in input_type:
        st.subheader("Step 2A: Upload Raw Excel")
        raw_excel = st.file_uploader("Choose Excel file", type=['xlsx', 'xls'], key='raw_excel')
        
        if raw_excel:
            file_id = f"{raw_excel.name}_{raw_excel.size}"
            
            if st.session_state.last_processed_file != file_id:
                progress = st.progress(0, "Reading file...")
                
                try:
                    progress.progress(30, "Processing...")
                    file_bytes = raw_excel.read()
                    coverage_df, basename = preprocess_excel_cached(file_bytes, raw_excel.name)
                    
                    progress.progress(80, "Generating CSV...")
                    st.session_state.coverage_data = coverage_df
                    st.session_state.file_basename = basename
                    st.session_state.last_processed_file = file_id
                    
                    progress.progress(100, "Complete!")
                    progress.empty()
                except Exception as e:
                    progress.empty()
                    st.error(f"❌ {str(e)}")
            
            if st.session_state.coverage_data is not None:
                st.success(f"✅ {len(st.session_state.coverage_data)} records")
                st.download_button(
                    "⬇️ Download CSV",
                    data=st.session_state.coverage_data.to_csv(index=False),
                    file_name=f"{st.session_state.file_basename}_coverage.csv",
                    mime="text/csv",
                    use_container_width=True
                )
    else:
        st.subheader("Step 2A: Upload CSV Files")
        coverage_file = st.file_uploader("Choose Coverage CSV", type=['csv'], key='coverage')
        
        if coverage_file:
            st.session_state.file_basename = os.path.splitext(coverage_file.name)[0]
            df = pd.read_csv(coverage_file)
            df.columns = df.columns.str.strip()
            
            coverage_col = next((col for col in ['% 1x', '%1x', '% 1x '] if col in df.columns), None)
            
            if coverage_col:
                st.session_state.coverage_data = df
                st.success(f"✅ {len(df)} records")
            else:
                st.error("❌ No coverage column found")
        
        # Mitochondrial file upload option
        st.markdown("---")
        st.markdown("**Optional: Mitochondrial Data**")
        mito_file = st.file_uploader(
            "Upload Mitochondrial Excel (Optional)",
            type=['xlsx', 'xls'],
            key='mito_file',
            help="Excel file with mitochondrial gene coverage (header in row 2)"
        )
        
        if mito_file:
            try:
                # Load mito file with header in row 2 (skip first row)
                mito_df = pd.read_excel(mito_file, header=1)
                
                # Strip all column names of leading/trailing spaces
                mito_df.columns = mito_df.columns.str.strip()
                
                # Find %1x column (more flexible matching)
                mito_cols = [c for c in mito_df.columns if '1x' in str(c).lower()]
                
                if not mito_cols:
                    st.error("❌ No coverage column found in mitochondrial file")
                    st.session_state.mito_data = None
                else:
                    mito_percent_col = mito_cols[0]
                    
                    # Rename to standard name
                    mito_df = mito_df.rename(columns={mito_percent_col: '% 1x'})
                    
                    # Extract gene names from Name column (before '/')
                    if 'Name' not in mito_df.columns:
                        st.error("❌ No 'Name' column found in mitochondrial file")
                        st.session_state.mito_data = None
                    else:
                        # Extract gene ID
                        mito_df['Gene_ID'] = mito_df['Name'].astype(str).str.split('/').str[0].str.strip()
                        
                        # Keep only needed columns
                        mito_df = mito_df[['Gene_ID', '% 1x']].copy()
                        
                        # Convert to numeric and round
                        mito_df['% 1x'] = pd.to_numeric(mito_df['% 1x'], errors='coerce')
                        
                        # Remove rows with NaN coverage
                        before_count = len(mito_df)
                        mito_df = mito_df.dropna(subset=['% 1x'])
                        after_count = len(mito_df)
                        
                        if before_count > after_count:
                            st.warning(f"⚠️ Removed {before_count - after_count} rows with invalid coverage values")
                        
                        # Round the values
                        mito_df['% 1x'] = mito_df['% 1x'].round(2)
                        
                        # Remove empty gene names
                        mito_df = mito_df[mito_df['Gene_ID'].str.strip() != '']

                        st.session_state.mito_data = mito_df
                        st.success(f"✅ Loaded {len(mito_df)} mitochondrial genes")
                        
            except Exception as e:
                st.error(f"❌ Error loading mitochondrial file: {str(e)}")
                import traceback
                st.code(traceback.format_exc())
                st.session_state.mito_data = None
        else:
            st.session_state.mito_data = None

with col2:
    st.subheader("Step 2B: Panel Genes")
    tab1, tab2 = st.tabs(["📁 Excel", "📝 Paste"])
    
    with tab1:
        panel_file = st.file_uploader("Choose Excel", type=['xlsx', 'xls'], key='panel')
        if panel_file:
            panel_df = pd.read_excel(panel_file)
            if 'GENE' in panel_df.columns:
                genes = panel_df['GENE'].dropna().unique().tolist()
                st.session_state.panel_genes = genes
                st.success(f"✅ {len(genes)} genes")
            else:
                st.error("❌ No GENE column")
    
    with tab2:
        gene_text = st.text_area(
            "Paste genes (newline, comma, or space separated)",
            height=150,
            placeholder="BRCA1 BRCA2 TP53\nEGFR, KRAS\nTP53"
        )
        
        if st.button("Load Genes", use_container_width=True):
            genes = parse_gene_list(gene_text)
            if genes:
                st.session_state.panel_genes = genes
                duplicates_removed = len(gene_text.split()) - len(genes)
                if duplicates_removed > 0:
                    st.success(f"✅ {len(genes)} unique genes ({duplicates_removed} duplicates removed)")
                else:
                    st.success(f"✅ {len(genes)} genes")
            else:
                st.warning("⚠️ No genes found")

# Process data
if st.session_state.coverage_data is not None and st.session_state.panel_genes:
    df = st.session_state.coverage_data.copy()
    panel_genes = [g.strip().upper() for g in st.session_state.panel_genes]

    df['Gene_Name'] = df['Gene_Name'].str.strip().str.upper()
    df['Ref Name'] = df['Ref Name'].str.strip().str.upper()

    coverage_col = next((col for col in ['% 1x', '%1x', '% 1x '] if col in df.columns), None)
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
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Genes", len(df_grouped))
    with col2:
        st.metric("Low Coverage", len(df_grouped[df_grouped['Perc_1x'] < 90]))
    with col3:
        st.metric("Avg Coverage", f"{df_grouped['Perc_1x'].mean():.2f}%")
    
    st.divider()
    
    col1, col2 = st.columns(2)
    

    with col1:
        if st.button("📄 Generate Word", type="primary", use_container_width=True):
            with st.spinner("Creating..."):
                try:
                    # Check if mito data exists
                    has_mito = 'mito_data' in st.session_state and st.session_state.mito_data is not None
                    
                    # Prepare final dataset
                    if has_mito:
                        # CRITICAL FIX: Rename panel genes column BEFORE concatenating
                        # df_grouped has column 'Perc_1x'
                        # mito_data has column '% 1x'
                        # They must match before concat!
                        
                        df_panel = df_grouped.copy()
                        df_panel = df_panel.rename(columns={'Perc_1x': '% 1x'})
                        df_panel = df_panel.sort_values('Gene_ID')

                        # Now both dataframes have the same column name: '% 1x'
                        df_final = pd.concat([
                            df_panel,
                            st.session_state.mito_data
                        ], ignore_index=True)

                        # Remove duplicates (keep first occurrence - panel genes take precedence)
                        df_final = df_final.drop_duplicates(subset='Gene_ID', keep='first')
                        df_final = df_final.sort_values('Gene_ID', ignore_index=True)
                        
                        st.info(f"📊 Combined: {len(df_panel)} panel genes + {len(st.session_state.mito_data)} mito genes = {len(df_final)} total")
                    else:
                        # No mito data - just use panel data
                        df_final = df_grouped.copy()
                        df_final = df_final.rename(columns={'Perc_1x': '% 1x'})
                    
                    word_filename = f"{st.session_state.file_basename}_report.docx"
                    doc = create_word_document_with_mito(df_final, word_filename)
                    st.session_state.doc_bytes = None
                    doc_bytes = io.BytesIO()
                    doc.save(doc_bytes)
                    doc_bytes.seek(0)
                    st.session_state.doc_bytes = doc_bytes
                    st.session_state.word_filename = word_filename
                    st.success("✅ Ready!")
                except Exception as e:
                    st.error(f"❌ {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())
        
        if 'doc_bytes' in st.session_state:
            st.download_button(
                f"⬇️ {st.session_state.word_filename}",
                data=st.session_state.doc_bytes,
                file_name=st.session_state.word_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary"
            )
    
    with col2:
        csv_filename = f"{st.session_state.file_basename}_filtered.csv"
        st.download_button(
            f"📊 {csv_filename}",
            data=df_grouped.to_csv(index=False),
            file_name=csv_filename,
            mime="text/csv",
            use_container_width=True
        )
    
    with st.expander("👁️ View Genes"):
        # Create styled HTML preview like before
        preview_html = '<div style="background: white; padding: 1rem; border-radius: 8px; border: 1px solid #dee2e6; max-height: 300px; overflow-y: auto;">'
        for _, row in df_grouped.iterrows():
            if row['Perc_1x'] < 90:
                preview_html += f'<span style="display: inline-block; padding: 0.3rem 0.6rem; margin: 0.2rem; background: #f8d7da; color: #721c24; border-radius: 4px; font-family: monospace; font-size: 0.9rem; font-weight: bold;"><i>{row["Gene_ID"]}</i> ({row["Perc_1x"]}%)</span>'
            else:
                preview_html += f'<span style="display: inline-block; padding: 0.3rem 0.6rem; margin: 0.2rem; background: #e9ecef; border-radius: 4px; font-family: monospace; font-size: 0.9rem;"><i>{row["Gene_ID"]}</i> ({row["Perc_1x"]}%)</span>'
        preview_html += '</div>'
        st.markdown(preview_html, unsafe_allow_html=True)
        st.caption("ℹ️ Genes with coverage below 90% are highlighted in red")

st.divider()

# Batch HTML Generator
with st.expander("📦 Batch HTML Report Generator"):
    st.markdown("""
    **Generate interactive HTML reports from multiple CSV files**
    
    Upload multiple CSV files and get a ZIP with beautiful HTML reports for each file.
    Each report includes search, sorting, and download features.
    """)
    
    batch_files = st.file_uploader(
        "Upload Multiple CSV Files",
        type=['csv'],
        accept_multiple_files=True,
        key='batch_csvs',
        help="Select all CSV files to convert to HTML"
    )
    
    if batch_files:
        st.info(f"✅ Loaded {len(batch_files)} files")
        
        if st.button("🎨 Generate HTML Reports", type="primary", use_container_width=True):
            progress_bar = st.progress(0, "Generating reports...")

            try:
                from datetime import datetime
                n = len(batch_files)

                if n == 1:
                    csv_file = batch_files[0]
                    patient_name = os.path.splitext(csv_file.name)[0]
                    df = pd.read_csv(csv_file)
                    html_content = generate_html_report(df, patient_name)
                    progress_bar.progress(1.0, "Done!")
                    progress_bar.empty()
                    st.success("✅ Report ready!")
                    safe_name = patient_name.replace(' ', '_').replace('/', '_')
                    st.download_button(
                        label=f"⬇️ Download {safe_name}.html",
                        data=html_content.encode('utf-8'),
                        file_name=f"coverage_{safe_name}.html",
                        mime="text/html",
                        use_container_width=True,
                        type="primary"
                    )
                else:
                    import zipfile
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                        for idx, csv_file in enumerate(batch_files):
                            patient_name = os.path.splitext(csv_file.name)[0]
                            df = pd.read_csv(csv_file)
                            html_content = generate_html_report(df, patient_name)
                            safe_name = patient_name.replace(' ', '_').replace('/', '_')
                            zip_file.writestr(f"coverage_{safe_name}.html", html_content)
                            progress_bar.progress((idx + 1) / n, f"Creating {idx + 1} of {n}...")
                    zip_buffer.seek(0)
                    progress_bar.empty()
                    st.success(f"✅ Generated {n} reports!")
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    st.download_button(
                        label="⬇️ Download All Reports (ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name=f"gene_coverage_reports_{timestamp}.zip",
                        mime="application/zip",
                        use_container_width=True,
                        type="primary"
                    )

            except Exception as e:
                progress_bar.empty()
                st.error(f"❌ Error: {str(e)}")
                import traceback
                st.code(traceback.format_exc())

def generate_html_report(df, patient_name):
    """Generate interactive HTML report matching the Batch HTML Generator design."""

    gene_col = next(
        (c for c in ['Gene_ID', 'Gene_Name', 'Gene', 'gene'] if c in df.columns),
        df.columns[0]
    )
    coverage_col = next(
        (c for c in ['% 1x', '% 1x ', 'Perc_1x', 'Coverage', 'Percentage'] if c in df.columns),
        df.columns[1]
    )

    # Sort alphabetically by gene name before rendering
    df = df.copy()
    df[gene_col] = df[gene_col].astype(str).str.strip().str.upper()
    df = df.sort_values(gene_col, ignore_index=True)

    rows_html = []
    for sno, (_, row) in enumerate(df.iterrows(), start=1):
        gene = str(row.get(gene_col, '') or '')
        pct = row.get(coverage_col, '')
        if pct == 0 or pct == '0' or pct == 0.0:
            pct_str = '0'
        elif pct != pct:
            pct_str = ''
        else:
            pct_str = str(pct)
        rows_html.append(f"<tr><td>{sno}</td><td>{gene}</td><td>{pct_str}</td></tr>")

    table_rows = '\n'.join(rows_html)

    LOGO_B64 = "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"
    
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Gene Coverage Report - {patient_name}</title>
  <style>
    body {{
      margin: 0;
      font-family: Arial, sans-serif;
    }}
    .header-wrapper {{
      background-color: white;
      position: relative;
      z-index: 2;
      height: 150px;
      box-sizing: border-box;
      display: flex;
      align-items: center;
      padding: 0 30px;
    }}
    .logo img {{
      width: 200px;
      height: auto;
      margin-right: 10px;
    }}
    .logotext {{
      color: black;
      font-weight: bold;
      font-size: 18px;
    }}
    .patient-header {{
      background-color: white;
      padding: 10px 30px;
      border-bottom: 2px solid orangered;
    }}
    .patient-header h3 {{
      margin: 0;
      color: #333;
      font-size: 18px;
    }}
    .container {{
      text-align: center;
      margin-top: 20px;
    }}
    .controls {{
      margin: 20px auto;
      max-width: 800px;
      display: flex;
      gap: 10px;
      justify-content: center;
      align-items: center;
      flex-wrap: wrap;
    }}
    input[type="text"] {{
      padding: 8px 12px;
      font-size: 14px;
      border: 1px solid #ccc;
      border-radius: 4px;
      min-width: 300px;
    }}
    button {{
      padding: 8px 16px;
      background-color: orangered;
      color: white;
      border: none;
      border-radius: 4px;
      cursor: pointer;
      font-size: 14px;
    }}
    button:hover {{
      background-color: #cc5500;
    }}
    button.active {{
      background-color: #8b0000;
    }}
    .table-wrapper {{
      max-height: 700px;
      overflow-y: auto;
      border: 1px solid #ccc;
      margin: 0 auto;
      width: 60%;
    }}
    table {{
      border-collapse: collapse;
      width: 100%;
    }}
    th, td {{
      border: 1px solid #000000;
      padding: 8px;
      font-weight: 100;
      font-size: 14px;
      text-align: center;
    }}
    th {{
      background-color: orangered;
      color: white;
      position: sticky;
      top: 0;
      cursor: pointer;
    }}
    @media only screen and (max-width: 600px) {{
      .table-wrapper {{ width: 95%; }}
      table {{ font-size: 12px; }}
    }}
  </style>
</head>
<body>
  <div class="header-wrapper">
    <div class="logo">
      <img src="data:image/png;base64,{LOGO_B64}" alt="Logo">
    </div>
    <div class="logotext">Gene Coverage Metrics</div>
  </div>

  <div class="patient-header">
    <h3>{patient_name}</h3>
  </div>

  <div class="container">
    <div class="controls">
      <input type="text" id="searchInput" placeholder="Search by Gene..." />
      <button onclick="sortTable(true)">Sort Ascending</button>
      <button onclick="sortTable(false)">Sort Descending</button>
      <button id="toggleLow" onclick="toggleLowCoverage()">Show Low Coverage Only (&lt;90%)</button>
      <button id="downloadBtn">Download Excel</button>
    </div>

    <div class="table-wrapper">
      <table id="dataTable">
        <thead>
          <tr>
            <th>S.No</th>
            <th>Gene</th>
            <th>Percentage of coding region covered</th>
          </tr>
        </thead>
        <tbody>
          {table_rows}
        </tbody>
      </table>
    </div>
  </div>

  <script>
    const searchInput = document.getElementById('searchInput');
    const tbody = document.querySelector('tbody');
    let rows = Array.from(tbody.querySelectorAll('tr'));

    let showLowOnly = false;

    function applyFilters() {{
      const searchText = searchInput.value.toLowerCase();
      rows.forEach(row => {{
        const cells = Array.from(row.getElementsByTagName('td'));
        const matchSearch = !searchText || cells.some(cell => cell.textContent.toLowerCase().includes(searchText));
        const pct = parseFloat(row.cells[2].textContent);
        const isLow = !isNaN(pct) && pct < 90;
        const matchToggle = !showLowOnly || isLow;
        row.style.display = (matchSearch && matchToggle) ? '' : 'none';
      }});
    }}

    searchInput.addEventListener('keyup', applyFilters);

    function toggleLowCoverage() {{
      showLowOnly = !showLowOnly;
      const btn = document.getElementById('toggleLow');
      btn.classList.toggle('active', showLowOnly);
      btn.textContent = showLowOnly ? 'Show All Genes' : 'Show Low Coverage Only (<90%)';
      applyFilters();
    }}

    function sortTable(ascending) {{
      const sortedRows = rows.slice().sort((a, b) => {{
        const aVal = parseFloat(a.cells[2].textContent) || 0;
        const bVal = parseFloat(b.cells[2].textContent) || 0;
        return ascending ? aVal - bVal : bVal - aVal;
      }});
      sortedRows.forEach((row, index) => {{
        row.cells[0].textContent = index + 1;
        tbody.appendChild(row);
      }});
      rows = sortedRows;
    }}

    document.getElementById('downloadBtn').addEventListener('click', function() {{
      const visibleRows = rows.filter(row => row.style.display !== 'none');
      let csvContent = "S.No,Gene,Percentage\\n";
      visibleRows.forEach(row => {{
        const cells = Array.from(row.cells).map(cell => '"' + cell.textContent.trim() + '"');
        csvContent += cells.join(',') + '\\n';
      }});
      const blob = new Blob([csvContent], {{ type: 'text/csv;charset=utf-8;' }});
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = 'gene_coverage_{patient_name}.csv';
      link.click();
    }});
  </script>
</body>
</html>"""
  
st.divider()
st.caption("All processing happens on the server. Data is not stored permanently.")
