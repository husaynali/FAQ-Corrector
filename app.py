import streamlit as st
import pandas as pd
import re
import unicodedata
from datetime import datetime
from io import BytesIO
from thefuzz import process
from typing import Optional, Dict, Tuple
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# -----------------------------
# Page Config
# -----------------------------
st.set_page_config(page_title="FAQ Processing Suite", page_icon="🔧", layout="wide")

# -----------------------------
# Enhanced CSS
# -----------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');

#MainMenu, footer, header {visibility: hidden;}

* {
    font-family: 'Inter', sans-serif;
}

.block-container {
    background: linear-gradient(135deg, #FFF5E6 0%, #FFE4D6 100%);
    border-radius: 16px;
    padding: 2.5rem;
    box-shadow: 0 4px 6px rgba(0,0,0,0.05);
}

h1 {
    color: #3D2C5C;
    font-weight: 700;
    font-size: 2.5rem;
    margin-bottom: 0.5rem;
}

h2, h3 {
    color: #574964;
    font-weight: 600;
}

.subtitle {
    color: #574964;
    font-size: 1.1rem;
    margin-bottom: 2rem;
}

p, span, div, label, li {
    color: #2C2C2C !important;
}

[data-testid="stFileUploader"] {
    background: white;
    border: 2px dashed #9F8383;
    border-radius: 12px;
    padding: 2rem;
    transition: all 0.3s ease;
}

[data-testid="stFileUploader"]:hover {
    border-color: #574964;
    box-shadow: 0 4px 12px rgba(87, 73, 100, 0.1);
}

.stButton>button, .stDownloadButton>button {
    background: linear-gradient(135deg, #27AE60 0%, #229954 100%);
    color: white !important;
    font-weight: 600;
    border: none;
    border-radius: 8px;
    padding: 0.75rem 2rem;
    font-size: 1rem;
    box-shadow: 0 4px 12px rgba(39, 174, 96, 0.3);
    transition: all 0.3s ease;
}

.stButton>button:hover, .stDownloadButton>button:hover {
    background: linear-gradient(135deg, #229954 0%, #1E8449 100%);
    box-shadow: 0 6px 16px rgba(39, 174, 96, 0.4);
    transform: translateY(-2px);
}

.stat-card {
    background: white;
    border-radius: 12px;
    padding: 1.5rem;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    border-left: 4px solid #574964;
    margin: 0.5rem 0;
}

.stat-title {
    color: #7B6B8E !important;
    font-size: 0.9rem !important;
    font-weight: 600 !important;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

.stat-value {
    color: #3D2C5C !important;
    font-size: 1.8rem !important;
    font-weight: 700 !important;
    margin-top: 0.5rem;
}

div[data-testid="stDataFrame"] {
    background: white;
    border-radius: 12px;
    padding: 1rem;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
}

.stAlert {
    border-radius: 10px;
    border-left: 4px solid;
    padding: 1rem;
    margin: 1rem 0;
}

input[type="text"], input[type="number"] {
    background-color: white !important;
    color: #2C2C2C !important;
    border: 2px solid #E0E0E0 !important;
    border-radius: 8px !important;
}

[data-baseweb="select"] > div {
    background-color: white !important;
    color: #2C2C2C !important;
    border: 2px solid #E0E0E0 !important;
    border-radius: 8px !important;
}

.form-region {
    border: 3px dotted #9F8383;
    border-radius: 16px;
    padding: 1.5rem;
    margin: 1rem 0;
    background: rgba(255, 255, 255, 0.5);
    backdrop-filter: blur(10px);
}

.region-title {
    color: #574964 !important;
    font-size: 1.2rem !important;
    font-weight: 700 !important;
    margin-bottom: 1rem;
}

.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
}

.stTabs [data-baseweb="tab"] {
    background-color: white;
    border-radius: 8px;
    padding: 8px 16px;
}

textarea {
    background-color: white !important;
    color: #2C2C2C !important;
    border: 2px solid #E0E0E0 !important;
    border-radius: 8px !important;
}
</style>
""", unsafe_allow_html=True)

# -----------------------------
# Helper Functions for FAQ Corrector
# -----------------------------
def clean_faq_levels(text):
    """Clean and parse FAQ levels from text"""
    if pd.isna(text):
        return [None] * 5
    
    text = str(text).replace('"', '').strip()
    text = text.replace('\n', '|')
    text = re.sub(r'(?<=[a-z])(?=[A-Z])', ' | ', text)
    text = re.sub(r'\s+', ' ', text)
    
    parts = [p.strip() for p in re.split(r'\|', text) if p.strip()]
    parts = parts[:5]
    
    while len(parts) < 5:
        parts.append(None)
    
    return parts

def generate_question(row):
    """Generate question from Level 4 or Level 5"""
    if pd.notna(row.get('Level_5')):
        return row['Level_5']
    if pd.notna(row.get('Level_4')):
        return row['Level_4']
    return None

def soft_clean_text(text):
    """Soft clean text for FAQ key"""
    if pd.isna(text):
        return ''
    
    text = str(text)
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r'[\r\n\t]+', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'\s+([?.!,;:])', r'\1', text)
    
    return text.strip()

def build_faq_key(df):
    """Build FAQ key from all levels"""
    df['FAQ_KEY'] = df[['Level_1', 'Level_2', 'Level_3', 'Level_4', 'Level_5']].apply(
        lambda row: ' '.join([soft_clean_text(x) for x in row if str(x).strip() != '' and str(x) != 'None']), 
        axis=1
    )
    return df

def rename_fail_pass_columns(df):
    """Rename Unnamed columns to fail/pass pairs"""
    new_cols = []
    cols = list(df.columns)
    i = 0

    while i < len(cols):
        col = str(cols[i]).strip()
        if i + 1 < len(cols) and str(cols[i+1]).startswith("Unnamed"):
            base = col.lower().strip()
            new_cols.append(f"{base}_fail")
            new_cols.append(f"{base}_pass")
            i += 2
        else:
            new_cols.append(col)
            i += 1

    df.columns = new_cols
    return df

def process_table(df, faq_column):
    """Process a single table"""
    df[['Level_1', 'Level_2', 'Level_3', 'Level_4', 'Level_5']] = df[faq_column].apply(
        lambda x: pd.Series(clean_faq_levels(x))
    )
    
    df['FAQ Category'] = df['Level_3']
    df['FAQ Description'] = df[['Level_4', 'Level_5']].apply(
        lambda x: ' - '.join([str(i) for i in x if pd.notna(i)]), 
        axis=1
    )
    
    df['Question'] = df.apply(generate_question, axis=1)
    df = build_faq_key(df)
    
    return df

def group_table(df):
    """Group table by FAQ_KEY and sum numeric columns"""
    group_cols = ["FAQ_KEY", "Level_1", "Level_2", "Level_3", "Level_4", "Level_5"]
    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    
    if numeric_cols:
        grouped = (
            df
            .groupby(group_cols, dropna=False)[numeric_cols]
            .sum()
            .reset_index()
        )
        return grouped
    else:
        return df

# -----------------------------
# Helper Functions for FAQ Mapper
# -----------------------------
def soft_clean_text_mapper(text) -> str:
    """Normalize and clean text for mapping"""
    if pd.isna(text):
        return ''
    text = str(text)
    text = unicodedata.normalize("NFKC", text)
    text = text.lower()
    text = re.sub(r'[\r\n\t]+', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'[^\w\s]', '', text)
    return text.strip()

def fuzzy_map_faq(raw_text: str, choices: list, threshold: int) -> Tuple[Optional[str], int]:
    """Fuzzy map raw FAQ to closest clean FAQ"""
    if not raw_text:
        return None, 0
    
    match, score = process.extractOne(raw_text, choices)
    if score >= threshold:
        return match, score
    return None, score

# -----------------------------
# Utility Functions
# -----------------------------
def to_excel_multi_sheet(sheets_dict):
    """Convert multiple dataframes to Excel with multiple sheets"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet_name, df in sheets_dict.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

# -----------------------------
# Session State Initialization
# -----------------------------
if "mode" not in st.session_state:
    st.session_state.mode = "FAQ Corrector"

# -----------------------------
# Header
# -----------------------------
st.markdown("# 🔧 FAQ Processing Suite")
st.markdown('<p class="subtitle">Complete FAQ processing toolkit: Corrector & Mapper</p>', unsafe_allow_html=True)

# -----------------------------
# Mode Selection
# -----------------------------
mode = st.radio(
    "Select Mode:",
    ["📝 FAQ Corrector", "🗺️ FAQ Mapper"],
    horizontal=True,
    label_visibility="collapsed"
)

st.markdown("---")

# =============================
# MODE 1: FAQ CORRECTOR
# =============================
if "FAQ Corrector" in mode:
    st.markdown("## 📝 FAQ Corrector")
    st.markdown("Process standard/template files with automatic FAQ parsing and grouping")
    
    with st.expander("📖 Instructions", expanded=False):
        st.markdown("""
        **Steps:**
        1. Upload **File C** and **File D** (standard/template files)
        2. Select the FAQ column for each file
        3. Click **Process Both Files**
        4. Download the corrected Excel file with 4 sheets
        
        **Output:** c_table, d_table, c_grouped, d_grouped
        """)
    
    # Upload Section
    st.markdown('<div class="form-region">', unsafe_allow_html=True)
    st.markdown('<div class="region-title">📤 Upload Files</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### File C (Standard/Template)")
        file_c = st.file_uploader(
            "Upload File C (.xlsx or .xls)", 
            type=["xlsx", "xls"],
            key="corrector_file_c"
        )
    
    with col2:
        st.markdown("#### File D (Standard/Template)")
        file_d = st.file_uploader(
            "Upload File D (.xlsx or .xls)", 
            type=["xlsx", "xls"],
            key="corrector_file_d"
        )
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    if file_c and file_d:
        try:
            df_c = pd.read_excel(file_c)
            df_d = pd.read_excel(file_d)
            
            st.success(f"✅ Files uploaded successfully!")
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("File C Rows", f"{len(df_c):,}")
            with col2:
                st.metric("File D Rows", f"{len(df_d):,}")
            
            # Column Selection
            st.markdown('<div class="form-region">', unsafe_allow_html=True)
            st.markdown('<div class="region-title">🎯 Select FAQ Columns</div>', unsafe_allow_html=True)
            
            col1, col2 = st.columns(2)
            
            with col1:
                possible_cols_c = [col for col in df_c.columns if 'faq' in col.lower() or 'count' in col.lower()]
                default_c = possible_cols_c[0] if possible_cols_c else df_c.columns[0]
                
                faq_col_c = st.selectbox(
                    "FAQ column for File C",
                    options=df_c.columns.tolist(),
                    index=df_c.columns.tolist().index(default_c) if default_c in df_c.columns else 0,
                    key="faq_col_c"
                )
            
            with col2:
                possible_cols_d = [col for col in df_d.columns if 'faq' in col.lower() or 'count' in col.lower()]
                default_d = possible_cols_d[0] if possible_cols_d else df_d.columns[0]
                
                faq_col_d = st.selectbox(
                    "FAQ column for File D",
                    options=df_d.columns.tolist(),
                    index=df_d.columns.tolist().index(default_d) if default_d in df_d.columns else 0,
                    key="faq_col_d"
                )
            
            st.markdown('</div>', unsafe_allow_html=True)
            
            if st.button("🚀 Process Both Files", use_container_width=True, type="primary", key="process_corrector"):
                with st.spinner("Processing files..."):
                    try:
                        df_c = rename_fail_pass_columns(df_c)
                        df_d = rename_fail_pass_columns(df_d)
                        
                        c_processed = process_table(df_c.copy(), faq_col_c)
                        d_processed = process_table(df_d.copy(), faq_col_d)
                        
                        c_grouped = group_table(c_processed.copy())
                        d_grouped = group_table(d_processed.copy())
                        
                        # Save to session
                        st.session_state.c_processed = c_processed
                        st.session_state.d_processed = d_processed
                        st.session_state.c_grouped = c_grouped
                        st.session_state.d_grouped = d_grouped
                        
                        st.success("✅ Processing complete!")
                        st.balloons()
                        
                    except Exception as e:
                        st.error(f"❌ Error: {str(e)}")
            
            # Display Results
            if "c_processed" in st.session_state and st.session_state.c_processed is not None:
                st.markdown("---")
                st.markdown("### 📊 Results")
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.markdown(f"""
                    <div class="stat-card">
                        <div class="stat-title">File C Processed</div>
                        <div class="stat-value">{len(st.session_state.c_processed):,}</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    st.markdown(f"""
                    <div class="stat-card" style="border-left-color: #3498DB;">
                        <div class="stat-title">File D Processed</div>
                        <div class="stat-value" style="color: #3498DB !important;">{len(st.session_state.d_processed):,}</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col3:
                    st.markdown(f"""
                    <div class="stat-card" style="border-left-color: #27AE60;">
                        <div class="stat-title">File C Grouped</div>
                        <div class="stat-value" style="color: #27AE60 !important;">{len(st.session_state.c_grouped):,}</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col4:
                    st.markdown(f"""
                    <div class="stat-card" style="border-left-color: #E74C3C;">
                        <div class="stat-title">File D Grouped</div>
                        <div class="stat-value" style="color: #E74C3C !important;">{len(st.session_state.d_grouped):,}</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                # Download
                sheets_dict = {
                    "c_table": st.session_state.c_processed,
                    "d_table": st.session_state.d_processed,
                    "c_grouped": st.session_state.c_grouped,
                    "d_grouped": st.session_state.d_grouped
                }
                
                excel_data = to_excel_multi_sheet(sheets_dict)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                
                st.download_button(
                    label=f"📥 Download Complete Excel (4 sheets)",
                    data=excel_data,
                    file_name=f"FAQ_Corrected_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")

# =============================
# MODE 2: FAQ MAPPER
# =============================
else:
    st.markdown("## 🗺️ FAQ Mapper")
    st.markdown("Map evaluation data to clean FAQ dictionary using fuzzy matching + keyword fallback")
    
    with st.expander("📖 Instructions", expanded=False):
        st.markdown("""
        **Steps:**
        1. Upload **Evaluation File** (contains FAQs to be mapped)
        2. Upload **FAQ Dictionary** (clean reference with level_1 through level_6 and full_faq)
        3. Set fuzzy matching threshold (80 recommended)
        4. Optionally add keyword mappings for fallback
        5. Click **Run Mapping**
        6. Download mapped results and unmapped rows for review
        
        **Output:** Evaluation file with mapped hierarchical levels + mapping status
        """)
    
    # Upload Section
    st.markdown('<div class="form-region">', unsafe_allow_html=True)
    st.markdown('<div class="region-title">📤 Upload Files</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### Evaluation File")
        eval_file = st.file_uploader(
            "Upload evaluation Excel", 
            type=["xlsx", "xls"],
            key="mapper_eval_file",
            help="File containing FAQs to be mapped"
        )
    
    with col2:
        st.markdown("#### FAQ Dictionary")
        faq_dict_file = st.file_uploader(
            "Upload FAQ dictionary Excel", 
            type=["xlsx", "xls"],
            key="mapper_dict_file",
            help="Clean FAQ reference with hierarchical levels"
        )
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    if eval_file and faq_dict_file:
        try:
            eval_df = pd.read_excel(eval_file)
            faq_dict = pd.read_excel(faq_dict_file)
            
            # Normalize columns
            faq_dict.columns = [c.lower().strip() for c in faq_dict.columns]
            
            st.success(f"✅ Files uploaded successfully!")
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Evaluation Rows", f"{len(eval_df):,}")
            with col2:
                st.metric("FAQ Dictionary Entries", f"{len(faq_dict):,}")
            
            # Validate FAQ Dictionary
            required_cols = ['level_1', 'level_2', 'level_3', 'level_4', 'level_5', 'level_6', 'full_faq']
            missing = [c for c in required_cols if c not in faq_dict.columns]
            
            if missing:
                st.error(f"❌ FAQ Dictionary missing columns: {', '.join(missing)}")
                st.stop()
            
            # Detect FAQ column in eval
            faq_cols = [c for c in eval_df.columns if 'faq' in c.lower()]
            if not faq_cols:
                st.error("❌ No FAQ column found in evaluation file")
                st.stop()
            
            faq_col = st.selectbox("Select FAQ column from evaluation file:", faq_cols)
            
            # Configuration
            st.markdown('<div class="form-region">', unsafe_allow_html=True)
            st.markdown('<div class="region-title">⚙️ Configuration</div>', unsafe_allow_html=True)
            
            col1, col2 = st.columns(2)
            
            with col1:
                threshold = st.slider(
                    "Fuzzy Matching Threshold",
                    min_value=50,
                    max_value=100,
                    value=80,
                    help="Confidence threshold for fuzzy matching (0-100)"
                )
            
            with col2:
                use_keywords = st.checkbox("Enable Keyword Fallback", value=True)
            
            if use_keywords:
                st.markdown("**Keyword Mappings** (format: keyword → target FAQ)")
                
                default_keywords = {
                    'food preparation': 'Food preparation is too slow',
                    'wrong item': 'Incorrect items were picked up',
                    'didnt receive': 'Order not handed over in person',
                    'unable to contact': 'Unable to contact customer'
                }
                
                keyword_text = st.text_area(
                    "Enter keyword mappings (one per line: keyword → target)",
                    value="\n".join([f"{k} → {v}" for k, v in default_keywords.items()]),
                    height=150
                )
                
                # Parse keyword mappings
                keyword_mappings = {}
                for line in keyword_text.strip().split('\n'):
                    if '→' in line:
                        parts = line.split('→')
                        if len(parts) == 2:
                            keyword_mappings[parts[0].strip()] = parts[1].strip()
            else:
                keyword_mappings = {}
            
            st.markdown('</div>', unsafe_allow_html=True)
            
            # Run Mapping
            if st.button("🚀 Run FAQ Mapping", use_container_width=True, type="primary", key="run_mapper"):
                with st.spinner("Running FAQ mapping pipeline..."):
                    try:
                        # Clean text
                        eval_df['faq_clean'] = eval_df[faq_col].apply(soft_clean_text_mapper)
                        faq_dict['faq_clean'] = faq_dict['full_faq'].apply(soft_clean_text_mapper)
                        
                        # Remove duplicates
                        faq_dict = faq_dict.drop_duplicates(subset=['faq_clean'], keep='first')
                        
                        # Fuzzy matching
                        st.info("🔍 Applying fuzzy matching...")
                        choices = faq_dict['faq_clean'].tolist()
                        
                        mapped = eval_df['faq_clean'].apply(lambda x: fuzzy_map_faq(x, choices, threshold))
                        eval_df['faq_mapped'] = mapped.apply(lambda x: x[0])
                        eval_df['mapping_score'] = mapped.apply(lambda x: x[1])
                        
                        # Merge hierarchical levels
                        eval_df = eval_df.merge(
                            faq_dict[['faq_clean', 'level_1', 'level_2', 'level_3',
                                     'level_4', 'level_5', 'level_6']],
                            left_on='faq_mapped',
                            right_on='faq_clean',
                            how='left',
                            suffixes=('', '_dict')
                        )
                        
                        if 'faq_clean_dict' in eval_df.columns:
                            eval_df = eval_df.drop(columns=['faq_clean_dict'])
                        
                        eval_df['mapping_status'] = eval_df['mapping_score'].apply(
                            lambda x: 'Mapped (Fuzzy)' if x >= threshold else 'Unmapped'
                        )
                        
                        fuzzy_mapped = (eval_df['mapping_status'] == 'Mapped (Fuzzy)').sum()
                        st.success(f"✅ Fuzzy matching mapped {fuzzy_mapped}/{len(eval_df)} rows")
                        
                        # Keyword fallback
                        if use_keywords and keyword_mappings:
                            st.info("🔑 Applying keyword fallback...")
                            
                            unmapped_mask = eval_df['mapping_status'] == 'Unmapped'
                            keyword_mapped = 0
                            
                            for idx in eval_df[unmapped_mask].index:
                                faq_text = eval_df.loc[idx, 'faq_clean']
                                
                                for keyword, target_faq in keyword_mappings.items():
                                    if keyword in faq_text:
                                        target_clean = soft_clean_text_mapper(target_faq)
                                        dict_row = faq_dict[faq_dict['faq_clean'] == target_clean]
                                        
                                        if not dict_row.empty:
                                            eval_df.loc[idx, 'faq_mapped'] = target_clean
                                            eval_df.loc[idx, 'mapping_status'] = 'Mapped (Keyword)'
                                            eval_df.loc[idx, 'mapping_score'] = 100
                                            
                                            for level in ['level_1', 'level_2', 'level_3', 'level_4', 'level_5', 'level_6']:
                                                eval_df.loc[idx, level] = dict_row[level].iloc[0]
                                            
                                            keyword_mapped += 1
                                            break
                            
                            st.success(f"✅ Keyword fallback mapped {keyword_mapped} additional rows")
                        
                        # Save to session
                        st.session_state.mapped_df = eval_df
                        
                        st.success("✅ Mapping pipeline complete!")
                        st.balloons()
                        
                    except Exception as e:
                        st.error(f"❌ Error: {str(e)}")
            
            # Display Results
            if "mapped_df" in st.session_state:
                df_final = st.session_state.mapped_df
                
                st.markdown("---")
                st.markdown("### 📊 Mapping Summary")
                
                fuzzy_count = (df_final['mapping_status'] == 'Mapped (Fuzzy)').sum()
                keyword_count = (df_final['mapping_status'] == 'Mapped (Keyword)').sum()
                unmapped_count = (df_final['mapping_status'] == 'Unmapped').sum()
                avg_score = df_final[df_final['mapping_status'] == 'Mapped (Fuzzy)']['mapping_score'].mean()
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.markdown(f"""
                    <div class="stat-card" style="border-left-color: #27AE60;">
                        <div class="stat-title">Fuzzy Mapped</div>
                        <div class="stat-value" style="color: #27AE60 !important;">{fuzzy_count:,}</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    st.markdown(f"""
                    <div class="stat-card" style="border-left-color: #3498DB;">
                        <div class="stat-title">Keyword Mapped</div>
                        <div class="stat-value" style="color: #3498DB !important;">{keyword_count:,}</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col3:
                    st.markdown(f"""
                    <div class="stat-card" style="border-left-color: #E74C3C;">
                        <div class="stat-title">Unmapped</div>
                        <div class="stat-value" style="color: #E74C3C !important;">{unmapped_count:,}</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col4:
                    st.markdown(f"""
                    <div class="stat-card" style="border-left-color: #F39C12;">
                        <div class="stat-title">Avg Fuzzy Score</div>
                        <div class="stat-value" style="color: #F39C12 !important;">{avg_score:.1f}%</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                # Download buttons
                st.markdown("### 💾 Download Results")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    excel_mapped = to_excel_multi_sheet({"mapped_evaluation": df_final})
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    
                    st.download_button(
                        label=f"📥 Download Mapped Evaluation ({len(df_final):,} rows)",
                        data=excel_mapped,
                        file_name=f"FAQ_Mapped_{timestamp}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                with col2:
                    if unmapped_count > 0:
                        unmapped_df = df_final[df_final['mapping_status'] == 'Unmapped']
                        excel_unmapped = to_excel_multi_sheet({"unmapped_for_review": unmapped_df})
                        
                        st.download_button(
                            label=f"📥 Download Unmapped ({unmapped_count} rows)",
                            data=excel_unmapped,
                            file_name=f"FAQ_Unmapped_{timestamp}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    else:
                        st.success("✅ All rows mapped successfully!")
                
                # Preview
                st.markdown("---")
                st.markdown("### 👀 Data Preview")
                
                tab1, tab2, tab3 = st.tabs(["✅ Fuzzy Mapped", "🔑 Keyword Mapped", "❌ Unmapped"])
                
                with tab1:
                    fuzzy_df = df_final[df_final['mapping_status'] == 'Mapped (Fuzzy)']
                    st.dataframe(fuzzy_df.head(100), use_container_width=True, height=400)
                    st.caption(f"Showing first 100 of {len(fuzzy_df):,} fuzzy mapped rows")
                
                with tab2:
                    keyword_df = df_final[df_final['mapping_status'] == 'Mapped (Keyword)']
                    if len(keyword_df) > 0:
                        st.dataframe(keyword_df.head(100), use_container_width=True, height=400)
                        st.caption(f"Showing first 100 of {len(keyword_df):,} keyword mapped rows")
                    else:
                        st.info("No rows mapped via keywords")
                
                with tab3:
                    if unmapped_count > 0:
                        unmapped_df = df_final[df_final['mapping_status'] == 'Unmapped']
                        st.warning(f"⚠️ {unmapped_count} rows require manual review")
                        st.dataframe(unmapped_df.head(100), use_container_width=True, height=400)
                        st.caption(f"Showing first 100 of {len(unmapped_df):,} unmapped rows")
                    else:
                        st.success("✅ No unmapped rows!")
                
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
