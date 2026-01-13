import streamlit as st
import pandas as pd
import re
import unicodedata
from datetime import datetime
from io import BytesIO

# -----------------------------
# Page Config
# -----------------------------
st.set_page_config(page_title="FAQ Corrector", page_icon="🔧", layout="wide")

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

input[type="text"] {
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
</style>
""", unsafe_allow_html=True)

# -----------------------------
# Helper Functions
# -----------------------------
def clean_faq_levels(text):
    """Clean and parse FAQ levels from text"""
    if pd.isna(text):
        return [None] * 5
    
    # Remove quotes
    text = str(text).replace('"', '').strip()
    
    # Replace line breaks with separator
    text = text.replace('\n', '|')
    
    # Add separator between camel/mixed text
    text = re.sub(r'(?<=[a-z])(?=[A-Z])', ' | ', text)
    
    # Replace multiple spaces
    text = re.sub(r'\s+', ' ', text)
    
    # Split potential levels
    parts = [p.strip() for p in re.split(r'\|', text) if p.strip()]
    
    # Keep only first 5 levels
    parts = parts[:5]
    
    # Ensure 5 columns
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
    
    # Normalize hidden characters
    text = unicodedata.normalize("NFKC", text)
    
    # Replace line breaks and tabs with space
    text = re.sub(r'[\r\n\t]+', ' ', text)
    
    # Remove multiple spaces
    text = re.sub(r'\s+', ' ', text)
    
    # Fix spaces before punctuation
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

        # If next column is Unnamed, treat as fail/pass pair
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
    """Process a single table (c_table or d_table)"""
    # Apply clean_faq_levels
    df[['Level_1', 'Level_2', 'Level_3', 'Level_4', 'Level_5']] = df[faq_column].apply(
        lambda x: pd.Series(clean_faq_levels(x))
    )
    
    # Set FAQ Category
    df['FAQ Category'] = df['Level_3']
    
    # Set FAQ Description
    df['FAQ Description'] = df[['Level_4', 'Level_5']].apply(
        lambda x: ' - '.join([str(i) for i in x if pd.notna(i)]), 
        axis=1
    )
    
    # Generate Question
    df['Question'] = df.apply(generate_question, axis=1)
    
    # Build FAQ Key
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
if "c_processed" not in st.session_state:
    st.session_state.c_processed = None
if "d_processed" not in st.session_state:
    st.session_state.d_processed = None
if "c_grouped" not in st.session_state:
    st.session_state.c_grouped = None
if "d_grouped" not in st.session_state:
    st.session_state.d_grouped = None

# -----------------------------
# Header
# -----------------------------
st.markdown("# 🔧 FAQ Corrector")
st.markdown('<p class="subtitle">Upload two standard/template files and process FAQ data with automatic grouping</p>', unsafe_allow_html=True)

# -----------------------------
# Instructions
# -----------------------------
with st.expander("📖 How to use", expanded=False):
    st.markdown("""
    **Steps:**
    1. Upload **File C** (e.g., "final c" or "c_table" sheet)
    2. Upload **File D** (e.g., "final d" or "d_table" sheet)
    3. Select the FAQ column for each file
    4. Click **Process Both Files**
    5. Review processed and grouped data
    6. Download the corrected Excel file with all sheets
    
    **What it does:**
    - Parses FAQ text into 5 hierarchical levels
    - Renames fail/pass columns (Unnamed columns)
    - Generates FAQ Category, Description, Question, and FAQ_KEY
    - Groups data by FAQ_KEY and sums numeric columns
    - Exports 4 sheets: c_table, d_table, c_grouped, d_grouped
    """)

# -----------------------------
# Upload Section
# -----------------------------
st.markdown('<div class="form-region">', unsafe_allow_html=True)
st.markdown('<div class="region-title">📤 Upload Files</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    st.markdown("#### File C (Standard/Template)")
    file_c = st.file_uploader(
        "Upload File C (.xlsx or .xls)", 
        type=["xlsx", "xls"],
        key="file_c",
        help="Upload the first file (e.g., 'final c' or 'c_table' sheet)"
    )

with col2:
    st.markdown("#### File D (Standard/Template)")
    file_d = st.file_uploader(
        "Upload File D (.xlsx or .xls)", 
        type=["xlsx", "xls"],
        key="file_d",
        help="Upload the second file (e.g., 'final d' or 'd_table' sheet)"
    )

st.markdown('</div>', unsafe_allow_html=True)

# -----------------------------
# Process Files
# -----------------------------
if file_c and file_d:
    try:
        # Read files
        df_c = pd.read_excel(file_c)
        df_d = pd.read_excel(file_d)
        
        st.success(f"✅ Files uploaded successfully!")
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("File C Rows", f"{len(df_c):,}")
            st.metric("File C Columns", len(df_c.columns))
        with col2:
            st.metric("File D Rows", f"{len(df_d):,}")
            st.metric("File D Columns", len(df_d.columns))
        
        # -----------------------------
        # Column Selection
        # -----------------------------
        st.markdown('<div class="form-region">', unsafe_allow_html=True)
        st.markdown('<div class="region-title">🎯 Select FAQ Columns</div>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### File C - FAQ Column")
            # Try to auto-detect FAQ column
            possible_cols_c = [col for col in df_c.columns if 'faq' in col.lower() or 'count' in col.lower()]
            default_c = possible_cols_c[0] if possible_cols_c else df_c.columns[0]
            
            faq_col_c = st.selectbox(
                "Select FAQ column for File C",
                options=df_c.columns.tolist(),
                index=df_c.columns.tolist().index(default_c) if default_c in df_c.columns else 0,
                key="faq_col_c"
            )
            
            if faq_col_c:
                sample_c = str(df_c[faq_col_c].iloc[0])[:100]
                st.caption(f"Sample: {sample_c}...")
        
        with col2:
            st.markdown("#### File D - FAQ Column")
            # Try to auto-detect FAQ column
            possible_cols_d = [col for col in df_d.columns if 'faq' in col.lower() or 'count' in col.lower()]
            default_d = possible_cols_d[0] if possible_cols_d else df_d.columns[0]
            
            faq_col_d = st.selectbox(
                "Select FAQ column for File D",
                options=df_d.columns.tolist(),
                index=df_d.columns.tolist().index(default_d) if default_d in df_d.columns else 0,
                key="faq_col_d"
            )
            
            if faq_col_d:
                sample_d = str(df_d[faq_col_d].iloc[0])[:100]
                st.caption(f"Sample: {sample_d}...")
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        # -----------------------------
        # Process Button
        # -----------------------------
        if st.button("🚀 Process Both Files", use_container_width=True, type="primary"):
            with st.spinner("Processing files... This may take a moment."):
                try:
                    # Rename fail/pass columns
                    df_c = rename_fail_pass_columns(df_c)
                    df_d = rename_fail_pass_columns(df_d)
                    
                    # Process tables
                    c_processed = process_table(df_c.copy(), faq_col_c)
                    d_processed = process_table(df_d.copy(), faq_col_d)
                    
                    # Group tables
                    c_grouped = group_table(c_processed.copy())
                    d_grouped = group_table(d_processed.copy())
                    
                    # Save to session state
                    st.session_state.c_processed = c_processed
                    st.session_state.d_processed = d_processed
                    st.session_state.c_grouped = c_grouped
                    st.session_state.d_grouped = d_grouped
                    
                    st.success("✅ Processing complete!")
                    st.balloons()
                    
                except Exception as e:
                    st.error(f"❌ Error during processing: {str(e)}")
                    st.info("💡 Make sure the FAQ column contains valid data")
        
    except Exception as e:
        st.error(f"❌ Error reading files: {str(e)}")
        st.info("💡 Make sure your files are valid Excel files (.xlsx or .xls)")

# -----------------------------
# Display Results
# -----------------------------
if st.session_state.c_processed is not None and st.session_state.d_processed is not None:
    st.markdown("---")
    
    # Statistics
    st.markdown("### 📊 Processing Statistics")
    
    col1, col2, col3, col4 = st.columns(4)
    
    c_total = len(st.session_state.c_processed)
    d_total = len(st.session_state.d_processed)
    c_grouped_count = len(st.session_state.c_grouped)
    d_grouped_count = len(st.session_state.d_grouped)
    
    with col1:
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-title">File C Processed</div>
            <div class="stat-value">{c_total:,}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="stat-card" style="border-left-color: #3498DB;">
            <div class="stat-title">File D Processed</div>
            <div class="stat-value" style="color: #3498DB !important;">{d_total:,}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="stat-card" style="border-left-color: #27AE60;">
            <div class="stat-title">File C Grouped</div>
            <div class="stat-value" style="color: #27AE60 !important;">{c_grouped_count:,}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div class="stat-card" style="border-left-color: #E74C3C;">
            <div class="stat-title">File D Grouped</div>
            <div class="stat-value" style="color: #E74C3C !important;">{d_grouped_count:,}</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Download Section
    st.markdown("### 💾 Download Processed Data")
    
    sheets_dict = {
        "c_table": st.session_state.c_processed,
        "d_table": st.session_state.d_processed,
        "c_grouped": st.session_state.c_grouped,
        "d_grouped": st.session_state.d_grouped
    }
    
    excel_data = to_excel_multi_sheet(sheets_dict)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    col1, col2 = st.columns([3, 1])
    with col1:
        st.download_button(
            label=f"📥 Download Complete Excel File (4 sheets)",
            data=excel_data,
            file_name=f"FAQ_Processed_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col2:
        file_size = len(excel_data) / 1024
        st.metric("File Size", f"{file_size:.1f} KB")
    
    st.info("📋 The Excel file contains 4 sheets: c_table, d_table, c_grouped, d_grouped")
    
    st.markdown("---")
    
    # Preview Tabs
    st.markdown("### 👀 Data Preview")
    
    tab1, tab2, tab3, tab4 = st.tabs(["📄 C Table", "📄 D Table", "📊 C Grouped", "📊 D Grouped"])
    
    with tab1:
        st.markdown("#### File C - Processed")
        display_cols_c = ['Level_1', 'Level_2', 'Level_3', 'Level_4', 'Level_5', 'FAQ Category', 'Question', 'FAQ_KEY']
        display_cols_c = [col for col in display_cols_c if col in st.session_state.c_processed.columns]
        st.dataframe(st.session_state.c_processed[display_cols_c].head(100), use_container_width=True, height=400)
        st.caption(f"Showing first 100 of {len(st.session_state.c_processed):,} rows")
    
    with tab2:
        st.markdown("#### File D - Processed")
        display_cols_d = ['Level_1', 'Level_2', 'Level_3', 'Level_4', 'Level_5', 'FAQ Category', 'Question', 'FAQ_KEY']
        display_cols_d = [col for col in display_cols_d if col in st.session_state.d_processed.columns]
        st.dataframe(st.session_state.d_processed[display_cols_d].head(100), use_container_width=True, height=400)
        st.caption(f"Showing first 100 of {len(st.session_state.d_processed):,} rows")
    
    with tab3:
        st.markdown("#### File C - Grouped by FAQ_KEY")
        st.dataframe(st.session_state.c_grouped.head(100), use_container_width=True, height=400)
        st.caption(f"Showing first 100 of {len(st.session_state.c_grouped):,} rows")
    
    with tab4:
        st.markdown("#### File D - Grouped by FAQ_KEY")
        st.dataframe(st.session_state.d_grouped.head(100), use_container_width=True, height=400)
        st.caption(f"Showing first 100 of {len(st.session_state.d_grouped):,} rows")

else:
    # Empty state
    st.info("👆 Upload both files to get started!")
    
    # Show example
    with st.expander("💡 Expected File Format"):
        st.markdown("""
        **Your files should contain:**
        - A column with FAQ hierarchy (e.g., "Count of FAQ" or "FAQ")
        - Optional: Unnamed columns next to named columns (will be renamed to fail/pass pairs)
        - Optional: Numeric columns (will be summed during grouping)
        
        **Example FAQ format:**
        ```
        Services|Mobile|Data Plans|4G Plans|Unlimited
        Support|Billing|Invoice|Download Invoice
        Products|Devices|Samsung|Galaxy S23
        ```
        """)
        
        example_data = {
            "Count of FAQ": [
                "Services|Mobile|Data Plans|4G Plans|Unlimited",
                "Support|Billing|Invoice|Download Invoice"
            ],
            "rc1_service_attitude": ["Pass", "Fail"],
            "Unnamed: 1": [10, 5]
        }
        st.dataframe(pd.DataFrame(example_data), use_container_width=True)
        st.caption("After processing, 'Unnamed: 1' will become 'rc1_service_attitude_pass' and 'rc1_service_attitude_fail'")
