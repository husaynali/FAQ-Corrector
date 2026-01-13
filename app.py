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

def process_dataframe(df, faq_column):
    """Process the uploaded dataframe"""
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

def to_excel(df):
    """Convert dataframe to Excel bytes"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Processed FAQs")
    return output.getvalue()

# -----------------------------
# Session State Initialization
# -----------------------------
if "processed_data" not in st.session_state:
    st.session_state.processed_data = None
if "search_term" not in st.session_state:
    st.session_state.search_term = ""
if "filter_option" not in st.session_state:
    st.session_state.filter_option = "All Records"

# -----------------------------
# Header
# -----------------------------
st.markdown("# 🔧 FAQ Corrector")
st.markdown('<p class="subtitle">Upload, process, and correct FAQ data with automatic level parsing</p>', unsafe_allow_html=True)

# -----------------------------
# Instructions
# -----------------------------
with st.expander("📖 How to use", expanded=False):
    st.markdown("""
    **Steps:**
    1. Upload your Excel file containing FAQ data
    2. Select the column that contains the FAQ text
    3. The app will automatically parse and clean FAQ levels (1-5)
    4. Review the processed data with search and filters
    5. Download the corrected Excel file
    
    **What it does:**
    - Parses FAQ text into 5 hierarchical levels
    - Generates FAQ Category (from Level 3)
    - Creates FAQ Description (from Level 4 & 5)
    - Extracts Question (from Level 4 or 5)
    - Builds FAQ Key for grouping
    """)

# -----------------------------
# Upload Section
# -----------------------------
st.markdown("### 📤 Upload FAQ File")
uploaded_file = st.file_uploader(
    "Upload Excel file (.xlsx or .xls)", 
    type=["xlsx", "xls"],
    help="Upload the file containing FAQ data to be processed"
)

if uploaded_file:
    try:
        # Read the file
        df = pd.read_excel(uploaded_file)
        
        st.success(f"✅ File uploaded successfully! Found {len(df):,} rows and {len(df.columns)} columns.")
        
        # Show column selection
        st.markdown("### 🎯 Select FAQ Column")
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            faq_column = st.selectbox(
                "Which column contains the FAQ text?",
                options=df.columns.tolist(),
                help="Select the column that contains the FAQ hierarchy (e.g., 'Count of FAQ' or 'FAQ')"
            )
        
        with col2:
            st.metric("Sample Value", "Preview")
            if faq_column:
                sample_val = str(df[faq_column].iloc[0])[:50] + "..." if len(str(df[faq_column].iloc[0])) > 50 else str(df[faq_column].iloc[0])
                st.caption(sample_val)
        
        # Process button
        if st.button("🚀 Process FAQ Data", use_container_width=True):
            with st.spinner("Processing FAQ data..."):
                processed_df = process_dataframe(df.copy(), faq_column)
                st.session_state.processed_data = processed_df
                st.success("✅ Processing complete!")
                st.balloons()
        
    except Exception as e:
        st.error(f"❌ Error reading file: {str(e)}")
        st.info("💡 Make sure your file is a valid Excel file (.xlsx or .xls)")

# -----------------------------
# Display Processed Data
# -----------------------------
if st.session_state.processed_data is not None:
    df_processed = st.session_state.processed_data
    
    st.markdown("---")
    
    # Statistics
    st.markdown("### 📊 Statistics")
    
    col1, col2, col3, col4 = st.columns(4)
    
    total_records = len(df_processed)
    complete_records = df_processed['Level_5'].notna().sum()
    incomplete_records = total_records - complete_records
    with_category = df_processed['FAQ Category'].notna().sum()
    
    with col1:
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-title">Total Records</div>
            <div class="stat-value">{total_records:,}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="stat-card" style="border-left-color: #27AE60;">
            <div class="stat-title" style="color: #229954 !important;">Complete (5 Levels)</div>
            <div class="stat-value" style="color: #27AE60 !important;">{complete_records:,}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="stat-card" style="border-left-color: #E74C3C;">
            <div class="stat-title" style="color: #C0392B !important;">Incomplete</div>
            <div class="stat-value" style="color: #E74C3C !important;">{incomplete_records:,}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div class="stat-card" style="border-left-color: #3498DB;">
            <div class="stat-title" style="color: #2980B9 !important;">With Category</div>
            <div class="stat-value" style="color: #3498DB !important;">{with_category:,}</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Search and Filter
    st.markdown("### 🔍 Search & Filter")
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        search = st.text_input(
            "Search in any field",
            value=st.session_state.search_term,
            placeholder="Type to search...",
            key="search_input"
        )
        st.session_state.search_term = search
    
    with col2:
        filter_opt = st.selectbox(
            "Filter by",
            ["All Records", "Complete (5 Levels)", "Incomplete"],
            key="filter_select"
        )
        st.session_state.filter_option = filter_opt
    
    # Apply filters
    filtered_df = df_processed.copy()
    
    if st.session_state.search_term:
        mask = filtered_df.astype(str).apply(
            lambda row: row.str.contains(st.session_state.search_term, case=False, na=False).any(), 
            axis=1
        )
        filtered_df = filtered_df[mask]
    
    if st.session_state.filter_option == "Complete (5 Levels)":
        filtered_df = filtered_df[filtered_df['Level_5'].notna()]
    elif st.session_state.filter_option == "Incomplete":
        filtered_df = filtered_df[filtered_df['Level_5'].isna()]
    
    st.info(f"📋 Showing {len(filtered_df):,} of {len(df_processed):,} records")
    
    # Download button
    st.markdown("### 💾 Download Processed Data")
    
    excel_data = to_excel(df_processed)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    st.download_button(
        label=f"📥 Download Corrected Excel ({len(df_processed):,} records)",
        data=excel_data,
        file_name=f"FAQ_Corrected_{timestamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    
    st.markdown("---")
    
    # Display table
    st.markdown("### 📋 Processed Data Preview")
    
    # Select columns to display
    display_cols = ['Level_1', 'Level_2', 'Level_3', 'Level_4', 'Level_5', 'FAQ Category', 'FAQ Description', 'Question', 'FAQ_KEY']
    display_cols = [col for col in display_cols if col in filtered_df.columns]
    
    st.dataframe(
        filtered_df[display_cols].head(100),
        use_container_width=True,
        height=600
    )
    
    if len(filtered_df) > 100:
        st.caption(f"ℹ️ Showing first 100 rows. Download the Excel file for complete data.")

else:
    # Empty state
    st.info("👆 Upload an Excel file to get started!")
    
    # Show example
    with st.expander("💡 Example FAQ Format"):
        example_data = {
            "Count of FAQ": [
                "Services|Mobile|Data Plans|4G Plans|Unlimited",
                "Support|Billing|Invoice|Download Invoice",
                "Products|Devices|Samsung|Galaxy S23"
            ]
        }
        st.dataframe(pd.DataFrame(example_data), use_container_width=True)
        st.caption("The app will parse these into Level_1, Level_2, Level_3, Level_4, Level_5")
