import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import tldextract
from fuzzywuzzy import fuzz
import re
from io import BytesIO

# Page configuration
st.set_page_config(
    page_title="HubSpot Leads Checker",
    layout="wide",
    initial_sidebar_state="expanded"
)

def is_file_empty(file):
    """Check if the uploaded file is empty"""
    pos = file.tell()
    file.seek(0, 2)
    size = file.tell()
    file.seek(pos)
    return size == 0

def extract_domain(email):
    """Extract base domain from email address, ignoring TLD"""
    if pd.isna(email) or not isinstance(email, str):
        return ''
    try:
        if '@' in email:
            email = email.split('@')[1]
        extracted = tldextract.extract(email.lower().strip())
        return extracted.domain
    except:
        return ''

def normalize_company_name(name):
    """Normalize company name for comparison"""
    if pd.isna(name) or not isinstance(name, str):
        return ''
    normalized = re.sub(r'[^\w\s]', '', name.lower())
    suffixes = [' gmbh', ' ag', ' ltd', ' llc', ' inc', ' bv', ' holding']
    for suffix in suffixes:
        normalized = normalized.replace(suffix, '')
    return normalized.strip()

def extract_email_from_text(text):
    """Extract email from text string"""
    if pd.isna(text) or not isinstance(text, str):
        return ''
    email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', text)
    return email_match.group(0) if email_match else ''

def fix_column_names(df):
    """Clean and standardize column names"""
    if df is None:
        return None
    
    columns = df.columns if isinstance(df.columns, pd.Index) else list(df.columns)
    
    clean_columns = []
    for col in columns:
        col = str(col).replace('\ufeff', '')
        col = col.strip('"').strip("'")
        col = col.strip()
        clean_columns.append(col)
    
    df.columns = clean_columns
    return df

def process_excel(uploaded_file):
    """Process uploaded Excel file with HubSpot format handling"""
    if uploaded_file is not None:
        try:
            df = pd.read_excel(
                uploaded_file,
                engine='openpyxl',
                dtype=str
            )
            df = fix_column_names(df)
            df = df.dropna(how='all').dropna(axis=1, how='all')
            return df
            
        except Exception as e:
            st.error(f"Error reading file: Please ensure it's a valid HubSpot Excel export.")
            with st.expander("Error Details"):
                st.error(str(e))
            return None
    return None

def clean_output_data(df):
    """Clean and prepare output data"""
    if df.empty:
        return df
    
    primary_columns = [
        'Vorname', 'Nachname', 'E-Mail-Adresse', 'Firma/Organisation',
        'First Name', 'Last Name', 'Email', 'Company',
        'Associated Contact', 'Associated Company',
        'Domain-Name des Unternehmens',
        'Reason'
    ]
    
    columns = [col for col in primary_columns if col in df.columns]
    if not columns:
        columns = df.columns.tolist()
    
    cleaned_df = df[columns].drop_duplicates()
    return cleaned_df

def save_to_excel(df, filename):
    """Save DataFrame to Excel with proper formatting"""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
        worksheet = writer.sheets['Sheet1']
        for idx, col in enumerate(df.columns):
            max_length = max(
                df[col].astype(str).apply(len).max(),
                len(str(col))
            ) + 2
            worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)
    
    return buffer.getvalue()

# NEUE FUNKTION BEGINNT
def check_leads(deals_df, alignment_df, new_leads_df):
    """Check leads against existing deals and alignments"""
    if deals_df is None or alignment_df is None or new_leads_df is None:
        return None, None, None

    # Store domains per company
    company_domains = {}
    if 'Associated Company' in deals_df.columns and 'Associated Contact' in deals_df.columns:
        for _, row in deals_df.iterrows():
            company = normalize_company_name(row['Associated Company'])
            email = extract_email_from_text(row['Associated Contact'])
            if company and email:
                domain = extract_domain(email)
                if domain:
                    if company not in company_domains:
                        company_domains[company] = set()
                    company_domains[company].add(domain)

    # Add domains from alignment check
    if 'Domain-Name des Unternehmens' in alignment_df.columns and 'Unternehmensname' in alignment_df.columns:
        for _, row in alignment_df.iterrows():
            company = normalize_company_name(row['Unternehmensname'])
            domain = extract_domain(row['Domain-Name des Unternehmens'])
            if company and domain:
                if company not in company_domains:
                    company_domains[company] = set()
                company_domains[company].add(domain)

    new_leads = []
    existing_leads = []
    double_check_leads = []

    total_leads = len(new_leads_df)
    progress_bar = st.progress(0)
    
    for idx, lead in new_leads_df.iterrows():
        progress = (idx + 1) / total_leads
        progress_bar.progress(progress)
        
        # Get lead email and domain
        lead_email = ''
        for email_col in ['E-Mail-Adresse', 'Email', 'E-Mail']:
            if email_col in lead and pd.notna(lead[email_col]):
                lead_email = lead[email_col].lower().strip()
                break
        
        lead_domain = extract_domain(lead_email) if lead_email else ''

        # Get lead company
        lead_company = ''
        for company_col in ['Firma/Organisation', 'Company', 'Firma']:
            if company_col in lead and pd.notna(lead[company_col]):
                lead_company = normalize_company_name(str(lead[company_col]))
                break

        match_found = False
        double_check = False
        reasons = []

        # Check if company exists and match domain
        for existing_company, domains in company_domains.items():
            if lead_company and normalize_company_name(lead_company) == existing_company:
                match_found = True
                reasons.append(f'Company already exists: {existing_company}')
            elif lead_domain and lead_domain in domains:
                match_found = True
                reasons.append(f'Email domain matches company: {existing_company}')

        # Add to appropriate list
        lead_dict = lead.to_dict()
        if reasons:
            lead_dict['Reason'] = ' & '.join(reasons)
            if match_found:
                existing_leads.append(lead_dict)
            else:
                double_check_leads.append(lead_dict)
        else:
            new_leads.append(lead_dict)

    progress_bar.empty()

    return (
        pd.DataFrame(new_leads) if new_leads else pd.DataFrame(),
        pd.DataFrame(existing_leads) if existing_leads else pd.DataFrame(),
        pd.DataFrame(double_check_leads) if double_check_leads else pd.DataFrame()
    )
# NEUE FUNKTION ENDET

# Streamlit UI
st.title("HubSpot Leads Checker")
st.markdown("---")

with st.expander("📋 Instructions", expanded=False):
    st.markdown("""
    1. Export and upload HubSpot Deals Excel file (containing Associated Contact and Company)
    2. Export and upload Deal Alignment Excel file (containing Domain-Name des Unternehmens)
    3. Upload your New Leads file
    4. Click Process to analyze the data
    5. Download the filtered results
    """)

col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("1. HubSpot Deals")
    deals_file = st.file_uploader("Upload HubSpot Deals export", type=['xlsx'])
    if deals_file:
        if is_file_empty(deals_file):
            st.error("The uploaded file is empty")
        else:
            deals_df = process_excel(deals_file)
            if deals_df is not None:
                st.success(f"✓ Loaded {len(deals_df)} deals")
                with st.expander("View loaded columns", expanded=False):
                    st.info("\n".join(deals_df.columns))

with col2:
    st.subheader("2. Deal Alignment")
    alignment_file = st.file_uploader("Upload Deal Alignment export", type=['xlsx'])
    if alignment_file:
        if is_file_empty(alignment_file):
            st.error("The uploaded file is empty")
        else:
            alignment_df = process_excel(alignment_file)
            if alignment_df is not None:
                st.success(f"✓ Loaded {len(alignment_df)} alignments")
                with st.expander("View loaded columns", expanded=False):
                    st.info("\n".join(alignment_df.columns))

with col3:
    st.subheader("3. New Leads")
    leads_file = st.file_uploader("Upload new leads file", type=['xlsx'])
    if leads_file:
        if is_file_empty(leads_file):
            st.error("The uploaded file is empty")
        else:
            leads_df = process_excel(leads_file)
            if leads_df is not None:
                st.success(f"✓ Loaded {len(leads_df)} leads")
                with st.expander("View loaded columns", expanded=False):
                    st.info("\n".join(leads_df.columns))

if st.button("🚀 Process Files", disabled=not (deals_file and alignment_file and leads_file)):
    with st.spinner("Processing leads..."):
        new_leads_df, existing_leads_df, double_check_df = check_leads(deals_df, alignment_df, leads_df)
        
        tab1, tab2, tab3 = st.tabs(["✨ New Leads", "🔄 Existing Leads", "⚠️ Double Check"])
        
        with tab1:
            st.subheader(f"New Leads ({len(new_leads_df)})")
            if not new_leads_df.empty:
                display_df = clean_output_data(new_leads_df)
                st.dataframe(display_df)
                
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                excel_data = save_to_excel(display_df, f"new_leads_{timestamp}.xlsx")
                
                st.download_button(
                    label="📥 Download New Leads Excel",
                    data=excel_data,
                    file_name=f"new_leads_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        with tab2:
            st.subheader(f"Existing Leads ({len(existing_leads_df)})")
            if not existing_leads_df.empty:
                display_df = clean_output_data(existing_leads_df)
                st.dataframe(display_df)
                
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                excel_data = save_to_excel(display_df, f"existing_leads_{timestamp}.xlsx")
                
                st.download_button(
                    label="📥 Download Existing Leads Excel",
                    data=excel_data,
                    file_name=f"existing_leads_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        with tab3:
            st.subheader(f"Double Check Required ({len(double_check_df)})")
            if not double_check_df.empty:
                display_df = clean_output_data(double_check_df)
                st.dataframe(display_df)
                
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                excel_data = save_to_excel(display_df, f"double_check_leads_{timestamp}.xlsx")
                
                st.download_button(
                    label="📥 Download Double Check Leads Excel",
                    data=excel_data,
                    file_name=f"double_check_leads_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

st.markdown("---")
st.markdown("Made with ❤️ for HubSpot lead management")
