import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import tldextract
from fuzzywuzzy import fuzz
import re
from io import BytesIO

st.set_page_config(
    page_title="HubSpot Leads Checker",
    layout="wide",
    initial_sidebar_state="expanded"
)

def is_file_empty(file):
    pos = file.tell()
    file.seek(0, 2)
    size = file.tell()
    file.seek(pos)
    return size == 0

def extract_domain(email):
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
    if pd.isna(name) or not isinstance(name, str):
        return ''
    normalized = re.sub(r'[^\w\s]', '', name.lower())
    suffixes = [' gmbh', ' ag', ' ltd', ' llc', ' inc', ' bv', ' holding']
    for suffix in suffixes:
        normalized = normalized.replace(suffix, '')
    return normalized.strip()

def extract_email_from_text(text):
    if pd.isna(text) or not isinstance(text, str):
        return ''
    email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', text)
    return email_match.group(0) if email_match else ''

def fix_column_names(df):
    if df is None:
        return None
    columns = df.columns if isinstance(df.columns, pd.Index) else list(df.columns)
    clean_columns = [str(col).replace('\ufeff', '').strip('"').strip("'").strip() for col in columns]
    df.columns = clean_columns
    return df

def process_excel(uploaded_file):
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file, engine='openpyxl', dtype=str)
            df = fix_column_names(df)
            df = df.dropna(how='all').dropna(axis=1, how='all')
            return df
        except Exception as e:
            st.error("Error reading file: Please ensure it's a valid HubSpot Excel export.")
            with st.expander("Error Details"):
                st.error(str(e))
            return None
    return None

def clean_output_data(df):
    if df.empty:
        return df
    primary_columns = [
        'Vorname', 'Nachname', 'Email', 'Unternehmen',
        'Associated Contact', 'Associated Company',
        'Domain-Name des Unternehmens',
        'Reason'
    ]
    columns = [col for col in primary_columns if col in df.columns]
    if not columns:
        columns = df.columns.tolist()
    return df[columns].drop_duplicates()

def save_to_excel(df, filename):
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

def check_leads(deals_df, alignment_df, new_leads_df):
    if deals_df is None or alignment_df is None or new_leads_df is None:
        return None, None

    # Store domains per company
    company_domains = {}
    if 'Associated Company' in deals_df.columns and 'Associated Contact' in deals_df.columns:
        for _, row in deals_df.iterrows():
            if pd.notna(row['Associated Company']) and pd.notna(row['Associated Contact']):
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
            if pd.notna(row['Unternehmensname']) and pd.notna(row['Domain-Name des Unternehmens']):
                company = normalize_company_name(row['Unternehmensname'])
                domain = str(row['Domain-Name des Unternehmens']).lower().strip()
                if company and domain:
                    if company not in company_domains:
                        company_domains[company] = set()
                    company_domains[company].add(domain)

    new_leads = []
    existing_leads = []

    total_leads = len(new_leads_df)
    progress_bar = st.progress(0)
    
    for idx, lead in new_leads_df.iterrows():
        progress = (idx + 1) / total_leads
        progress_bar.progress(progress)
        
        # Required fields check
        if not all(field in lead and pd.notna(lead[field]) for field in ['Vorname', 'Nachname', 'Email', 'Unternehmen']):
            continue

        lead_email = lead['Email'].lower().strip()
        lead_domain = extract_domain(lead_email)
        lead_company = normalize_company_name(lead['Unternehmen'])

        match_found = False
        reasons = []

        # Check company and domain matches
        for existing_company, domains in company_domains.items():
            if lead_company and lead_company == existing_company:
                match_found = True
                reasons.append(f'Company already exists: {existing_company}')
            elif lead_domain and (
                lead_domain in domains or 
                any(d.split('.')[0] == lead_domain.split('.')[0] for d in domains)  # Check base domain
            ):
                match_found = True
                reasons.append(f'Email domain matches company: {existing_company}')

        lead_dict = lead.to_dict()
        if match_found:
            lead_dict['Reason'] = ' & '.join(reasons)
            existing_leads.append(lead_dict)
        else:
            new_leads.append(lead_dict)

    progress_bar.empty()

    return (
        pd.DataFrame(new_leads) if new_leads else pd.DataFrame(),
        pd.DataFrame(existing_leads) if existing_leads else pd.DataFrame()
    )

# Streamlit UI
st.title("HubSpot Leads Checker")
st.markdown("---")

with st.expander("📋 Instructions", expanded=False):
    st.markdown("""
    Required files from HubSpot:
    1. Export "alle deals" list as XLSX
       - Contains: Associated Contact and Company information
    2. Export "Deal Alignment" list as XLSX
       - Contains: Domain-Name des Unternehmens data
    3. Your new leads file (XLSX) must contain:
       - Vorname
       - Nachname
       - Email
       - Unternehmen
    
    Steps:
    1. Upload your HubSpot Deals XLSX
    2. Upload Deal Alignment XLSX
    3. Upload New Leads XLSX
    4. Click Process to analyze
    5. Download filtered results
    """)

col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("1. HubSpot Deals")
    deals_file = st.file_uploader("Upload HubSpot Deals export (XLSX)", type=['xlsx'])
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
    alignment_file = st.file_uploader("Upload Deal Alignment export (XLSX)", type=['xlsx'])
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
    leads_file = st.file_uploader("Upload new leads file (XLSX)", type=['xlsx'])
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
        new_leads_df, existing_leads_df = check_leads(deals_df, alignment_df, leads_df)
        
        tab1, tab2 = st.tabs(["✨ New Leads", "🔄 Existing Leads"])
        
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

st.markdown("---")
st.markdown("Made with ❤️ for HubSpot lead management")
