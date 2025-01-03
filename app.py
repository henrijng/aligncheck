import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import tldextract
from fuzzywuzzy import fuzz
import re

# Page configuration
st.set_page_config(
    page_title="HubSpot Leads Checker",
    layout="wide",
    initial_sidebar_state="expanded"
)

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

def are_companies_similar(comp1, comp2, threshold=85):
    """Check if two company names are similar using fuzzy matching"""
    if pd.isna(comp1) or pd.isna(comp2):
        return False
    norm1 = normalize_company_name(comp1)
    norm2 = normalize_company_name(comp2)
    if not norm1 or not norm2:
        return False
    return fuzz.ratio(norm1, norm2) >= threshold

def fix_column_names(df):
    """Clean and standardize column names from HubSpot export"""
    if df is None:
        return None
        
    # Remove any BOM characters and clean column names
    df.columns = df.columns.str.replace('\ufeff', '')
    df.columns = df.columns.str.strip()
    return df

def process_csv(uploaded_file):
    """Process uploaded CSV with HubSpot format handling"""
    if uploaded_file is not None:
        try:
            # Try reading with comma delimiter first
            df = pd.read_csv(uploaded_file, encoding='utf-8', dtype=str)
            df = fix_column_names(df)
            # Check if we got only one column (might be semicolon separated)
            if len(df.columns) == 1:
                raise pd.errors.EmptyDataError
            return df
        except (pd.errors.EmptyDataError, UnicodeDecodeError):
            try:
                # Try with semicolon delimiter
                df = pd.read_csv(uploaded_file, sep=';', encoding='utf-8', dtype=str)
                df = fix_column_names(df)
                return df
            except UnicodeDecodeError:
                try:
                    # Final attempt with CP1252 encoding
                    df = pd.read_csv(uploaded_file, sep=';', encoding='cp1252', dtype=str)
                    df = fix_column_names(df)
                    return df
                except:
                    st.error("Error reading file. Please ensure it's a valid CSV.")
                    return None
    return None

def check_leads(deals_df, alignment_df, new_leads_df):
    """Enhanced lead checking for HubSpot export formats"""
    if deals_df is None or alignment_df is None or new_leads_df is None:
        return None, None

    # Create sets for lookups with proper column names
    existing_emails = set()
    email_columns = ['Email', 'Associated Email', 'E-Mail']
    for col in email_columns:
        if col in deals_df.columns:
            existing_emails.update(deals_df[col].dropna().apply(lambda x: x.lower().strip()))

    # Handle company names from both files
    existing_companies = []
    company_columns = ['Associated Company', 'Company', 'Unternehmensname']
    
    for df in [deals_df, alignment_df]:
        for col in company_columns:
            if col in df.columns:
                existing_companies.extend(df[col].dropna().tolist())

    # Handle domains from alignment file
    existing_domains = set()
    if 'Domain-Name des Unternehmens' in alignment_df.columns:
        existing_domains.update(
            alignment_df['Domain-Name des Unternehmens']
            .dropna()
            .apply(extract_domain)
        )

    new_leads = []
    existing_leads = []

    total_leads = len(new_leads_df)
    progress_bar = st.progress(0)
    
    for idx, lead in new_leads_df.iterrows():
        progress = (idx + 1) / total_leads
        progress_bar.progress(progress)
        
        email = ''
        for email_col in ['E-Mail-Adresse', 'Email', 'E-Mail']:
            if email_col in lead and pd.notna(lead[email_col]):
                email = lead[email_col].lower().strip()
                break

        company = ''
        for company_col in ['Firma/Organisation', 'Company', 'Firma']:
            if company_col in lead and pd.notna(lead[company_col]):
                company = lead[company_col]
                break

        match_found = False
        reason = []

        if email in existing_emails:
            match_found = True
            reason.append('Email exists in deals')

        if not match_found and email:
            lead_domain = extract_domain(email)
            if lead_domain in existing_domains:
                match_found = True
                reason.append('Company domain exists')

        if not match_found and company:
            for existing_company in existing_companies:
                if are_companies_similar(company, existing_company):
                    match_found = True
                    reason.append(f'Similar company exists: {existing_company}')
                    break

        lead_dict = lead.to_dict()
        if match_found:
            lead_dict['Reason'] = ' & '.join(reason)
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
    1. Export and upload your HubSpot Deals CSV file (no modifications needed)
    2. Export and upload your Deal Alignment CSV file (no modifications needed)
    3. Upload your New Leads CSV file
    4. Click Process to analyze the data
    5. Download the filtered results
    """)

col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("1. HubSpot Deals")
    deals_file = st.file_uploader("Upload HubSpot Deals export", type=['csv'])
    if deals_file:
        deals_df = process_csv(deals_file)
        if deals_df is not None:
            st.success(f"✓ Loaded {len(deals_df)} deals")
            st.info("Preview of loaded columns: " + ", ".join(deals_df.columns[:3]) + "...")

with col2:
    st.subheader("2. Deal Alignment")
    alignment_file = st.file_uploader("Upload Deal Alignment export", type=['csv'])
    if alignment_file:
        alignment_df = process_csv(alignment_file)
        if alignment_df is not None:
            st.success(f"✓ Loaded {len(alignment_df)} alignments")
            st.info("Preview of loaded columns: " + ", ".join(alignment_df.columns[:3]) + "...")

with col3:
    st.subheader("3. New Leads")
    leads_file = st.file_uploader("Upload new leads CSV", type=['csv'])
    if leads_file:
        leads_df = process_csv(leads_file)
        if leads_df is not None:
            st.success(f"✓ Loaded {len(leads_df)} leads")
            st.info("Preview of loaded columns: " + ", ".join(leads_df.columns[:3]) + "...")

if st.button("🚀 Process Files", disabled=not (deals_file and alignment_file and leads_file)):
    with st.spinner("Processing leads..."):
        new_leads_df, existing_leads_df = check_leads(deals_df, alignment_df, leads_df)
        
        tab1, tab2 = st.tabs(["✨ New Leads", "🔄 Existing Leads"])
        
        with tab1:
            st.subheader(f"New Leads ({len(new_leads_df)})")
            if not new_leads_df.empty:
                st.dataframe(new_leads_df)
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                csv = new_leads_df.to_csv(sep=';', index=False, encoding='utf-8')
                st.download_button(
                    label="📥 Download New Leads CSV",
                    data=csv.encode('utf-8'),
                    file_name=f"new_leads_{timestamp}.csv",
                    mime="text/csv"
                )
        
        with tab2:
            st.subheader(f"Existing Leads ({len(existing_leads_df)})")
            if not existing_leads_df.empty:
                st.dataframe(existing_leads_df)
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                csv = existing_leads_df.to_csv(sep=';', index=False, encoding='utf-8')
                st.download_button(
                    label="📥 Download Existing Leads CSV",
                    data=csv.encode('utf-8'),
                    file_name=f"existing_leads_{timestamp}.csv",
                    mime="text/csv"
                )

st.markdown("---")
st.markdown("Made with ❤️ for HubSpot lead management")
