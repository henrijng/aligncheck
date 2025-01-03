import streamlit as st
import pandas as pd
import numpy as np
import io
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

# Utility functions
def extract_domain(email):
    """Extract base domain from email address, ignoring TLD"""
    if pd.isna(email) or not isinstance(email, str):
        return ''
    try:
        # Remove everything before @ if it's an email
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
    # Convert to lowercase and remove special characters
    normalized = re.sub(r'[^\w\s]', '', name.lower())
    # Remove common company suffixes
    suffixes = [' gmbh', ' ag', ' ltd', ' llc', ' inc', ' bv']
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

def process_csv(uploaded_file):
    """Process uploaded CSV file with encoding detection"""
    if uploaded_file is not None:
        try:
            # Try CP1252 first (common for German files)
            df = pd.read_csv(uploaded_file, sep=';', encoding='cp1252')
        except UnicodeDecodeError:
            try:
                # Fallback to UTF-8
                df = pd.read_csv(uploaded_file, sep=';', encoding='utf-8')
            except:
                st.error("Error reading file. Please ensure it's a valid CSV.")
                return None
        return df
    return None

def validate_columns(df, required_columns, file_type):
    """Validate that required columns are present"""
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        st.error(f"Missing required columns in {file_type}: {', '.join(missing_columns)}")
        return False
    return True

def check_leads(deals_df, alignment_df, new_leads_df):
    """Enhanced lead checking with domain and fuzzy company matching"""
    if deals_df is None or alignment_df is None or new_leads_df is None:
        return None, None

    # Create sets and dictionaries for lookups
    existing_emails = set(
        deals_df['Associated Email']
        .dropna()
        .apply(lambda x: x.lower().strip())
    )
    
    existing_domains = set(
        deals_df['Associated Email']
        .dropna()
        .apply(extract_domain)
    )

    # Add domains from alignment file
    existing_domains.update(
        alignment_df['Domain-Name des Unternehmens']
        .dropna()
        .apply(extract_domain)
    )

    # Prepare company names
    existing_companies = list(
        pd.concat([
            deals_df['Associated Company'].dropna(),
            alignment_df['Unternehmensname'].dropna()
        ])
    )

    new_leads = []
    existing_leads = []

    # Process each lead with detailed matching
    total_leads = len(new_leads_df)
    progress_bar = st.progress(0)
    
    for idx, lead in new_leads_df.iterrows():
        progress = (idx + 1) / total_leads
        progress_bar.progress(progress)
        
        email = lead['E-Mail-Adresse'].lower().strip() if pd.notna(lead['E-Mail-Adresse']) else ''
        company = lead['Firma/Organisation'] if pd.notna(lead['Firma/Organisation']) else ''
        
        match_found = False
        reason = []

        # Check exact email match
        if email in existing_emails:
            match_found = True
            reason.append('Email exists in deals')

        # Check domain match
        if not match_found and email:
            lead_domain = extract_domain(email)
            if lead_domain in existing_domains:
                match_found = True
                reason.append('Company domain exists in deals')

        # Check company name match
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
    1. Upload your HubSpot Deals CSV file (alle Deals.csv)
    2. Upload your Deal Alignment CSV file (Deal alignment check.csv)
    3. Upload your New Leads CSV file
    4. Click Process to analyze the data
    5. Download the filtered results
    
    **Note:** The tool handles various domain extensions (.com, .de, .nl) and similar company names.
    """)

# File uploaders in columns
col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("1. HubSpot Deals")
    deals_file = st.file_uploader("Upload alle Deals.csv", type=['csv'])
    if deals_file:
        deals_df = process_csv(deals_file)
        if deals_df is not None and validate_columns(deals_df, ['Associated Email', 'Associated Company'], 'HubSpot Deals'):
            st.success(f"✓ Loaded {len(deals_df)} deals")
            st.info(f"Found {deals_df['Associated Email'].notna().sum()} emails and {deals_df['Associated Company'].notna().sum()} companies")

with col2:
    st.subheader("2. Deal Alignment")
    alignment_file = st.file_uploader("Upload Deal alignment check.csv", type=['csv'])
    if alignment_file:
        alignment_df = process_csv(alignment_file)
        if alignment_df is not None and validate_columns(alignment_df, ['Unternehmensname', 'Domain-Name des Unternehmens'], 'Deal Alignment'):
            st.success(f"✓ Loaded {len(alignment_df)} alignments")
            st.info(f"Found {alignment_df['Domain-Name des Unternehmens'].notna().sum()} domains")

with col3:
    st.subheader("3. New Leads")
    leads_file = st.file_uploader("Upload new leads CSV", type=['csv'])
    if leads_file:
        leads_df = process_csv(leads_file)
        if leads_df is not None and validate_columns(leads_df, ['E-Mail-Adresse', 'Firma/Organisation'], 'New Leads'):
            st.success(f"✓ Loaded {len(leads_df)} leads")
            st.info(f"Found {leads_df['E-Mail-Adresse'].notna().sum()} emails")

# Process button
if st.button("🚀 Process Files", disabled=not (deals_file and alignment_file and leads_file)):
    with st.spinner("Processing leads..."):
        new_leads_df, existing_leads_df = check_leads(deals_df, alignment_df, leads_df)
        
        # Display results in tabs
        tab1, tab2 = st.tabs(["✨ New Leads", "🔄 Existing Leads"])
        
        with tab1:
            st.subheader(f"New Leads ({len(new_leads_df)})")
            if not new_leads_df.empty:
                st.dataframe(new_leads_df)
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                csv = new_leads_df.to_csv(sep=';', index=False, encoding='cp1252')
                st.download_button(
                    label="📥 Download New Leads CSV",
                    data=csv.encode('cp1252'),
                    file_name=f"new_leads_{timestamp}.csv",
                    mime="text/csv"
                )
        
        with tab2:
            st.subheader(f"Existing Leads ({len(existing_leads_df)})")
            if not existing_leads_df.empty:
                st.dataframe(existing_leads_df)
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                csv = existing_leads_df.to_csv(sep=';', index=False, encoding='cp1252')
                st.download_button(
                    label="📥 Download Existing Leads CSV",
                    data=csv.encode('cp1252'),
                    file_name=f"existing_leads_{timestamp}.csv",
                    mime="text/csv"
                )

# Footer
st.markdown("---")
st.markdown("Made with ❤️ for HubSpot lead management")
