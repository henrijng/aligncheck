import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime

st.set_page_config(page_title="HubSpot Leads Checker", layout="wide")

def normalize_string(s):
    """Normalize strings for comparison"""
    if pd.isna(s):
        return ''
    return str(s).lower().strip()

def process_csv(uploaded_file):
    """Process uploaded CSV file with proper encoding"""
    if uploaded_file is not None:
        try:
            # Try reading with CP1252 encoding
            df = pd.read_csv(uploaded_file, sep=';', encoding='cp1252')
            return df
        except UnicodeDecodeError:
            # Fallback to UTF-8 if CP1252 fails
            df = pd.read_csv(uploaded_file, sep=';', encoding='utf-8')
            return df
    return None

def check_leads(deals_df, alignment_df, new_leads_df):
    """Check new leads against existing deals and alignments"""
    if deals_df is None or alignment_df is None or new_leads_df is None:
        return None, None

    # Create sets of existing emails and companies
    existing_emails = set(
        deals_df['Associated Email']
        .apply(normalize_string)
        .dropna()
        .unique()
    )
    
    existing_companies = set(
        pd.concat([
            deals_df['Associated Company'].apply(normalize_string),
            alignment_df['Unternehmensname'].apply(normalize_string)
        ])
        .dropna()
        .unique()
    )

    # Initialize lists for results
    new_leads = []
    existing_leads = []

    # Process each lead
    for _, lead in new_leads_df.iterrows():
        email = normalize_string(lead['E-Mail-Adresse'])
        company = normalize_string(lead['Firma/Organisation'])
        
        if email in existing_emails:
            lead_dict = lead.to_dict()
            lead_dict['Reason'] = 'Email exists in deals'
            existing_leads.append(lead_dict)
        elif company in existing_companies:
            lead_dict = lead.to_dict()
            lead_dict['Reason'] = 'Company exists in deals or alignment list'
            existing_leads.append(lead_dict)
        else:
            new_leads.append(lead.to_dict())

    return (
        pd.DataFrame(new_leads) if new_leads else pd.DataFrame(),
        pd.DataFrame(existing_leads) if existing_leads else pd.DataFrame()
    )

def get_download_link(df, filename):
    """Generate a download link for a DataFrame"""
    if df is not None and not df.empty:
        csv = df.to_csv(sep=';', index=False, encoding='cp1252')
        b64 = pd.util.testing._bytes_to_str(csv.encode('cp1252'))
        href = f'<a href="data:file/csv;base64,{b64}" download="{filename}">Download {filename}</a>'
        return href
    return ""

# Streamlit UI
st.title("HubSpot Leads Checker")

with st.expander("Instructions", expanded=True):
    st.markdown("""
    1. Upload your HubSpot Deals CSV file (alle Deals.csv)
    2. Upload your Deal Alignment CSV file (Deal alignment check.csv)
    3. Upload your New Leads CSV file
    4. Click Process to analyze the data
    5. Download the filtered results
    """)

# File uploaders
col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("HubSpot Deals")
    deals_file = st.file_uploader("Upload alle Deals.csv", type=['csv'])
    if deals_file:
        deals_df = process_csv(deals_file)
        if deals_df is not None:
            st.success(f"Loaded {len(deals_df)} deals")

with col2:
    st.subheader("Deal Alignment")
    alignment_file = st.file_uploader("Upload Deal alignment check.csv", type=['csv'])
    if alignment_file:
        alignment_df = process_csv(alignment_file)
        if alignment_df is not None:
            st.success(f"Loaded {len(alignment_df)} alignments")

with col3:
    st.subheader("New Leads")
    leads_file = st.file_uploader("Upload new leads CSV", type=['csv'])
    if leads_file:
        leads_df = process_csv(leads_file)
        if leads_df is not None:
            st.success(f"Loaded {len(leads_df)} leads")

# Process button
if st.button("Process Files", disabled=not (deals_file and alignment_file and leads_file)):
    with st.spinner("Processing..."):
        new_leads_df, existing_leads_df = check_leads(deals_df, alignment_df, leads_df)
        
        # Display results
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader(f"New Leads ({len(new_leads_df)})")
            if not new_leads_df.empty:
                st.dataframe(new_leads_df)
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                csv = new_leads_df.to_csv(sep=';', index=False, encoding='cp1252')
                st.download_button(
                    label="Download New Leads CSV",
                    data=csv.encode('cp1252'),
                    file_name=f"new_leads_{timestamp}.csv",
                    mime="text/csv"
                )
        
        with col2:
            st.subheader(f"Existing Leads ({len(existing_leads_df)})")
            if not existing_leads_df.empty:
                st.dataframe(existing_leads_df)
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                csv = existing_leads_df.to_csv(sep=';', index=False, encoding='cp1252')
                st.download_button(
                    label="Download Existing Leads CSV",
                    data=csv.encode('cp1252'),
                    file_name=f"existing_leads_{timestamp}.csv",
                    mime="text/csv"
                )
