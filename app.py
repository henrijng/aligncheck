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
    file.seek(0, 2)  # Seek to end of file
    size = file.tell()
    file.seek(pos)  # Seek back to original position
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

def are_companies_similar(comp1, comp2, threshold=85):
    """Check if two company names are similar using fuzzy matching"""
    if pd.isna(comp1) or pd.isna(comp2):
        return False
    norm1 = normalize_company_name(comp1)
    norm2 = normalize_company_name(comp2)
    if not norm1 or not norm2:
        return False
    return fuzz.ratio(norm1, norm2) >= threshold

def validate_lead(lead):
    """Validate if a lead should be included based on business rules"""
    required_fields = {
        'email': ['E-Mail-Adresse', 'Email', 'E-Mail'],
        'company': ['Firma/Organisation', 'Company', 'Firma'],
        'name': ['Vorname', 'First Name', 'First']
    }
    
    # Check for required fields
    for field_type, possible_columns in required_fields.items():
        has_field = False
        for col in possible_columns:
            if col in lead and pd.notna(lead[col]) and str(lead[col]).strip():
                has_field = True
                break
        if not has_field:
            return False
    
    return True

def fix_column_names(df):
    """Clean and standardize column names"""
    if df is None:
        return None
    
    # Handle both series and list of column names
    if isinstance(df.columns, pd.Index):
        columns = df.columns.tolist()
    else:
        columns = list(df.columns)
    
    # Clean each column name
    clean_columns = []
    for col in columns:
        # Remove BOM characters and clean column names
        col = str(col).replace('\ufeff', '')
        # Remove quotes
        col = col.strip('"').strip("'")
        # Remove leading/trailing whitespace
        col = col.strip()
        clean_columns.append(col)
    
    df.columns = clean_columns
    return df

def process_excel(uploaded_file):
    """Process uploaded Excel file with HubSpot format handling"""
    if uploaded_file is not None:
        try:
            # Read Excel file
            df = pd.read_excel(
                uploaded_file,
                engine='openpyxl',
                dtype=str
            )
            
            # Fix column names
            df = fix_column_names(df)
            
            # Remove empty rows and columns
            df = df.dropna(how='all').dropna(axis=1, how='all')
            
            return df
            
        except Exception as e:
            st.error(f"Error reading file: Please ensure it's a valid HubSpot Excel export.")
            with st.expander("Error Details"):
                st.error(str(e))
            return None
    return None

def process_csv(uploaded_file):
    """Process uploaded CSV file"""
    if uploaded_file is not None:
        try:
            df = pd.read_csv(
                uploaded_file,
                encoding='utf-8',
                dtype=str,
                on_bad_lines='skip'
            )
            df = fix_column_names(df)
            return df
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(
                    uploaded_file,
                    encoding='cp1252',
                    dtype=str,
                    on_bad_lines='skip'
                )
                df = fix_column_names(df)
                return df
            except Exception as e:
                st.error(f"Error reading CSV file. Please check the file format.")
                return None
    return None

def clean_output_data(df):
    """Clean and prepare output data"""
    if df.empty:
        return df
        
    # List of columns to keep (add or modify based on your needs)
    columns_to_keep = [
        'Vorname', 'Nachname', 'E-Mail-Adresse', 'Firma/Organisation',
        'First Name', 'Last Name', 'Email', 'Company',
        'Reason'  # Keep reason column for existing leads
    ]
    
    # Keep only columns that exist in the DataFrame
    columns = [col for col in columns_to_keep if col in df.columns]
    
    # Select columns and remove duplicates
    cleaned_df = df[columns].drop_duplicates()
    
    return cleaned_df

def save_to_excel(df, filename):
    """Save DataFrame to Excel with proper formatting"""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
        
        # Auto-adjust columns width
        worksheet = writer.sheets['Sheet1']
        for idx, col in enumerate(df.columns):
            max_length = max(
                df[col].astype(str).apply(len).max(),
                len(str(col))
            ) + 2
            worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)
    
    return buffer.getvalue()

def check_leads(deals_df, alignment_df, new_leads_df):
    """Check leads against existing deals and alignments"""
    if deals_df is None or alignment_df is None or new_leads_df is None:
        return None, None

    # Create sets for lookups with proper column names
    existing_emails = set()
    email_columns = ['Email', 'Associated Email', 'E-Mail', 'E-Mail-Adresse']
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
    domain_columns = ['Domain-Name des Unternehmens', 'Website', 'Domain']
    for col in domain_columns:
        if col in alignment_df.columns:
            existing_domains.update(
                alignment_df[col]
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
        
        # Extract email and company
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

        # Check if contact is already in deals
        if email in existing_emails:
            match_found = True
            reason.append('Email exists in deals')

        # Check company domain
        if not match_found and email:
            lead_domain = extract_domain(email)
            if lead_domain and lead_domain in existing_domains:
                match_found = True
                reason.append('Company domain exists')

        # Check similar company names
        if not match_found and company:
            for existing_company in existing_companies:
                if are_companies_similar(company, existing_company):
                    match_found = True
                    reason.append(f'Similar company exists: {existing_company}')
                    break

        # Check if lead is valid
        if not validate_lead(lead):
            match_found = True
            reason.append('Missing required information')

        # Add lead to appropriate list
        lead_dict = lead.to_dict()
        if match_found:
            lead_dict['Reason'] = ' & '.join(reason)
            existing_leads.append(lead_dict)
        else:
            new_leads.append(lead_dict)

    progress_bar.empty()

    # Create DataFrames and clean them
    new_leads_df = pd.DataFrame(new_leads) if new_leads else pd.DataFrame()
    existing_leads_df = pd.DataFrame(existing_leads) if existing_leads else pd.DataFrame()
    
    return new_leads_df, existing_leads_df

# Streamlit UI
st.title("HubSpot Leads Checker")
st.markdown("---")

with st.expander("📋 Instructions", expanded=False):
    st.markdown("""
    1. Export and upload your HubSpot Deals Excel file (alle deals.xlsx)
    2. Export and upload your Deal Alignment Excel file (deal alignment check.xlsx)
    3. Upload your New Leads file (Excel or CSV)
    4. Click Process to analyze the data
    5. Download the filtered results
    """)

col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("1. HubSpot Deals")
    deals_file = st.file_uploader("Upload HubSpot Deals export (alle deals.xlsx)", type=['xlsx'])
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
    alignment_file = st.file_uploader("Upload Deal Alignment export (deal alignment check.xlsx)", type=['xlsx'])
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
    leads_file = st.file_uploader("Upload new leads file", type=['xlsx', 'csv'])
    if leads_file:
        if is_file_empty(leads_file):
            st.error("The uploaded file is empty")
        else:
            if leads_file.name.endswith('.xlsx'):
                leads_df = process_excel(leads_file)
            else:
                leads_df = process_csv(leads_file)
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
                # Clean and prepare data for display
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
                # Clean and prepare data for display
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
