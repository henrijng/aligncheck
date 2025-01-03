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

def get_local_part(email):
    """Get the local-part (before the @) of an email"""
    if pd.isna(email) or not isinstance(email, str) or '@' not in email:
        return ''
    return email.lower().split('@')[0].strip()

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

# ---------------------------------------------------------------------------------
# Erweiterte check_leads-Funktion:
# 1. Mehrere mögliche Firmen-Spalten
# 2. Fuzzy Matching bei Firmenname
# 3. Fuzzy Matching der Email-Domain
# 4. Prüfung auf Duplicate Email Names (local-part)
# ---------------------------------------------------------------------------------
def check_leads(deals_df, alignment_df, new_leads_df):
    """
    Check leads against existing deals and alignments
    - Mehrere Firmen-Spalten
    - Fuzzy Matching für Firma
    - Fuzzy Matching für Domain
    - Prüfung lokaler Teil (E-Mail) auf Duplikate
    """
    if deals_df is None or alignment_df is None or new_leads_df is None:
        return None, None, None

    # Schwellenwerte für Fuzzy Matching
    HIGH_COMPANY_THRESHOLD = 85  # Firmenname Fuzzy
    MID_COMPANY_THRESHOLD  = 70
    HIGH_DOMAIN_THRESHOLD  = 90  # Domain Fuzzy
    MID_DOMAIN_THRESHOLD   = 70

    # ---------------------------
    # 1) Aus deals_df und alignment_df die bekannten Daten einsammeln
    # ---------------------------
    # 1a) E-Mail local-parts + domain + normalisierte Firmen
    existing_local_parts = set()  # für exakte Übereinstimmung
    company_domains = {}          # firma -> set(domains)
    known_companies = set()       # für fuzzy-check

    if 'Associated Company' in deals_df.columns and 'Associated Contact' in deals_df.columns:
        for _, row in deals_df.iterrows():
            assoc_company = row.get('Associated Company', '')
            assoc_contact = row.get('Associated Contact', '')
            
            # Firmen
            normalized_company = normalize_company_name(assoc_company)
            if normalized_company:
                known_companies.add(normalized_company)
            
            # Local Part der E-Mail
            email = extract_email_from_text(assoc_contact)
            if email:
                local_part = get_local_part(email)
                if local_part:
                    existing_local_parts.add(local_part)
                
                # Domain
                domain = extract_domain(email)
                if domain and normalized_company:
                    if normalized_company not in company_domains:
                        company_domains[normalized_company] = set()
                    company_domains[normalized_company].add(domain)

    # 1b) Aus alignment_df Firmen & Domains übernehmen
    # z.B. "Unternehmensname" + "Domain-Name des Unternehmens"
    if ('Domain-Name des Unternehmens' in alignment_df.columns 
        and 'Unternehmensname' in alignment_df.columns):
        for _, row in alignment_df.iterrows():
            alignment_company = row.get('Unternehmensname', '')
            normalized_al_comp = normalize_company_name(alignment_company)
            if normalized_al_comp:
                known_companies.add(normalized_al_comp)
            domain_col = row.get('Domain-Name des Unternehmens', '')
            domain = extract_domain(domain_col)
            if domain and normalized_al_comp:
                if normalized_al_comp not in company_domains:
                    company_domains[normalized_al_comp] = set()
                company_domains[normalized_al_comp].add(domain)

    # ---------------------------
    # 2) Neue Leads prüfen
    # ---------------------------
    new_leads = []
    existing_leads = []
    double_check_leads = []

    total_leads = len(new_leads_df)
    progress_bar = st.progress(0)

    # Liste aller möglichen Firmenspaltentitel, falls mehrere existieren
    possible_company_cols = [
        'Firma/Organisation', 'Company', 'Firma', 
        'Alt Company', 'Alternative Company', 'Weitere Firma'
        # Nach Bedarf ergänzen
    ]

    for idx, lead in new_leads_df.iterrows():
        progress = (idx + 1) / total_leads
        progress_bar.progress(progress)
        
        # (a) E-Mail aus dem Lead holen
        lead_email = ''
        for email_col in ['E-Mail-Adresse', 'Email', 'E-Mail']:
            if email_col in lead and pd.notna(lead[email_col]):
                lead_email = lead[email_col].lower().strip()
                break
        
        lead_local_part = get_local_part(lead_email) if lead_email else ''
        lead_domain_raw = extract_domain(lead_email) if lead_email else ''

        # (b) Mehrere Firmen-Spalten checken - wir nehmen alle vorhandenen
        # und normalisieren sie jeweils
        lead_companies_normalized = []
        for col in possible_company_cols:
            if col in lead and pd.notna(lead[col]):
                nc = normalize_company_name(str(lead[col]))
                if nc:
                    lead_companies_normalized.append(nc)
        
        # Wenn keine Firma da, leere Liste
        if not lead_companies_normalized:
            lead_companies_normalized = []

        # ---------------------------------------------------------
        # Duplikat-Check: E-Mail local-part
        # ---------------------------------------------------------
        # wenn local-part bereits vorhanden -> likely existing
        email_name_match_found = False
        if lead_local_part and (lead_local_part in existing_local_parts):
            email_name_match_found = True

        # ---------------------------------------------------------
        # Domain-Check mit Fuzzy Matching
        # ---------------------------------------------------------
        domain_match_found = False
        domain_double_check = False
        best_domain_score = 0
        matched_domain_company = ""

        if lead_domain_raw:
            # Wir gehen alle existierenden Firmen + deren Domains durch
            for existing_company, domains in company_domains.items():
                for d in domains:
                    # Fuzzy-Vergleich
                    domain_score = fuzz.ratio(lead_domain_raw, d)
                    if domain_score > best_domain_score:
                        best_domain_score = domain_score
                        matched_domain_company = existing_company
        
            # interpretieren, ob wir domain_match_found oder double_check
            if best_domain_score >= HIGH_DOMAIN_THRESHOLD:
                domain_match_found = True
            elif best_domain_score >= MID_DOMAIN_THRESHOLD:
                domain_double_check = True

        # ---------------------------------------------------------
        # Firmen-Check mit Fuzzy
        # ---------------------------------------------------------
        company_match_found = False
        company_double_check = False
        reasons_company = []

        for lead_company_normalized in lead_companies_normalized:
            # (1) Check: exakte Übereinstimmung?
            if lead_company_normalized in known_companies:
                company_match_found = True
                reasons_company.append(
                    f'Firma exakte Übereinstimmung: "{lead_company_normalized}"'
                )
                break  # wir brauchen nur 1 exakten Treffer
            
            # (2) Fuzzy Check: 
            # Wir schauen, ob es eine ~hohe Übereinstimmung zu einer existierenden Firma gibt.
            best_score = 0
            best_match_firm = ""
            for existing_company in known_companies:
                score = fuzz.ratio(lead_company_normalized, existing_company)
                if score > best_score:
                    best_score = score
                    best_match_firm = existing_company
            
            if best_score >= HIGH_COMPANY_THRESHOLD:
                company_match_found = True
                reasons_company.append(
                    f'Fuzzy Firmen-Match {best_score}% -> "{best_match_firm}"'
                )
                break  # High Score reicht uns, wir sind sicher
            elif best_score >= MID_COMPANY_THRESHOLD:
                # Nur wenn wir KEINEN High Score vorher gefunden haben
                # => double check
                company_double_check = True
                reasons_company.append(
                    f'Firma ähnlich (Fuzzy {best_score}%) -> "{best_match_firm}"'
                )
            # Falls wir mehrere company_cols haben, gucken wir alle durch.
            # Möglicherweise kommt später ein HIGH_SCORE -> break nicht vergessen.

        # ---------------------------------------------------------
        # Endgültige Einordnung:
        #
        # Wir betrachten jetzt alle Teil-Checks:
        # 1) email_name_match_found? => existing
        # 2) domain_match_found? => existing
        # 3) domain_double_check? => double_check
        # 4) company_match_found? => existing
        # 5) company_double_check? => double_check
        #
        # "Neue Leads" nur, wenn nichts existiert und kein double_check
        # ---------------------------------------------------------
        lead_dict = lead.to_dict()
        reasons = []

        # Email-Name
        if email_name_match_found:
            reasons.append("Lokaler E-Mail-Name bereits bekannt")

        # Domain
        if domain_match_found:
            reasons.append(
                f"Domain fuzzy-match ~{best_domain_score}% mit Firma: '{matched_domain_company}'"
            )
        elif domain_double_check:
            reasons.append(
                f"Domain ähnlich (~{best_domain_score}%) -> prüfen"
            )

        # Company
        reasons += reasons_company  # falls wir Texte aus dem Fuzzy-Vergleich haben

        # Sammeln, ob existing oder double_check
        if email_name_match_found or domain_match_found or company_match_found:
            lead_dict['Reason'] = " & ".join(reasons)
            existing_leads.append(lead_dict)
        else:
            # Falls wir nicht in existing sind, checken wir double_check
            if domain_double_check or company_double_check:
                lead_dict['Reason'] = " & ".join(reasons)
                double_check_leads.append(lead_dict)
            else:
                # Sonst: New
                lead_dict['Reason'] = " & ".join(reasons)  # i.d.R. leer
                new_leads.append(lead_dict)

    progress_bar.empty()

    # DataFrames zurückgeben
    return (
        pd.DataFrame(new_leads) if new_leads else pd.DataFrame(),
        pd.DataFrame(existing_leads) if existing_leads else pd.DataFrame(),
        pd.DataFrame(double_check_leads) if double_check_leads else pd.DataFrame()
    )

# ------------------------------------------------------------
# Streamlit UI
# ------------------------------------------------------------
st.title("HubSpot Leads Checker")
st.markdown("---")

with st.expander("📋 Instructions", expanded=False):
    st.markdown("""
    1. Export and upload HubSpot Deals Excel file (containing Associated Contact and Company)
    2. Export and upload Deal Alignment Excel file (containing Domain-Name des Unternehmens + Unternehmensname)
    3. Upload your New Leads file (mit mehreren möglichen Firmenspalten)
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

# Button zum Verarbeiten
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
