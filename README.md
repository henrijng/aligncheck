# HubSpot Leads Checker

An advanced Streamlit application for checking new leads against existing HubSpot deals and company alignments. The app handles email domain variations and performs fuzzy company name matching.

## Features

- Smart email domain matching (handles variations like .com, .de, .nl)
- Fuzzy company name matching to catch slight variations
- Supports large CSV files with efficient processing
- Handles German character encoding (CP1252)
- Separate download for new and existing leads
- Detailed matching reasons in output

## Required CSV Formats

### HubSpot Deals CSV (alle Deals.csv):
- Required columns:
  - Associated Email
  - Associated Company
  - Deal-Name
  - Associated Contact
  - Associated Company IDs

### Deal Alignment CSV (Deal alignment check.csv):
- Required columns:
  - Unternehmensname (Company name)
  - Domain-Name des Unternehmens (Company domain)

### New Leads CSV:
- Required columns:
  - E-Mail-Adresse (Email address)
  - Firma/Organisation (Company/Organization)
  - Vorname (First name)
  - Nachname (Last name)

## Setup

1. Install requirements:
```bash
pip install -r requirements.txt
```

2. Run the app locally:
```bash
streamlit run app.py
```

## Usage

1. Upload your HubSpot Deals CSV file
2. Upload your Deal Alignment CSV file
3. Upload your New Leads CSV file
4. Click Process to analyze the data
5. Download the filtered results

## Error Handling

The app includes:
- CSV format validation
- Encoding detection (CP1252/UTF-8)
- Missing column checks
- Domain normalization
- Company name normalization

## Output Files

1. New Leads CSV:
   - Contains leads with no existing deals
   - Maintains original data format
   - Includes timestamp in filename

2. Existing Leads CSV:
   - Shows matching leads
   - Includes reason for match
   - Details about existing deal/company

## Notes

- The app handles domain variations (.com, .de, .nl, etc.)
- Company names are matched using fuzzy logic to catch slight variations
- All processing is done locally for data security
- Large files are processed efficiently with progress indicators
