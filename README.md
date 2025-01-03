# HubSpot Leads Checker

An advanced Streamlit application for checking new leads against existing HubSpot deals and company alignments. The app handles email domain variations and performs fuzzy company name matching.

## Features

- Direct processing of HubSpot CSV exports without modification
- Smart email domain matching (handles variations like .com, .de, .nl)
- Fuzzy company name matching to catch slight variations
- Supports large CSV files with efficient processing
- Separate download for new and existing leads
- Detailed matching reasons in output

## Required CSV Files (Direct HubSpot Exports)

### HubSpot Deals CSV:
- Export directly from HubSpot deals
- No modifications needed
- Default export format supported

### Deal Alignment CSV:
- Export directly from HubSpot
- Supports default columns:
  - "Unternehmensname"
  - "Domain-Name des Unternehmens"

### New Leads CSV:
- Required columns:
  - E-Mail-Adresse
  - Firma/Organisation
  - Vorname
  - Nachname

## Setup

1. Install requirements:
```bash
pip install -r requirements.txt
