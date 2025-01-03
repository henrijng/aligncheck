# HubSpot Leads Checker

A Streamlit application for checking new leads against existing HubSpot deals and company alignments.

## Features

- Direct processing of HubSpot Excel exports
- **Extended** fuzzy domain and company name matching
- **Multiple** possible company columns recognized
- **Email local-part** check for duplicates
- Support for large files
- Separate download for new, existing, and double-check leads

## Required Files

### HubSpot Deals Excel (alle deals.xlsx)
- Export directly from HubSpot (must contain columns **Associated Contact** and **Associated Company**)

### Deal Alignment Excel (deal alignment check.xlsx)
- Export directly from HubSpot (must contain columns **Unternehmensname** and **Domain-Name des Unternehmens**)

### New Leads (Excel)
- Should include:
  - Email address (e.g. `E-Mail-Adresse`, `Email`, `E-Mail`)
  - One or more Company columns (e.g. `Firma/Organisation`, `Company`, `Alt Company`, etc.)
  - First name / Last name (optional but recommended)

## Setup

1. Clone the repository.
2. Install the requirements:
   ```bash
   pip install -r requirements.txt
