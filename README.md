# HubSpot Leads Checker

A Streamlit application for checking new leads against existing HubSpot deals and company alignments.

## Features
- Direct processing of HubSpot Excel exports
- Email domain matching
- Company and contact verification
- Support for large files
- Three-way classification: New, Existing, and Double Check

## Required Files
### HubSpot Deals Excel (alle deals.xlsx):
- Associated Contact (Name and Email)
- Associated Company

### Deal Alignment Excel (deal alignment check.xlsx):
- Domain-Name des Unternehmens
- Unternehmensname

### New Leads Excel:
- Email address
- Company name
- Contact details

## Setup
1. Install requirements:
```bash
pip install -r requirements.txt
