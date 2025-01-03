# HubSpot Leads Checker

A Streamlit application for checking new leads against existing HubSpot deals and company alignments.

## Features

- Direct processing of HubSpot Excel exports  
- Smart email domain matching  
- Fuzzy company name matching  
- Support for large files  
- Separate download for new and existing leads  

## Required Files

### HubSpot Deals Excel (alle deals.xlsx):
- Export directly from HubSpot

### Deal Alignment Excel (deal alignment check.xlsx):
- Export directly from HubSpot

### New Leads (Excel or CSV):
- Must contain:  
  - Email address  
  - Company name  
  - First name  
  - Last name  

## Setup

1. Install requirements:
   ```bash
   pip install -r requirements.txt
