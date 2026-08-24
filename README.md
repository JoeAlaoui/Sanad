# SANAD

SANAD is a modular Flask application designed to consolidate operational
and business reporting activities within a unified interface.

## Overview

The application provides reporting modules for:

- Tarkhiss support activities
- Moussanada support activities
- Regulatory file reporting
- Excel and PDF report generation
- Administrative configuration

## Background

SANAD was designed to address recurring reporting needs involving data
collection, consolidation, analysis, visualization and report generation.

The public repository contains a demonstration-oriented version of the
application. Any operational data, personal information and organization-
specific configuration have been excluded.

## Features

- Monthly operational dashboards
- Excel and CSV data import
- KPI calculation
- Interactive charts
- Excel report generation
- PDF report generation
- Configurable statuses
- Modular reporting architecture
- Administration interface

## Technology Stack

- Python
- Flask
- JavaScript
- HTML5
- CSS3
- OpenPyXL
- Chart.js
- ReportLab

## Project Structure

```text
sanad/
├── app.py
├── glpi_import.py
├── pchc_import.py
├── requirements.txt
├── static/
│   ├── app.js
│   └── style.css
└── templates/
    └── index.html
