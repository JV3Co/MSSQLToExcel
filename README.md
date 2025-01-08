# MSSQL to Excel Exporter

This Python script exports specified SQL Server tables to an Excel file.

## Features
- Export tables with raw data or convert non-binary fields to text.

## Prerequisites
- Python 3.x installed.
- Microsoft ODBC Driver for SQL Server installed.

## Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/export-script.git
   cd export-script

## Example Terminal Commands
1. Export to Excel with raw field types

    python MSSQLToExcel.py \
        --server localhost \
        --database SampleDB \
        --username sa \
        --password password123 \
        --tables "[Employees] [Departments]" \
        --output exported_data.xlsx


2. Export to Excel with all fields converted to text. 
    **VARBINARY fields will not be converted, but will still be included in the output.

    python MSSQLToExcel.py \
        --server localhost \
        --database SampleDB \
        --username sa \
        --password password123 \
        --tables "[Employees] [Departments]" \
        --output exported_data.xlsx \
        --export-as-text

