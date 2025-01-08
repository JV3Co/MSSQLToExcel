import argparse
import pandas as pd
from sqlalchemy import create_engine
import re

# Function to parse tables provided in square brackets
def parse_tables(table_list):
    pattern = r"\[([^\]]+)\]"
    return re.findall(pattern, table_list)

# Function to convert non-binary columns to text
def convert_to_text(df):
    for column in df.columns:
        # Only convert non-binary fields to text
        if df[column].dtype == "object" or pd.api.types.is_string_dtype(df[column]):
            try:
                df[column] = df[column].astype(str).apply(
                    lambda x: x.encode("utf-8", "replace").decode("utf-8") if isinstance(x, str) else x
                )
            except Exception as e:
                print(f"Error converting column '{column}' to text. Retaining original format. Error: {e}")
        else:
            print(f"Leaving raw data for column: {column}")
    return df

# Set up argument parser
parser = argparse.ArgumentParser(description="Export SQL Server tables to an Excel file.")
parser.add_argument("--server", required=True, help="SQL Server address (e.g., localhost, 127.0.0.1)")
parser.add_argument("--database", required=True, help="Name of the SQL Server database")
parser.add_argument("--username", required=True, help="Username for SQL Server")
parser.add_argument("--password", required=True, help="Password for SQL Server")
parser.add_argument("--tables", required=True, help="List of table names in square brackets (e.g., [Table1] [Table2])")
parser.add_argument("--output", required=True, help="Path to the output Excel file")
parser.add_argument(
    "--export-as-text", action="store_true", 
    help="Convert non-binary fields to text while retaining binary fields in raw format"
)

# Parse arguments
args = parser.parse_args()

# Extract arguments
server = args.server
database = args.database
username = args.username
password = args.password
tables_to_export = parse_tables(args.tables)
output_file = args.output
export_as_text = args.export_as_text

# Create SQLAlchemy engine
engine = create_engine(
    f"mssql+pyodbc://{username}:{password}@{server}/{database}?driver=ODBC+Driver+18+for+SQL+Server&TrustServerCertificate=yes"
)

# Export tables to Excel
try:
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for table in tables_to_export:
            query = f"SELECT * FROM [{table}]"
            try:
                # Fetch data
                df = pd.read_sql(query, engine)

                # Convert non-binary fields to text if requested
                if export_as_text:
                    df = convert_to_text(df)

                if not df.empty:
                    df.to_excel(writer, sheet_name=table[:31], index=False)  # Excel sheet name max 31 chars
                    print(f"Exported table: {table}")
                else:
                    print(f"No data found in table: {table}")
            except Exception as e:
                print(f"Error exporting table {table}: {e}")
    print(f"Data exported successfully to {output_file}")
except Exception as e:
    print(f"Failed to export data: {e}")
