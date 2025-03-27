import os
import re
import pandas as pd
import mysql.connector
from mysql.connector import Error
from decimal import Decimal, ROUND_HALF_UP
import logging
from datetime import datetime

# Set up logging
log_file = f"cis_price_import_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file),
        logging.StreamHandler()
    ]
)

# Define the folder paths
base_folder = r"C:\Users\cs.lee.ADVANTECH\Desktop\CIS_PirceInfo"
excel_folder_path = os.path.join(base_folder, "ExcelData")

# Define the database connection parameters
db_config = {
    'host': 'localhost',      # Update with your MySQL host
    'database': 'cis_schema', # Your database name
    'user': 'root',           # Update with your MySQL username
    'password': 'yuioHJK&345'  # Update with your MySQL password
}

# Function to extract date from file name
def extract_date_from_filename(filename):
    date_pattern = r'CIS-Raw Data (\d{4}-\d{2}-\d{2})'
    match = re.search(date_pattern, filename)
    if match:
        return match.group(1)
    return None

# Fixed function to find column name regardless of case and type
def find_column(df, possible_names):
    # First try exact match
    for name in possible_names:
        if name in df.columns:
            return name
    
    # Then try case-insensitive match
    columns_dict = {}
    for col in df.columns:
        try:
            # Convert to string first to handle non-string column names
            col_str = str(col)
            columns_dict[col_str.lower()] = col
        except Exception as e:
            logging.warning(f"Error processing column name '{col}': {e}")
    
    for name in possible_names:
        try:
            name_lower = name.lower()
            if name_lower in columns_dict:
                return columns_dict[name_lower]
        except Exception as e:
            logging.warning(f"Error comparing column name '{name}': {e}")
    
    # Print available columns for debugging
    logging.info(f"Available columns: {', '.join([str(col) for col in df.columns])}")
    return None

# Function to establish database connection
def get_db_connection():
    try:
        connection = mysql.connector.connect(**db_config)
        return connection
    except Error as e:
        logging.error(f"Error connecting to MySQL database: {e}")
        return None

# Function to convert string to decimal with exactly 5 decimal places
def convert_to_decimal(value_str):
    if pd.isna(value_str) or str(value_str).strip() == '':
        return Decimal('0.00000')
    
    # Convert to string first to handle any type
    value_str = str(value_str).strip()
    
    # Remove currency symbols, commas, and other non-numeric characters except decimal point and minus sign
    value_clean = re.sub(r'[^\d.-]', '', value_str)
    
    try:
        # Convert to Decimal with exactly 5 decimal places
        value_decimal = Decimal(value_clean).quantize(Decimal('0.00001'), rounding=ROUND_HALF_UP)
        return value_decimal
    except Exception as e:
        logging.warning(f"Could not convert '{value_str}' to decimal: {e}")
        return Decimal('0.00000')

# Function to process each Excel file
def process_excel_file(file_path, file_date):
    logging.info(f"Starting to process file: {os.path.basename(file_path)}")
    
    # Initialize counters
    records_inserted = 0
    records_updated = 0
    records_skipped = 0
    
    connection = None
    cursor = None
    
    try:
        # Read Excel file, skip the header row (starting from row 2)
        try:
            df = pd.read_excel(file_path )
            logging.info(f"Excel file read successfully with {len(df)} rows")
        except Exception as e:
            logging.error(f"Error reading Excel file {file_path}: {e}")
            return 0, 0, 0
        
        # Log the column names to help with debugging
        logging.info(f"Excel columns found: {', '.join([str(col) for col in df.columns])}")
        
        # Find the column names (case-insensitive)
        part_number_col = find_column(df, ["Part_Number", "PN", "Part Number", "PartNumber"])
        cost_usd_col = find_column(df, ["Cost_USD", "Cost", "USD", "Price", "Cost (USD)"])
        
        if not part_number_col:
            logging.error(f"Part Number column not found in {file_path}.")
            return 0, 0, 0
        
        if not cost_usd_col:
            logging.error(f"Cost USD column not found in {file_path}.")
            return 0, 0, 0
        
        logging.info(f"Using columns: Part Number = '{part_number_col}', Cost USD = '{cost_usd_col}'")
        
        # Connect to MySQL database
        connection = get_db_connection()
        if not connection:
            return 0, 0, 0
        
        cursor = connection.cursor()
        
        # SQL query for inserting/updating records
        upsert_query = """
        INSERT INTO part_cost_history_new (Part_Number, Cost_USD, Date)
        VALUES (%s, %s, %s)
        ON DUPLICATE KEY UPDATE Cost_USD = VALUES(Cost_USD)
        """
        

        # Process rows starting from the second row (index 1)
        for index, row in df.iloc[1:].iterrows():
            try:
                # Get Part_Number
                part_number = row[part_number_col]
                
                # Skip if Part_Number is blank or null
                if pd.isna(part_number) or str(part_number).strip() == '':
                    records_skipped += 1
                    continue
                
                # Clean up part number
                part_number = str(part_number).strip()
                
                # Get the Cost_USD value as string and ensure it has exactly 5 decimal places
                cost_usd_str = row[cost_usd_col]
                cost_usd = convert_to_decimal(cost_usd_str)
                
                # For logging/debugging - show original and converted values for the first few rows
                if index < 5:
                    logging.info(f"Row {index+2}: PN: '{part_number}', Original Cost: '{cost_usd_str}', Converted: {cost_usd}")
                
                # Check if record already exists (to differentiate between inserts and updates)
                check_query = """
                SELECT id FROM part_cost_history_new 
                WHERE Part_Number = %s AND Date = %s
                """
                cursor.execute(check_query, (part_number, file_date))
                existing_record = cursor.fetchone()
                
                # Insert or update record
                cursor.execute(upsert_query, (part_number, str(cost_usd), file_date))
                
                # Count as insert or update
                if existing_record:
                    records_updated += 1
                else:
                    records_inserted += 1
                
            except Exception as e:
                logging.error(f"Error processing row {index + 2}: {e}")
                records_skipped += 1
                continue
        
        # Commit changes to the database
        connection.commit()
        
        return records_inserted, records_updated, records_skipped
    
    except Exception as e:
        logging.error(f"Unexpected error processing file {file_path}: {e}")
        if connection:
            connection.rollback()
        return 0, 0, 0
    
    finally:
        # Close database connection
        if cursor:
            cursor.close()
        if connection:
            connection.close()

# Main function to process all Excel files
def main():
    logging.info("=" * 80)
    logging.info("Starting CIS Price Import process")
    logging.info("=" * 80)
    
    # List all Excel files in the folder
    try:
        all_files = os.listdir(excel_folder_path)
        excel_files = [f for f in all_files if f.startswith("CIS-Raw Data ") and (f.endswith(".xlsx") or f.endswith(".xls"))]
        logging.info(f"Found {len(excel_files)} Excel files to process out of {len(all_files)} total files")
    except Exception as e:
        logging.error(f"Error accessing folder {excel_folder_path}: {e}")
        return
    
    if not excel_files:
        logging.warning("No Excel files found in the specified folder.")
        return
    
    # Process each Excel file
    total_records_inserted = 0
    total_records_updated = 0
    total_records_skipped = 0
    
    for excel_file in excel_files:
        file_path = os.path.join(excel_folder_path, excel_file)
        file_date = extract_date_from_filename(excel_file)
        
        if not file_date:
            logging.warning(f"Could not extract date from file name: {excel_file}. Skipping...")
            continue
        
        logging.info(f"Processing file: {excel_file} with date {file_date}")
        records_inserted, records_updated, records_skipped = process_excel_file(file_path, file_date)
        
        logging.info(f"Results for file: {excel_file}")
        logging.info(f"  - Records inserted: {records_inserted}")
        logging.info(f"  - Records updated: {records_updated}")
        logging.info(f"  - Records skipped: {records_skipped}")
        logging.info("-" * 50)
        
        total_records_inserted += records_inserted
        total_records_updated += records_updated
        total_records_skipped += records_skipped
    
    logging.info("=" * 80)
    logging.info("Summary:")
    logging.info(f"Total records inserted: {total_records_inserted}")
    logging.info(f"Total records updated: {total_records_updated}")
    logging.info(f"Total records skipped: {total_records_skipped}")
    logging.info("CIS Price Import process completed")
    logging.info("=" * 80)

if __name__ == "__main__":
    main()