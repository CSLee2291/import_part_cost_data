# CIS Price Information Importer

## Overview
This tool imports part cost information from Excel files into a MySQL database. It processes Excel files with the naming pattern "CIS-Raw Data YYYY-MM-DD.xlsx" located in the ExcelData directory, extracts relevant information, and stores it in the `part_cost_history` table.

## Features
- Automatically extracts date information from filenames
- Identifies required columns in Excel files (with flexible column name matching)
- Validates and cleans data before insertion
- Handles missing or NULL values appropriately
- Enforces database column length constraints
- Avoids duplicate entries
- Provides detailed progress and error reporting

## Prerequisites
- Python 3.6 or higher
- Required Python packages:
  - pandas
  - mysql-connector-python
  - openpyxl (for reading Excel files)
- MySQL database server

## Installation

### Install Required Libraries
```
pip install pandas mysql-connector-python openpyxl
```

### Database Setup
The database schema should include the following table:

```sql
CREATE TABLE part_cost_history (
    id INT AUTO_INCREMENT PRIMARY KEY,
    Part_Number VARCHAR(50) NOT NULL,
    Manufacture VARCHAR(100) NOT NULL,
    Manufacture_Part_Number VARCHAR(50) NOT NULL,
    Cost_USD DECIMAL(10, 2) NOT NULL,
    Date DATE NOT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    modified_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    INDEX (Part_Number),
    INDEX (Manufacture),
    INDEX (Date),
    UNIQUE KEY unique_part_info (Part_Number, Manufacture, Manufacture_Part_Number, Date)
) COMMENT 'Stores part cost information at specific dates';
```

## File Structure
```
C:\Users\cs.lee.ADVANTECH\Desktop\CIS_PirceInfo\
│
├── import_part_cost_data.py     # Main Python script
├── README.md                    # This documentation file
│
└── ExcelData\                   # Directory containing Excel files
    ├── CIS-Raw Data 2023-01-01.xlsx
    ├── CIS-Raw Data 2023-02-01.xlsx
    └── ...
```

## Configuration
Before running the script, update the MySQL connection details in the script:

```python
mysql_config = {
    'host': 'localhost',
    'user': 'your_username',     # Replace with your MySQL username
    'password': 'your_password', # Replace with your MySQL password
    'database': 'your_database'  # Replace with your MySQL database name
}
```

## Usage

### Running the Script
1. Open Command Prompt
2. Navigate to the script directory:
   ```
   cd C:\Users\cs.lee.ADVANTECH\Desktop\CIS_PirceInfo
   ```
3. Run the script:
   ```
   python import_part_cost_data.py
   ```

### Expected Output
The script will display progress information as it processes each Excel file:
- Number of files found and processed
- Column mapping information
- Number of rows inserted, skipped, and values truncated
- Any errors encountered
- Summary of the import process

## Data Processing Details

### Column Mapping
The script tries to find the required columns in the Excel files using the following strategies:
1. Exact column name matches
2. Case-insensitive matches
3. Alternative column names:
   - **Part_Number**: PartNumber, Part Number, Part#, PN
   - **Manufacture**: Manufacturer, Vendor, Supplier
   - **Manufacture_Part_Number**: ManufacturePartNumber, Manufacture Part Number, MPN, Vendor Part Number
   - **Cost_USD**: Cost, Price, USD, Cost(USD)

### Data Validation
- Part_Number is required - rows with blank Part_Number are skipped
- Missing values for Manufacture and Manufacture_Part_Number are replaced with "Unknown"
- Missing Cost_USD values are set to 0.0
- Strings longer than their column limits are truncated:
  - Part_Number: 50 characters
  - Manufacture: 100 characters
  - Manufacture_Part_Number: 50 characters

## Troubleshooting

### Common Issues
1. **Database Connection Errors**:
   - Verify MySQL connection settings
   - Ensure MySQL server is running
   - Check user permissions

2. **File Access Issues**:
   - Ensure Excel files exist in the correct directory
   - Verify file naming follows the pattern "CIS-Raw Data YYYY-MM-DD.xlsx"

3. **Data Import Errors**:
   - Check Excel column names and formats
   - Look for unexpected data types in the Excel files

### Debugging
If errors occur:
1. Check the console output for specific error messages
2. Verify the Excel file structure matches expected format
3. Try running the script with a single Excel file for testing

## Maintenance
- Periodically check for duplicate entries in the database
- Update the script if Excel file formats change
- Consider creating a backup of the database before large imports
