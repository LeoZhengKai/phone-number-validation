# README: Phone Number Cleaning and Validation Script

# Overview
This Python script processes an Excel file to clean and validate phone numbers based on country-specific rules. It uses the pandas library for data manipulation, phonenumbers for validation, and openpyxl for working with Excel files.

# Features
1. Load Excel Files: Reads data from user-specified Excel sheets.
2. Phone Number Cleaning: Cleans and standardizes phone numbers by removing non-numeric characters.
3. Validation: Validates phone numbers using the phonenumbers library and custom rules for specific countries.
4. Country Detection: Automatically detects the country based on the phone number's prefix.
5. Output: Saves cleaned and validated phone numbers to a new Excel file.

# Installation
Before running the script, ensure the following Python packages are installed:
1. pandas
2. phonenumbers
3. openpyxl
4. You can install them using:

bash
Copy code
pip install pandas phonenumbers openpyxl

# Usage
1. Specify the file path: Update the file_path variable with the path to your Excel file.

2. Run the script:
The script displays the available sheets in the Excel file.
Input the name of the sheet you want to process.

3. Validation Rules:
Phone numbers are cleaned and validated based on predefined country rules stored in the country_data dictionary.

4. Output:
The cleaned and validated phone numbers, along with their country codes and validation status, are saved to a new Excel file.

# Example Output
The script generates an output file named cleaned_phone_numbers_with_country_code.xlsx with the following columns:

country_code: Detected country code.
cleaned_phone_number: Cleaned phone number without the country code.
status: Validation status (e.g., "Valid", "Invalid").

# Notes
Update the phone_column variable to match the column name in your Excel file containing phone numbers.
Modify country_data to add or adjust rules for additional countries if needed.

