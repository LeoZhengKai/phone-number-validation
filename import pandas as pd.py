import pandas as pd
import phonenumbers
import re
from openpyxl import load_workbook


# Step 1: Load the Excel File
file_path = '/Users/leozhengkai/Downloads/sans-vip-sin 15 Oct 2024 Updated (2).xlsx'
df = pd.read_excel(file_path, engine='openpyxl')

# Step 2: Define the column that contains phone numbers
phone_column = '电话'  # Replace 'phone_number' with your actual column name if needed

# Display available sheet names
xls = pd.ExcelFile(file_path)
print("Available sheets:", xls.sheet_names)

# Choose a sheet to work on
sheet_name = input("Enter the sheet name you want to process: ")

# Load the selected sheet into a DataFrame
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Step 3: Country Data for Validation
country_data = {
    "+374": ["Armenia", [8], "0-9"],
    "+61": ["Australia", [9], "4"],  # Mobile starts with 4, 9 digits
    "+880": ["Bangladesh", [10], "1-9"],  # Bangladesh, 10 digits
    "+975": ["Bhutan", [8], "1-9"],  # Bhutan, 8 digits
    "+673": ["Brunei Darussalam", [7], "2-8"],  # Brunei, 7 digits
    "+855": ["Cambodia", [9], "1-9"],  # Cambodia, 9 digits
    "+86": ["China", [11], "1-9"],  # China, 11 digits
    "+91": ["India", [10], "7-9"],  # India, 10 digits, starts with 7, 8, or 9
    "+62": ["Indonesia", [10, 11, 12, 13], "8"],  # Indonesia, 10-12 digits, mobile starts with 8
    "+972": ["Israel", [9], "5"],  # Israel, 9 digits, mobile starts with 5
    "+81": ["Japan", [10], "1-9"],  # Japan, 10 digits
    "+962": ["Jordan", [9], "7"],  # Jordan, 9 digits, mobile starts with 7
    "+7": ["Kazakhstan", [10], "7"],  # Kazakhstan, 10 digits, mobile starts with 7
    "+965": ["Kuwait", [8], "1-9"],  # Kuwait, 8 digits
    "+856": ["Laos", [8], "2-9"],  # Laos, 8 digits
    "+60": ["Malaysia", [9], "1-9"],  # Malaysia, 9 digits
    "+960": ["Maldives", [7], "1-9"],  # Maldives, 7 digits
    "+977": ["Nepal", [10], "9"],  # Nepal, 10 digits, mobile starts with 9
    "+92": ["Pakistan", [10], "3"],  # Pakistan, 10 digits, mobile starts with 3
    "+63": ["Philippines", [10], "9"],  # Philippines, 10 digits, mobile starts with 9
    "+974": ["Qatar", [8], "3-7"],  # Qatar, 8 digits, starts with 3-7
    "+966": ["Saudi Arabia", [9], "5"],  # Saudi Arabia, 9 digits, mobile starts with 5
    "+65": ["Singapore", [8], "8-9"],  # Singapore, 8 digits, starts with 8 or 9
    "+82": ["South Korea", [10], "1-9"],  # South Korea, 10 digits
    "+94": ["Sri Lanka", [9], "7"],  # Sri Lanka, 9 digits, mobile starts with 7
    "+886": ["Taiwan", [9], "9"],  # Taiwan, 9 digits, mobile starts with 9
    "+66": ["Thailand", [9], "8"],  # Thailand, 9 digits, mobile starts with 8
    "+971": ["United Arab Emirates", [9], "5"],  # UAE, 9 digits, mobile starts with 5
    "+1": ["USA/Canada", [10], "2-9"],  # USA, Canada, 10 digits, starts with 2-9
    "+44": ["United Kingdom", [10], "7"],  # UK, 10 digits, mobile starts with 7
    "+55": ["Brazil", [10], "9"],  # Brazil, 10 digits, mobile starts with 9
    "+52": ["Mexico", [10], "1-9"],  # Mexico, 10 digits
    "+33": ["France", [9], "6-7"],  # France, 9 digits, mobile starts with 6 or 7
    "+49": ["Germany", [10], "1-9"],  # Germany, 10 digits
    "+39": ["Italy", [9], "3"],  # Italy, 9 digits, mobile starts with 3
    "+34": ["Spain", [9], "6"],  # Spain, 9 digits, mobile starts with 6
    "+27": ["South Africa", [9], "7-8"],  # South Africa, 9 digits, starts with 7 or 8
    "+84": ["Vietnam", [9], "1-9"],  # Vietnam, 9 digits
    "+93": ["Afghanistan", [9], "7"],  # Afghanistan, 9 digits, mobile starts with 7
    "+994": ["Azerbaijan", [9], "1-9"],  # Azerbaijan, 9 digits
    "+973": ["Bahrain", [8], "3-9"],  # Bahrain, 8 digits
    "+961": ["Lebanon", [8], "3-9"],  # Lebanon, 8 digits
    "+95": ["Myanmar", [8], "1-9"],  # Myanmar, 8 digits
    "+968": ["Oman", [8], "9"],  # Oman, 8 digits, mobile starts with 9
    "+964": ["Iraq", [10], "7"],  # Iraq, 10 digits, mobile starts with 7
    "+996": ["Kyrgyzstan", [9], "1-9"],  # Kyrgyzstan, 9 digits
    "+90": ["Turkey", [10], "5"],  # Turkey, 10 digits, mobile starts with 5
    "+993": ["Turkmenistan", [8], "1-9"],  # Turkmenistan, 8 digits
    "+998": ["Uzbekistan", [9], "9"],  # Uzbekistan, 9 digits, mobile starts with 9
    "+850": ["North Korea", [8], "1-9"],  # North Korea, 8 digits
    "+852": ["Hong Kong", [8], "9"],  # Hong Kong, 8 digits, mobile starts with 9
    "+853": ["Macau", [8], "6"],  # Macau, 8 digits, mobile starts with 6
    "+976": ["Mongolia", [8], "8"],  # Mongolia, 8 digits, mobile starts with 8
    "+964": ["Yemen", [9], "7"],  # Yemen, 9 digits, mobile starts with 7
}

# Step 4: Function to Clean and Validate Phone Numbers Using `country_data`
def clean_phone_number(number, default_country="SG", country_data=country_data):
    """
    Cleans and standardizes phone numbers, extracts the country code,
    phone number without the country code, and validation status.
    
    Args:
    - number: The raw phone number string.
    - default_country: Default country code if no country code is provided.
    - country_data: A dictionary of country codes and validation rules.
    
    Returns:
    - Tuple: (country_code, phone_number, validation_status)
    """
    
    # Step 1: Remove all non-numeric characters except the '+' sign
    cleaned_number = re.sub(r'[^\d+]', '', str(number))
    
    # Step 2: Check for numbers starting with 00 (international dialing without +)
    if cleaned_number.startswith("00"):
        cleaned_number = "+" + cleaned_number[2:]  # Replace 00 with +

    # Step 3: Handle local numbers with leading 0
    if cleaned_number.startswith("0") and not cleaned_number.startswith("+"):
        # Assume it's a local number and add default country code
        cleaned_number = "+" + str(phonenumbers.country_code_for_region(default_country)) + cleaned_number[1:]

    # Step 4a: If number starts with a country code without the '+' sign, try to detect it
    for code in country_data.keys():
        if cleaned_number.startswith(code.replace('+', '')):
            cleaned_number = '+' + cleaned_number  # Add '+' to the detected country code
            break

    # Step 4b: Parse the cleaned number using phonenumbers library
    try:
        # If the number starts with +, treat it as international
        if cleaned_number.startswith("+"):
            parsed_number = phonenumbers.parse(cleaned_number, None)
        else:
            # Otherwise, assume it's a local number for the default country
            parsed_number = phonenumbers.parse(cleaned_number, default_country)

        # Step 5: Validate the phone number using phonenumbers
        if phonenumbers.is_valid_number(parsed_number):
            # Extract country code and the actual phone number (without the country code)
            country_code = "+" + str(parsed_number.country_code)
            actual_phone_number = phonenumbers.national_significant_number(parsed_number)
            
            # Step 6: Second check using country_data
            if country_code in country_data:
                country_info = country_data[country_code]
                valid_lengths = country_info[1]
                valid_starts = country_info[2]

                # Check if the phone number length is valid
                if len(actual_phone_number) not in valid_lengths:
                    return country_code, actual_phone_number, f"Invalid (Length should be {valid_lengths})"
                
                # Check if the phone number starts with the expected digits
                if not re.match(f"^[{valid_starts}]", actual_phone_number):
                    return country_code, actual_phone_number, f"Invalid (Should start with {valid_starts})"
                
                # If both checks pass
                return country_code, actual_phone_number, "Valid"
            else:
                return country_code, actual_phone_number, "Valid (No additional country check)"
        
        else:
            return "", cleaned_number, "Invalid (Invalid number format)"
    
    except phonenumbers.phonenumberutil.NumberParseException:
        return "", cleaned_number, "Invalid (Parsing error)"

# Step 7: Apply the cleaning and validation function to the DataFrame
# Extract the country code, phone number (without country code), and validation status
df['country_code'], df['cleaned_phone_number'], df['status'] = zip(*df[phone_column].apply(lambda x: clean_phone_number(x)))


# Step 6: Save the cleaned and validated data to a new Excel file
output_file = '/Users/leozhengkai/Downloads/cleaned_phone_numbers_with_country_code.xlsx'
df.to_excel(output_file, sheet_name=f'{sheet_name}_Processed', index=False)

print(f"Phone numbers cleaned and saved to {output_file}")