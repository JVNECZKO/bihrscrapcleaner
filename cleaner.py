import pandas as pd
import csv
import re

def clean_and_update_excel_with_product_id(input_file_path, reference_data_path, output_csv_path):
    # Import data from sheet1
    sheet1_df = pd.read_excel(input_file_path, sheet_name=0, dtype={'Text2': str})

    # Get reference code from source
    reference_df = pd.read_excel(reference_data_path, dtype={'Reference code': str})

    # Clean Datasheet1
    columns_to_clear = ['Image_URL', 'Image_URL1', 
                        'Text3', 'Text4', 'Field', 'Field1', 'Text5', 
                        'Text6', 'Text7', 'Text8', 'Field2', 'Field3']
    for col in columns_to_clear:
        if col in sheet1_df.columns:
            sheet1_df[col] = ''

    # Reference Clean
    def clean_text2(text):
        text = re.sub(r'Bihr Part No:|NORMAL_PRODUCT|STATIC_KIT', '', text)
        return text.strip()

    sheet1_df['Text2'] = sheet1_df['Text2'].apply(clean_text2)

    # Remove [,] from model
    sheet1_df['Field4'] = sheet1_df['Field4'].apply(lambda x: '' if '[' in str(x) or ']' in str(x) else x)

    # Search by ID for reference
    def find_product_id(reference, ref_df):
        matching_rows = ref_df[ref_df['Reference code'] == reference]
        if not matching_rows.empty:
            return matching_rows.iloc[0]['Product ID']
        else:
            return float('nan')

    # Match ID and Reference
    sheet1_df['Product ID'] = sheet1_df['Text2'].apply(lambda x: find_product_id(x, reference_df))

    # Move data and columns
    sheet1_df['Make'] = sheet1_df['Field4']
    sheet1_df['Model'] = sheet1_df['Field5']
    sheet1_df['Year'] = sheet1_df['Field7']

    # Remove origin that are replaced
    sheet1_df.drop(columns=['Text', 'Text1', 'Text9', 'Image_URL', 'Field4', 'Field5', 'Field7'], inplace=True)

    # Change position
    columns_order = ['Product ID', 'Make', 'Model', 'Year'] + [col for col in sheet1_df.columns if col not in ['Product ID', 'Make', 'Model', 'Year']]
    sheet1_df = sheet1_df[columns_order]

    # Save output to CSV file
    sheet1_df.to_csv(output_csv_path, quoting=csv.QUOTE_NONNUMERIC, index=False)

# Function imput
input_file_path = 'C:\\Users\\lukas\\Desktop\\ALL\\3\\1.xlsx'  # Path to file
reference_data_path = 'C:\\Users\\lukas\\Desktop\\reference_data.xlsx'  # Path to file with with master data from Prestashop
output_csv_path = 'C:\\Users\\lukas\\Desktop\\Gotowe\\TEST335.csv'  # Output save place and filename

clean_and_update_excel_with_product_id(input_file_path, reference_data_path, output_csv_path)
