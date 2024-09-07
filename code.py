## remember  to install --- pip install pandas openpyxl##

import pandas as pd
from urllib.parse import quote

# Function to encode URLs
def encode_url(url):
    return quote(url, safe='')

# Load the Excel file from the second sheet (index 1)
file_path = 'Resources.xlsx'  # Replace with your actual file path
df = pd.read_excel(file_path, sheet_name='WHA')  # sheet_name=1 specifies the second sheet

# Create a list to hold all the rows for the CSV
csv_data = []

# Generate formatted text and prepare CSV data
for index, row in df.iterrows():
    # Encode the Resources link
    encoded_link = encode_url(row['Resources link'])

    # Create the formatted text
    formatted_text = f'''
    [vc_row equal_height="yes" type="vc_default" css=".vc_custom_1575291876546{{margin-bottom: -70px !important;}}"][vc_column width="1/2" css=".vc_custom_1575291908923{{padding-bottom: 50px !important;}}" offset="vc_col-lg-4 vc_col-md-5"][dt_fancy_image type="from_url" image="{row['Image Link']}"][vc_column_text css=".vc_custom_1723999004127{{margin-bottom: 40px !important;border-left-width: 3px !important;padding-left: 20px !important;border-left-style: solid !important;border-color: rgba(0,0,0,0.1) !important;}}"]VCA-Lebanon[/vc_column_text][dt_default_button link="url:{encoded_link}|target:_blank" size="big" btn_width="btn_full_width" icon_type="none"]Read Document[/dt_default_button][/vc_column][vc_column width="1/2" css=".vc_custom_1575291917463{{padding-bottom: 50px !important;}}" offset="vc_col-lg-8 vc_col-md-7"][ultimate_heading main_heading="{row['Title']}" heading_tag="h3" alignment="left" main_heading_margin="margin-bottom:30px;" main_heading_style="font-weight:bold;"][/ultimate_heading][vc_column_text css=".vc_custom_1575291964331{{padding-bottom: 30px !important;}}"]{row['Description']}[/vc_column_text][/vc_column][/vc_row]
    '''


    # Prepare the row for the CSV
    csv_row = {
        'Title': row['Title'],
        'Content': formatted_text,
        'Post Type': 'dt_portfolio',
        'Image URL': row['Image Link'],
        'Image Featured': row['Image Link'],
        'Portfolio Categories': 'WHA'
    }

    # Append the row to the CSV data list
    csv_data.append(csv_row)

# Create a DataFrame for the CSV
csv_df = pd.DataFrame(csv_data)

# Save the DataFrame to a CSV file
csv_file_path = 'output.csv'  # You can change the file name if needed
csv_df.to_csv(csv_file_path, index=False)

print(f"CSV file '{csv_file_path}' has been created successfully.")
