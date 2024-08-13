import pandas as pd
import warnings
from openpyxl import Workbook

# Suppress openpyxl UserWarning about default style
warnings.simplefilter("ignore", UserWarning)

# Define the function to transform the registration list with seating fallback
def transform_registration_list_with_seating_fallback(file_path, seating_file_path, output_file_path):
    # Load the Excel file
    excel_data = pd.ExcelFile(file_path)
    seating_data = pd.ExcelFile(seating_file_path)

    # Load the data from the main sheet
    sheet_name = [s for s in excel_data.sheet_names if 'Convention' in s][0]
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    
    # Load the seating chart data
    seating_chart_df = pd.read_excel(seating_file_path)

    # Strip any leading or trailing whitespace from the column names
    df.columns = df.columns.str.strip()
    seating_chart_df.columns = seating_chart_df.columns.str.strip()

    # Remove the apostrophe from event names
    df['Event'] = df['Event'].str.replace('’', "'")

    # Filter the dataframe to include only rows where the status is "Paid"
    paid_df = df[df['Status Reason'] == 'Paid']

    # Create a list of unique events
    unique_events = paid_df['Event'].unique()

    # Create a new DataFrame to store the transformed data
    transformed_df = paid_df[['Existing Contact', 'Date of Birth (Existing Contact) (Contact)', 'First Name (Existing Contact) (Contact)', 'Last Name (Existing Contact) (Contact)']].drop_duplicates().reset_index(drop=True)

    # Add each event as a new column
    for event in unique_events:
        transformed_df[event] = transformed_df.apply(
            lambda row: event if event in paid_df[(paid_df['Existing Contact'] == row['Existing Contact']) & 
                                                  (paid_df['Date of Birth (Existing Contact) (Contact)'] == row['Date of Birth (Existing Contact) (Contact)'])]['Event'].values 
            else None, axis=1
        )

    # Extract relevant columns from the seating chart
    seating_info_df = seating_chart_df[['Contact', 'Date of Birth (Contact) (Contact)', 'Table', 'Name', 'Event']].rename(
        columns={'Contact': 'Existing Contact', 'Date of Birth (Contact) (Contact)': 'Date of Birth (Existing Contact) (Contact)'})

    # Merge the transformed data with the seating chart data
    final_df = pd.merge(transformed_df, seating_info_df, on=['Existing Contact', 'Date of Birth (Existing Contact) (Contact)'], how='left')

    # If 'Table' is NaN, fallback to matching with 'Name'
    fallback_df = pd.merge(transformed_df, seating_chart_df[['Name', 'Table', 'Event']], left_on='Existing Contact', right_on='Name', how='left', suffixes=('', '_fallback'))

    # Update 'Table' column with fallback values where applicable
    final_df['Table'] = final_df['Table'].combine_first(fallback_df['Table'])
    final_df['Event'] = final_df['Event'].combine_first(fallback_df['Event'])

    # Add special guests for Banquet night to the existing event column
    banquet_guests_df = seating_chart_df[seating_chart_df['Contact'].isna() & (seating_chart_df['Name'].notna())].copy()
    banquet_guests_df = banquet_guests_df[['Name', 'Table', 'Event']].rename(columns={'Name': 'Existing Contact'})

    # Combine the final DataFrame with the banquet guests
    final_df = pd.concat([final_df, banquet_guests_df], ignore_index=True, sort=False)

    # Sort the final DataFrame by last name
    final_df = final_df.sort_values(by=['Last Name (Existing Contact) (Contact)', 'First Name (Existing Contact) (Contact)'])

    # Save the final dataframe to a new Excel file
    final_df.to_excel(output_file_path, index=False)

    return final_df

# Define the function to remove (extract) specific names
def extract_specific_names(df, names_to_extract, output_file_path):
    if names_to_extract:
        # Filter the DataFrame to include only the specified names
        filtered_df = df[df['Existing Contact'].isin(names_to_extract)]
        # Save the filtered DataFrame to a new CSV file
        filtered_df.to_csv(output_file_path, index=False)
    else:
        filtered_df = pd.DataFrame()
    
    return filtered_df

# Define the file paths
file_path = './Jacksonville Convention 2024 Registration List 7-15-2024 2-20-21 PM.xlsx'
seating_file_path = './Jacksonville Convention Seating Chart 7-12-2024 2-47-26 PM.xlsx'
output_file_path_transformed = './MAIL_MERGEABLE_Transformed_Registration_List.xlsx'
output_file_path_filtered = './Filtered_Names_List.csv'

# Run the transformation function with fallback
final_df_with_seating_fallback = transform_registration_list_with_seating_fallback(file_path, seating_file_path, output_file_path_transformed)

# Define the list of names to remove
names_to_extract = [
  "SampleFirst SampleLast"
]

# Extract specific names from the transformed data
final_filtered_df = extract_specific_names(final_df_with_seating_fallback, names_to_extract, output_file_path_filtered)
