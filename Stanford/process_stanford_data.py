import pandas as pd

year = 24

# Load the Excel file
df = pd.read_excel('Table_1_Authors_career_' + str(year + 2000) + '.xlsx', sheet_name='Data')

# Filter rows where 'Období' column equals 'H' + year
filtered_df = df[ 
    (df['inst_name'] == ('Czech Technical University in Prague')) 
].copy()

# Save the filtered data to a new Excel file
filtered_df.to_excel(str(year) + '_stanford_filtered.xlsx', index=False)