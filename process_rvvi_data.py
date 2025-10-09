import pandas as pd

year = 24

# Load the Excel file
df = pd.read_excel('24_evaluations.xlxs', header=2)

# Filter rows where 'Období' column contains 'H20'
filtered_df = df[ 
    (df['Období'] == 'H24') & 
    (df['Organizační jednotka'] == 'České vysoké učení technické v Praze/Fakulta stavební') 
]

filtered_df.sort_values(by=['Obor (Ford)', 'Název výsledku', 'Vypracoval'], inplace=True)

# Save the filtered data to a new Excel file
filtered_df.to_excel('24_evaluations_filtered.xlsx', index=False)