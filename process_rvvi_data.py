import pandas as pd

year = 24

# Load the Excel file
df = pd.read_excel('M1_VO_posudky_H20-H24.xlsx', header=2)

# Filter rows where 'Období' column equals 'H' + year
filtered_df = df[ 
    (df['Období'] == ('H' + str(year))) & 
    (df['Organizační jednotka'] == 'České vysoké učení technické v Praze/Fakulta stavební') 
].copy()

filtered_df.sort_values(by=['Obor (Ford)', 'Název výsledku', 'Vypracoval'], inplace=True)

# Save the filtered data to a new Excel file
filtered_df.to_excel(str(year) + '_evaluations_filtered.xlsx', index=False)