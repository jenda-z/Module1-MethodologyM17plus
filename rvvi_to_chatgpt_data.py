import pandas as pd
import docx 
import os

year = 24

# Load the Excel file
df = pd.read_excel('24_evaluations_filtered.xlsx')

oldColumns = ['Druh ', 'Kritérium', 'Autoři', 'Název výsledku', 'Obor (Ford)', 'Zdůvodnění', 'Finální známka']
newColumns = [ 'Text1', 'Text2', 'Text3', 'Text4']

filtered_df = pd.DataFrame(columns = oldColumns + newColumns)

print(filtered_df)

df = df[oldColumns + ['Text']]
results = df['Název výsledku'].unique()

for res in results:
    res_df = df[df['Název výsledku'] == res].dropna().drop_duplicates()

    # Add result justification
    fileName = res_df[['Zdůvodnění']].dropna().values[0][0]
    doc = docx.Document(fileName)
    
    fullText = ''    
    fullText = '\n'.join([para.text for para in doc.paragraphs])

    res_df['Zdůvodnění'] = fullText

    # Add grade justifications
    texts = df[df['Název výsledku'] == res]['Text'].tolist()
    texts = (texts + [''] * 4)[:4] # Pad or trim the list to exactly 4 elements

    for i, col in enumerate(newColumns):
        res_df[col] = texts[i]

    res_df = res_df.drop(columns=['Text']) # Delete the extra 'Text' column in res_df

    filtered_df = pd.concat([filtered_df, res_df], ignore_index=True)

# Save the filtered data to a new Excel file
filtered_df.to_excel('24_fce_evaluations_chatgpt.xlsx', index=False)