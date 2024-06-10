import pandas as pd
from jiwer import wer

# Load the Excel file
file_path = '/home/pradyumn/Downloads/whisperdecode(1).xlsx'
df = pd.read_excel(file_path)

# Column indices (0-based index, so 11 for 'L', 12 for 'M')
reference_index = 11
hypothesis_index = 12

# Calculate WER for each row and store it in a new column 'N'
df['N'] = df.apply(lambda row: wer(row.iloc[reference_index], row.iloc[hypothesis_index]), axis=1)

# Save the updated DataFrame back to Excel
df.to_excel(file_path, index=False)
