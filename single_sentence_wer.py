import pandas as pd
from jiwer import wer, Compose, ToLowerCase, RemovePunctuation, RemoveMultipleSpaces, Strip

# Load the Excel file
file_path = '/home/pradyumn/Downloads/whisperdecode(1).xlsx'
df = pd.read_excel(file_path)

# Column indices (0-based index, so 11 for 'L', 12 for 'M')
reference_index = 11
hypothesis_index = 12

# Define the transformation to remove punctuation
transform = Compose([
    ToLowerCase(),
    RemovePunctuation(),
    RemoveMultipleSpaces(),
    Strip()
])

# Function to calculate WER with preprocessing
def calculate_wer(reference, hypothesis):
    transformed_reference = transform(reference)
    transformed_hypothesis = transform(hypothesis)
    return wer(transformed_reference, transformed_hypothesis)

# Calculate WER for each row and store it in a new column 'N'
df['N'] = df.apply(lambda row: calculate_wer(row.iloc[reference_index], row.iloc[hypothesis_index]), axis=1)

# Save the updated DataFrame back to Excel
df.to_excel(file_path, index=False)
