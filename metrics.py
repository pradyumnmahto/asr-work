import openpyxl
import jiwer
from jiwer import Compose, ToLowerCase, RemovePunctuation, RemoveMultipleSpaces, Strip

# Define the file path
file_path = '/home/pradyumn/Documents/whisper_decode_wer(Copy).xlsx'

# Load the Excel workbook and select the sheet
wb = openpyxl.load_workbook(file_path)
sheet1 = wb['Sheet1']
sheet2 = wb['Sheet2']

# Define the cell ranges for the reference and hypothesis sentences
ranges = [
    (2, 6),
    (7, 11),
    (12, 16),
    (17, 21),
    (22, 26),
    (27, 31),
    (32, 36)
]

# Lists to store the WER results and error metrics
wer_results = []
substitutions_results = []
insertions_results = []
deletions_results = []

# Helper function to get cell values from a column within a specified range
def get_cell_values(sheet, column, start_row, end_row):
    return [sheet[f'{column}{row}'].value for row in range(start_row, end_row + 1)]

# Define the transformation to remove punctuation
transform = Compose([
    ToLowerCase(),
    RemovePunctuation(),
    RemoveMultipleSpaces(),
    Strip()
])

# Helper function to preprocess text using the transformation
def preprocess(text_list):
    return [transform(text) for text in text_list]

# Calculate WER and error metrics for each range
for start, end in ranges:
    reference = preprocess(get_cell_values(sheet1, 'L', start, end))
    hypothesis = preprocess(get_cell_values(sheet1, 'M', start, end))
    
    # Compute measures after preprocessing
    measures = jiwer.compute_measures(" ".join(reference), " ".join(hypothesis))
    
    wer_results.append(measures['wer'])
    substitutions_results.append(measures['substitutions'])
    insertions_results.append(measures['insertions'])
    deletions_results.append(measures['deletions'])

# Write the WER results and error metrics to Sheet2 at the specified cells
for i, (wer, subs, ins, dels) in enumerate(zip(wer_results, substitutions_results, insertions_results, deletions_results), start=2):
    sheet2[f'C{i}'] = wer
    sheet2[f'D{i}'] = subs
    sheet2[f'E{i}'] = ins
    sheet2[f'F{i}'] = dels

# Save the workbook
wb.save(file_path)

print("WER results and error metrics written to Sheet2 in the specified cells.")

