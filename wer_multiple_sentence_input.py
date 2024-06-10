import pandas as pd
import openpyxl
import jiwer

# Define the file path
file_path = '/home/pradyumn/Documents/whisper_decode_wer.xlsx'

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

# List to store the WER results
wer_results = []

# Helper function to get cell values from a column within a specified range
def get_cell_values(sheet, column, start_row, end_row):
    return [sheet[f'{column}{row}'].value for row in range(start_row, end_row + 1)]

# Calculate WER for each range
for start, end in ranges:
    reference = get_cell_values(sheet1, 'L', start, end)
    hypothesis = get_cell_values(sheet1, 'M', start, end)
    overall_wer = jiwer.wer(" ".join(reference), " ".join(hypothesis))
    wer_results.append(overall_wer)

# Write the WER results to Sheet2 at the specified cells (C2 to C8)
for i, wer in enumerate(wer_results, start=2):
    sheet2[f'C{i}'] = wer

# Save the workbook
wb.save(file_path)

print("WER results written to Sheet2 in the specified cells.")
