import pandas as pd
import jiwer
from jiwer import Compose, RemovePunctuation, RemoveMultipleSpaces, Strip, ToLowerCase

# Custom transformation to normalize Unicode characters
def normalize_unicode(text):
    import unicodedata
    return unicodedata.normalize('NFKC', text)

# Function to split a string into a list of words for the Excel cells
def split_into_cells(sentence, label):
    return [label] + sentence.split('|')

# Function to process and visualize alignments
def visualize_alignments(file_path, output_excel_path, num_sentences=None):
    # Load the Excel file
    df = pd.read_excel(file_path)

    # Ensure the columns are named 'reference' and 'hypothesis'
    reference_sentences = df['reference'].tolist()
    hypothesis_sentences = df['hypothesis'].tolist()

    # If num_sentences is specified, limit the number of sentences to process
    if num_sentences is not None:
        reference_sentences = reference_sentences[:num_sentences]
        hypothesis_sentences = hypothesis_sentences[:num_sentences]

    # Preprocess sentences
    transformation = Compose([
        ToLowerCase(),            # Convert to lowercase
        RemovePunctuation(),      # Remove punctuation
        RemoveMultipleSpaces(),   # Remove multiple spaces
        Strip(),                  # Strip leading/trailing spaces
        lambda s: normalize_unicode(s)  # Normalize Unicode characters
    ])
    hypothesis_sentences = [transformation(sentence) for sentence in hypothesis_sentences]
    reference_sentences = [transformation(sentence) for sentence in reference_sentences]

    # List to store the rows for the new DataFrame
    rows = []

    # Process and visualize alignment for each pair of sentences
    for index, (reference, hypothesis) in enumerate(zip(reference_sentences, hypothesis_sentences)):
        # Compute the alignment
        alignment = jiwer.process_words(reference, hypothesis)
        
        # Visualize the alignment
        alignment_visual = jiwer.visualize_alignment(alignment)
        
        # Identify lines starting with "REF" and take the next two lines
        alignment_lines = alignment_visual.split('\n')
        ref_line, hyp_line, align_line = None, None, None
        for i, line in enumerate(alignment_lines):
            if line.startswith("REF:"):
                ref_line = alignment_lines[i]
                hyp_line = alignment_lines[i + 1]
                align_line = alignment_lines[i + 2]
                break
        
        if ref_line is None or hyp_line is None or align_line is None:
            continue

        # Prepare rows for the new DataFrame
        ref_cells = split_into_cells(ref_line.replace('REF: ', ''), 'REF:')
        hyp_cells = split_into_cells(hyp_line.replace('HYP: ', ''), 'HYP:')
        align_cells = [''] + [cell if cell.strip() else 'C' for cell in align_line.split('|')]
        
        # Add the rows to the list
        rows.append(ref_cells)
        rows.append(hyp_cells)
        rows.append(align_cells)

    # Create a new DataFrame from the rows
    new_df = pd.DataFrame(rows)

    # Save the new DataFrame to an Excel file
    new_df.to_excel(output_excel_path, index=False, header=False)

    print(f"Alignment data saved to {output_excel_path}")

# File paths
file_path = "/home/pradyumn/Downloads/cv-corpus-9.0-2022-04-27-hi/cv-corpus-9.0-2022-04-27/hi/corpus_ref_hyp.xlsx"
output_excel_path = "/home/pradyumn/Downloads/cv-corpus-9.0-2022-04-27-hi/cv-corpus-9.0-2022-04-27/hi/alignment_output_7.xlsx"

# Number of sentences to visualize (set to None for processing all)
num_sentences = None  # Change this to a specific number if needed, e.g., 20

# Run the visualization
visualize_alignments(file_path, output_excel_path, num_sentences)

