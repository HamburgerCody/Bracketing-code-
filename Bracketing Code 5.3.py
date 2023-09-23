import docx
import os
import logging
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import simpledialog

logging.basicConfig(level=logging.ERROR)

# ... Other functions ...

def process_word_document(docx_filename, word_list, result_filename):
    try:
        doc_text = extract_text_from_docx(docx_filename)
    except Exception as e:
        logging.error(f"Error: Unable to read the Word document. Error: {str(e)}")
        return False

    paragraphs = doc_text.split('\n')
    processed_paragraphs = []

    for paragraph in paragraphs:
        processed_paragraph = add_double_brackets_to_words_in_text(paragraph, word_list)
        processed_paragraphs.append(processed_paragraph)

    result_text = '\n'.join(processed_paragraphs)

    # Create a new document to store the result
    result_doc = docx.Document()
    result_doc.add_paragraph(result_text)

    # Determine the full output path
    full_output_path = os.path.join(result_filename)

    # Check if the file already exists and prompt for overwrite
    if os.path.exists(full_output_path):
        confirm = messagebox.askyesno("File Exists", f"The file '{full_output_path}' already exists. Do you want to overwrite it?")
        if not confirm:
            print("Processing canceled.")
            return False

    try:
        # Save the result to a temporary Word document first
        temp_result_filename = os.path.splitext(result_filename)[0] + "_temp.docx"
        result_doc.save(temp_result_filename)

        # Remove the original file if it exists
        if os.path.exists(result_filename):
            os.remove(result_filename)

        # Rename the temporary file to the desired output filename
        os.rename(temp_result_filename, result_filename)

        return True
    except Exception as e:
        logging.error(f"Error: Unable to save the Word document. Error: {str(e)}")
        return False

# ... Rest of the code ...
