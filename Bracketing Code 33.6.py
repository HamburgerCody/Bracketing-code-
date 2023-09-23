import docx
import os
import tkinter as tk
from tkinter import filedialog, messagebox, Text
import re
import logging

# Initialize the saved word list as empty
saved_word_list = []

# Variable to track whether processing should occur
processing_enabled = False

# Function to add double brackets to words in text (with a maximum of two brackets)
def add_double_brackets_to_words_in_text(input_text, word_list):
    result_text = input_text

    for word in word_list:
        # Check if the word already contains double brackets
        if f"[[{word}]]" not in result_text:
            # Use a case-sensitive replacement
            result_text = result_text.replace(word, f"[[{word}]]")
        else:
            # If double brackets are already present, replace any occurrences of triple brackets with double brackets
            result_text = result_text.replace(f"[[[{word}]]]", f"[[{word}]]")

    return result_text

# Function to clean up extra brackets beside the two that should be there
def clean_up_extra_brackets(input_text):
    cleaned_text = input_text

    # Use a regular expression to remove extra brackets beside the two that should be there
    cleaned_text = re.sub(r'\[\[+([^\[\]]+)\]\]+', r'[[\1]]', cleaned_text)

    return cleaned_text

# Function to process a Word document
def process_word_document(input_filename, word_list, result_folder, num_iterations=3, overwrite=False):
    try:
        # Load the Word document
        doc_text = extract_text_from_docx(input_filename)

        # Initialize an empty list to store the processed paragraphs
        processed_paragraphs = []

        for _ in range(num_iterations):
            # Process the document text
            processed_text = add_double_brackets_to_words_in_text(doc_text, word_list)
            processed_text = clean_up_extra_brackets(processed_text)
            processed_paragraphs.append(processed_text)

        result_text = '\n'.join(processed_paragraphs)

        # Get the original file name without the extension
        original_file_name = os.path.splitext(os.path.basename(input_filename))[0]

        # Create a sanitized output file name by replacing problematic characters and appending "_Processed"
        output_filename = re.sub(r'[\/:*?"<>|]', '_', original_file_name) + "_Processed.docx"

        # Ensure the output folder exists
        os.makedirs(result_folder, exist_ok=True)

        full_output_path = os.path.join(result_folder, output_filename)

        # Check if the file already exists and if not set to overwrite, ask for confirmation
        if os.path.exists(full_output_path) and not overwrite:
            confirm = messagebox.askyesno("File Exists", f"The file '{full_output_path}' already exists. Do you want to overwrite it?")
            if not confirm:
                logging.error("Processing canceled.")
                return False

        # Create a new Word document to store the result
        result_doc = docx.Document()
        result_doc.add_paragraph(result_text)

        # Save the result to the specified Word document
        result_doc.save(full_output_path)

        return True
    except Exception as e:
        logging.error(f"Error: An error occurred while processing the document. Error: {str(e)}")
        return False

# Function to check for and limit the number of brackets around words (maximum of two brackets)
def limit_brackets(word_list):
    for i in range(len(word_list)):
        word = word_list[i]
        # Use a while loop to remove extra brackets until only a maximum of two brackets remain
        while "[[" in word:
            word = word.replace("[[", "[")
        while "]]" in word:
            word = word.replace("]]", "]")
        word_list[i] = word

# Function to process a folder
def process_folder(input_folder, output_folder):
    try:
        global processing_enabled  # Access the global variable

        if processing_enabled:  # Check if processing is enabled
            # Iterate over files in the input folder
            for root, _, files in os.walk(input_folder):
                for file in files:
                    file_path = os.path.join(root, file)
                    if file.endswith(".docx"):
                        # Process the Word document and display a message based on success or failure
                        if process_word_document(file_path, saved_word_list, output_folder, overwrite=True):
                            print(f"Processing completed successfully for: {file_path}")
                        else:
                            print(f"Processing failed for: {file_path}")

    except Exception as e:
        logging.error(f"Error: An error occurred while processing the folder. Error: {str(e)}")

# Function to update word list
def update_word_list(text_box):
    global saved_word_list, processing_enabled
    saved_word_list = text_box.get("1.0", "end-1c").splitlines()
    # Check and limit the number of brackets around words
    limit_brackets(saved_word_list)

    # Enable processing if there are words in the list
    if saved_word_list:
        processing_enabled = True
    else:
        processing_enabled = False
        messagebox.showwarning("Warning", "Word list is empty. Processing is disabled.")

# Function to save the word list to a text file
def save_word_list(text_box):
    global saved_word_list
    file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
    if file_path:
        with open(file_path, "w") as file:
            for word in saved_word_list:
                file.write(word + "\n")

# Function to load the word list from a text file
def load_word_list(text_box):
    global saved_word_list
    file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
    if file_path:
        with open(file_path, "r") as file:
            saved_word_list = [line.strip() for line in file]
            # Check and limit the number of brackets around words
            limit_brackets(saved_word_list)
            update_word_list(text_box)  # Update the text box with the loaded word list

# Function to extract text from a .txt file
def extract_text_from_txt(txt_filename):
    try:
        with open(txt_filename, "r", encoding="utf-8") as txt_file:
            text = txt_file.read()
        return text
    except Exception as e:
        logging.error(f"Error: Unable to read the text file. Error: {str(e)}")
        return ""

# Function to load words from a .txt file
def load_words_from_txt(txt_filename):
    try:
        with open(txt_filename, "r", encoding="utf-8") as txt_file:
            word_list = [line.strip() for line in txt_file.readlines()]
        return word_list
    except Exception as e:
        logging.error(f"Error: Unable to read the word list from the text file. Error: {str(e)}")
        return []

# Function to start the GUI application
def start_gui():
    # Create the main application window
    app = tk.Tk()
    app.title("Word Document Processor")

    # Create and configure widgets
    text_box = Text(app, height=10, width=40, font=("Arial", 12))
    text_box.pack(padx=20, pady=10)

    button_frame = tk.Frame(app)
    button_frame.pack(pady=20)

    # Button to update the word list
    update_word_list_button = tk.Button(button_frame, text="Update Word List", font=("Arial", 14),
                                       command=lambda: update_word_list(text_box))
    update_word_list_button.grid(row=1, column=0, columnspan=2, pady=10)

    # Buttons to save and load the word list
    save_word_list_button = tk.Button(button_frame, text="Save Word List", font=("Arial", 14),
                                      command=lambda: save_word_list(text_box))
    save_word_list_button.grid(row=2, column=0, pady=5)

    load_word_list_button = tk.Button(button_frame, text="Load Word List", font=("Arial", 14),
                                      command=lambda: load_word_list(text_box))
    load_word_list_button.grid(row=2, column=1, pady=5)

    # Buttons to select a document and select a folder
    select_document_button = tk.Button(button_frame, text="Select Document", font=("Arial", 14),
                                       command=select_document)
    select_document_button.grid(row=0, column=0, pady=5)

    select_folder_button = tk.Button(button_frame, text="Select Folder", font=("Arial", 14),
                                     command=select_folder)
    select_folder_button.grid(row=0, column=1, pady=5)

    # Label to display the current working directory
    global current_directory_label
    current_directory_label = tk.Label(app, text="", font=("Arial", 10))
    current_directory_label.pack()

    # Start the GUI application
    app.mainloop()

# Initialize the logging configuration
logging.basicConfig(filename='processing.log', level=logging.ERROR)

# Start the GUI application
start_gui()
