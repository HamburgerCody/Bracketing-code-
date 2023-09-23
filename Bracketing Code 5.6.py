import docx
import os
import logging
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import simpledialog  # For user input

# Configure logging without creating a log file
logging.basicConfig(level=logging.ERROR)

# Initialize the saved word list as empty
saved_word_list = []

def add_double_brackets_to_words_in_text(input_text, word_list):
    words = input_text.split()
    result = []

    for word in words:
        # Check if the word matches any word in the provided list (case-insensitive)
        if any(word.lower() == item.lower() for item in word_list):
            result.append(f"[[{word}]]")  # Double brackets
        else:
            result.append(word)

    result_text = ' '.join(result)
    return result_text

def process_text(input_text, word_list):
    # Process the input text
    processed_text = add_double_brackets_to_words_in_text(input_text, word_list)
    return processed_text

# Create the main application window
app = tk.Tk()
app.title("Word Processor")
app.configure(bg="#808080")  # Set the background color to gray

# Create and configure widgets
label = tk.Label(app, text="Word List:", fg="white", bg="#808080", font=("Arial", 14))
text_box = tk.Text(app, height=10, width=40, font=("Arial", 12))
update_word_list_button = tk.Button(app, text="Update Word List", fg="white", bg="#007AFF", font=("Arial", 14), command=update_word_list)
process_button = tk.Button(app, text="Process", fg="white", bg="#007AFF", font=("Arial", 14), command=process_text_and_display)

output_label = tk.Label(app, text="Processed Text:", fg="white", bg="#808080", font=("Arial", 14))
processed_text_box = tk.Text(app, height=10, width=40, font=("Arial", 12))

# Place widgets on the window
label.pack(pady=20)
text_box.pack(padx=20, pady=10)
update_word_list_button.pack(pady=10)
process_button.pack(pady=10)
output_label.pack(pady=20)
processed_text_box.pack(padx=20, pady=10)

# Function to process the text and display the result
def process_text_and_display():
    global saved_word_list
    word_list = text_box.get("1.0", "end-1c").splitlines()
    if not word_list:
        messagebox.showerror("Error", "Word list is empty.")
        return

    input_text = text_box.get("1.0", "end-1c")
    processed_text = process_text(input_text, word_list)
    processed_text_box.delete("1.0", tk.END)
    processed_text_box.insert("1.0", processed_text)

# Function to update word list
def update_word_list():
    global saved_word_list
    saved_word_list = text_box.get("1.0", "end-1c").splitlines()

# Start the GUI application
app.mainloop()
