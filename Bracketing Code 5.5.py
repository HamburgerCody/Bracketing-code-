import docx
import os
import logging
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import simpledialog

logging.basicConfig(level=logging.ERROR)

# ... Other functions ...

def process_text(input_text, word_list):
    paragraphs = input_text.split('\n')
    processed_paragraphs = []

    for paragraph in paragraphs:
        processed_paragraph = add_double_brackets_to_words_in_text(paragraph, word_list)
        processed_paragraphs.append(processed_paragraph)

    result_text = '\n'.join(processed_paragraphs)
    return result_text

# ... Rest of the code ...

def start_application():
    # ... Other GUI setup ...

    def process_text_input():
        word_list = saved_word_list
        input_text = text_box.get("1.0", "end-1c")
        if not input_text:
            messagebox.showerror("Error", "Input text is empty.")
            return

        result_text = process_text(input_text, word_list)
        result_text_box.delete("1.0", tk.END)
        result_text_box.insert(tk.END, result_text)

    process_text_button = tk.Button(app, text="Process Text", fg="white", bg="#007AFF", font=("Arial", 14), command=process_text_input)
    process_text_button.pack(pady=10)

    app.mainloop()

saved_word_list = []

if __name__ == "__main__":
    start_application()
