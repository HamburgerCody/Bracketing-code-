# Function to process the document
def process_document():
    global saved_word_list
    word_list = text_box.get("1.0", "end-1c").splitlines()
    if not word_list:
        messagebox.showerror("Error", "Word list is empty.")
        return

    input_filename = filedialog.askopenfilename(filetypes=[("Word Document", "*.docx")])
    if not input_filename:
        return

    output_filename = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
    if not output_filename:
        return

    # Number of processing iterations
    num_iterations = 3  # You can change this to 2 or 3

    for _ in range(num_iterations):
        # Process the Word document and display a message based on success or failure
        if process_word_document(input_filename, word_list, output_filename):
            messagebox.showinfo("Success", f"Processing completed successfully. Result saved as '{output_filename}'")
        else:
            messagebox.showerror("Error", "Processing failed.")
