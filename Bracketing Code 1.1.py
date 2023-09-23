import docx

def read_word_list_from_file(file_path):
    try:
        with open(file_path, 'r') as file:
            word_list = [line.strip() for line in file.readlines()]
            return word_list
    except FileNotFoundError:
        print(f"Error: The file '{file_path}' was not found.")
        return []

def bracket_matching_words(input_text, word_list):
    # Rest of the code remains the same

def process_word_document(docx_filename, word_list, result_filename):
    # Rest of the code remains the same

if __name__ == "__main__":
    # Specify the path of the file containing the word list
    word_list_file = "word_list.txt"
    
    # Read the word list from the file
    word_list = read_word_list_from_file(word_list_file)
    word_set = set(word.lower() for word in word_list)

    # Specify the new path for the input Word document
    docx_filename = r"C:\Users\HamburgerC\Desktop\Bracketing code\Unbracketed.docx"
    
    # Specify the new path for the output processed Word document
    result_filename = r"C:\Users\HamburgerC\Desktop\Bracketing code\Processed_Bracketed.docx"

    # Call the function to process the Word document
    process_word_document(docx_filename, word_set, result_filename)
