def read_word_list_from_txt(txt_filename):
    try:
        with open(txt_filename, 'r', encoding='utf-8') as file:
            word_list = [line.strip() for line in file.readlines()]
        return word_list
    except FileNotFoundError:
        print(f"Error: The file '{txt_filename}' was not found.")
        return []

if __name__ == "__main__":
    while True:
        # Prompt the user to input the path to the word list text file
        word_list_path = input("Enter the path to the word list text file (e.g., C:\\path\\to\\wordlist.txt): ").strip()

        # Read the word list from the text file
        word_list = read_word_list_from_txt(word_list_path)

        if not word_list:
            print("Word list is empty or could not be read.")
            continue

        # ... Rest of your script ...
