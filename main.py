from ExtractSentences import ExtractSentences
from ConvertDocxToDoc import convert_docx_to_doc

if __name__ == '__main__':
    continue_bool = 'Yes'

    # while continue_bool.lower() != "no":
    search_words = input('Enter the words you want to search for: ')

    extractSentences = ExtractSentences('slogans2025.docx')
    extractSentences.create_reverse_index()

    for word in search_words.split():
        extractSentences.create_new_doc(word)

    extractSentences.save_new_doc(search_words)
    # convert_docx_to_doc(f"{search_words}.docx", f"{search_words}.doc")
        # continue_bool = input("Do you want to continue: ")