from ExtractSentences import ExtractSentences


if __name__ == '__main__':
    continue_bool = 'Yes'

    # while continue_bool.lower() != "no":
    search_words = input('Enter the words you want to search for: ')

    extractSentences = ExtractSentences('slogans2025.docx')
    extractSentences.create_reverse_index()

    for word in search_words.split():
        extractSentences.create_new_doc(word)

    extractSentences.save_new_doc(search_words)
        # continue_bool = input("Do you want to continue: ")