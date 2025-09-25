from ExtractSentences import ExtractSentences


if __name__ == '__main__':
    extractSentences = ExtractSentences('slogans2025.docx')
    extractSentences.create_reverse_index()
    extractSentences.create_new_doc('God')