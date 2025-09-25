from docx import Document

class ExtractSentences:
    def __init__(self, filename):
        self.sentence_list = []
        self.word_index = {}
        self.doc = Document(filename)

    def create_reverse_index(self):
        for i in range(0, len(self.doc.paragraphs)):
            line = self.doc.paragraphs[i].text

            line_list = line.lower().split()
            for word in line_list:
                if self.word_index.get(word):
                    self.word_index[word].append(i)
                else:
                    self.word_index[word] = [i]

    def create_new_doc(self, target_word):
        new_doc = Document()
        new_doc.add_heading(f"Sentences with the word {target_word}")

        line_numbers = self.word_index.get(target_word.lower())

        for line_number in line_numbers:
            new_doc.add_paragraph(self.doc.paragraphs[line_number].text)

        new_doc.save(f"Sentences with the word {target_word}.docx")