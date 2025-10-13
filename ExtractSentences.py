from docx import Document

class ExtractSentences:
    def __init__(self, filename):
        self.sentence_list = []
        self.word_index = {}
        self.doc = Document(filename)
        self.new_doc = Document()

    def create_reverse_index(self):
        for i in range(0, len(self.doc.paragraphs)):
            line = self.doc.paragraphs[i].text

            line = re.sub(r'[^\w\s]', '', line)
            line_list = [word.lower() for word in line.split()]

            for word in line_list:
                if self.word_index.get(word):
                    self.word_index[word].append(i)
                else:
                    self.redis_client.set(word, json.dumps([i]))

    def create_new_doc(self, target_word):
        self.new_doc.add_heading(f"Sentences with the word {target_word}")

        line_numbers = self.word_index.get(target_word.lower())

        for line_number in line_numbers:
            self.new_doc.add_paragraph(self.doc.paragraphs[line_number].text)

    def save_new_doc(self, target_word):
        self.new_doc.save(f"{target_word}.docx")