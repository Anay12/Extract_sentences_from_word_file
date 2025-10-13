import win32com.client
from spire.doc import *
from spire.doc.common import *

def convert_docx_to_doc(input_path, output_path):
    # word_app = win32com.client.gencache.EnsureDispatch("Word.Application")
    # doc = word_app.Documents.Open(input_path)
    # doc.SaveAs(output_path, FileFormat=0)
    # doc.Close()
    # word_app.Quit()

    document = Document()
    document.LoadFromFile(input_path)
    document.SaveToFile(output_path, FileFormat.Doc)
    document.Close()



# class ConvertDocxToDoc:
#     def __init__(self, input_path, output_path):
#         self.input_path = input_path
#         self.output_path = output_path

    # @staticmethod





