import os
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT



if __name__ == "__main__":
    #filename = "2023-01-TIC1-Test-1.docx"
    #filename = "2023-01-TIC1-Test-2.docx"
    filename = "2023-01-TIC1-Test-3.docx"
    from student import Student
    stud = Student()
    # Créer une instance de l'application Word
    #word_app = win32com.client.Dispatch("Word.Application")

    # Ouvrir le document

    #document = word_app.Documents.Open(document_path)

    print(filename)
    path = os.getcwd()
    filename = path+'/'+filename
    print(filename)

    document = Document(filename)

