import os
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

def verifier_entetes_pieds_de_page_word(document, student, key = "piedDePage"):
    # document : pydocx
    # TODO : GROSSES ERREURS dans les points : > au max

    # TODO : ne vérifie pas les alignements ==> à faire manuellement
    max_points = student.max_points[key]
    raison = ""
    to_check_manually = "alignement GMD + en-tête : en haut à droite"
    group = "unknown"
    points = 0
    # for sec in document.sections:
    #    print(sec)
    section = document.sections[0]
    # nbPages = document.page_count
    # TODO : split in 4 functions : function headers
    header = ""
    i = 0
    for h in section.header.paragraphs:
        header += h.text
        i += 1

    if "NPS" in header:
        points += max_points / 4
        group = "PS"
    elif "NP" in header:
        points += max_points / 4
        group = "NP"
    else:
        raison += "Pas de section écrite correctement en en-tête."
        group = "unknown"
    # print("points après headers", points)

    # TODO : split in 4 functions : function footer_middle
    footer = section.footer.paragraphs[0].text
    if "Examen TICE".lower() in footer.lower() and "B1".lower() in footer.lower():
        points += max_points / 4
        # TODO : vérifier qu'il est bien au milieu
    else:
        raison += "Pas de Examen TICE - B1 écrit au milieu du pied de page."
        # print("milieu OK")
    # print("points après pied de page - milieu", points)

    # TODO : split in 4 functions : function footer_right (and 2 sub-functions : page_number et total_pages)
    # footers =  section.Footers
    #modified  if "page" in footer.lower() and ("sur" in footer.lower() or ("de" in footer.lower())):
    #    points += max_points / 4
    #elif "page" in footer.lower():
    #    points += max_points / 8
    #    raison += "Pas de nombre total de pages. "
    #    to_check_manually = "Vérifier nombre de pages."
    #else:
    #    to_check_manually = "Vérifier nombre de pages."
    #    raison += "Pas de numérotation de pages. "
    #    check_page_number_Word(document_pywin32, total_pages)
    # print("points après pied de page - droite", points)

    # TODO : vérifier qu'il est bien à droite

    # TODO : split in 4 functions : function footer_left_name
    if (student.name.lower() in footer.lower()) \
            and ((student.firstname.lower() in footer.lower()) or (student.firstname[0].lower() in footer.lower())):
        points += max_points / 4
        # TODO : vérifier qu'il est bien à gauche"
    else:
        raison += "Le nom ne se trouve pas en pied de page."
    # print("points après pied de page - gauche", points)

    student.scores[key] = points
    student.reasons[key] = raison
    student.to_check_manually += to_check_manually
    if student.scores[key] < student.max_points[key]:
        student.to_check.add(key)
    return group




if __name__ == "__main__":
    #filename = "2023-01-TIC1-Test-1.docx"
    filename = "2023-01-TIC1-Test-2.docx"
    #filename = "2023-01-TIC1-Test-3.docx"
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

    verifier_entetes_pieds_de_page_word(document, stud)
