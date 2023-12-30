import sys
import os
import win32com.client
### Macro's Definition
def define_macros():
    macro_check_links = """
Function ContainsExternalHyperlinkWithDisplayText() As Boolean
    Dim hyperlink As Hyperlink
    Dim fieldCode As String
    Dim displayText As String

    ' Iterate through all hyperlinks in the document
    For Each hyperlink In ActiveDocument.Hyperlinks
        ' Get the field code and display text
        fieldCode = hyperlink.SubAddress
        displayText = hyperlink.TextToDisplay

        ' Check if the hyperlink is an external link and display text is different from the field code
        If hyperlink.Address <> "" And displayText <> fieldCode Then
            ContainsExternalHyperlinkWithDisplayText = True
            Exit Function
        End If
    Next hyperlink

    ' No such hyperlink found
    ContainsExternalHyperlinkWithDisplayText = False
End Function"""
    macro_check_page_num_and_nbr = """
Function VerifierNumeroEtNombrePagesPiedDePage() As Integer
    Dim section As Section
    Dim footer As Range
    Dim pageField As Field
    Dim pageCountField As Field
    Dim numPagesFound As Boolean
    Dim pageCountFound As Boolean

    ' Initialiser les indicateurs
    numPagesFound = False
    pageCountFound = False

    ' Parcourir toutes les sections du document
    For Each section In ActiveDocument.Sections
        ' Accéder au pied de page de la section
        Set footer = section.Footers(wdHeaderFooterPrimary).Range

        ' Réinitialiser les indicateurs pour chaque section
        numPagesFound = False
        pageCountFound = False

        ' Rechercher le champ de numéro de page
        For Each pageField In footer.Fields
            If pageField.Type = wdFieldPage Then
                ' Le champ de numéro de page est présent
                numPagesFound = True
            End If
        Next pageField

        ' Rechercher le champ du nombre total de pages
        For Each pageCountField In footer.Fields
            If pageCountField.Type = wdFieldNumPages Then
                ' Le champ du nombre total de pages est présent
                pageCountFound = True
            End If
        Next pageCountField

        ' Évaluer les indicateurs et renvoyer le résultat approprié
        If numPagesFound And pageCountFound Then
            ' Les deux champs sont présents
            VerifierNumeroEtNombrePagesPiedDePage = 2
            Exit Function
        ElseIf numPagesFound Or pageCountFound Then
            ' Seul l'un des champs est présent
            VerifierNumeroEtNombrePagesPiedDePage = 1
            Exit Function
        End If
    Next section

    ' Aucun champ de numéro de page ou de nombre total de pages trouvé
    VerifierNumeroEtNombrePagesPiedDePage = 0
End Function"""
    macro_tot_page_number = """
Function VerifierNbrePagesTotPiedDePage() As Boolean
    Dim section As Section
    Dim footer As Range
    Dim pageField As Field

    ' Parcourir toutes les sections du document
    For Each section In ActiveDocument.Sections
        ' Accéder au pied de page de la section
        Set footer = section.Footers(wdHeaderFooterPrimary).Range

        ' Rechercher le champ de numéro de page
        For Each pageField In footer.Fields
            If pageField.Type = wdFieldNumPages Then
                ' Le champ de numéro de page est présent
                VerifierNbrePagesTotPiedDePage = True
                Exit Function
            End If
        Next pageField
    Next section

    ' Aucun champ de numéro de page trouvé
    VerifierNbrePagesTotPiedDePage = False
End Function"""

    macro_page_number = """
Function VerifierNombrePagesPiedDePage() As Boolean
    Dim section As Section
    Dim footer As Range
    Dim pageField As Field

    ' Parcourir toutes les sections du document
    For Each section In ActiveDocument.Sections
        ' Accéder au pied de page de la section
        Set footer = section.Footers(wdHeaderFooterPrimary).Range

        ' Rechercher le champ de numéro de page
        For Each pageField In footer.Fields
            If pageField.Type = wdFieldPage Then
                ' Le champ de numéro de page est présent
                VerifierNombrePagesPiedDePage = True
                Exit Function
            End If
        Next pageField
    Next section

    ' Aucun champ de numéro de page trouvé
    VerifierNombrePagesPiedDePage = False
End Function"""

    macro_count_words = """
Function CompterMots() As Long
    ' Stocker le nombre de mots dans une variable
    Dim wordCount As Long
    wordCount = ActiveDocument.Words.Count

    ' Renvoyer le nombre de mots à Python
    CompterMots = wordCount
End Function"""

    macros = [macro_count_words, macro_page_number, macro_tot_page_number, macro_check_page_num_and_nbr, macro_check_links]
    return macros

def print_debug(debug, message):
    if debug:
        print(message)
def check_hyperlinks(word_app, student, key = "lien", debug=False):

    max_points = student.max_points[key]
    why = ""
    to_check_manually = ""
    group = "unknown"
    score = 0

    try:
        links = word_app.Run("ContainsExternalHyperlinkWithDisplayText")
        if links:
            score+=max_points
        else:
            print_debug(debug, "KO ! NON pour les hyperliens avec un texte différent")
            why+="pas d'hyperliens avec un texte différent"
    except Exception as e:
        sys.stderr.write("error in page number and page total word_macros.py\check_hyperlinks " + str(e))
    #finally:
        # Fermer le document en enregistrant les modifications
    #    document.Close(True)

    student.scores[key] = score
    student.reasons[key] = why
    student.to_check_manually += to_check_manually
    if student.scores[key] < student.max_points[key]:
        student.to_check.add(key)
    print_debug(debug, "fin check_hyperlinks "+ str(links))
    return {}

def check_Page_num_and_tot_word(word_app, student, key = "numEtNbrPages", debug=False):

    max_points = student.max_points[key]
    why = ""
    to_check_manually = ""
    group = "unknown"
    score = 0

    try:
        #add_word_macro(document, define_macros())
        pied_de_page = word_app.Run("VerifierNombrePagesPiedDePage")
        if pied_de_page:
            score+=1
        else:
            print_debug(debug, "KO ! NON pour le nombre de page")
            why+="nombre de page non indiqué de manière automatique"
        pied_de_page = word_app.Run("VerifierNbrePagesTotPiedDePage")
        if pied_de_page:
            score+=1
        else:
            print_debug(debug, "KO ! NON pour le nombre de page total")
            why += "nombre total de page non indiqué de manière automatique"
        if debug:
            pts = word_app.Run("VerifierNumeroEtNombrePagesPiedDePage")
            print(pts, "sur 2")
    except Exception as e:
        sys.stderr.write("error in page number and page total word_macors.py\check_page_num_and_tot_word " + str(e))
    #finally:
        # Fermer le document en enregistrant les modifications
    #    document.Close(True)

    student.scores[key] = score
    student.reasons[key] = why
    student.to_check_manually += to_check_manually
    if student.scores[key] < student.max_points[key]:
        student.to_check.add(key)
    print_debug(debug, "fin check_Page_num_and_tot_word")
    return {}

#def add_word_macros_pywin32():
def add_word_macro(document, debug=False):
    macros = define_macros()
    try:
        new_module = document.VBProject.VBComponents.Add(1)  # 1 correspond à vbext_ct_StdModule

        # Ajouter le code de la macro au module
        for m in macros:
            new_module.CodeModule.AddFromString(m)
    except Exception as e:
        print(f"Une erreur s'est produite dans l'ajout de la macro : {e}")
    print_debug(debug, "macros ajoutées")
    #try:
    #    word_count = doc.Run("CompterMots")
    #    print(f"Le nombre de mots dans le document est : {word_count}")
    #except Exception as e:
    #    print(f"Une erreur s'est produite dans le retour de la macro : {e}")
    #finally:
        # doc.Close(True)
    # Fermer l'application Word
    #word_app.Quit()


if __name__ == "__main__":
    filename = "2023-01-TIC1-Test-1.docx"
    #filename = "2023-01-TIC1-Test-2.docx"
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
    # Créer une instance de l'application Word
    word_app = win32com.client.Dispatch("Word.Application")

    # Ouvrir le document
    try:
        document = word_app.Documents.Open(filename)
    except Exception as e :
        for el in e:
            print(el)


    add_word_macro(document)
    check_Page_num_and_tot_word(word_app, stud, debug=True)
    check_hyperlinks(word_app, stud, debug=True)
    #add_word_macro(document, define_macros())
    document.Close(True)
    word_app.Quit()