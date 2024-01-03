import sys
import os
import win32com.client

from tkinter import messagebox
import time

import re
remove_non_english = lambda s: re.sub(r'[^a-zA-Z0-9]', '', s)

### Macro's Definition
def define_macros():
    macros=[]
    macros.append("""
Function HasTable() As Boolean
    Dim doc As Document
    Dim tbl As Table

    ' Reference to the active document
    Set doc = ActiveDocument

    ' Check if there is at least one table in the document
    If doc.Tables.Count > 0 Then
        HasTable = True
    Else
        HasTable = False
    End If
End Function""")
    macros.append("""
Function GetFooterText() As String
    Dim footerText As String
    Dim numSections As Integer

    ' Get the number of sections in the document
    numSections = ActiveDocument.Sections.Count

    ' Check if the document has at least one section
    If numSections < 1 Then
        GetFooterText = "The document does not have any sections."
        Exit Function
    End If

    ' Access the footer based on the DifferentFirstPageHeaderFooter setting
    If ActiveDocument.PageSetup.DifferentFirstPageHeaderFooter Then
        ' DifferentFirstPageHeaderFooter is True, consider the second section
        If numSections >= 2 Then
            ' Retrieve the footer of the second section
            footerText = ActiveDocument.Sections(2).Footers(wdHeaderFooterPrimary).Range.Text
        Else
            ' There is no second section, consider the first section
            footerText = ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Text
        End If
    Else
        ' DifferentFirstPageHeaderFooter is False, consider the first section
        footerText = ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Text
    End If

    ' Set the result to the footer text
    GetFooterText = footerText
End Function""")
    macros.append("""
Function GetRightAlignedHeaderText() As String
    Dim headerText As String
    Dim paragraph As Paragraph

    ' Check if the document has at least one section
    If ActiveDocument.Sections.Count < 1 Then
        GetRightAlignedHeaderText = "The document does not have any sections."
        Exit Function
    End If

    ' Access the header of the first section
    Set paragraph = ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range.Paragraphs(1)

    ' Check if the paragraph is right-aligned
    If paragraph.Alignment = wdAlignParagraphRight Then
        ' Get the text from the right-aligned paragraph
        headerText = paragraph.Range.Text
    Else
        headerText = "No right-aligned text found in the header."
    End If

    ' Set the result to the header text
    GetRightAlignedHeaderText = headerText
End Function""")
    macros.append("""
Function GetHeaderOfFirstSection() As String
    Dim headerText As String

    ' Check if the document has at least one section
    If ActiveDocument.Sections.Count < 1 Then
        GetHeaderOfFirstSection = "The document does not have any sections."
        Exit Function
    End If

    ' Access the header of the first section
    headerText = ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range.Text

    ' Set the result to the header text
    GetHeaderOfFirstSection = headerText
End Function""")
    macros.append("""
Function EvaluateHyperlinkConditions() As Double
    Dim hyperlink As Hyperlink
    Dim fieldCode As String
    Dim displayText As String
    Dim hasExternalLink As Boolean
    Dim hasSameDisplayAndFieldCode As Boolean

    ' Initialize flags
    hasExternalLink = False
    hasSameDisplayAndFieldCode = False

    ' Iterate through all hyperlinks in the document
    For Each hyperlink In ActiveDocument.Hyperlinks
        ' Get the field code and display text
        fieldCode = hyperlink.SubAddress
        displayText = hyperlink.TextToDisplay

        ' Check conditions
        If hyperlink.Address <> "" Then
            hasExternalLink = True
            If displayText = fieldCode Then
                hasSameDisplayAndFieldCode = True
            End If
        End If
    Next hyperlink

    ' Evaluate and return the result
    If hasExternalLink And Not hasSameDisplayAndFieldCode Then
        EvaluateHyperlinkConditions = 2
    ElseIf hasExternalLink Then
        EvaluateHyperlinkConditions = 0.5
    Else
        EvaluateHyperlinkConditions = 0
    End If
End Function
""")
    # TODO : check why Test-5 is working : spaces to align
    # TODO : check why Test-6 is not working : number of pages not found (1st page ? )
    # TODO : check why Test-7 is not working : floating page number (not in footer ?)
    macros.append( """
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
End Function""")
    macros.append( """
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
End Function""")

    macros.append("""
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
End Function""")

    macros.append( """
Function CompterMots() As Long
    ' Stocker le nombre de mots dans une variable
    Dim wordCount As Long
    wordCount = ActiveDocument.Words.Count

    ' Renvoyer le nombre de mots à Python
    CompterMots = wordCount
End Function""")

    return macros

def print_debug(debug, message):
    if debug:
        print(message)

def check_footer_left_text(student, complete_footer, left_footer, max_score, debug=False):
    score = 0
    why = ""
    to_check_manually=""
    if (student.name.lower() in left_footer.lower()) :#\
            #and ((student.firstname.lower() in left_footer.lower()) or (
            #student.firstname[0].lower() in footer_text[0].lower())):
        score = max_score / 4
        print_debug(debug, "name " + student.name + "in left footer "+str(max_score))

    elif (student.name.lower() in complete_footer.lower()): # \
            #and ((student.firstname.lower() in footer.lower()) or (student.firstname[0].lower() in footer.lower())):
        score = max_score / 8
        print_debug(debug, "name " + student.name + "in footer but not on the left")
        to_check_manually = "Nom en bas à gauche ? si oui, +" + str(max_score / 8)
        # TODO : vérifier qu'il est bien à gauche"
    else:
        why = "Le nom ne se trouve pas en pied de page."
        to_check_manually = "Nom en bas à gauche ? "
        print_debug(debug, "No name in footer")
    return(score, why, to_check_manually)


def check_footer_middle_text(complete_footer, footer_middle, middle_text, max_score, debug=False):
    """ Only [A-Za-z-0-9] chars will be check because some dash are automatically replaced by En dash, or Em dash or...
    and some students put more than one space."""
    score = 0
    to_check_manually = ""
    why=""
    mt =remove_non_english(middle_text).lower()
    ft = remove_non_english(footer_middle).lower()
    cf = remove_non_english(complete_footer).lower()

    if mt in ft :
        score = max_score / 4
        print_debug(debug, "middle_text OK")
        # TODO : vérifier qu'il est bien au milieu
    elif mt in cf :
        score = max_score / 8
        print_debug(debug, "middle_text OK mais pas vraiment au milieu")
        # TODO : vérifier qu'il est bien au milieu
        why = middle_text+" écrit, mais pas au milieu du pied de page."
        to_check_manually = "vérifier le pied de page (milieu)"
        print_debug(debug, middle_text+" in footer but not center")
    else :
        why = "Pas de "+middle_text+" écrit au milieu du pied de page."
        to_check_manually = "vérifier le pied de page (milieu)"
        print_debug(debug, "No "+middle_text+" in footer")
    print_debug(debug, "middle_text score ="+str(score)+"/"+str(max_score))
    return(score, why, to_check_manually)

def check_Page_num_and_tot_word(word_app, student, key = "piedDePage", debug=False):

    max_points = student.max_points[key]
    why = ""
    to_check_manually = ""
    score = 0

    try:
        #add_word_macro(document, define_macros())
        pied_de_page = word_app.Run("VerifierNombrePagesPiedDePage")
        if pied_de_page:
            score+=max_points/8
        else:
            print_debug(debug, "KO ! NON pour le nombre de page")
            why+="nombre de page non indiqué de manière automatique"
        pied_de_page = word_app.Run("VerifierNbrePagesTotPiedDePage")
        if pied_de_page:
            score+=max_points/8
        else:
            print_debug(debug, "KO ! NON pour le nombre de page total")
            why += "nombre total de page non indiqué de manière automatique"
        print_debug(debug, "nbr_tot_pages : "+str(score)+"/"+str(max_points))
    except Exception as e:
        sys.stderr.write("error in page number and page total word_macors.py\check_page_num_and_tot_word " + str(e))
    print_debug(debug, "fin check_Page_num_and_tot_word")
    return (score, why, to_check_manually)
def get_footer_parts(complete_footer, debug):
    left_footer=""
    middle_footer=""
    #TODO check Test 5 : why is it working ? split by spaces but retrieve points ?
    if "\t" in complete_footer:
        footer_text = complete_footer.split("\t")
        if len(footer_text) >= 3:
            left_footer = footer_text[0]
            middle_footer = footer_text[1]
            right_footer = footer_text[2]
        elif len(footer_text) == 2:
            left_footer = footer_text[0]
            middle_footer = footer_text[1]
    elif "0" in complete_footer:  # in some cases, The "tab" characters ar returned as 0 !? Ex : Test3.docx
        footer_text = complete_footer.split("0")
        if len(footer_text) >= 3:
            left_footer = footer_text[0]
            middle_footer = footer_text[1]
            right_footer = footer_text[2]
        elif len(footer_text) == 2:
            left_footer = footer_text[0]
            middle_footer = footer_text[1]
    return(left_footer, middle_footer)

def check_right_header(word_app, student, key = "piedDePage", debug=False):
    # TODO : adapt scores and heading text (make it generic/parameter ?)
    # TODO : how to be OK if the position come from 1 or 2 tabs ?
    max_scores = student.max_points[key]
    why = ""
    to_check_manually = ""
    group = "unknown"
    score = 0
    links = ""

    try:
        header_text = word_app.Run("GetRightAlignedHeaderText")
        if "NPS" in header_text:
            score += max_scores / 4
            group = "PS"
            print_debug(debug, "NPS OK")
        elif "NP" in header_text:
            score += max_scores / 4
            group = "NP"
            print_debug(debug, "NP OK")
        else:
            why += "Pas de section écrite correctement en en-tête."
            print_debug(debug, "Pas de section écrite correctement en en-tête.")

            group = "unknown"

    except Exception as e:
        sys.stderr.write("error in word_macros.py\check_header " + str(e))

    student.scores[key] = score
    student.reasons[key] = why
    student.to_check_manually += to_check_manually
    if student.scores[key] < student.max_points[key]:
        student.to_check.add(key)
    print_debug(debug, "fin check_header "+ str(links))
    return {}
def check_header(word_app, student, key = "piedDePage", debug=False):
    # TODO : adapter les points et le texte des en-têtes (généraliser ?)
    max_scores = student.max_points[key]
    why = ""
    to_check_manually = ""
    group = "unknown"
    score = 0

    try:
        header_text = word_app.Run("GetHeaderOfFirstSection")
        if "NPS" in header_text:
            score += max_scores / 4
            group = "PS"
            print_debug(debug, "NPS OK")
        elif "NP" in header_text:
            score += max_scores / 4
            group = "NP"
            print_debug(debug, "NP OK")
        else:
            why += "Pas de section écrite correctement en en-tête."
            print_debug(debug, "Pas de section écrite correctement en en-tête.")

            group = "unknown"

    except Exception as e:
        sys.stderr.write("error in word_macros.py\check_header " + str(e))

    print_debug(debug, "check_header "+ str(score)+"/"+str(max_scores)+" group : "+group)
    return (score, why, to_check_manually, group)
def check_header_and_footer(word_app, student, middle_text_asked="S2-B1 - Numérique", key = "piedDePage", debug=False):
    # TODO : adapt scores and heading text (make it generic/parameter ?)
    # TODO : how to be OK if the position come from 1 or 2 tabs ?
    max_scores = student.max_points[key]
    why = ""
    to_check_manually = ""
    score = 0
    footer_text=[]
    right_footer=""
    # ---------------------------- Header ------------------------------
    (sc, wh, tch, group) = check_header(word_app, student, key, debug)
    score += sc
    why += wh
    to_check_manually += tch
    # ---------------------------- Footer ------------------------------
    try:
        complete_footer = word_app.Run("GetFooterText")
        print_debug(debug, repr(complete_footer))
        # split parts of the footer if possible
        # split by tab or 0
        (left_footer, middle_footer) = get_footer_parts(complete_footer, debug)

        print_debug(debug, footer_text)
        print_debug(debug, "Name = "+student.name)
        print_debug(debug, "Firstname = "+student.firstname)

        # name in left footer
        (sc, wh, tch) = check_footer_left_text(student, complete_footer, left_footer, max_scores, debug)
        score += sc
        why += wh
        to_check_manually += tch

        # check middle text.
        (sc, wh, tch) = check_footer_middle_text(complete_footer, middle_footer, middle_text_asked, max_scores, debug)
        score += sc
        why += wh
        to_check_manually += tch

        # The right part is done with "check_Page_num_and_tot_word" function
        (sc, wh, tch) = check_Page_num_and_tot_word(word_app, student,key, debug)
        score += sc
        why += wh
        to_check_manually += tch

    except Exception as e:
        sys.stderr.write("error in word_macros.py\check_header_and_footer " + str(e))

    student.scores[key] = score
    student.reasons[key] = why
    student.to_check_manually += to_check_manually
    if student.scores[key] < student.max_points[key]:
        student.to_check.add(key)
    print_debug(debug, "header_and_footer : "+str(score)+"/"+str(max_scores))
    print_debug(debug, "fin check_header_and_footer ")
    return (group)


def check_hyperlinks(word_app, student, key = "lien", debug=False):

    max_scores = student.max_points[key]
    why = ""
    to_check_manually = ""
    score = 0
    links = ""

    try:
        score = word_app.Run("EvaluateHyperlinkConditions")
        if score == 0.5:
            print_debug(debug, "KO ! hyperlien présent, mais sans un texte différent. ")
            why += "hyperlien présent, mais sans un texte différent. "
        elif score <= 0:
            print_debug(debug, "KO ! NON pour les hyperliens avec un texte différent")
            why+="pas d'hyperliens avec un texte différent"
        elif score == max_scores:
            print_debug(debug, "Tout va bien dans les hyperliens")
        else:
            print_debug("choses étrange dans check_hyperlink")
    except Exception as e:
        sys.stderr.write("error in page number and page total word_macros.py\check_hyperlinks " + str(e))

    student.scores[key] = score
    student.reasons[key] = why
    student.to_check_manually += to_check_manually
    if student.scores[key] < student.max_points[key]:
        student.to_check.add(key)
    print_debug(debug, "fin check_hyperlinks "+ str(links))
    return {}
def check_tables(word_app, student, key = "tableau", debug=False):

    max_scores = student.max_points[key]
    why = ""
    to_check_manually = ""
    score = 0

    try:
        table = word_app.Run("HasTable")
        if table:
            print_debug(debug, "OK, Tableau présent. ")
            score=max_scores
        else :
            print_debug(debug, "pas de tableau")
            why += "pas de tableau dans le document. "
            to_check_manually+="vérifier tableau - "
    except Exception as e:
        sys.stderr.write("error in word_macros.py\check_tables " + str(e))

    student.scores[key] = score
    student.reasons[key] = why
    student.to_check_manually += to_check_manually
    if student.scores[key] < student.max_points[key]:
        student.to_check.add(key)
    print_debug(debug, "fin check_tables ")
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
    global debug
    debug = True
    filename = "2023-01-TIC1-Test-1.docx"
    #filename = "2023-01-TIC1-Test-2.docx"
    #filename = "2023-01-TIC1-Test-3.docx"
    #filename = "2023-01-TIC1-Test-4.docx"
    filename = "2023-01-TIC1-Test-6.docx"
    from student import Student
    stud = Student()
    stud.name="Test"
    stud.firstname="1"
    # Créer une instance de l'application Word
    #word_app = win32com.client.Dispatch("Word.Application")

    # Ouvrir le document

    #document = word_app.Documents.Open(document_path)

    print(filename)
    file=filename
    path = os.getcwd()
    filename = path+'/'+filename
    print(filename)
    # Créer une instance de l'application Word
    word_app = win32com.client.Dispatch("Word.Application")


    command_executed=False
    error=False
    while not command_executed:
        if os.path.exists(file):
            try:
                os.rename(file, file)
                print_debug(debug, 'Access on file "' + file + '" is available!')
                time.sleep(1)
                if error:
                    word_app = win32com.client.Dispatch("Word.Application")
                command_executed = True
            except OSError as e:
                print('Access-error on file "' + file + '"! \n' + str(e))
                messagebox.showinfo(title="Script de correction automatique",
                                    message="Fermer le document " + file + " pour la nouvelle correction")
                error=True
    # Ouvrir le document
    try:
        document = word_app.Documents.Open(filename)
    except Exception as e :
        print("erreur dans l'ouverture du document"+str(e))


    add_word_macro(document)
    #
    #check_hyperlinks(word_app, stud, debug=True)

    #check_header(word_app, stud, key="piedDePage", debug=True)
    #check_header_and_footer(word_app, stud, middle_text_asked="Examen TICE – B1", key="piedDePage", debug=True)
    check_tables(word_app, stud, key="tableau", debug=debug)
    document.Close(True)
    word_app.Quit()