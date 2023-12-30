class Student:
    max_points = {"format": 2, "nomFichiers": 2,"poids": 2, "orthographe": 0, "pages": 2, "styles": 4, "piedDePage": 2,
        "numEtNbrPages":2, "espaces": 4, "TDM": 4, "listes": 2, "citation": 2, "noteBasPage": 2, "lien": 2,
        "images": 4, "section": 0}

    def reset(self):
        self.scores = {"format": 0, "nomFichiers": 0, "poids": 0, "orthographe": 0, "pages": 0,
                       "styles": 0, "piedDePage": 0, "numEtNbrPages":0, "espaces": 0, "TDM": 0, "listes": 0, "citation": 0,
                       "noteBasPage": 0, "lien": 0, "images": 0, "section": 0}
        self.reasons = {"format": "", "nomFichiers": "", "poids": "", "orthographe": "", "pages": "",
                        "styles": "", "piedDePage": "", "numEtNbrPages": "", "espaces": "", "TDM": "", "listes": "", "citation": "",
                        "noteBasPage": "", "lien": "", "images": "", "section": ""}
        self.name = ""
        self.firstname = ""
        self.group = "Unknown"
        self.to_check_manually = ""
        self.to_check = set()

    def __init__(self):
        self.scores = {"format": 0, "nomFichiers": 0, "poids": 0, "orthographe": 0, "pages": 0,
                       "styles": 0, "piedDePage": 0, "numEtNbrPages":0, "espaces": 0, "TDM": 0, "listes": 0, "citation": 0,
                       "noteBasPage": 0, "lien": 0, "images": 0, "section": 0}
        self.reasons = {"format": "", "nomFichiers": "", "poids": "", "orthographe": "", "pages": "",
                        "styles": "", "piedDePage": "", "numEtNbrPages": "", "espaces": "", "TDM": "", "listes": "", "citation": "",
                        "noteBasPage": "", "lien": "", "images": "", "section": ""}
        self.name = ""
        self.firstname = ""
        self.group = "Unknown"
        self.to_check_manually = ""
        self.to_check = set()
