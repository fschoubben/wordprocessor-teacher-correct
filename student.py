class Student:
    max_points = {"format": 2, "nomFichiers": 2,"poids": 2, "orthographe": 0, "pages": 0,
                  "styles": 4, "piedDePage": 4, "espaces": 4, "TDM": 2, "section": 2, "listes": 2,
                  "tableau": 2,  "citation": 2, "noteBasPage": 2, "lien": 2, "images": 4}

    def reset(self):
        self.scores = {"format": 0, "nomFichiers": 0, "poids": 0, "orthographe": 0, "pages": 0,
                       "styles": 0, "piedDePage": 0, "espaces": 0, "TDM": 0, "section": 0, "listes": 0,
                       "tableau": 2,  "citation": 0, "noteBasPage": 0, "lien": 0, "images": 0 }
        self.reasons = {"format": "", "nomFichiers": "", "poids": "", "orthographe": "", "pages": "",
                        "styles": "", "piedDePage": "", "espaces": "", "TDM": "", "section": "", "listes": "",
                        "tableau": "", "citation": "", "noteBasPage": "", "lien": "", "images": ""}
        self.name = ""
        self.firstname = ""
        self.group = "Unknown"
        self.to_check_manually = ""
        self.to_check = set()

    def __init__(self):
        self.reset()
        # self.scores =  {"format": 0, "nomFichiers": 0, "poids": 0, "orthographe": 0, "pages": 0,
        #                "styles": 0, "piedDePage": 0, "espaces": 0, "TDM": 0, "section": 0, "listes": 0,
        #                "tableau":2,  "citation": 0, "noteBasPage": 0, "lien": 0, "images": 0 }
        # self.reasons = {"format": "", "nomFichiers": "", "poids": "", "orthographe": "", "pages": "",
        #                 "styles": "", "piedDePage": "", "espaces": "", "TDM": "", "listes": "", "citation": "",
        #                 "noteBasPage": "", "lien": "", "images": "", "section": ""}
        # self.name = ""
        # self.firstname = ""
        # self.group = "Unknown"
        # self.to_check_manually = ""
        # self.to_check = set()
if __name__ == "__main__":
    st=Student()
    total = 0
    i=0
    for key, value in st.scores.items():
       #print(key," --> ", value, "/", Student.max_points[key])
        print('{:<15}  --> {:>3} / {:>3}'.format(key, value, Student.max_points[key]))
        total+=st.max_points[key]
        i+=1
    print("max scrore = ",total,", ",i,"elements")