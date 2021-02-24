# Class for Clients...holds needed information
# - Name
# - Folder Path
# - Excel Sheet Path
# - Office365 email
from PyQt5.QtWidgets import QPushButton

class Client:
    def __init__(self, name, folderPath, excelPath, office365Email):
        self.name = name
        self.folderPath = folderPath
        self.excelPath = excelPath
        self.office365Email = office365Email

    def info(self):
        print(self.name)
        print("\t", self.folderPath)
        print("\t", self.excelPath)
        print("\t", self.office365Email)
        print()

    def isValid(self):
        params = (self.name, self.folderPath, self.excelPath, self.office365Email)

        # Check for empty strings, or only whitespace strings. Exclude email
        for i in params[0:2]:
            if not i or i.isspace():
                return False

        # Check for space in email if one exists
        if self.office365Email is not None:
            if ' ' in self.office365Email:
                return False

        return True
