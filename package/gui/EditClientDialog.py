import os
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from ..Client import Client


class EditClientDialog(QDialog):
    def __init__(self, resources_dir, client, parent=None):
        super(EditClientDialog, self).__init__(parent)
        self.resource_dir = resources_dir
        self.client = client

        # ***Window init*** #
        winIcon = QIcon()
        winIcon.addFile(os.path.join(self.resource_dir, 'images\\cas_logo_2.png'), QSize(24, 24))
        self.setWindowIcon(winIcon)
        self.setWindowTitle("Edit Client Prompt")
        self.resize(400, 400)

        layout = QGridLayout()

        # Widget for editing name
        self.nameLabel = QLabel('Client Name:')
        layout.addWidget(self.nameLabel, 0, 0)
        self.nameBox = QLineEdit(self)
        self.nameBox.setText(self.client.name)
        layout.addWidget(self.nameBox, 0, 1)

        # Widget for editing root dir with browse button
        self.dirLabel = QLabel('Client Folder:')
        layout.addWidget(self.dirLabel, 1, 0)
        self.dirBox = QLineEdit(self)
        self.dirBox.setText(self.client.folderPath)
        layout.addWidget(self.dirBox, 1, 1)
        self.browseDir = QPushButton("Select Dir")
        self.browseDir.clicked.connect(self.browseDirPath)
        layout.addWidget(self.browseDir, 1, 3)

        # Widget for editing excel path
        self.excelLabel = QLabel('Client Excel:')
        layout.addWidget(self.excelLabel, 2, 0)
        self.excelBox = QLineEdit(self)
        self.excelBox.setText(self.client.excelPath)
        layout.addWidget(self.excelBox, 2, 1)
        self.browseFile = QPushButton("Select File")
        self.browseFile.clicked.connect(self.browseFilePath)
        layout.addWidget(self.browseFile, 2, 3)

        # Widget for editing email
        self.emailLabel = QLabel('Client Email:')
        layout.addWidget(self.emailLabel, 3, 0)
        self.emailBox = QLineEdit(self)
        self.emailBox.setText(self.client.office365Email)
        layout.addWidget(self.emailBox, 3, 1)



        # Save and Cancel buttons
        self.buttons = QDialogButtonBox(
            QDialogButtonBox.Save | QDialogButtonBox.Cancel,
            Qt.Horizontal, self)
        layout.addWidget(self.buttons, 4, 1)

        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)

        self.setLayout(layout)
        self.resize(layout.sizeHint().width(), layout.sizeHint().height())

        # get current date and time from the dialog

    def browseDirPath(self):
        dirPath = QFileDialog.getExistingDirectory(self, 'Select Folder')
        self.dirBox.setText(dirPath)

    def browseFilePath(self):
        filePath = QFileDialog.getOpenFileName(self, 'Select File')
        self.excelBox.setText(filePath[0])

    @staticmethod
    def getInputs(resource_dir, client, parent=None):
        dialog = EditClientDialog(resource_dir, client, parent)
        result = dialog.exec_()

        name = dialog.nameBox.text()
        dirPth = dialog.dirBox.text()
        filePth = dialog.excelBox.text()
        email = dialog.emailBox.text()
        if email.isspace() or not email:
            email = None

        return name, dirPth, filePth, email, result == QDialog.Accepted




