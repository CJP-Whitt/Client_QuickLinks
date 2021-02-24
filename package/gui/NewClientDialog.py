import os
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *


class NewClientDialog(QDialog):
    def __init__(self, resources_dir, parent=None):
        super(NewClientDialog, self).__init__(parent)
        self.resource_dir = resources_dir

        # ***Window init*** #
        winIcon = QIcon()
        winIcon.addFile(os.path.join(self.resource_dir, 'images\\cas_logo_2.png'), QSize(24, 24))
        self.setWindowIcon(winIcon)
        self.setWindowTitle("New Client Prompt")
        self.resize(400, 400)

        layout = QGridLayout()

        # Widget for editing name
        self.nameBox = QLineEdit(self)
        self.nameBox.setPlaceholderText("Client Name")
        layout.addWidget(self.nameBox, 0, 0)

        # Widget for editing root dir with browse button
        self.dirBox = QLineEdit(self)
        self.dirBox.setPlaceholderText("Client Root Dir Path")
        layout.addWidget(self.dirBox, 1, 0)
        self.browseDir = QPushButton("Select Dir")
        self.browseDir.clicked.connect(self.browseDirPath)
        layout.addWidget(self.browseDir, 1, 2)

        # Widget for editing excel path
        self.excelBox = QLineEdit(self)
        self.excelBox.setPlaceholderText("Client Excel File Path")
        layout.addWidget(self.excelBox, 2, 0)
        self.browseFile = QPushButton("Select File")
        self.browseFile.clicked.connect(self.browseFilePath)
        layout.addWidget(self.browseFile, 2, 2)

        # Widget for editing email
        self.emailBox = QLineEdit(self)
        self.emailBox.setPlaceholderText("Client Email")
        layout.addWidget(self.emailBox, 3, 0)

        # Save and Cancel buttons
        self.buttons = QDialogButtonBox(
            QDialogButtonBox.Save | QDialogButtonBox.Cancel,
            Qt.Horizontal, self)
        layout.addWidget(self.buttons, 4, 0)

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
    def getInputs(resource_dir, parent=None):
        dialog = NewClientDialog(resource_dir, parent)
        result = dialog.exec_()

        name = dialog.nameBox.text()
        dirPth = dialog.dirBox.text()
        filePth = dialog.excelBox.text()
        email = dialog.emailBox.text()
        if email.isspace() or not email:
            email = None

        return name, dirPth, filePth, email, result == QDialog.Accepted




