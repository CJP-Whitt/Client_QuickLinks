import os, sys, subprocess
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from ..Client import Client
from ..ClientFileParser import ClientFileParser
from .NewClientDialog import NewClientDialog
from .EditClientDialog import EditClientDialog
from package.utils import is_frozen, frozen_temp_path


def url_infuse(emailAddy):
    url = r'https://login.microsoftonline.com/common/oauth2/authorize?client_id=4345a7b9-9a63-4910-a426-35363201d503' \
          r'&redirect_uri=https%3A%2F%2Fwww.office.com%2Flanding&response_type=code%20id_token&scope=openid%20profile' \
          r'&response_mode=form_post&nonce=637268985401811218' \
          r'.NjdjMmM0MDEtZTY3MC00YmMwLWJmNWMtYmQyN2E5YTI1MmU3NzkyNDQxODEtYzM5OS00ZjA1LThiMGUtNzk5MjI3MGRiOWZk' \
          r'&ui_locales=en-US&mkt=en-US&client-request-id=5a5903ca-62b6-4489-a362-200ae46d3e1b&login_hint=' + \
          emailAddy + r'&state=bSINFLudzeImYZ55rZPUVQlLVzLYjwhBd4PV5JSpry2WYcSI_pJrUzZwl' \
                      r'-FDED6opyApJpecivrKJXmWS8zSK8VPxCIFfnMBOwmumTzQUjxN6VQ3OidM--FWwyVdk' \
                      r'-vcC6820LOnM1uQKBVx5gIsreRaTSJW0xUWtDe_j8V9v2DKg__2Zm6B4q-DLMrEhDngMRmgzvREa0A_xGJs0LXg' \
                      r'-EVx41D39WoPHdSjNYOMUT5pCbVQaORSQcKJHWu8Hzv98aO_1DN3_P9cQ3qwGcGFAIgLhgiVS' \
                      r'-6jQZiy2RDNCylWdcIqPNsqqVY3nZ1VgqZL3HLELVjiAK2VYdPW8nV8mw&x-client-SKU=ID_NETSTANDARD2_0&x' \
                      r'-client-ver=6.5.0.0 '
    return url


class Application:
    def __init__(self, parser, resource_dir):
        self.myParser = parser
        self.resource_dir = resource_dir

        # App/MainWinodw init
        self.app = QApplication(sys.argv)

        # Win init
        self.win = QMainWindow()
        self.menuBar = QMenuBar()
        self.fileMenu = QMenu("&File", )
        self.newClient = QAction("New Client")  # New client btn
        self.exportList = QAction("Export")  # Export clients btn
        self.importList = QAction("Import")  # Import clients btn
        self.statusBar = QStatusBar()
        self.win.setStatusBar(self.statusBar)

        self.create_window()

        # ***Table init*** #
        self.myTable = QTableWidget()

        self.myTable.setRowCount(len(self.myParser.curClientList))
        self.myTable.setColumnCount(5)
        self.myTable.setHorizontalHeaderLabels((["Client Excel", "Root", "O365 Login", "", ""]))
        self.myTable.horizontalHeader().sectionPressed.disconnect()
        self.load_table()

        # App start, exit when done
        sys.exit(self.app.exec_())

    def create_window(self):
        # Win init
        winIcon = QIcon()
        winIcon.addFile(os.path.join(self.resource_dir, 'images\\cas_logo_2.png'), QSize(24, 24))
        self.win.setWindowIcon(winIcon)
        self.win.setWindowTitle("CAS Client Quick Links")

        # ***Menubar init*** #
        addIcon = QIcon()
        addIcon.addFile(os.path.join(self.resource_dir, 'images\\add_client.png'), QSize(16, 16))
        self.newClient.setIcon(addIcon)
        self.newClient.triggered.connect(self.new_client)

        exportIcon = QIcon()
        exportIcon.addFile(os.path.join(self.resource_dir, 'images\\export_icon.png'), QSize(16, 16))
        self.exportList.setIcon(exportIcon)
        self.exportList.triggered.connect(self.exp_file)

        importIcon = QIcon()
        importIcon.addFile(os.path.join(self.resource_dir, 'images\\import_icon.png'), QSize(24, 24))
        self.importList.setIcon(importIcon)
        self.importList.triggered.connect(self.exp_file)

        # Menubar actions
        self.fileMenu.addAction(self.newClient)
        self.fileMenu.addSeparator()
        self.fileMenu.addAction(self.importList)
        self.fileMenu.addAction(self.exportList)
        self.fileMenu.addSeparator()
        self.menuBar.addMenu(self.fileMenu)
        self.win.setMenuBar(self.menuBar)

        # ***App Start Menu Bar Start Msg*** #
        self.statusBar.showMessage('Welcome to CAS Client Quick Links', 5000)

        self.win.show()

    def load_table(self):
        # ***Icons init*** #
        editIcon = QIcon()
        editIcon.addFile(os.path.join(self.resource_dir, 'images\\edit.png'), QSize(24, 24))

        delIcon = QIcon()
        delIcon.addFile(os.path.join(self.resource_dir, 'images\\delete.png'), QSize(24, 24))

        dirIcon = QIcon()
        dirIcon.addFile(os.path.join(self.resource_dir, 'images\\directory.png'), QSize(24, 24))

        emailIcon = QIcon()
        emailIcon.addFile(os.path.join(self.resource_dir, 'images\\email_icon.jpg'), QSize(24, 24))

        # ***Buttons Init*** #
        print(len(self.myParser.curClientList))
        for i in range(len(self.myParser.curClientList)):  # range(len(self.myParser.curClientList)):
            # Client name and excel file
            btn0 = QPushButton(self.myParser.curClientList[i].name)
            btn0.setStyleSheet("font: 13px Arial, sans-serif;")
            btn0.clicked.connect(lambda *args, row=i, column=0: self.open_excel(row, column))
            self.myTable.setCellWidget(i, 0, btn0)
            # Client root dir
            btn1 = QPushButton()
            btn1.setIcon(dirIcon)
            btn1.setIconSize(QSize(24, 24))
            btn1.clicked.connect(lambda *args, row=i, column=1: self.open_dir(row, column))
            self.myTable.setCellWidget(i, 1, btn1)
            # Client email link
            email = self.myParser.curClientList[i].office365Email
            if not email:  # Only make button if an email exists
                btn2 = QPushButton()
                btn2.setEnabled(False)
            else:
                btn2 = QPushButton()
                btn2.setIcon(emailIcon)
                btn2.clicked.connect(lambda *args, row=i, column=2: self.open_email(row, column))
            self.myTable.setCellWidget(i, 2, btn2)
            # Edit client button
            btn3 = QPushButton()
            btn3.setIcon(editIcon)
            btn3.clicked.connect(lambda *args, row=i, column=3: self.edit_client(row, column))
            self.myTable.setCellWidget(i, 3, btn3)
            # Del client button
            btn4 = QPushButton()
            btn4.setIcon(delIcon)
            btn4.clicked.connect(lambda *args, row=i, column=4: self.del_client(row, column))
            self.myTable.setCellWidget(i, 4, btn4)

        self.myTable.setAlternatingRowColors(False)
        self.win.setCentralWidget(self.myTable)

        # ***Extra settings*** #
        self.myTable.horizontalHeader().setStyleSheet("font: bold 16px Arial, sans-serif;")
        self.myTable.verticalHeader().setHidden(True)
        self.myTable.setWordWrap(True)
        self.myTable.setSizeAdjustPolicy(QAbstractScrollArea.AdjustToContents)
        self.myTable.resizeColumnsToContents()
        self.myTable.verticalScrollBar().setValue(True)
        self.myTable.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.win.resize(self.myTable.sizeHint().width(), 400)
        self.myTable.show()
        self.win.show()

    def edit_client(self, r, c):
        self.statusBar.showMessage('Editing client: ' + self.myParser.curClientList[r].name, 5000)

        name, dirPth, filePth, email, state = EditClientDialog.getInputs(self.resource_dir, self.myParser.curClientList[r])
        if not state:
            self.statusBar.showMessage('Client edit cancelled', 5000)
            return

        newClient = Client(name, dirPth, filePth, email)
        if not newClient.isValid():
            self.statusBar.showMessage('Aborting Client Edit...Invalid Parameters', 5000)
            return

        self.myParser.edit_client(self.myParser.curClientList[r].name, newClient)
        self.statusBar.showMessage('Client ' + name + ' edited!', 5000)

        self.load_table()
        self.win.resize(self.myTable.sizeHint().width(), 400)

    def del_client(self, r, c):
        clientName = self.myParser.curClientList[r].name
        self.statusBar.showMessage('Deleting client: ' + clientName, 5000)
        reply = QMessageBox.question(self.win, 'Delete Client',
                                     "Do you want to delete " + clientName + "?",
                                     QMessageBox.Yes | QMessageBox.Cancel, QMessageBox.Cancel)
        if reply == QMessageBox.Yes:
            self.myParser.remove_client(clientName)
            self.load_table()
            self.statusBar.showMessage('Client deleted ' + clientName, 5000)
        else:
            self.statusBar.showMessage('Client not deleted ' + clientName, 5000)

        self.win.resize(self.myTable.sizeHint().width(), 400)

    def new_client(self):
        self.statusBar.showMessage('Creating new client...', 5000)

        name, dirPth, filePth, email, state = NewClientDialog.getInputs(self.resource_dir)

        if not state:
            self.statusBar.showMessage('Client creation cancelled', 5000)
            return
        newClient = Client(name, dirPth, filePth, email)
        if not newClient.isValid():
            self.statusBar.showMessage('Aborting client creation...Invalid paramaters', 5000)
            return

        self.myParser.write_client(newClient)

        self.load_table()
        self.statusBar.showMessage('Client ' + name + ' added!', 5000)
        self.fit_window()

    def imp_file(self):
        pass

    def exp_file(self):
        string = self.myParser.save_recovery()
        self.statusBar.showMessage('Exporting to file: ' + string)

    def open_excel(self, r, c):
        self.statusBar.showMessage('Opening excel for: ' + self.myParser.curClientList[r].name, 5000)
        print('start excel ' + '"' + self.myParser.curClientList[r].excelPath + '"')
        os.system('start excel ' + '"' + self.myParser.curClientList[r].excelPath + '"')

    def open_dir(self, r, c):
        self.statusBar.showMessage('Opening root dir for: ' + self.myParser.curClientList[r].name, 5000)
        pathString = r"{}".format(self.myParser.curClientList[r].folderPath)
        print('start "" ' + '"' + pathString + '"')
        os.system('start "" ' + '"' + pathString + '"')

    def open_email(self, r, c):
        # self.statusBar.showMessage('Email copied to clipboard: ' + self.myParser.curClientList[r].office365Email, 5000)
        # os.system('start https://login.microsoftonline.com/')
        # os.system('echo ' + self.myParser.curClientList[r].office365Email + '|clip')
        self.statusBar.showMessage('Opening O365 with: ' + self.myParser.curClientList[r].office365Email, 5000)
        url = url_infuse(self.myParser.curClientList[r].office365Email)
        print('start "" "' + url + '"')
        os.system('start "" "' + url + '"')

    def fit_window(self):
        self.win.resize(self.myTable.sizeHint().width(), 400)

