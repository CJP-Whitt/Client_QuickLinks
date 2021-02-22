# Class for Client XML File Parsing
# - reads xml file
# - writes xml file
# - alphabetizes
import os
import xml.etree.ElementTree as ET
import xml.dom.minidom
import shutil

from .Client import Client
from datetime import date


class ClientFileParser:
    def __init__(self, resourceDir, verbose=False):
        self.resource_dir = resourceDir
        self.filename = os.path.join(self.resource_dir, 'database\\ClientList.xml')

        # Check if default xml exists, if not create it
        if not os.path.exists(self.filename):
            text = r'<?xml version="1.0" ?><data></data>'
            xmlFile = open(self.filename, "w+")
            xmlFile.write(text)
            xmlFile.close()

        self.tree = ET.parse(self.filename)
        self.root = self.tree.getroot()
        self.verbose = verbose

        self.curClientList = self.read_clients()

    def read_clients(self):
        # Read all clients from the xml Client List file and return list of objs
        clients = []

        if self.verbose:
            print("XML Client Objects List:")

        for child in self.root:
            tempClient = Client(child[0].text, child[1].text, child[2].text, child[3].text)
            clients.append(tempClient)
            if self.verbose:
                tempClient.info()

        self.curClientList = clients

        return clients

    def write_client(self, client):
        # Add single new client
        newClient = ET.Element("client")
        name = ET.SubElement(newClient, "name")
        name.text = client.name
        foldPth = ET.SubElement(newClient, "folderPath")
        foldPth.text = client.folderPath
        excelPth = ET.SubElement(newClient, "excelPath")
        excelPth.text = client.excelPath
        email = ET.SubElement(newClient, "office365Email")
        email.text = client.office365Email
        self.root.append(newClient)
        self.tree.write(self.filename)
        self.alphabetize()

    def edit_client(self, targetName, client):
        # Edit via saving changes and target client, deleting client and adding new one
        # with changes
        self.remove_client(targetName)
        self.write_client(client)
        self.read_clients()

    def remove_client(self, clientName):
        # Remove client from xml doc
        for child in self.root:
            if child[0].text == clientName:
                self.root.remove(child)
                self.tree.write(self.filename)
                self.prettifyXML()
                self.read_clients()
                return

        if self.verbose:
            print("ClientFileParser.delete_client: Error, no matching client found...\n")

    def alphabetize(self):
        # Create xml tree from sorted list of client objects
        clientList = self.read_clients()
        sortedClients = sorted(clientList, key=lambda x: x.name.lower())

        # Create new XML tree from sorted objects and write to filename
        start = ET.Element('data')
        tempTree = ET.ElementTree(start)
        for client in sortedClients:
            newClient = ET.Element("client")
            name = ET.SubElement(newClient, "name")
            name.text = client.name
            foldPth = ET.SubElement(newClient, "folderPath")
            foldPth.text = client.folderPath
            excelPth = ET.SubElement(newClient, "excelPath")
            excelPth.text = client.excelPath
            email = ET.SubElement(newClient, "office365Email")
            email.text = client.office365Email
            start.append(newClient)

        # Make new tree and write it and save root
        self.tree = tempTree
        self.tree.write(self.filename)
        self.root = self.tree.getroot()
        # Pretty write XML
        self.prettifyXML()
        # Reread updated client list
        self.read_clients()

    def save_recovery(self):
        # Save extra data file of clientList for backup, uses current date for naming scheme
        # Get date for naming scheme
        today = date.today()
        dateFormatted = today.strftime("%m-%d-%Y")

        # Rename file with current date
        backupName = os.path.join(self.resource_dir, 'database\\ClientList_' + str(dateFormatted) + '.xml')
        shutil.copyfile(self.filename, backupName)
        return backupName

    def prettifyXML(self):
        # Correctly indent XML file
        dom = xml.dom.minidom.parse(self.filename)
        pretty_xml_string = dom.toprettyxml()
        with open(self.filename, 'w') as filetowrite:
            filetowrite.write(pretty_xml_string)
