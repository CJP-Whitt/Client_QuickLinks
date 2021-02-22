# **********************************************************************************************************
# Client_QuickLinks - (For IT Applications)
# Version 1.0 (1/15/2021)
# Created by Carson Whitt
#
# This application is for listing client excel and microsoft login data with quick link buttons. Has the
# ability to add new clients and delete old ones.
# **********************************************************************************************************
import os
from .ClientFileParser import ClientFileParser
from .gui.Application import Application


def main(resource_dir, QuickLinks_Persist, verbose=False):
    if verbose:
        print('Resources Path:', resource_dir)
        print('Database Path:', QuickLinks_Persist)

    # Create file parser for ClientList.xml
    myParser = ClientFileParser(QuickLinks_Persist)

    ### Create Window for Application ###
    myApp = Application(myParser, resource_dir)
