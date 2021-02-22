# These variables are set by pyinstaller if running from a frozen
import sys
import os
is_frozen = getattr(sys, 'frozen', False)
frozen_temp_path = getattr(sys, '_MEIPASS', '')

AppData = os.path.expanduser('~')
QuickLinks_AppData = os.path.join(AppData, 'AppData\\Local\\Quick Links')
QuickLinks_Resources = os.path.join(QuickLinks_AppData, 'resources')


if is_frozen:
    # If running from exe, look for AppData Local folder, create if needed
    if not os.path.exists(QuickLinks_AppData):
        os.makedirs(QuickLinks_AppData)
        os.makedirs(QuickLinks_Resources)
        os.makedirs(os.path.join(QuickLinks_Resources, 'database'))
