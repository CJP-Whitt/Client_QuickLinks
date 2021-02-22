import os
from package.main import main
from package.utils import is_frozen, frozen_temp_path, \
    AppData, QuickLinks_AppData, QuickLinks_Resources

# Var for verbose outputs on/off
verbose = True

if is_frozen:
    # Use database from Persistent App Data\Quick Links folder
    basedir = frozen_temp_path
else:
    # Use python local runtime folder
    basedir = os.path.dirname(os.path.abspath(__file__))
    QuickLinks_Resources = os.path.join(basedir, 'resources')

resource_dir = os.path.join(basedir, 'resources')


main(resource_dir, QuickLinks_Resources, verbose)
