import os
import glob

import polib


def get_pofiles(folder):
    po_files = os.path.join(folder, '*', 'LC_MESSAGES', '*.po')
    for path in glob.glob(po_files):
        locale = os.path.basename(os.path.dirname(os.path.dirname(path)))
        yield locale, polib.pofile(path)
