"""
A collection of utilities to add functionality to the gettext commands.
"""
import os
import glob

import polib


__version__ = '0.0.1'
__url__ = 'https://github.com/GaretJax/i18n-utils'


def get_pofiles(folder):
    po_files = os.path.join(folder, '*', 'LC_MESSAGES', '*.po')
    for path in glob.glob(po_files):
        locale = os.path.basename(os.path.dirname(os.path.dirname(path)))
        yield locale, polib.pofile(path)
