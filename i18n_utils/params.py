import os
import glob

import click

import polib


class LocaleFolderParamType(click.Path):
    name = 'locale-folder'

    def __init__(self):
        super(LocaleFolderParamType, self).__init__(
            exists=True, file_okay=False, dir_okay=True, writable=True,
            readable=True, resolve_path=True)

    def convert(self, value, param, ctx):
        path = super(LocaleFolderParamType, self).convert(value, param, ctx)
        return LocaleFolder(path)


class LocaleFolder(object):
    def __init__(self, path):
        self.path = path

    def locales(self):
        locales = os.path.join(self.path, '*', 'LC_MESSAGES')

        for path in glob.glob(locales):
            locale = os.path.basename(os.path.dirname(path))
            yield locale

    def pofiles(self, *args, **kwargs):
        po_files = os.path.join(self.path, '*', 'LC_MESSAGES', '*.po')
        for path in glob.glob(po_files):
            locale = os.path.basename(os.path.dirname(os.path.dirname(path)))
            yield locale, polib.pofile(path, *args, **kwargs)
