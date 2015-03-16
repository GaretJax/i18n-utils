import os
import glob

import click

import polib


def get_pofile_paths(folder):
    po_file = os.path.join(folder, '*', 'LC_MESSAGES', '*.po')
    for path in glob.glob(po_file):
        yield path


@click.command()
@click.option('--wrap-width', '-w', 'wrapwidth', default=78)
@click.argument('folders', nargs=-1)
def main(folders, wrapwidth):
    """Rewraps the po files found in the given locale folders.

    Useful to remove unecessary noise from source commits.
    """
    for f in folders:
        for path in get_pofile_paths(f):
            pofile = polib.pofile(path, wrapwidth=wrapwidth)
            pofile.save()


if __name__ == '__main__':
    main()
