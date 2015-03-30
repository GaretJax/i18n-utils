import click

from i18n_utils import params, normalization


@click.command()
@click.option('--wrap-width', '-w', 'wrapwidth',
              default=normalization.DEFAULT_WRAPPING_WIDTH)
@click.argument('folders', nargs=-1, type=params.LocaleFolderParamType())
def main(folders, wrapwidth):
    """Normalizes the PO files in the given locale folders.

    Useful to remove unecessary noise from source commits. The normalization
    process applies the following operations:

    * The lines are rewrapped to the given wrapwidth
    * Sort the entries in the file by their `msgstr` (obsoleted entries are
      sorted in line with current entries and not grouped together in the end)
    * Each occurrence is put on a line by itself, even if the wrapwidth would
      allow for more to be put on a single line
    """
    for f in folders:
        for locale, pofile in f.pofiles(
                wrapwidth=wrapwidth, klass=normalization.NormalizedPOFile):
            pofile.save()
