import click


@click.command()
@click.argument('pofile1', type=click.File())
@click.argument('pofile2', type=click.File())
def main(pofile1, pofile2):
    pass
