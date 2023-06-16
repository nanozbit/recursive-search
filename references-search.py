import os
import re
import openpyxl
import click


def search_match(filename, target_searched):
    targets = []
    with open(filename, 'r') as f:
        for num, line in enumerate(f, 1):
            pattern = re.compile(target_searched, re.IGNORECASE)
            match = re.search(pattern, line)
            if match:
                targets.append([num, line.strip(), filename])
    return targets


@click.command()
@click.option('--directory', help='Name directory')
@click.option('--target', prompt='Target Name', help='Word to search')
def main(directory, target):
    book_sheet = openpyxl.Workbook()
    sheet = book_sheet.active
    sheet.title = target
    header = ['Line Number', "Line", "Path"]
    sheet.append(header)

    for directories, subdirectories, files in os.walk(directory):

        for file in files:
            path_file = os.path.join(directories, file)
            if os.path.isfile(path_file) and path_file[-3:] == '.py':
                matches = search_match(path_file, target)
                for item in matches:
                    sheet.append([item[0], item[1], item[2]])

    book_sheet.save(f'{target}.xlsx')
    print('Completed Successfully')


if __name__ == '__main__':
    main()
