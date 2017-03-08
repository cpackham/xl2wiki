import argparse
import locale
import sys

import xlrd


def content(book, cell):
    """Get the content of a Cell converting as needed"""
    if cell.ctype == xlrd.XL_CELL_DATE:
        date = xlrd.xldate.xldate_as_datetime(cell.value, book.datemode)
        return date.strftime("%x")
    else:
        return "{}".format(cell.value)


def isempty(row):
    return sum(cell.ctype for cell in row) == 0


def simple_conversion(book, sheetname):
    sheet = book.sheet_by_name(sheetname)

    print "{|"
    for row in sheet.get_rows():
        if isempty(row):
            continue
        print "| " + " || ".join([content(book, cell) for cell in row])
        print "|-"
    print "|}"


def print_row(book, row, level):
    if level == 0:
        print "! colspan={} | {}".format(len(row),
                                         "".join([content(book, cell)
                                                  for cell in row]))
        return level + 1
    elif level == 1:
        print "! " + " !! ".join([content(book, cell) for cell in row])
        return level + 1
    else:
        print "| " + " || ".join([content(book, cell) for cell in row])
        return level


def title_conversion(book, sheetname):
    sheet = book.sheet_by_name(sheetname)
    level = 0

    print "{|"
    for row in sheet.get_rows():
        if isempty(row):
            continue
        level = print_row(book, row, level)
        print "|-"
    print "|}"


if __name__ == "__main__":
    locale.setlocale(locale.LC_ALL, '')

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("-s", "--sheet", type=str, default="Sheet1",
                        help="Name of worksheet default:%(default)s")
    parser.add_argument("--format", type=str, choices=["simple", "title"],
                        default="simple",
                        help="Output format default:%(default)s")
    parser.add_argument("file", type=str)

    args = parser.parse_args()

    try:
        book = xlrd.open_workbook(args.file)
    except IOError as e:
        sys.stdout.write("{}\n".format(e))
        sys.exit(1)

    if args.format == "title":
        title_conversion(book, args.sheet)
    else:
        simple_conversion(book, args.sheet)
