#!/usr/bin/env python
import argparse
import locale
import sys

import xlrd

import simple
import title

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
        title.conversion(book, args.sheet)
    else:
        simple.conversion(book, args.sheet)
