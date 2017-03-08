from util import content, isempty


def __print_row(book, row, level):
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


def convert(book, sheetname):
    sheet = book.sheet_by_name(sheetname)
    level = 0

    print "{|"
    for row in sheet.get_rows():
        if isempty(row):
            continue
        level = __print_row(book, row, level)
        print "|-"
    print "|}"
