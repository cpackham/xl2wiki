from util import content, isempty


def conversion(book, sheetname):
    sheet = book.sheet_by_name(sheetname)

    print "{|"
    for row in sheet.get_rows():
        if isempty(row):
            continue
        print "| " + " || ".join([content(book, cell) for cell in row])
        print "|-"
    print "|}"
