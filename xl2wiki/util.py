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
