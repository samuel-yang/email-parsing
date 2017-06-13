import xlrd, xlwt, pdfminer, csv, shutil, os, xlutils, sys
# from cstringIO import stringIO
from CurrencyConverter import *
from decimal import *
from Google_API_Manipulation import *
from datetime import *
from Database import *

# global database, today, tomorrow
reload(sys)
sys.setdefaultencoding('utf-8')

# """TESTING FILE"""
filename = 'Test KDDI Global - SMS A-Z Rates 201705242207.xlsx'

# """ C3ntro """
def c3ntro(filename, root, database):
    source = 'c3ntro'
    filename1 = format().excel_format(filename, source)
    # bst().source_build(root, filename1)
    # bst().write(root, database)

def horisen(filename, root, database):
    # """ BUILD SUPPORT FOR MCC MNC - separators and combiners
    source = 'HORISEN'
    filename1 = format().excel_format(filename, source)
    #BUILD DOESNT WORK BECAUSE FREAKING ENCODINGS SUCK
    # bst().source_build(root, filename1)
    # # bst().write(root, database)

def sms_az(filename, root, database):
    source = 'SMS A-Z'
    filename1 = convert().excel_to_csv(filename)
    filename2 = convert().csv_to_excel(filename1)
    filename3 = format().excel_format(filename2, source)
    bst().source_build(root, filename3)
    # bst().write(root, database)


def main():
    database = 'Rates for 06-12.xls'
    today = str(date.today())[-5:]

    version = database[-9:]
    version = version[:5]
    if not version == today:
        database = 'Rates for ' + today + '.xls'

    title = [0000000000000000000, 'Country', 'Network', 'MCC', 'MNC', 'MCCMNC', 'Rate', 'CURR', 'Converted Rate', 'Source', 'Effective Date', 0]
    header = bst().node(title[0], title)
    bst().database_build(database, header)
    sms_az(filename, header, database)
    bst().in_order_print(header)

if __name__ == '__main__':
    main()
    