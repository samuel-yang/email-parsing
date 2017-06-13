import xlrd, xlwt, pdfminer, csv, shutil, os, xlutils, sys
# from cstringIO import stringIO
from CurrencyConverter import *
from decimal import *
from Google_API_Manipulation import *
from datetime import *
from Database_Manipulation import *

# global database, today, tomorrow
reload(sys)
sys.setdefaultencoding('utf-8')

# """TESTING FILE"""
# filename = 'Test C3ntro Telecom - C3ntro Telecom Price Change Notification for Hookmobile-20170602.xlsx'
filename = 'Test Mitto AG - CoverageList_20170606_1000_hookmob1.xlsx'
# filename = 'Test.xlsx'

# """ C3ntro """
def c3ntro(filename, root, database):
    source = 'c3ntro'
    filename1 = format().excel_format(filename, source)
    bst().source_build(root, filename1)
    bst().write(root, database)

def horisen(filename, root, database):
    # """ BUILD SUPPORT FOR MCC MNC - separators and combiners
    source = 'HORISEN'
    filename1 = format().excel_format(filename, source)
    #BUILD DOESNT WORK BECAUSE FREAKING ENCODINGS SUCK
    # bst().source_build(root, filename1)
    # # bst().write(root, database)

def mitto(filename, root, database):
    source = 'Mitto AG'
    filename1 = format().excel_format(filename, source)
    bst().source_build(root, filename1)
    bst().write(root, database)

def sms_az(filename, root, database):
    source = 'SMS A-Z'
    filename1 = convert().excel_to_csv(filename)
    filename2 = convert().csv_to_excel(filename1)
    filename3 = format().excel_format(filename2, source)
    bst().source_build(root, filename3)
    # bst().write(root, database)


def main():
    # """Defining dates for use in methods"""
    today = str(date.today())[-5:]
    tomo = date.today() + timedelta(days = 1)
    tomorrow = str(tomo)[-5:]
    yester = date.today() - timedelta(days = 1)
    yesterday = str(yester)[-5:]
    database = 'Data/Rates for ' + today + '.xls'

    if not os.path.isfile(database):
        old_database = 'Data/Rates for ' + yesterday + '.xls'
        shutil.copy2(old_database, database)
        print "new file made"

    title = [0000000000000000000, 'Country', 'Network', 'MCC', 'MNC', 'MCCMNC', 'Rate', 'CURR', 'Converted Rate', 'Source', 'Effective Date', 0]
    header = bst().node(title[0], title)
    bst().database_build(database, header)
    mitto(filename, header, database)


if __name__ == '__main__':
    main()
    