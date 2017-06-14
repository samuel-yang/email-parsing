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
# filename = 'Test Mitto AG - CoverageList_20170606_1000_hookmob1.xlsx'
# filename = 'UPM_SMSR-1_HOOK MOBILE_USD_2017-06-12 FORMATTED.xls'
# filename = 'Test.xls'

# """Monty Mobile"""
def monty(filename, root, database, source):
    filename1 = convert().csv_to_excel(filename)
    filename2 = format().excel_format(filename1, source, 0)
    if filename2 == -1:
        move_to_noRates(filename)
        return 'No rate in document.'    
    bst().source_build(root, filename2) 
    bst().write(root, database)   
    file_clean(filename)

# """Tata"""
def tata(filename, root, database, source):
    filename1 = format().excel_format(filename, source, 1)
    if filename1 == -1:
        move_to_noRates(filename)
        return 'No rate in document.'    
    bst().source_build(root, filename1)
    bst().write(root, database)
    file_clean(filename)

# """Tedexis"""
def tedexis(filename, root, database, source):
    filename1 = format().excel_format(filename, source, 0)
    if filename1 == -1:
        move_to_noRates(filename)
        return 'No rate in document.'
    filename2 = format().excel_filter(filename1)
    bst().source_build(root, filename2)
    bst().write(root, database)
    file_clean(filename)

# """General Use Case"""
# support for C3ntro, Mitto, MMD, UPM, Wavecell
def general(filename, root, database, source):
    filename1 = format().excel_format(filename, source, 0)
    if filename1 == -1:
        move_to_noRates(filename)
        return 'No rate in document.'
    bst().source_build(root, filename1)
    bst().write(root, database)
    file_clean(filename)
    # move_to_processed(filename)



# """ ------------------------------------------- MAIN CODE HERE --------------------------------------------------------------------------------------------"""
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

    # company_list = dl_folder('0BzlU44AWMToxNkdCVXEzWndLT1U')
    # print company_list
    # temp_list = get_email_attachment_list(company_list)
    # print temp_list
    # temp = temp_list.pop()
    # temp = company_list.pop()
    temp = '20170421 - Tedexis_Pricing_List_PREMIUM.xlsx'
    # emaildate = temp[3]
    bst().database_build(database, header)
    status = tedexis(temp, header, database, 'Tedexis')
    # status = general(temp[0], header, database, 'Openmarket')
    print status


if __name__ == '__main__':
    main()
    