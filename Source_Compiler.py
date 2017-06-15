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
def monty(filename, root, database, source, edate):
    filename1 = convert().csv_to_excel(filename)
    filename2 = format().excel_format(filename1, source, 0, edate)
    if filename2 == -1:
        move_to_noRates(filename)
        return 'No rate in document.'    
    bst().source_build(root, filename2) 
    status = bst().write(root, database)   
    # file_clean(filename)
    index = filename.rfind('.')
    short = filename[:index]
    index = len(filename) - index
    ext = filename[-index:]
    newname = short + ' ' + str(edate) + ext
    # rename_file(filename, newname)
    # move_to_processed(filename)
    return status

# """Tata"""
def tata(filename, root, database, source, edate):
    filename1 = format().excel_format(filename, source, 1, edate)
    if filename1 == -1:
        move_to_noRates(filename)
        return 'No rate in document.'    
    bst().source_build(root, filename1)
    status = bst().write(root, database)
    # file_clean(filename)
    index = filename.rfind('.')
    short = filename[:index]
    index = len(filename) - index
    ext = filename[-index:]
    newname = short + ' ' + str(edate) + ext
    # rename_file(filename, newname)
    # move_to_processed(filename)
    return status

# """Tedexis"""
def tedexis(filename, root, source, edate, upload_list):
    bst.database_build(root, edate)
    filename1 = format().excel_format(filename, source, 0, edate)
    if filename1 == -1:
        move_to_noRates(filename)
        return 'No rate in document.'
    filename2 = format().excel_filter(filename1)
    bst().source_build(root, filename2)
    status = bst().write(root, edate, upload_list)
    # file_clean(filename)
    index = filename.rfind('.')
    short = filename[:index]
    index = len(filename) - index
    ext = filename[-index:]
    newname = short + ' ' + str(edate) + ext
    # rename_file(filename, newname)
    # move_to_processed(filename)
    return status

# """General Use Case"""
# support for mmd, UPM, wavecell, mitto, monty, centro, tata, tedexis, bics, openmarket
def general(filename, root, source, edate, upload_list):
    bst().database_build(root, edate)
    filename1 = format().excel_format(filename, source, 0, edate)
    if filename1 == -1:
        move_to_noRates(filename)
        return 'No rate in document.'
    bst().source_build(root, filename1)
    status = bst().write(root, edate, upload_list)
    # file_clean(filename)
    index = filename.rfind('.')
    short = filename[:index]
    index = len(filename) - index
    ext = filename[-index:]
    newname = short + ' ' + str(edate) + ext
    # rename_file(filename, newname)
    # move_to_day_folder(neswname, edate, 'Processed')
    return status


# """ ------------------------------------------- MAIN CODE HERE --------------------------------------------------------------------------------------------"""
def main():

    # if not os.path.isfile(database):
    #     old_database = 'Data/Rates for ' + yesterday + '.xls'
    #     if not os.path.isfile(old_database):
    #         book = xlwt.Workbook()
    #         sheet = book.add_sheet("sheet", cell_overwrite_ok = True)
    #         book.save(database)
    #     else:
    #         shutil.copy2(old_database, database)
    #     print "new file made"

    general_dictionary = ['MMDSmart', 'UPM Telecom', 'OpenMarket', 'Wavecell', 'Bics', 'Mitto AG', 'C3ntro Telecom']
    special_dictionary = ['Tedexis', 'Monty Mobile', 'Tata Communications']

    title = [0000000000000000000, 'Country', 'Network', 'MCC', 'MNC', 'MCCMNC', 'Rate', 'CURR', 'Converted Rate', 'Source', 'Effective Date', 0]
    header = bst().node(title[0], title)

    dl_list = dl_folder('0BzlU44AWMToxNkdCVXEzWndLT1U')
    upload_list = []
    
    if len(dl_list) == 0:
        print "No new files to be processed."
        return
    else:
        print "\nDownload list is: ", dl_list

    company_list = get_email_attachment_list(dl_list)
    print "Email attachment list is: ", company_list

    for i in range(len(company_list)):
        file_to_process = company_list.pop()
        print "\nFile currently being processed is: ", file_to_process[0]
        print "Remaining number of files to be processed is: ", len(company_list)
        for j in range(len(general_dictionary)):
            # """General use case scenario"""
            if file_to_process[1] == general_dictionary[j]:
                status = general(file_to_process[0], header, file_to_process[1], file_to_process[3], upload_list)
                print "Status of: ", file_to_process[0], ' is: ', status
            # """Special use case scenario"""
            elif j <= len(special_dictionary):
                # """Tedexis"""
                if file_to_process[1] == special_dictionary[0]:
                    status = tedexis(file_to_process[0], header, file_to_process[1], file_to_process[3], upload_list)
                    print "Status of: ", file_to_process[0], ' is: ', status
                # """Monty Mobile"""
                elif file_to_process[1] == special_dictionary[1]:
                    status = tedexis(file_to_process[0], header, file_to_process[1], file_to_process[3], upload_list)
                    print 'Status of: ', file_to_process[0], ' is: ', status
                # """Tata Communications"""
                elif file_to_process[1] == special_dictionary[2]:
                    status = tata(file_to_process[0], header, file_to_process[1], file_to_process[3], upload_list)
                    print 'Status of: ', file_to_process[0], ' is: ', status
                # """Not special case"""
                else:
                    pass
            # """Case not yet tested:::
            elif j == len(general_dictionary):
                print "The final check has been entered"
                if not file_to_process[i] == general_dictionary[j]:
                    print 'The file: ', file_to_process[0], ' is currently not supported.  Contact the developer to build support for this document.'

    print '\nNow uploading compiled data flies'
    for i in range(len(upload_list)):
        filename = upload_list.pop()
        filename = 'Rates for ' + filename + '.xls'
        to_delete = find_file_id(filename)
        if not to_delete == None:
            delete_file(to_delete)
        upload_as_gsheet('Data/' + filename, filename)
        move_to_folder(filename, '0BzlU44AWMToxdlJKMWFncWJzMVk')

if __name__ == '__main__':
    main()
    