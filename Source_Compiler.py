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

# """CLX Networks"""
def clx(filename, root, source, edate, upload_list):
    filename1 = convert().excel_tsv_to_csv(filename)
    filename2 = convert().csv_to_excel(filename1)
    filename3 = format().excel_format(filename2, source, 0, edate)


# """Monty Mobile"""
def monty(filename, root, source, edate, upload_list):
    filename1 = convert().csv_to_excel(filename)
    filename2 = format().excel_format(filename1, source, 0, edate)
    if filename2 == -1:
        move_to_noRates(filename)
        return 'No rate in document.'    
    bst().source_build(root, filename2) 
    status = bst().write(root, database)   
    index = filename.rfind('.')
    short = filename[:index]
    index = len(filename) - index
    ext = filename[-index:]
    newname = short + ' ' + str(edate) + ext
    # rename_file(filename, newname)
    # move_to_processed(filename)
    # file_clean(filename)    
    return status

# """Support to delete first row"""
def silverstreet(filename, root, source, edate, upload_list):
    book = xlrd.open_workbook(filename, 'rb')
    sheet = book.sheet_by_index(0)
    # """Check is to see if thit contains a random row value and modify it"""
    if sheet.cell(1,1).value == 'Catch all':
        new_book = xlwt.Workbook()
        sheet_wr = new_book.add_sheet("sheet", cell_overwrite_ok = True)
        rownum = sheet.nrows
        colnum = sheet.ncols
        for i in range(rownum):
            for j in range(colnum):
                if i == 0:
                    value = sheet.cell(i,j).value
                    sheet_wr.write(i,j,value)
                elif i > 1 and i < rownum:
                    value = sheet.cell(i,j).value
                    sheet_wr.write(i-1,j,value)
                else:
                    pass
        new_book.save(filename)

    bst().database_build(root, edate)
    filename1 = format().excel_format(filename, source, 0, edate)
    if filename1 == -1:
        move_to_noRates(filename)
        return 'No rate in document.'
    bst().source_build(root, filename1)
    status = bst().write(root, edate, upload_list)
    index = filename.rfind('.')
    short = filename[:index]
    index = len(filename) - index
    ext = filename[-index:]
    newname = short + ' ' + str(edate) + ext
    # rename_file(filename, newname)
    # move_to_day_folder(neswname, edate, 'Processed')
    file_clean(filename)
    return status

# """Tata"""
def tata(filename, root, database, source, edate):
    filename1 = format().excel_format(filename, source, 1, edate)
    if filename1 == -1:
        move_to_noRates(filename)
        return 'No rate in document.'    
    bst().source_build(root, filename1)
    status = bst().write(root, database)
    index = filename.rfind('.')
    short = filename[:index]
    index = len(filename) - index
    ext = filename[-index:]
    newname = short + ' ' + str(edate) + ext
    # rename_file(filename, newname)
    # move_to_processed(filename)
    # file_clean(filename)
    return status

# """Tedexis"""
def tedexis(filename, root, source, edate, upload_list):
    bst().database_build(root, edate)
    filename1 = format().excel_format(filename, source, 0, edate)
    if filename1 == -1:
        move_to_noRates(filename)
        return 'No rate in document.'
    filename2 = format().excel_filter(filename1)
    bst().source_build(root, filename2)
    status = bst().write(root, edate, upload_list)
    index = filename.rfind('.')
    short = filename[:index]
    index = len(filename) - index
    ext = filename[-index:]
    newname = short + ' ' + str(edate) + ext
    # rename_file(filename, newname)
    # move_to_processed(filename)
    # file_clean(filename)
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
    index = filename.rfind('.')
    short = filename[:index]
    index = len(filename) - index
    ext = filename[-index:]
    newname = short + ' ' + str(edate) + ext
    # rename_file(filename, newname)
    # move_to_day_folder(neswname, edate, 'Processed')
    # file_clean(filename)
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
    special_dictionary = ['Tedexis', 'Monty Mobile', 'Tata Communications', 'Silverstreet', 'CLX Networks']

    title = [0000000000000000000, 'Country', 'Network', 'MCC', 'MNC', 'MCCMNC', 'Rate', 'CURR', 'Converted Rate', 'Source', 'Effective Date', 0]
    header = bst().node(title[0], title)

    # """Folder ID is for Test Files Folder"""
    dl_list = dl_folder('0BzlU44AWMToxNkdCVXEzWndLT1U')
    upload_list = []
    
    if len(dl_list) == 0:
        print "No new files to be processed."
        return
    else:
        print "\nDownload list is: ", dl_list

    company_list = get_email_attachment_list(dl_list)
    print "Email attachment list is: ", company_list 
    if len(company_list) != len(dl_list):
        print ("Not all files downloaded for processing were located as an attachment in the emails.  'New' label status of email may have been removed.")

    for i in range(len(company_list)):
        file_to_process = company_list.pop()
        processed = False
        print "\nFile currently being processed is: ", file_to_process[0]
        print "Remaining number of files to be processed is: ", len(company_list)
        for j in range(len(general_dictionary)):
            # """General use case scenario"""
            if file_to_process[1] == general_dictionary[j]:
                status = general(file_to_process[0], header, file_to_process[1], file_to_process[3], upload_list)
                print "Status of: ", file_to_process[0], ' is: ', status
                processed = True
            # """Special use case scenario"""
            elif j in range(len(special_dictionary)):
                # """Tedexis"""
                if file_to_process[1] == special_dictionary[j] and j == 0:
                    status = tedexis(file_to_process[0], header, file_to_process[1], file_to_process[3], upload_list)
                    print "Status of: ", file_to_process[0], ' is: ', status
                    processed = True                
                # """Monty Mobile"""
                elif file_to_process[1] == special_dictionary[j] and j == 1:
                    status = tedexis(file_to_process[0], header, file_to_process[1], file_to_process[3], upload_list)
                    print 'Status of: ', file_to_process[0], ' is: ', status
                    processed = True
                # """Tata Communications"""
                elif file_to_process[1] == special_dictionary[j] and j == 2:
                    status = tata(file_to_process[0], header, file_to_process[1], file_to_process[3], upload_list)
                    print 'Status of: ', file_to_process[0], ' is: ', status
                    processed = True
                # """Silverstreet"""
                elif file_to_process[1] == special_dictionary[j] and j == 3:
                    status = silverstreet(file_to_process[0], header, file_to_process[1], file_to_process[3], upload_list)
                    print 'Status of: ', file_to_process[0], ' is: ', status
                # """CLX Networks"""
                elif file_to_process[1] == special_dictionary[j] and j == 4:
                    status = clx(file_to_process[0], header, file_to_process[1], file_to_process[3], upload_list)
                    print 'Status of: ', file_to_process[0], ' is: ', status
                # """Not special case"""
                else:
                    pass
            # """Case not yet tested:::
        if not processed:
            file_source = file_to_process[1]
            if file_to_process[1] == None:
                file_source = 'None'
            print ('The file: ' + file_to_process[0] + ' is currently not supported.  Source of file is: ' + file_source + 
                   '. Contact the developer to build support for this document.')

    print '\nNow uploading compiled data flies'
    for i in range(len(upload_list)):
        filename = upload_list.pop()
        filename = 'Rates for ' + filename + '.xls'
        to_delete = find_file_id(filename)
        if not to_delete == None:
            delete_file(to_delete)
        upload_as_gsheet('Data/' + filename, filename)
        move_to_folder(filename, '0BzlU44AWMToxdlJKMWFncWJzMVk')

    print "\nSource_Compiler has succesfully run to completion.\n\n\n"

if __name__ == '__main__':
    main()
    