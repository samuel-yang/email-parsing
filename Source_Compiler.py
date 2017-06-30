import xlrd, xlwt, pdfminer, csv, shutil, os, xlutils, sys
# import win32com.client
# from cstringIO import stringIO
from CurrencyConverter import *
from decimal import *
from Google_API_Manipulation import *
from time import sleep
from datetime import *
from Database_Manipulation import *
from gspread import *

# global database, today, tomorrow
global client
reload(sys)
sys.setdefaultencoding('utf-8')
client = authorize(get_credentials())

# """CLX Networks"""
def clx(filename, root, source, edate, upload_list):
    file_id = find_file_id_using_parent(filename, '0BzlU44AWMToxZnh5ekJaVUJUc2c') # Looks in "Files" folder
    filename1 = convert().excel_tsv_to_csv(filename)
    filename2 = convert().csv_to_excel(filename1)
    filename3 = format().excel_format(filename2, source, 0, edate)
    if filename3 == -1:
        move_to_folder(file_id, '0BzlU44AWMToxeFhld1pfNWxDTWs') # Moves to "NoRates" folder
        return 'No rate in document.'    
    bst().source_build(root, filename3)
    index = filename.rfind('.')
    short = filename[:index]
    index = len(filename) - index
    ext = filename[-index:]
    newname = short + ' ' + str(edate) + ext
    move_to_day_folder(file_id, edate, '0BzlU44AWMToxVU8ySkNBQzJQeFE') # Moves to date folder within "Processed" folder
    rename_file(file_id, newname)
    file_clean(filename)
# """Monty Mobile"""
def monty(filename, root, source, edate, upload_list):
    file_id = find_file_id_using_parent(filename, '0BzlU44AWMToxZnh5ekJaVUJUc2c') # Looks in "Files" folder
    filename1 = convert().csv_to_excel(filename)
    filename2 = format().excel_format(filename1, source, 0, edate)
    if filename2 == -1:
        move_to_folder(file_id, '0BzlU44AWMToxeFhld1pfNWxDTWs') # Moves to "NoRates" folder
        return 'No rate in document.'    
    filename3 = format().monty_is_special(filename2, filename1)
    bst().source_build(root, filename3) 
    index = filename.rfind('.')
    short = filename[:index]
    index = len(filename) - index
    ext = filename[-index:]
    newname = short + ' ' + str(edate) + ext
    move_to_day_folder(file_id, edate, '0BzlU44AWMToxVU8ySkNBQzJQeFE') # Moves to date folder within "Processed" folder
    rename_file(file_id, newname)
    file_clean(filename)

# """Support to delete first row"""
def silverstreet(filename, root, source, edate, upload_list):
    file_id = find_file_id_using_parent(filename, '0BzlU44AWMToxZnh5ekJaVUJUc2c') # Looks in "Files" folder
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

    filename1 = format().excel_format(filename, source, 0, edate)
    if filename1 == -1:
        move_to_folder(file_id, '0BzlU44AWMToxeFhld1pfNWxDTWs') # Moves to "NoRates" folder
        return 'No rate in document.'
    bst().source_build(root, filename1)
    index = filename.rfind('.')
    short = filename[:index]
    index = len(filename) - index
    ext = filename[-index:]
    newname = short + ' ' + str(edate) + ext
    move_to_day_folder(file_id, edate, '0BzlU44AWMToxVU8ySkNBQzJQeFE') # Moves to date folder within "Processed" folder
    rename_file(file_id, newname)
    file_clean(filename)

# """Tata"""
def tata(filename, root, source, edate, upload_list):
    file_id = find_file_id_using_parent(filename, '0BzlU44AWMToxZnh5ekJaVUJUc2c') # Looks in "Files" folder
    filename1 = format().excel_format(filename, source, 1, edate)
    if filename1 == -1:
        move_to_folder(file_id, '0BzlU44AWMToxeFhld1pfNWxDTWs') # Moves to "NoRates" folder
        return 'No rate in document.'    
    bst().source_build(root, filename1)
    index = filename.rfind('.')
    short = filename[:index]
    index = len(filename) - index
    ext = filename[-index:]
    newname = short + ' ' + str(edate) + ext
    move_to_day_folder(file_id, edate, '0BzlU44AWMToxVU8ySkNBQzJQeFE') # Moves to date folder within "Processed" folder
    rename_file(file_id, newname)
    # move_to_folder(file_id, '0BzlU44AWMToxVU8ySkNBQzJQeFE') # Moves to "Processed" folder
    file_clean(filename)

# """Tedexis"""
def tedexis(filename, root, source, edate, upload_list):
    file_id = find_file_id_using_parent(filename, '0BzlU44AWMToxZnh5ekJaVUJUc2c') # Looks in "Files" folder
    filename1 = format().excel_format(filename, source, 0, edate)
    if filename1 == -1:
        move_to_folder(file_id, '0BzlU44AWMToxeFhld1pfNWxDTWs') # Moves to "NoRates" folder
        return 'No rate in document.'
    filename2 = format().excel_filter(filename1)
    bst().source_build(root, filename2)
    index = filename.rfind('.')
    short = filename[:index]
    index = len(filename) - index
    ext = filename[-index:]
    newname = short + ' ' + str(edate) + ext
    move_to_day_folder(file_id, edate, '0BzlU44AWMToxVU8ySkNBQzJQeFE') # Moves to date folder within "Processed" folder
    rename_file(file_id, newname)
    file_clean(filename)

# """Agile Telecom"""
def agile(filename, root, source, edate, upload_list):
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    cd_path = os.getcwd()
    wb = excel.Workbooks.Open(cd_path + '\\' + filename)
    ws = wb.Worksheets(1)
    for shape in ws.Shapes:
        shape.Delete()
    print("Deleted all images from %s" % filename)
    ws.Rows(ws.UsedRange.Rows.Count).Delete()
    print("Deleted last row")    
    wb.Save()
    print("Saved %s" % filename)
    excel.Quit()
    
    file_id = find_file_id_using_parent(filename, '0BzlU44AWMToxZnh5ekJaVUJUc2c') # Looks in "Files" folder
    filename1 = format().excel_format(filename, source, 0, edate)
    if filename1 == -1:
        move_to_folder(file_id, '0BzlU44AWMToxeFhld1pfNWxDTWs') # Moves to "NoRates" folder
        return 'No rate in document.'
    bst().source_build(root, filename1)
    index = filename.rfind('.')
    short = filename[:index]
    index = len(filename) - index
    ext = filename[-index:]
    newname = short + ' ' + str(edate) + ext
    move_to_day_folder(file_id, edate, '0BzlU44AWMToxVU8ySkNBQzJQeFE') # Moves to date folder within "Processed" folder
    rename_file(file_id, newname)
    file_clean(filename)

# """General Use Case"""
# support for mmd, UPM, wavecell, centro, mitto, bics, openmarket, kddi, horisen, calltrade
def general(filename, root, source, edate, upload_list):
    file_id = find_file_id_using_parent(filename, '0BzlU44AWMToxZnh5ekJaVUJUc2c') # Looks in "Files" folder
    filename1 = format().excel_format(filename, source, 0, edate)
    if filename1 == -1:
        move_to_folder(file_id, '0BzlU44AWMToxeFhld1pfNWxDTWs') # Moves to "NoRates" folder
        return 'No rate in document.'
    bst().source_build(root, filename1)
    index = filename.rfind('.')
    short = filename[:index]
    index = len(filename) - index
    ext = filename[-index:]
    newname = short + ' ' + str(edate) + ext
    move_to_day_folder(file_id, edate, '0BzlU44AWMToxVU8ySkNBQzJQeFE') # Moves to date folder within "Processed" folder
    rename_file(file_id, newname)


# """ ------------------------------------------- MAIN CODE HERE --------------------------------------------------------------------------------------------"""
def main():

    general_dictionary = ['MMDSmart', 'UPM Telecom', 'OpenMarket', 'Wavecell', 'Bics', 'Mitto AG', 'C3ntro Telecom', 'HORISEN', 'KDDI Global']
    special_dictionary = ['Tedexis', 'Monty Mobile', 'Tata Communications', 'Silverstreet', 'CLX Networks', 'Agile Telecom']

    # title = [0000000000000000000, 'Country', 'Network', 'MCC', 'MNC', 'MCCMNC', 'Rate', 'CURR', 'Converted Rate', 'Source', 'Effective Date', 0]
    title = [0000000000000000000, 'Country', 'Network', 'MCC', 'MNC', 'MCCMNC', 'Rate', 'CURR', 'Converted Rate', 'Source', 'Effective Date', 'Price Change']
    header = bst().node(title[0], title)

    # """Folder ID is for Test Files Folder"""
    dl_list = dl_folder('0BzlU44AWMToxZnh5ekJaVUJUc2c')
    
    if len(dl_list) == 0:
        print "No new files to be processed."
        print "\nSource_Compiler has succesfully run to completion.\n\n\n"
        return
    else:
        print "\nDownload list is: ", dl_list

    # """Catches already processed files and modifies name so that emails can be found from the filenames"""
    for i in range(len(dl_list)):
        name = dl_list[i]
        index = name.rfind('.')
        hyphen1 = index - 3
        hyphen2 = index - 6
        if name[hyphen1] == '-' and name[hyphen2] == '-':
            date_removed = name[:index-11]
            ext = name[index:]
            new_name = date_removed + ext
            dl_list[i] = new_name
            os.rename(name, new_name)
            file_id = find_file_id(name)
            rename_file(file_id, new_name)

    company_list = get_email_attachment_list(dl_list)

    if company_list==None:
        company_list=[]
        print ("No 'New' messages in the Inbox")
        return
    else:
        print "Email attachment list is: ", company_list 

    if len(company_list) != len(dl_list):
        print ("Not all files downloaded for processing were located as an attachment in the emails.  'New' label status of email may have been removed.")

    index = len(company_list) - 1
    check_date = company_list[index][3]
    # as long as there is something in the company list

    # first build of database here
    bst().database_build(header, check_date, client) 
    upload_list = []

    while len(company_list) >= 0:
        
        try:
            file_to_process = company_list.pop()
        except IndexError:
            print ("No more files to be processed")
        # date change enters into if statement and builds last case
        if check_date != file_to_process[3] or len(company_list) == 0:
            #write to document here
            bst().write(header, check_date, client)
            while check_date <= file_to_process[3]:
                check_date = check_date + timedelta(days=1)
                bst().database_build(header, check_date, client)
                if check_date == file_to_process[3]:
                    break
                else:
                    bst().write(header, check_date, client)

        processed = False
        print "\nFile currently being processed is: ", file_to_process[0]
        print "Remaining number of files to be processed is: ", len(company_list)

        for j in range(len(general_dictionary)):
            # """General use case scenario"""
            if file_to_process[1] == general_dictionary[j]:
                index = file_to_process[0].rfind('.')
                index = len(file_to_process[0]) - index
                ext = file_to_process[0][-index:]
                if ext == '.xls' or ext == '.xlsx':
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
                    status = monty(file_to_process[0], header, file_to_process[1], file_to_process[3], upload_list)
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
                    processed = True
                # """CLX Networks"""
                elif file_to_process[1] == special_dictionary[j] and j == 4:
                    status = clx(file_to_process[0], header, file_to_process[1], file_to_process[3], upload_list)
                    print 'Status of: ', file_to_process[0], ' is: ', status
                    processed = True
                # """Agile Telecom"""
                elif file_to_process[1] == special_dictionary[j] and j == 5:
                    status = agile(file_to_process[0], header, file_to_process[1], file_to_process[3], upload_list)
                    print 'Status of: ', file_to_process[0], ' is: ', status
                    processed = True            
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
            file_id = find_file_id_using_parent(file_to_process[0], '0BzlU44AWMToxZnh5ekJaVUJUc2c')
            move_to_day_folder(file_id, file_to_process[3], '0BzlU44AWMToxOGtyYWZzSVAyNkE') # Moves to date folder in "NotProcessed"
            os.remove(file_to_process[0])

        check_date = file_to_process[3]

    # BUILDS TO CURRENT DAY
    while check_date < date.today():
        check_date = check_date + timedelta(days=1)
        bst().database_build(header, check_date, client)
        bst().write(header, check_date, client)

# def main():
#     title = [0000000000000000000, 'Country', 'Network', 'MCC', 'MNC', 'MCCMNC', 'Rate', 'CURR', 'Converted Rate', 'Source', 'Effective Date', 'Price Change']
#     header = bst().node(title[0], title)
#     bst().database_build(header, date.today() - timedelta(days=1), client)
#     bst().write(header, date.today() - timedelta(days=1), client)
#     print ("all done")

if __name__ == '__main__':
    main()
