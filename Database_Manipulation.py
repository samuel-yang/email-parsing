"""Imported libraries uses will be listed below:
    xlrd - opening and reading excel workbooks
    xlwt - writing excel workbooks
    xlutils - copying excel documents
    pdfminer - extracting data from pdfs
    csv - manipulating and working with csv documents
    shutil - 
    os - having operating software functions
    cstringIO - stringIO objects
    forex_python.converter - exchange rate converter
    Google_Drive_Manipulation - using the Google Dirve and Sheets API"""

import xlrd, xlwt, pdfminer, csv, shutil, os, xlutils, sys, openpyxl, gspread
# from cstringIO import stringIO
from CurrencyConverter import *
from decimal import *
from Google_API_Manipulation import *
from datetime import *
from xlutils.copy import copy
#from gspread import *

reload(sys)
sys.setdefaultencoding('utf-8')

""" Currency Rate List defined here, and called so that it is only called once per program iteration"""
global currency_rate, currency_dictionary, currency_list
currency_rate = get_rates()

currency_dictionary =  ({'USD': ['Rate', 'Price', 'New Price', 'New Price (USD)', 'Rate - USD']},
                        {'EUR': ['New Price(Euro)', 'Price Euro', 'New Price EUR', 'New Price (EUR)', 'Price \nEUR/SMS', 'Price in EUR']},
                        {'GBP': ['Price in GBP']},
                        {'CNY': []},
                        {'MXN': []},
                        {'AUD': ['Price in AUD']},
                        {'GW': ['GW0', 'GW111']})

# """Support only exists for USD, EUR right now, need to define dictionary for others"""
currency_list = ['USD', 'EUR', 'GBP', 'CNY', 'MXN', 'AUD', 'GW']

#value for date to be entered
global today, tomorrow
tomo = date.today() + timedelta(days = 1)
tomorrow = str(tomo)[-5:]
today = str(date.today())[-5:]

"""BST class is used for building binary search tree datastructure, used to pull from database
    and track data changes"""
class bst():
    class node():
        def __init__(self, hashkey, val):
            self.l_child = None
            self.r_child = None
            self.key = hashkey
            self.data = val

    def size(self, root):
        count = 1
        if root == None:
            return 0
        else:
            count = count + self.size(root.l_child)
            count = count + self.size(root.r_child)
            return count

    # """Database build differs from source build in that it extracts cell formatting for certian conditions,
    # to test and see if cells are properly highlighted, works directly with google sheets"""
    def database_build(self, root, edate, client):
        filename = 'Rates for ' + str(edate)
        # file_id = find_file_id(filename)
        # if file_id != None:
        #     delete_file(file_id)

        day_before = edate
        days = 0
        filename_old = filename
        found = False
        while days <10 and not found:
            try:
                book = client.open(filename_old)
                sheet = book.get_worksheet(0)
                found = True
                print ("Rate sheet for %s found." %str(day_before))
            except gspread.exceptions.SpreadsheetNotFound:
                print ('No sheet for %s found.' %str(day_before))
                day_before = day_before - timedelta(days=1)
                filename_old = 'Rates for ' + str(day_before)
                days = days + 1

        if not found or day_before != edate:
            newbook = client.create(filename) 
            newsheet = newbook.get_worksheet(0)
            file_id = find_file_id(filename)
            format_cell_alignment(file_id)
            print ("New sheet created. Sheet created is for %s" %str(edate))
            if not found:
                return

        rownum = sheet.row_count
        colnum = sheet.col_count
        full = sheet.get_all_values()
        if full == []:
            return
        full.pop(0)
        for i in range(len(full)):
            temp = full.pop(0)
            if temp[0] == '':
                pass
            else:
                provider = [0]
                provider = provider + temp
                string = ''
                if len(temp[3]) == 1:
                    provider[4] = '0' + str(provider[4])                
                for j in range(5):
                    string = string + str(provider[j+1]).decode('utf-8')
                
                provider[0] = hash(string)
                if provider[7] != 'USD':
                    for j in range(len(currency_list)):
                        if provider[7] in currency_list[j]:
                            curr = j
                            break
                    converted = currency_rate[curr]*float(provider[6])
                    provider[8] = converted
            try:
                if provider[11] == '':
                    provider[11] = '-----'
            except IndexError:
                provider.append('-----')

            provider[10] = convert_date(provider[10])
            if provider[10] < edate:
                provider[11] = '-----'
            
            for i in range(len(provider)):
                if i == 0 or i == 8 or i == 10 or i == 11:
                    pass
                else:
                    temp = str(provider[i]).decode('utf-8')
                    provider[i] = temp

            self.insert(root, self.node(provider[0], provider))



    def insert(self, root, node):
        if root is None:
            root = node
        else:
            """if statements are based on hash key of the strings built"""
            if root.key > node.key:
                if root.l_child is None:
                    root.l_child = node
                else:
                    self.insert(root.l_child, node)
            elif root.key < node.key:
                if root.r_child is None:
                    root.r_child = node
                else:
                    self.insert(root.r_child, node)
            elif root.key == node.key:
                if root.data[8] == 'Converted Rate':
                    pass
                # """Comparing rates of various nodes, typecasting to float"""
                # """Price decreased"""
                elif float(root.data[8]) > float(node.data[8]):
                    node.data[11] = 'Decrease'
                    root.data = node.data
                # """Price increased"""
                elif float(root.data[8]) < float(node.data[8]):
                    node.data[11] = 'Increase'
                    root.data = node.data
                # """no change"""
                else:
                    node.data[11] = '------'
                    root.data = node.data
                    
    def in_order_print(self, root):
        if not root:
            return
        self.in_order_print(root.l_child)
        print root.data
        self.in_order_print(root.r_child)

    def pre_order_print(self, root):
        if not root:
            return
        print root.data
        self.pre_order_print(root.l_child)
        self.pre_order_print(root.r_child)

    """Builds BST structure for all sources in filename that is taken in.  Structure built off of 
        root taken in as argument"""
    def source_build(self, root, filename):
        book = xlrd.open_workbook(filename, 'rb')
        sheet = book.sheet_by_index(0)
        rownum = sheet.nrows #should be 10
        colnum = sheet.ncols
        for i in range(rownum-1):
            i = i + 1
            string = ''
            """provider = [hash key, country, network, mcc, mnc, mccmnc, rates, curr, converted rate, source, date, change]"""
            provider = [0]
            for j in range(colnum):
                provider.append(sheet.cell(i,j).value) 
                if j < 5:
                    string = string + str(sheet.cell(i,j).value).decode("utf-8")
                else:
                    pass
            
            provider[0] = hash(string)
            provider[10] = convert_date(provider[10])
            provider.append('-----')
            self.insert(root, self.node(provider[0], provider))

    """Takes in node, and list.  Builds a pre-order list of node.data and stores in list taken in"""
    def to_database(self, root, templist):
        if not root:
            return
        templist.append(root.data)
        self.to_database(root.l_child, templist)
        self.to_database(root.r_child, templist)

    def write(self, root, edate, client):
        filename = 'Rates for ' + str(edate)
        book = client.open(filename)
        sheet = book.get_worksheet(0)
        sheet.clear()
        final_list = []
        self.to_database(root, final_list)
        rowcount = len(final_list)
        sheet.resize(rows=rowcount, cols=11)
        cell_list = sheet.range(1,1,1,10)
        full_update = []
        for i in range(rowcount):
            i = i + 1
            cell_list = sheet.range(i,1,i,11)
            provider = final_list.pop(0)
            index = 1
            for cell in cell_list:
                cell.value = provider[index]
                index = index + 1
            full_update = full_update + cell_list

            #try:
                #sheet.update_cells(cell_list)
            #except gspread.exceptions.RequestError:
                #print('Error entered')
                #provider.pop(0)
                #sheet.insert_row(provider, i)
                #sheet.delete_row(i+1)
        
        sheet.update_cells(full_update)
        #cell_list = sheet.range(1,1,1,11)
        #provider = final_list.pop(0)
        #index = 1
        #for cell in cell_list:
        #    cell.value = provider[index]
        #    index = index + 1
        #sheet.update_cells(cell_list)           
        
        #for i in range(len(final_list)):
        #    rows = i + 2
        #    provider = final_list.pop(0)
        #    provider.pop(0)
        #    sheet.insert_row(provider, rows)
            

        file_id = find_file_id(filename)
        conditional_format(file_id)

"""Convert class performs all file conversions"""
class convert():
    """CSV_TO_EXCEL takes in a string argument of the filename, and returns
        string with filename of converted document, removes original document"""
    def csv_to_excel(self, filename):
        file = open(filename, 'rb')
        read = csv.reader((file), delimiter = ',')
        book = xlwt.Workbook()
        sheet = book.add_sheet("Sheet 1")

        for rowi, row in enumerate(read):
            for coli, value in enumerate(row):
                sheet.write(rowi,coli,value)

        index = filename.rfind('.')
        filename1 = filename[:index]
        filename1 = filename1 + '.xls'
        book.save(filename1)
        file.close()
        # os.remove(filename)
        return filename1    




    """ EXCEL_TO_CSV takes in a string argument of the filename, and returns
        string with filename of converted document, removes original document"""
    def excel_to_csv(self, filename):
        book = xlrd.open_workbook(filename)
        sheet = book.sheet_by_index(0)
        index = filename.rfind('.')
        filename1 = filename[:index] 
        filename1 = filename1 + '.csv'
        csvfile = open(filename1, 'wb')
        write = csv.writer(csvfile, quoting=csv.QUOTE_ALL)

        for row in range(sheet.nrows):
            write.writerow(sheet.row_values(row))

        csvfile.close()
        # os.remove(filename)
        return filename1

    """ EXCEL_TSV_TO_CSV takes in a string argument of the filename, and returns
        string with filename of converted document, removes original document.  Converts
        excel files found in tsv format to csv"""
    def excel_tsv_to_csv(self, filename):
        tsvfile = open(filename, 'rb')
        read = csv.reader(tsvfile, dialect = csv.excel_tab)
        index = filename.rfind('.')
        filename1 = filename[:index]
        filename1 = filename1 + '.csv'
        csvfile = open(filename1, 'wb')
        write = csv.writer(csvfile, dialect = csv.excel)

        for row in read:
            write.writerow(row)

        tsvfile.close()
        csvfile.close()
        # os.remove(filename)
        return filename1

    """ PDF_TO_CSV takes in a string argument of filename, and returns string with 
        filename to converted document, removes original document."""
    # def pdf_to_csv(self, filename):
        # reference Email_Parser.py for this section


"""Format class contains all formatting methods that will be used"""
class format():
    global column_dictionary, column_list
    column_dictionary =({'Country': ['Country', 'CountryName']},
                        {'Network': ['Network', 'OperatorName', 'Operator']},
                        {'Country/Network': ['Country/Operator', 'Region/Operator', 'Country/Network']},
                        {'MCC': ['MCC']},
                        {'MNC': ['MNC']},
                        {'MCCMNC': ['MCCMNC', 'Network code', 'IMSI', 'MCC MNC']},
                        ({'Rate': ['Rate', 'Price', 'New Price', 'New Price(Euro)', 'Price Euro', 'New Price EUR', 
                            'New Price (EUR)', 'Price \nEUR/SMS', 'New Price (USD)', 'Rate - USD', 'Price in GBP',
                            'Price in AUD', 'Price in EUR', 'GW0', 'GW111']}))

    column_list = ['Country', 'Network', 'Country/Network', 'MCC', 'MNC', 'MCCMNC', 'Rate'] # , 'CURR', 'Source']

    # """ excel_filter takes and removes empty rows from a FORMATTED document """
    def excel_filter(self, filename):
        book = xlrd.open_workbook(filename)
        sheet = book.sheet_by_index(0)
        new_book = xlwt.Workbook()
        sheet_wr = new_book.add_sheet("sheet", cell_overwrite_ok = True)

        for i in range(sheet.nrows):
            if sheet.cell(i,0).value == '':
                break
            for j in range(sheet.ncols):
                value = sheet.cell(i,j).value
                sheet_wr.write(i,j,value)

        index = filename.rfind('.')
        filename1 = filename[:index]
        filename1 = filename1 + ' and FILTERED.xls'
        new_book.save(filename1)
        return filename1
        # os.remove(filename)


    # """ EXCEL_FORM takes in both .xls or .xlsx and rearranges the columns to be
    #   correctly ordered.  takes in filename as string, returns new filename."""
    def excel_format(self, filename, source, sheetindex, edate):
        book = xlrd.open_workbook(filename)
        sheet = book.sheet_by_index(sheetindex)
        new_book = xlwt.Workbook()
        sheet_wr = new_book.add_sheet("sheet", cell_overwrite_ok = True)

        """Forcing the header of the excel format"""
        sheet_wr.write(0,0, 'Country')
        sheet_wr.write(0,1, 'Network')
        sheet_wr.write(0,2, 'MCC')
        sheet_wr.write(0,3, 'MNC')
        sheet_wr.write(0,4, 'MCCMNC')
        sheet_wr.write(0,5, 'Rate')
        sheet_wr.write(0,6, 'CURR')
        # sheet_wr.write(0,7, 'Source')
        sheet_wr.write(0,7, 'Converted Rate')
        sheet_wr.write(0,8, 'Source')
        sheet_wr.write(0,9, 'Effective Date')
        """Freezing Header Row"""
        sheet_wr.set_panes_frozen(True)
        sheet_wr.set_horz_split_pos(1)

        rownum = sheet.nrows
        colnum = sheet.ncols

        edate_effective = edate + timedelta(days = 1)
        rate_present = False
        mccmnc_absent=True
        mcc_absent=True
        mnc_absent=True
        mnc_val=[0]
        mcc_val=[0]
        mccmnc_val=[0]


        for row in range(rownum):
            for y in range(colnum):
                # """Country"""
                if sheet.cell(row,y).value in column_dictionary[0][column_list[0]]:
                    for x in range(row+1, rownum):
                        value = sheet.cell(x,y).value
                        sheet_wr.write(x-row,0,value)
                # """Network"""        
                elif sheet.cell(row,y).value in column_dictionary[1][column_list[1]]:
                    for x in range(row+1, rownum):
                        value = sheet.cell(x,y).value
                        sheet_wr.write(x-row,1,value)
                # """Country/Network"""
                elif sheet.cell(row,y).value in column_dictionary[2][column_list[2]]:
                    for x in range(row+1, rownum):
                        value = self.separator(sheet.cell(x,y).value)
                        sheet_wr.write(x-row,0,value[0])
                        sheet_wr.write(x-row,1,value[1])
                # """MCC"""
                elif sheet.cell(row,y).value in column_dictionary[3][column_list[3]]:
                    mcc_absent=False
                    for x in range(row+1, rownum):
                        value = sheet.cell(x,y).value
                        mcc_val.append(value)
                        sheet_wr.write(x-row,2,value)
                # """MNC"""
                elif sheet.cell(row,y).value in column_dictionary[4][column_list[4]]:
                    mnc_absent=False
                    for x in range(row+1, rownum):
                        value = sheet.cell(x,y).value
                        mnc_val.append(value)
                        sheet_wr.write(x-row,3,value)
                # """MCCMNC"""
                elif sheet.cell(row,y).value in column_dictionary[5][column_list[5]]:
                    mccmnc_absent=False
                    for x in range(row+1, rownum):
                        value = sheet.cell(x,y).value
                        mccmnc_val.append(value)
                        sheet_wr.write(x-row,4,value)

                # """Rate"""
                elif sheet.cell(row,y).value in column_dictionary[6][column_list[6]]:
                    rate_present = True
                    for x in range(len(currency_list)):
                        if sheet.cell(row,y).value in currency_dictionary[x][currency_list[x]]:
                            i = x
                            break
                    for x in range(row+1, rownum):
                        if sheet.cell(x,y).value == '-':
                            value = 0
                        else:
                            value = sheet.cell(x,y).value
                        if sheet.cell(x,y).value == '':
                            pass
                        else:
                            sheet_wr.write(x-row,5,value)
                            if currency_list[i] == 'GW':
                                currency = 'USD'      
                                # """Adjust converted value - for GW0 and GW111"""
                                converted = float(str(value)[-4:])/10000
                            elif not currency_list[i] == 'USD':
                                currency = currency_list[i]
                                converted = currency_rate[i]*float(value)
                            else:
                                currency = currency_list[i]
                                converted = value
                            sheet_wr.write(x-row,6,currency)
                            sheet_wr.write(x-row,7,converted)
                            sheet_wr.write(x-row,8,source)
                            sheet_wr.write(x-row,9,str(edate_effective))                                
                else:
                    pass
                
        # """ Computing missing MNC, MCC or MCCMNC Values"""
        
        # # """MCCMNC is absent"""
        if mcc_absent == False and mnc_absent==False and mccmnc_absent==True:
            for i in range(1,len(mcc_val)):
                if "," not in str(mnc_val[i]) and "/" not in str(mnc_val[i]):
                    ind1=str(mcc_val[i]).rfind(".")
                    ind2=str(mnc_val[i]).rfind(".")
                    if ind1 != -1 and ind2 != -1:
                        val=str(mnc_val[i])[:ind2]
                        if len(val)==1:
                            val="0"+val
                        value=str(mcc_val[i])[:ind1]+val                  
                        sheet_wr.write(i,4,value)
                    else:
                        val1 = str(mcc_val[i])
                        val2 = str(mnc_val[i])
                        value = val1 + val2
                        sheet_wr.write(i,4,value)
                else:
                    value=""
                    sheet_wr.write(i,4,value)
                    
        # """MNC and MNC individual columns are absent"""
        if mccmnc_absent==False and mcc_absent== True and mnc_absent==True:
            for i in range(1,len(mccmnc_val)):
                value_mcc=str(mccmnc_val[i])[:3]
                sheet_wr.write(i,2,value_mcc)
                value_mnc=str(mccmnc_val[i])[3:]
                if "." in value_mnc:
                    ind3=value_mnc.index(".")
                    value_mnc=value_mnc[:ind3]
                sheet_wr.write(i,3,value_mnc)

        if not rate_present:
            move_to_day_folder(filename, edate, 'NoRates')
            return -1

        index = filename.rfind('.')
        filename1 = filename[:index]
        filename1 = filename1 + ' FORMATTED.xls'
        new_book.save(filename1)
        print "File has been properly formatted."
        # os.remove(filename)
        return filename1

    # """Monty_is_special - formats the Rate to EUR, as it is not labeled properly """
    def monty_is_special(self, filename, og_file):
        book = xlrd.open_workbook(filename)
        sheet = book.sheet_by_index(0)
        og_book = xlrd.open_workbook(og_file)
        og_sheet = og_book.sheet_by_index(0)
        newbook = xlutils.copy.copy(book)
        sheet_wr = newbook.get_sheet(0)
        for x in range(sheet.nrows):
            if x > 0:
                curr = og_sheet.cell(x,10).value
                sheet_wr.write(x,6,curr)
                rate = sheet.cell(x,5).value
                for i in range(len(currency_list)):
                    if curr in currency_list[i]:
                        j = i
                        break

                converted = currency_rate[j]*float(rate)
                sheet_wr.write(x,7,converted)

        newbook.save(filename)
        return filename

    """ Parses through strings with multiple components and returns list with separated strings"""
    def separator(self, cell_val):
        c=0
        start=False
        start_ind=0
        end_ind=0
        for i in range(len(cell_val)):
            k=cell_val[i].isalpha()
            m=cell_val[i].isspace()
            if i==0 and k==False and m==False:
                start=True
                start_ind=c+1
                c=c+1
            elif k==False and m==False:
                end_ind=c
                if start==False:
                    end_ind=c+1
                break
            c=c+1
        cell_val1=cell_val[end_ind:].strip()
        c1=end_ind
        start_ind1=c1
        for j in range(len(cell_val1)):
            k=cell_val[c1].isalpha()
            m=cell_val[i].isspace()
            if k==True and m==True:
                start_ind1=c1
            else:
                c1=c1+1       
        #country
        str1=cell_val[start_ind:end_ind-1]
        #operator
        str2=cell_val[start_ind1:].strip()
        return [str1,str2]

def file_clean(filename):
    index = filename.rfind('.')
    short = filename[:index]
    if os.path.isfile(short + '.xls'):
        os.remove(short + '.xls')

    if os.path.isfile(short + '.xlsx'):
        os.remove(short + '.xlsx')

    if os.path.isfile(short + '.csv'):
        os.remove(short + '.csv')

    if os.path.isfile(short + ' FORMATTED.xls'):
        os.remove(short + ' FORMATTED.xls')

    if os.path.isfile(short + ' FORMATTED and FILTERED.xls'):
        os.remove(short + ' FORMATTED and FILTERED.xls')

    print "All file versions of ", short, "have been deleted."
##print(seperator(str1))
##print(seperator(str2))



"""-------------------------------------------------------------------------Main Code ------------------------------------------------------------"""
# format().excel_format('test.xls')
# print str(date.today())
# # date.timedelta(days=1)
# tomorrow = date.today() + timedelta(days=1)
