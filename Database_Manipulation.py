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

import xlrd, xlwt, pdfminer, csv, shutil, os, xlutils, sys, openpyxl
# from cstringIO import stringIO
from CurrencyConverter import *
from decimal import *
from Google_API_Manipulation import *
from datetime import *
from xlutils.copy import copy

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

    # """Database build differs from source build in that it extracts cell formatting for certian conditions,
    # to test and see if cells are properly highlighted"""
    # def database_build(self, root, edate):
    #     yesterday = edate - timedelta(days = 1)
    #     filename = 'Data/Rates for ' + str(edate) + '.xls'
    #     if not os.path.isfile(filename):
    #         old_filename = 'Data/Rates for ' + str(yesterday) + '.xls'
    #         if not os.path.isfile(old_filename):
    #             book = xlwt.Workbook()
    #             sheet = book.add_sheet("sheet", cell_overwrite_ok = True)
    #             book.save(filename)
    #         else:
    #             shutil.copy2(old_filename, filename)
    #         print "New file made: ", filename

    #     book = xlrd.open_workbook(filename, formatting_info = True)
    #     sheet = book.sheet_by_index(0)
    #     rownum = sheet.nrows #should be 10
    #     colnum = sheet.ncols
    #     for i in range(rownum-1):
    #         i = i + 1
    #         string = ''
    #         """provider = [hash key, country, network, mcc, mnc, mccmnc, rates, curr, converted rate, source, date, change]"""
    #         provider = [0]
    #         for j in range(colnum):
    #             provider.append(sheet.cell(i,j).value)
    #             if j < 5:
    #                 string = string + str(sheet.cell(i,j).value).encode("utf-8")
    #             else:
    #                 pass

    #         provider[0] = hash(string)
    #         provider.append(0)
    #         if not provider[10] == today:
    #             xfx = sheet.cell_xf_index(i, 7)
    #             xf = book.xf_list[xfx]
    #             bgx = xf.background.pattern_colour_index
    #             ## RED = 10, GREEN = 17
    #             if bgx == 10:
    #                 provider[11] = 1
    #             elif bgx ==17:
    #                 provider[11] = -1

    #         self.insert(root, self.node(provider[0], provider))

    # """Database build NEW. downloads the latest rates sheet from the google drive, and extracts the information,
    # to biuld from there.  If file not found, it pulls form the oldest possible version of the rates sheet"""

    def database_build(self, root, edate):
        filename = 'Rates for ' + str(edate)
        # """Attempts to locate file using the filename in the 'Compiled Data Folder' """"
        file_id = find_file_id_using_parent(filename, '0BzlU44AWMToxYmdRR1hHVXJiQ1E')
        day_before = edate
        days = 0
        new_book = False
        filename_old = filename

        while(file_id == None):
            print ("No file found")
            day_before = day_before - timedelta(days = 1)
            filename_old = 'Rates for ' + str(day_before)
            file_id = find_file_id_using_parent(filename_old, '0BzlU44AWMToxYmdRR1hHVXJiQ1E')
            days = days + 1
            if days > 10:
                #if not os.path.isfile(filename_old + '.xlsx'):
                print ("New Rates sheet created.  Either no previous versions or most recent version is more than 10 days old.")
                book = openpyxl.Workbook()
                filename_old = filename
                book.save(filename_old + '.xlsx')
                new_book = True
                break
            
        if not new_book:
            export_sheet(file_id)

        filename = filename_old + '.xlsx'
        book = openpyxl.load_workbook(filename)
        sheet = book.active
        rownum = sheet.max_row
        colnum = sheet.max_column
        for i in range(rownum-1):
            i = i + 1
            if sheet.cell(row=i+1, column=1).value == None:
                break
            string = ''
            """provider = [hash key, country, network, mcc, mnc, mccmnc, rates, curr, converted rate, source, date, change]"""
            provider = [0]
            for j in range(colnum):
                if j == 7:
                    if provider[7] == 'CURR':
                        provider.append(sheet.cell(row=i+1, column=j+1).value)
                    elif not provider[7] == 'USD':
                        curr = 0
                        for x in range(len(currency_list)):
                            if provider[7] in currency_list[x]:
                                curr = x
                                break

                        converted = currency_rate[curr]*float(provider[6])
                        provider.append(converted)
                    else:
                        provider.append(sheet.cell(row=i+1,column=j+1).value)
                elif j == 9:
                    provider.append(str(sheet.cell(row=i+1, column=j+1).value))
                else:
                    provider.append(sheet.cell(row=i+1, column=j+1).value)
                if j < 5:
                    string = string + str(sheet.cell(row=i+1, column=j+1).value).encode("utf-8")
                else:
                    pass

            provider[0] = hash(string)
            provider.append(0)
            if not provider[10] == today:
                cell_fill = str(sheet.cell(row=i+1, column=8).fill)
                index = cell_fill.rfind('rgb=')+4
                color = cell_fill[index:index+10]
                # """RED = 'FFFF0000' and GREEN = 'FF008000'"""
                if color == 'FFFF0000':
                    provider[11] = 1
                elif color == 'FF008000':
                    provider[11] = -1

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
                    node.data[11] = -1
                    root.data = node.data
                # """Price increased"""
                elif float(root.data[8]) < float(node.data[8]):
                    node.data[11] = 1
                    root.data = node.data
                # """no change"""
                else:
                    pass
                    
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
                    string = string + str(sheet.cell(i,j).value).encode("utf-8")
                else:
                    pass

            provider[0] = hash(string)
            provider.append(0)
            self.insert(root, self.node(provider[0], provider))
        # os.remove(filename)

    """Takes in node, and list.  Builds a pre-order list of node.data and stores in list taken in"""
    def to_database(self, root, templist):
        if not root:
            return
        templist.append(root.data)
        self.to_database(root.l_child, templist)
        self.to_database(root.r_child, templist)

    def write(self, root, edate, upload_list):
        book = xlwt.Workbook(style_compression=2)
        sheet = book.add_sheet("sheet",cell_overwrite_ok=True)
        filename = 'Rates for ' + str(edate) + '.xls'
        final_list = []
        length = 10 #lenght of provider list - 2 (hash key and change value)

        self.to_database(root, final_list)
        for x in range(len(final_list)):
            provider = final_list.pop(0)
            # print len(final_list)
            for k in range(length):
                if x == 0:
                    st = xlwt.easyxf('align: horiz center')
                    sheet.write(x,k,provider[k+1],st)
                else:
                    if k == 5:
                        st = xlwt.easyxf('align: horiz right')
                        sheet.write(x,k,provider[k+1],st)
                    elif k == 7:
                        # price increased
                        if provider[11] == 1:
                            st = xlwt.easyxf('pattern: pattern solid, fore_color red; align: horiz right')
                            sheet.write(x,k,float(provider[k+1]),st)
                            # print "marker 1"
                        # price decreased
                        elif provider[11] == -1:
                            st = xlwt.easyxf('pattern: pattern solid, fore_color green; align: horiz right')
                            sheet.write(x,k,float(provider[k+1]),st)
                            # print "marker 2"
                        else:
                            st = xlwt.easyxf('align: horiz right')
                            sheet.write(x,k,provider[k+1],st)
                            # print "marker 3"
                    else:
                        st = xlwt.easyxf('align: horiz left')
                        sheet.write(x,k,provider[k+1],st)
                        # print "marker 4"

        sheet.col(0).width = 6500
        sheet.col(1).width = 8000
        sheet.col(2).width = 2500
        sheet.col(6).width = 2500 
        sheet.col(8).width = 5000
        sheet.set_panes_frozen(True)
        sheet.set_horz_split_pos(1)
        book.save(filename)

        for x in range(len(upload_list)):
            if str(edate) == upload_list[x]:
                return 'Succesfully written. Data for ', str(edate), 'has alredy been queued to upload.'

        # """If edate isn't found already in list, add it to list to upload"""
        upload_list.append(str(edate))
        return 'Successfully written. Data for ', str(edate), 'is now queued to upload.'

 
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
                # # """Rate"""
                # elif sheet.cell(row,y).value in column_dictionary[6][column_list[6]]:
                #     rate_present = True
                #     for x in range(len(currency_list)):
                #         if sheet.cell(row,y).value in currency_dictionary[x][currency_list[x]]:
                #             i = x
                #             break
                #     for x in range(row+1, rownum):
                #         if sheet.cell(x,y).value == '-':
                #             value = 0
                #         else:
                #             value = sheet.cell(x,y).value
                #         sheet_wr.write(x-row,5,value)
                #         sheet_wr.write(x-row,6,currency_list[i])
                #         # """ Converting currencies here"""
                #         if not currency_list[i] == 'USD':
                #             converted = currency_rate[i]*float(value)
                #         else:
                #             converted = value
                #         sheet_wr.write(x-row,7,converted)
                #         sheet_wr.write(x-row,8, source)
                #         sheet_wr.write(x-row,9, tomorrow)
                # """Modified Rate
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
