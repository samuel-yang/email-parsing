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

import xlrd, xlwt, pdfminer, csv, shutil, os, xlutils, sys
# from cstringIO import stringIO
from CurrencyConverter import *
from decimal import *
from Google_API_Manipulation import *
from datetime import *

reload(sys)
sys.setdefaultencoding('utf-8')

""" Currency Rate List defined here, and called so that it is only called once per program iteration"""
global currency_rate
currency_rate = get_rates()

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
    def database_build(self, filename, root):
         book = xlrd.open_workbook(filename, formatting_info = True)
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
            if not provider[10] == today:
                xfx = sheet.cell_xf_index(i, 7)
                xf = book.xf_list[xfx]
                bgx = xf.background.pattern_colour_index
                ## RED = 10, GREEN = 17
                if bgx == 10:
                    provider[11] = 1
                elif bgx ==17:
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

    def write(self, root, filename):
        book = xlwt.Workbook()
        sheet = book.add_sheet("sheet")
        final_list = []
        length = 10 #lenght of provider list - 2 (hash key and change value)
        self.to_database(root, final_list)
        for x in range(len(final_list)):
            provider = final_list.pop(0)
            # print len(final_list)
            for k in range(length):
                if k == 7:
                    # price increased
                    if provider[11] == 1:
                        st = xlwt.easyxf('pattern: pattern solid, fore_color red;')
                        sheet.write(x,k,provider[k+1],st)
                        # print "marker 1"
                    # price decreased
                    elif provider[11] == -1:
                        st = xlwt.easyxf('pattern: pattern solid, fore_color green;')
                        sheet.write(x,k,provider[k+1],st)
                        # print "marker 2"
                    else:
                        sheet.write(x,k,provider[k+1])
                        # print "marker 3"
                else:
                    sheet.write(x,k,provider[k+1])
                    # print "marker 4"

        book.save(filename)
        print "Successfully written"

 
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
        os.remove(filename)
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
        os.remove(filename)
        return filename1

    """ PDF_TO_CSV takes in a string argument of filename, and returns string with 
        filename to converted document, removes original document."""
    # def pdf_to_csv(self, filename):
        # reference Email_Parser.py for this section


"""Format class contains all formatting methods that will be used"""
class format():
    global column_dictionary, column_list, currency_dictionary, currency_list, currency_rate
    column_dictionary =({'Country': ['Country']},
                        {'Network': ['Network']},
                        {'Country/Network': ['Country/Operator', 'Region/Operator', 'Country/Network']},
                        {'MCC': ['MCC']},
                        {'MNC': ['MNC']},
                        {'MCCMNC': ['MCCMNC', 'Network code']},
                        {'Rate': ['Rate', 'Price', 'New Price', 'New Price(Euro)', 'Price Euro', 'New Price EUR', 'New Price (EUR)', 'Price \nEUR/SMS']})
                        # {'CURR': ['USD', 'EUR' 'GBP', 'CNY', 'MXN']},
                        # {'Source': ['Source']})

    column_list = ['Country', 'Network', 'Country/Network', 'MCC', 'MNC', 'MCCMNC', 'Rate'] # , 'CURR', 'Source']

    currency_dictionary =  ({'USD': ['Rate', 'Price', 'New Price']},
                            {'EUR': ['New Price(Euro)', 'Price Euro', 'New Price EUR', 'New Price (EUR)', 'Price \nEUR/SMS']})

    # """Support only exists for USD, EUR right now, need to define dictionary for others"""
    currency_list = ['USD', 'EUR', 'GBP', 'CNY', 'MXN']

    rate_present = False

    """ EXCEL_FORM takes in both .xls or .xlsx and rearranges the columns to be
        correctly ordered.  takes in filename as string, returns new filename."""
    def excel_format(self, filename, source):
        mccmnc_absent=True
        mcc_absent=True
        mnc_absent=True
        mnc_val=[0]
        mcc_val=[0]
        mccmnc_val=[0]
        book = xlrd.open_workbook(filename)
        sheet = book.sheet_by_index(0)
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
                    print "Rate detected"
                    for x in range(len(currency_list)):
                        if sheet.cell(row,y).value in currency_dictionary[x][currency_list[x]]:
                            i = x
                            break
                    for x in range(row+1, rownum):
                        value = sheet.cell(x,y).value
                        sheet_wr.write(x-row,5,value)
                        sheet_wr.write(x-row,6,currency_list[i])
                        # """ Converting currencies here"""
                        if not currency_list[i] == 'USD':
                            converted = currency_rate[i]*float(value)
                        else:
                            converted = value
                        sheet_wr.write(x-row,7,converted)
                        sheet_wr.write(x-row,8, source)
                        sheet_wr.write(x-row,9, tomorrow)
                else:
                    pass

                
        """ Computing missing MNC, MCC or MCCMNC Values"""
        
        # """MCCMNC is absent"""
        if mcc_absent== False and mnc_absent==False and mccmnc_absent==True:
            for i in range(1,len(mcc_val)):
                if "," not in str(mnc_val[i]) and "/" not in str(mnc_val[i]):
                    ind1=str(mcc_val[i]).index(".")
                    ind2=str(mnc_val[i]).index(".")
                    val=str(mnc_val[i])[:ind2]
                    if len(val)==1:
                        val="0"+val
                    value=str(mcc_val[i])[:ind1]+val                  
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

        index = filename.rfind('.')
        filename1 = filename[:index]
        filename1 = filename1 + ' FORMATTED.xls'
        new_book.save(filename1)
        print "File has been properly formatted."
        # os.remove(filename)
        return filename1

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

##print(seperator(str1))
##print(seperator(str2))



"""-------------------------------------------------------------------------Main Code ------------------------------------------------------------"""
# format().excel_format('test.xls')
# print str(date.today())
# # date.timedelta(days=1)
# tomorrow = date.today() + timedelta(days=1)
