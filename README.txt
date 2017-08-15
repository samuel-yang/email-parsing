Documentation for Aggregator Compiler Project
________________


Purpose
Compile rates of various messaging routes from a variety of aggregators, across multiple file formats, currencies and other variables and generate a single file that easily displays all data.
Setup
1. Download the script file, located at https://github.com/samuel-yang/email-parsing in the /dist directory into the desired directory.  This will be the location where local files will be saved to.  Also required will be the client_secret.json file.
2. Launch the script, which should start a browser session, requesting a log-in for the files.  Find details of the log-in information in the e-mails.  
3. Shared rate sheet access can be found at https://drive.google.com/drive/folders/0BzlU44AWMToxNEtxSWROcjkzYVE?usp=sharing 
4. Editable documents and changes must be uploaded as workbooks in the following folder https://drive.google.com/open?id=0BzlU44AWMToxYmdRR1hHVXJiQ1E 
How it works
1. E-mails received with file attachments that detail the various rates from aggregators.
2. Attachments are automatically pulled into Google Drive every hour, using a Drive add-on.
3. Files are downloaded locally, to be processed by Source_Compiler.py
4. File names are referenced against the inbox messages to determine the sender of each file.
5. Sender is referenced against a google sheet that specifies Source Name by e-mail.
6. Date of e-mail, source, filename are handed off for processing.
7. If new date detected, rate sheet is generated for the last date and written to drive. 
8. Updates to current day.  
9. Rate sheets are saved locally, and uploaded to the Drive.


Excel documents used for reference are in Compiled Data.  - DO NOT DELETE
Sheets for viewing are found in Rate Sheets.  
Relevant Files:
        Source_Compiler.py
        Database_Manipulation.py(c)
        CurrencyConverterNew.py(c)
        Google_API_Manipulation.py(c)
        Email_Notifications.py(c)
        write_log.py(c)
________________


Source_Compiler.py
Main program that handles all the processing.  All other files only contain 
methods that are called upon by this program.
________________


Special Cases:
The following methods are used for processing files from specific
providers.  These providers required special formatting and cases and required individual processing, rather than general scenarios.
agile(filename, root, source, edate, upload_list, change_header):
        Agile is a special case scenario where the file needs to be processed prior to formatting.  The method takes filename in this WINDOWS only pre-processing method, as it uses the Win32 package to access the file, delete an image, and then saves the excel workbook so that it can be formatted correctly.  This necessity arises from the fact that the cross-platform formatting methods are unable to handle Agile’s files due to their size.  The file is formatted, using edate as the effective date written into the document, and using a file id based lookup for Google Drive, it is then moved to the processed folder, and nested within the corresponding date’s file folder.  All routes are added as individual nodes in a binary search tree (BST) with its root being root.  All changed rates are added to a separate BST with root change_header.  It then renames the file, and removes it from local storage. 
Returns a string with filename and the status if processed correctly.  


calltrade(filename, root, source, edate, upload_list, change_header):
Inputs: (String, node, String, Date object, list, node)
Calltrade includes various currencies in their provided rate sheets.  This requires a unique formatting method in format().calltrade() that takes in filename to correctly build the nodes for the binary search tree (BST).  After, using edate as the effective date written into the document, and using a file id based lookup for Google Drive, it is then moved to the processed folder, and nested within the corresponding date’s file folder.  All routes are added as individual nodes in a binary search tree (BST) with its root being root.  All changed rates are added to a separate BST with root change_header.  It then renames the file, and removes it from local storage. Returns a string with filename and the status if processed correctly.  


clx(filename, root, source, edate, upload_list, change_header):
Inputs: (String, node, String, Date object, list, node)
        CLX documents come as an ‘.xlsx’ or ‘.xls’ extension, but are actually formatted as ‘.tsv’ or tab separated values.  An additional set of methods [convert().excel_tsv_to_csv() and convert().csv_to_excel()] are required to convert filename from ‘.tsv’ to ‘.csv’ and ultimately back to an excel document.  The file is formatted, using edate as the effective date written into the document, and using a file id based lookup for Google Drive, it is then moved to the processed folder, and nested within the corresponding date’s file folder.  All routes are added as individual nodes in a binary search tree (BST) with its root being root.  All changed rates are added to a separate BST with root change_header.  It then renames the file, and removes it from local storage. Returns a string with filename and the status if processed correctly.  




identidad(filename, root, source, edate, upload_list, change_header, wholesale_header):
Inputs: (String, node, String, Date object, list, node, node)
        Identidad provides both premium and wholesale route documents.  Both files are given to the formatting method through the filename, using edate as the effective date.  Using a file id based lookup for Google Drive, it is then moved to the processed folder, and nested within the corresponding date’s file folder.  Premium routes are added as individual nodes in a binary search tree (BST) with its root being root, while wholesale routes are added to wholesale_header.  The distinction is made based on a ‘wholesale_name’ identifier in the method.  All changed rates are added to a separate BST with root change_header.  It then renames the file, and removes it from local storage. Returns a string with filename and the status if processed correctly.  








mitto(filename, root, source, edate, upload_list, change_header, wholesale_header):
Inputs: (String, node, String, Date object, list, node, node)
Mitto provides both premium and wholesale route documents.  Both files are given to the formatting method through the filename, using edate as the effective date.  Using a file id based lookup for Google Drive, it is then moved to the processed folder, and nested within the corresponding date’s file folder.  Premium routes are added as individual nodes in a binary search tree (BST) with its root being root, while wholesale routes are added to wholesale_header.  The distinction is made based on a ‘wholesale_name’ identifier in the method.  All changed rates are added to a separate BST with root change_header.  It then renames the file, and removes it from local storage. Returns a string with filename and the status if processed correctly.  
monty(filename, root, source, edate, upload_list, change_header):
Inputs: (String, node, String, Date object, list, node, node)
        Monty takes the filename and converts it using convert().csv_to_excel() to an excel format.  The file is then formatted, using edate as the effective date written into the document, and using a file id based lookup for Google Drive, it is then moved to the processed folder, and nested within the corresponding date’s file folder.  All routes are added as individual nodes in a binary search tree (BST) with its root being root.  All changed rates are added to a separate BST with root change_header.  It then renames the file, and removes it from local storage. Returns a string with filename and the status if processed correctly.  
silverstreet(filename, root, source, edate, upload_list, change_header):
Inputs: (String, node, String, Date object, list, node, node)
        Silverstreet requires an additional catch process that takes in filename and eliminates invalid values from the document.  The file is then formatted, using edate as the effective date written into the document, and using a file id based lookup for Google Drive, it is then moved to the processed folder, and nested within the corresponding date’s file folder.  All routes are added as individual nodes in a binary search tree (BST) with its root being root.  All changed rates are added to a separate BST with root change_header.  It then renames the file, and removes it from local storage. Returns a string with filename and the status if processed correctly. 
tata(filename, root, source, edate, upload_list, change_header):
Inputs: (String, node, String, Date object, list, node, node)
        Tata’s route information is stored on the second sheet in the workbook.  This has to be specified prior to further processes.  The file is then formatted, using edate as the effective date written into the document, and using a file id based lookup for Google Drive, it is then moved to the processed folder, and nested within the corresponding date’s file folder.  All routes are added as individual nodes in a binary search tree (BST) with its root being root.  All changed rates are added to a separate BST with root change_header.  It then renames the file, and removes it from local storage. Returns a string with filename and the status if processed correctly. 
tedexis(filename, root, source, edate, upload_list, change_header):
Inputs: (String, node, String, Date object, list, node, node)
The file filename is formatted, using edate as the effective date written into the document.  Tedexis is filtered post-formatting to remove erroneous routes, and using a file id based lookup for Google Drive, it is then moved to the processed folder, and nested within the corresponding date’s file folder.  All routes are added as individual nodes in a binary search tree (BST) with its root being root.  All changed rates are added to a separate BST with root change_header.  It then renames the file, and removes it from local storage. Returns a string with filename and the status if processed correctly. 


General Cases:
This method is used for all other aggregators, it should also be easily able to handle most new providers.  New providers should be tested on a development version, and then moved to a full production version.  
general(filename, root, source, edate, upload_list, change_header):
Inputs: (String, node, String, Date object, list, node, node)
The file filename is formatted, using edate as the effective date written into the document, and using a file id based lookup for Google Drive, it is then moved to the processed folder, and nested within the corresponding date’s file folder.  All routes are added as individual nodes in a binary search tree (BST) with its root being root.  All changed rates are added to a separate BST with root change_header.  It then renames the file, and removes it from local storage. Returns a string with filename and the status if processed correctly. 


main():
Main method is run every 30 minutes or 1800 seconds, using pythons sleep function.
Important variables listed first, followed by psuedo code below.


SOURCE DICTIONARIES
All sources that are not being processed using the template, must be listed in the following:
general_dictionary
general_dictionary = (List) of (String) arguments, arguments are verified compatible provider general case sources, as found in the "Aggregator Source Sheet" in Google Drive, that is downloaded as an xlsx document for local use.
special_dictionary
special_dictionary = (List) of (String) arguments, arguments are verified compatible special case sources, as found in the "Aggergator Source Sheet" in Google Drive.  


Psuedo code
1. Downloads currency exchange workbook if non existent.  This workbook contains all rates for specified currencies from the initialization of this program to current day. 
2. Generates 3 root nodes, a premium and wholesale root with the data list in the same format, and an additional change root that tracks changes in pricing.
3. Downloads files using dl_folder(folder_id), which returns a list of all files downloaded.  
   1. If no files downloaded, program completes and sleeps.
   2. Since all files are renamed with the effective date appended into the name, a check is performed to rename these files back to their original if they are being re-processed.  
1. List of downloaded files is given to get_email_attachment_list(dl_list) which returns a list of lists of all necessary information.
   1. Files are searched for in the attachments of all messages marked ‘NEW’ in the inbox.  A check is performed to see if there are no messages that match this requirement.  (Should only occur when re-running the program on already processed files) Completes the program and puts it to sleep.
1. From the oldest effective date that is returned above, program attempts to download the last 7 days of rates.  This accounts for downtime scenarios, where no new rate sheets are generated for a period of time.  If no files are found that match this criteria, it creates a new workbook, with the date associated with the file.  
2. Binary search tree (BST) is built using the most current (with respect to the oldest file being processed) rate sheet.  
3. List is removed one at a time from the returned list from step 3, file sources are compared to general_dictionary and special_dictionary to determine if the file is able to processed.  If file cannot be processed, a message is generated and logged and file is moved into a corresponding Not Processed folder.  
   1. Each file’s effective date is checked with regards to the previously processed document.  If the effective date has changed, that rate sheet is written to, and uploaded to the drive.  The new date’s database is built on top of the new existing data, so that all data is preserved.
1. If the date of the newest file is behind the date the program is being processed, it builds to current day to update all records.
2. Currency exchange sheet is uploaded with new information. 
3. Sleeps for 30 minutes or 1800 seconds.
________________
Database_Manipulation.py
Contains numerous methods for manipulating the data, from formatting files correctly, to building them into nodes and adding them to the binary search tree.   Additionally, contains file conversion methods. 
Important variables
currency_dictionary: contains all possible strings associated with particular currency types.
currency_list: contains all currently supported currencies
BST
Class node():
Class node is used to create node structures that contain certain values.  Nodes all contain a key, which is a hash of a string including their country, network, mcc, mnc, and mccmnc.  Nodes contain data, which throughout this program, is stored in the form of lists.  Each node also contains a left child and right child.  
size(self, root):
Input: (node)
        bst().size takes a root node and calculates the number of nodes that exist in that binary search tree (BST) which it then returns as an integer. 
changes_to_email(self, notify_list, email, column_title):
Input: (list, String, list)
        This method builds a workbook from notify_list with column headings column_title and saves it to an excel document with the name ‘Changes.xlsx’. 
database_build(self, root, edate, change_root, wholesale_root):
Input: (node, Date object, node, node)
        Database_build takes in edate and attempts to locate a rate sheet for that day locally.  It then builds the premium BST with root root, the wholesale BST with root wholesale_root and the price changes BST with root change_root.  Price changes are only added to the BST if the effective date of that route is greater than current day.  
insert(self, root, node, change_root):
Input: (node, node, node)
        Inserts nodes into the BST given by root.  If price change is detected, a new node with the same information is used to form a new node, whose list structure is different, which is then passed off to insert_new() along with change_root.
insert_new(self, root, node):
Input: (node, node)
        Inserts node into root, but has no additional checks.  Is only used to insert into the price change BST. 
insert_price(self, root, node, notify_list):
Input: (node, node, list)
        Inserts nodes from the pricing sheet into the price change BST.  At the point of insertion, all new prices should have already populated this BST, allowing comparisons of previous price changes.  If the profit margin drops below .2, it is appended to the notify_list.  
in_order_print(self, root):
Input: (node)
        Prints out all the nodes in a BST, starting from the lowest node.key value to the highest. 
pre_order_print(self, root):
Input: (node)
        Prints out all the nodes in a BST, printing them in the order that they were added.
price_build(self, root, filename):
Input: (node, String)
        Builds a binary search tree onto root using the route data from filename.  Generally associated with the pricing sheet that is provided.  
source_build(self, root, filename, change_root):
Input: (node, String, node)
        Takes filename, and extracts the data to create new nodes, which are then built onto the root given by root.  Change_root is passed along as an argument into the insert() methods. 
to_database(self, root, templist):
Input: (node, list)
        Takes in a binary search tree given by root, and appends each one’s data values onto the templist so that it can be written to the document. 
write(self, root, edate, wholesale_root):
Input: (node, Date object, node)
        Takes in both the premium root and wholesale_root, and writes all information to an excel document with the name generated from the edate.  The method then uploads the document as a google sheet.
write_price(self, change_root, wholesale_root):
Input: (node, node)
        Works off of the change_root and wholesale_root to build the pricing sheet, as well as building notify list.  
Convert
csv_to_excel(self, filename):
Input: (String)
        Converts csv file with filename to an excel file with extension ‘.xls’
excel_to_csv(self, filename):
Input: (String)
        Converts excel file with filename to a csv file.
excel_tsv_to_csv(self, filename):
Input: (String)
        Converts excel in formatted tsv to a csv file. 
Format
calltrade(self, filename, edate):
Input: (String, Date object)
        Calltrade is a specialized formatting method for workbooks from calltrade.  It catches the multiple currency types that Calltrade sends us in a single document and processes the rates correctly so that they can then be further formatted and entered into the database.
excel_filter(self, filename):
Input: (String)
        This method takes in a workbook given by filename and removes all empty rows from the file that is found within the data cells.  Some providers provide data with errors or that is incomplete, causing these errors.  When complete it saves the already formatted document by overwriting it. 
excel_format(self, filename, source, sheetindex, edate):
Input: (String, String, integer, Date object):
        This is the main method used for isolating column data for each aggregator and remapping it to be in a single format that can then be processed by the program.  Filename is taken for the original file and the information is extracted as best as populated to populate a new document that is labeled with an appended ‘FORMATTED’.  This method calls upon the separator() method to populate the MCC, MNC, and MCCMNC columns if the data is incomplete. The source given is used to label all of these routes as being made available from that provider.  Sheetindex determines which sheet of the document by index to be looking at to extract the data.  Lastly, the edate is used as the effective date for the data and it is populated into this new workbook to reflect that.  
monty_is_special(self, filename, og_file, edate):
Input: (String, String, Date object)
        Monty rates are provided in EUR, but the document makes no such note of this.  This method is pre-applied to Monty documents so that the rates are not mistaken for USD.  The edate is used to determine the date of which the conversion rate must be looked at.
quicksort(self, to_sort, index):
Input: (list, integer)
        Quicksort sorts the a list and puts it in order, to_sort is a list of lists, index is the item in each list to use to sort.
separator(self, cell_val):
Input: (cell)
        Separator is used to extract the string data from files that are currently being formatted to parse out the MCC and MNC if only the combined MCCMNC is given.  
file_clean(filename):
Input: (String)
        Filename is parsed to remove the extension, and then all possible file versions of that document are checked to see if they exist, and then removed from the local drive. 
________________


Google_API_Manipulation.py
get_credentials():
        Gets valid user credentials from storage. If nothing has been stored, or if the stored credentials are invalid, the OAuth2 flow is completed to obtain the new credentials.  The credentials are saved at the directory ~/.credentials/googleapis.com-python.json                Returns: Credentials, the obtained credential.


initialize_drive_service():
Initializes a Google Drive service instance.
Returns: drive_service, a Google Drive service instance.
initialize_gmail_service():
Initializes a Gmail service instance.
Returns: gmail_service, a Gmail service instance.
initialize_sheets_service():
Initializes a Google Sheets service instance.
Returns: sheets_service, a Google Sheets service instance.
create_folder(name):
Creates a folder in Google Drive.
Args: 
    name: name of the folder.
delete_file(file_id):
Permanently delete a file from Google Drive using file ID, skipping the trash.
Args:
    file_id: ID of the file to delete. 
clean_folder(folder_id):
Permanently deletes files from a Google Drive folder using folder ID, skipping the trash. Only leaves .xls, .xlsx, .csv, and .pdf files, deleting all other file types, including Google app files (sheets, docs, slides, etc.).
Args:
    folder_id: ID of the folder to delete from.
rename_file(file_id, newname):
Renames a file in Google Drive using file ID.
Args:
    file_id: ID of the file to rename.
    newname: new name of the file.
rename_file_using_name(filename, newname):
Renames a file in Google Drive using file name.  Assumes there is only one file with that name in the entire Drive.
Args:
    filename: name of the file to rename.
    newname: new name of the file.
dl_file(file_id, file_name):
Downloads a non-Google app file from Google Drive.
Args:
    file_id: ID of the file to download.
    file_name: name of the file to download.
export_sheet(file_id):
Downloads a Google sheet from Google Drive as an Excel file.
Args:
    file_id: ID of the file to download.
dl_folder(folder_id):
Downloads files from a folder from Google Drive. Only downloads .xls, .xlsx, .csv, and .pdf files.
Args:
    folder_id: ID of the folder to download from.
Returns:  a list of all downloaded files
get_filenames_in_folder(folder_id):
Gets a list of file names from a folder from Google Drive.  Only records .xls, .xlsx, .csv, and .pdf files.
Args:
    folder_id: ID of the folder to download from.
Returns:  file_list: list of file names.
find_file_id(filename):
Gets file ID of a file in Google Drive using file name. Assumes there is only one file or folder with that name in the entire Drive.
Args:
    filename: name of the file.
Returns:  file_id, file ID of the file.
find_file_id_using_parent(filename, parent_id):
Gets file ID of a file in Google Drive using file name and parent ID.  Assumes there is only one file with that name in the specified folder.
Args:
    filename: name of the file.
    parent_id: ID of the parent.
Returns: file_id, file ID of the file.
find_file_name(file_id):
Gets name of a file in Google Drive using its ID.
Args:
    file_id: ID of the file.
Returns: file_name, name of the file.
move_to_folder(file_id, folder_id):
Moves a file to a folder in Google Drive using file ID.
Args:
    file_id: ID of the file to move.
    folder_id: ID of the destination folder.
move_to_folder_using_name(filename, folder_id):
Moves a file to a folder in Google Drive using file name. Assumes there is only one file with that name in the entire Drive.
Args:
    filename: name of the file to move.
    folder_id: ID of the destination folder.
upload_as_gsheet(file_to_upload, filename):
Uploads a file as a Google Sheet to Google Drive. Default path is working directory. Takes Excel files (.xls, .xlsx) only.
Args:
    file_to_upload: full path of the file to upload, including extension.
    filename: name of the file to be displayed on Google Drive.
get_email(ind):
Gets an email message resource from the Gmail inbox.  It only looks at emails in the inbox with the label "New".
Args:
    ind: index of the email.
Returns: mail, an email resource.
get_email_date(ind):
Gets the date of an email message from the Gmail inbox. It only looks at emails in the inbox with the label "New".
Args:
    ind: index of the email.     
Returns: date, the date of the email.


get_email_sender(ind):
Gets the sender of an email message from the Gmail inbox. It only looks at emails in the inbox with the label "New".
Args:
    ind: index of the email.        
Returns: sender, the source of an email and the email address.
get_email_attachment(ind):
Gets the name of an attachment of an email message from the Gmail inbox.  It only looks at emails in the inbox with the label "New".
Args:
    ind: index of the email.
Returns: file, a list containing the file name of the attachment of an email.
part_id(part):
Finds file name of an attachment using partId.
Args:
    part: dictionary "part" of the email.
Returns: filename, the name of the file.
part_find(part):
Recursive method that finds file name of an attachment using partId, looking through parts.
Args:
    part: dictionary "part" of the email.
Returns: filename, the name of the file.
get_email_attachment_list(dl_list):
Gets a list of attachments from Google Drive. It only looks at emails in the inbox with the label "New".
Args:
    dl_list: list of downloaded files.
Returns: attach_list, a list of lists with each sublist containing the file name, source, email address, and date sent.
find_source_from_email(email_string):
Finds the source of an email by using the email address. Uses the following spreadsheet to determine the source:
https://docs.google.com/spreadsheets/d/1rJlhCxJIy1DyYlzp8G9aVap505QBwxcTmiH9zleZzG4
Args:
    email_string: string containing the sender name and email address in the following format:
    "Sender Name" <email@email.com>
Returns: source, the name of the source.
convert_date_email(date):
Returns a date-time object using a string. Used for the email methods.
Args:
    date: string of the date to be converted. Format example: 20 Jun 2017
Returns: date_obj, the date-time object.    
convert_date(date):
Returns a date-time object using a string.
Args:
    date: string of the date to be converted. Format example: 2017-06-20
Returns: date_obj, the date-time object.    
move_to_day_folder(file_id, datetime_obj, parent_id):
Organizes the files according to their respective "day" folders and places the day folder into the parent folder. If a day folder does not exist, it will be created. Only looks at files in the "Files" folder.
Args:
    file_id: string of the ID of the file to move.
    datetime_obj: date of the email as date-time object.
    parent_id: string of the ID of the parent folder.
move_to_day_folder_using_names(filename, datetime_obj, parent_name):
Organizes the files according to their respective "day" folders and places the day folder into the parent folder. If a day folder does not exist, it will be created. Only looks at files in the "Files" folder.
Args:
    filename: string of the name of the file to move.
    datetime_obj: date of the email as date-time object.
     parent_name: string of the name of the parent folder.
remove_label(ind):
Removes the "New" label and adds the "Processed" label.
Args:
    ind: Index of the message from which the attachments have been pulled.
format_cell_alignment(sheet_id):
Formats the alignment of a spreadsheet cell.
Args:
    sheet_id: ID of the sheet.
conditional_format(spreadsheet_id):
Conditionally formats the colors of spreadsheet rows based on a value in a cell.
Args:
    spreadsheet_id: ID of the sheet.
freeze_first_row(spreadsheet_id, rowCount):
Freezes the first row of the spreadsheet.
Args:
    spreadsheet_id: ID of the sheet.
    rowCount: number of rows.
unfreeze_first_row(spreadsheet_id, rowCount):
Unfreezes the first row of the spreadsheet.
Args:
    spreadsheet_id: ID of the sheet.
    rowCount: number of rows.
upload_excel(file_name):
Uploads an Excel (.xlsx) file to Drive. Default path is working directory. Takes Excel files (.xlsx) only.
Args:
    file_name: full name of the file including extension.
upload_log(file_name):
Uploads a log (.log) file to Drive. Default path is working directory. Takes plain text files only.
Args:
    file_name: full name of the file including extension.
________________


CurrencyConverterNew.py
get_currency(currency, rate_in, date):
Finds conversion rate of currency to rate_in for specified date.
Args:
    Currency: String representing currency
    Rate_in: String - in this instance ‘USD’
    Date: date object
Returns: conversion rate for specified date as a float
get_rates(rate_in, date):
Finds conversion rates for all currencies listed in base_currencies to the one given by rate_in.
Args:
    Rate_in: String - ‘USD’
    Date: date object
Returns: list of floats, corresponding to conversion rates. 
get_rate_for_date(checkdate):
Given a date by checkdate, it will search through an excel sheet that stores all retrieved conversion rates.  If one is not located, it will call upon other methods to calculate the conversion rates for that day.
Args:
    Checkdate: date object
Returns: list of floats, corresponding to conversion rates for given checkdate
________________


Email_Notifications.py
get_credentials():
Gets valid user credentials from storage. If nothing has been stored, or if the stored credentials are invalid, the OAuth2 flow is completed to obtain the new credentials. The credentials are saved at the directory ~/.credentials/googleapis.com-python.json
Returns: Credentials, the obtained credential.
initialize_gmail_service():
Initializes a Gmail service instance.
Returns: gmail_service, a Gmail service instance.
create_message(sender, to, subject, message_text):
Create a message for an email.
Args:
   sender: Email address of the sender.
   to: Email address of the receiver.
   subject: The subject of the email message.
   message_text: The text of the email message.
Returns: An object containing a base64url encoded email object.
create_message_with_attachment(sender, to, subject, message_text, file):
Create a message for an email with an attachment.
Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.
    file: The path to the file to be attached.
Returns: An object containing a base64url encoded email object.  
send_message(user_id, message):
Send an email message.
Args:
    user_id: User's email address. The special value "me" can be used to indicate the authenticated user.
    message: Message to be sent.
Returns: Sent Message.