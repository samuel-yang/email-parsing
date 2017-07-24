from apiclient import discovery
from apiclient.http import MediaFileUpload, MediaIoBaseDownload
from apiclient import errors
from oauth2client import client, tools
from oauth2client.file import Storage

import base64, googleapiclient, httplib2, os, io, xlrd, datetime, re, logging

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/googleapis.com-python.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/gmail.modify'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Rates'

def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.
    The credentials are saved at the directory ~/.credentials/googleapis.com-python.json

    Returns:
        Credentials, the obtained credential.
    """
    
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        flow.params['access_type'] = 'offline'
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        logging.info("Storing credentials to " + credential_path)
        print("Storing credentials to " + credential_path)
    return credentials

def initialize_drive_service():
    """Initializes a Google Drive service instance.
    
    Returns:
        drive_service, a Google Drive service instance.
    """    
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())    
    drive_service = discovery.build('drive', 'v3', http=http)    
    
    return drive_service

def initialize_gmail_service():
    """Initializes a Gmail service instance.
    
    Returns:
        gmail_service, a Gmail service instance.
    """        
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    gmail_service = discovery.build('gmail', 'v1', http=http)
    
    return gmail_service

def initialize_sheets_service():
    """Initializes a Google Sheets service instance.
    
    Returns:
        sheets_service, a Google Sheets service instance.
    """        
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    sheets_service = discovery.build('sheets', 'v4', http=http)
    
    return sheets_service

def create_folder(name):
    """Creates a folder in Google Drive.

      Args:
        name: name of the folder. 
      """     
    drive_service = initialize_drive_service()
    
    file_metadata = {
        'name' : name,
      'mimeType' : 'application/vnd.google-apps.folder'
    }
    file = drive_service.files().create(body=file_metadata,
                                        fields='id').execute()
    logging.info("Created folder \"" + name + "\" (ID: %s)" % file.get('id'))
    print("Created folder \"" + name + "\" (ID: %s)" % file.get('id'))  

def delete_file(file_id):
    """Permanently delete a file from Google Drive using file ID, skipping the trash.
    
      Args:
        file_id: ID of the file to delete. 
      """       
    drive_service = initialize_drive_service()
    file_name = find_file_name(file_id)
    
    try:
        drive_service.files().delete(fileId=file_id).execute()
        logging.info("Deleted file: " + file_name)
        print("Deleted file: ", file_name)
    except errors.HttpError, error:
        logging.error("An error occurred during delete attempt: %s" % error)
        print("An error occurred: %s" % error)    

def clean_folder(folder_id):
    """Permanently deletes files from a Google Drive folder using folder ID, skipping the trash.
    
    Only leaves .xls, .xlsx, .csv, and .pdf files, deleting all other file types, including Google app files (sheets, docs, slides, etc.).

      Args:
        folder_id: ID of the folder to delete from.
      """      
    drive_service = initialize_drive_service()  
    parent_id = folder_id
    folder_name = find_file_name(parent_id)
    page_token = None
    delete = False
    
    #Delete Google files
    while True:
        response = drive_service.files().list(q='"%s" in parents and (mimeType contains "google-apps")' % (parent_id), 
                                        spaces='drive', fields='nextPageToken, files(id, name)', pageToken=page_token).execute()
    
        for file in response.get('files', []):
            # Process change
            file_id = file.get('id')
            file_name_full = file.get('name')
            
            #print("Deleted Google file: %s (ID: %s)" % (file_name_full, file_id))
            delete_file(file_id)
            delete = True
                
        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break;    
    
    #Delete files that are not xls, xlsx, csv, or pdf format
    while True:
        response = drive_service.files().list(q='"%s" in parents' % (parent_id), spaces='drive', 
                                        fields='nextPageToken, files(id, name)', pageToken=page_token).execute()
    
        for file in response.get('files', []):
            # Process change
            file_id = file.get('id')
            file_name_full = file.get('name')
            file_name = (".").join(file.get('name').split(".")[:-1])
            extension = file.get('name').split(".")[-1]
            
            if (extension != 'xls') and (extension != 'xlsx') and (extension != 'csv') and (extension != 'pdf'):
                #print("Deleted file: %s (ID: %s)" % (file_name_full, file_id))
                delete_file(file_id)
                delete = True
                
        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break;
    
    if (delete == False):
        logging.info("No files found to remove from " + folder_name + " (ID: " + parent_id + ").")
        print("No files found to remove from " + folder_name + " (ID: " + parent_id + ").")
    else:
        logging.info("Finished cleaning folder from " + folder_name + " (ID: " + parent_id + ").")  
        print("Finished cleaning folder from " + folder_name + " (ID: " + parent_id + ").")     

def rename_file(file_id, newname):
    """Renames a file in Google Drive using file ID.
    
    Args:
        file_id: ID of the file to rename.
        newname: new name of the file.
    """

    drive_service = initialize_drive_service()
    filename = find_file_name(file_id)
    file_metadata = {
        'name' : newname
        }
    
    try:
        file = drive_service.files().update(fileId=file_id, body=file_metadata, fields='id').execute()
        logging.info("File \"{0}\" renamed as: {1} (ID: {2}).".format(filename, newname, file_id))
        print("File \"{0}\" renamed as: {1} (ID: {2}).".format(filename, newname, file_id))
    except TypeError:
        logging.error("Could not rename file.")
        print("Could not rename file.")
        pass

def rename_file_using_name(filename, newname):
    """Renames a file in Google Drive using file name.
    
    Assumes there is only one file with that name in the entire Drive.
    
    Args:
        filename: name of the file to rename.
        newname: new name of the file.
    """

    drive_service = initialize_drive_service()
    file_id = find_file_id(filename)
    file_metadata = {
        'name' : newname
        }
    
    try:
        file = drive_service.files().update(fileId=file_id, body=file_metadata, fields='id').execute()
        logging.info(": File \"{0}\" renamed as: {1} (ID: {2}).".format(filename, newname, file_id))        
        print("File \"{0}\" renamed as: {1} (ID: {2}).".format(filename, newname, file_id))
    except TypeError:
        logging.error("Could not rename file.")
        print("Could not rename file.")        
        pass     
        
def dl_file(file_id, file_name):
    """Downloads a non-Google app file from Google Drive.
    
    Args:
        file_id: ID of the file to download.
        file_name: name of the file to download.
    """    
    drive_service = initialize_drive_service()
    
    request = drive_service.files().get_media(fileId=file_id)
    #fh = io.BytesIO()
    fh = io.FileIO(file_name, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    logging.info("Downloading " + file_name + " (ID: " + file_id + ").")
    print("Downloading " + file_name + " (ID: " + file_id + ").")
    while done is False:
        status, done = downloader.next_chunk()
        logging.info("Download %d%%." % int(status.progress() * 100))
        print("Download %d%%." % int(status.progress() * 100))
     
def export_sheet(file_id):
    """Downloads a Google sheet from Google Drive as an Excel file.
    
    Args:
        file_id: ID of the file to download.
    """
    drive_service = initialize_drive_service()
    file_name = find_file_name(file_id).encode('utf-8')
    
    request = drive_service.files().export_media(fileId=file_id,
                                                 mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    #fh = io.BytesIO()
    fh = io.FileIO(file_name + '.xlsx', 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    logging.info("Downloading " + file_name + " (ID: " + file_id + ").")
    print("Downloading " + file_name + " (ID: " + file_id + ").")
    while done is False:
        status, done = downloader.next_chunk()
        logging.info("Download %d%%." % int(status.progress() * 100))
        print "Download %d%%." % int(status.progress() * 100)    

def dl_folder(folder_id):
    """Downloads files from a folder from Google Drive.
    
    Only downloads .xls, .xlsx, .csv, and .pdf files.
    
    Args:
        folder_id: ID of the folder to download from.
    """    
    drive_service = initialize_drive_service()
    parent_id = folder_id
    folder_name = find_file_name(parent_id)
    page_token = None
    files_exist = False
    company_list = []
    
    #Searches within folder for non-Google files with xls, xlsx, csv, or pdf extensions
    while True:
        response = drive_service.files().list(q='"%s" in parents and (not mimeType contains "google-apps")' % (parent_id), 
                                        spaces='drive', fields='nextPageToken, files(id, name)', pageToken=page_token).execute()
    
        for file in response.get('files', []):
            # Process change
            file_id = file.get('id')
            file_name = file.get('name').encode('utf-8')
            file_name_no_extension = (".").join(file.get('name').split(".")[:-1]).encode('utf-8')
            extension = file.get('name').split(".")[-1].encode('utf-8')
            
            if (extension == 'xls') or (extension == 'xlsx') or (extension == 'csv') or (extension == 'pdf'):
                files_exist = True
                logging.info("Found file: %s (ID: %s)" % (file_name, file_id))
                print ("Found file: %s (ID: %s)" % (file_name, file_id))
                count = 0
                temp = file_name_no_extension
                multiple = False
                #while True:
                    #if not os.path.isfile(temp + "." + extension):
                        #break
                    #else:
                        #count = count + 1
                        #if temp[len(temp) - 3] == '(' and temp[len(temp) - 1] == ')':
                            #temp = temp[:len(temp) - 3]                        
                        #temp = temp + "(" + str(count) + ")"
                        #multiple = True
                #if multiple:
                    #company_list.append(temp + "." + extension)
                #while count > 0:
                    #if count == 1:
                        #os.rename(file_name_no_extension + "." + extension, temp + "." + extension)
                        #rename_file(file_id, temp + "." + extension)
                    #else:
                        #temp = file_name_no_extension + "(" + str(count - 1) + ")"
                        #temp2 = file_name_no_extension + "(" + str(count) + ")"
                        #os.rename(temp + "." + extension, temp2 + "." + extension)
                        #temp_file_id = find_file_id(temp + "." + extension)
                        #rename_file(temp_file_id, temp2 + "." + extension)
                    #count = count - 1                    
                    
                while True:
                    if not os.path.isfile(temp + "." + extension):
                        break
                    count = count + 1
                    if temp[len(temp) - 3] == '(' and temp[len(temp) - 1] == ')':
                        temp = temp[:len(temp) - 3]                        
                    temp = temp + "(" + str(count) + ")"
                    multiple = True
                
                dl_file(file_id, temp + "." + extension)
                if multiple:
                    rename_file(file_id, temp + "." + extension)
                    company_list.append(temp + "." + extension)
                    count = count - 1
                    while count > 0:
                        temp = temp[:len(temp) - 3]
                        temp = temp + "(" + str(count) + ")"
                        file_id = find_file_id(temp + "." + extension)                     
                        rename_file(file_id, temp + "." + extension)
                        count = count - 1
                    file_id = find_file_id(temp[:len(temp) - 3] + "." + extension)
                    rename_file(file_id, temp[:len(temp) - 3] + "." + extension)
                        
                else:
                    company_list.append(file_name)
                    
        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break;
        
    items = response.get('files', [])
    if (not items or files_exist == False):
        logging.info("No files found to download from " + folder_name + " (ID: " + parent_id + ").")
        print("No files found to download from " + folder_name + " (ID: " + parent_id + ").")
    else:
        logging.info("Finished downloading files from " + folder_name + " (ID: " + parent_id + ").")
        print("Finished downloading files from " + folder_name + " (ID: " + parent_id + ").")
        #print("List of Files:")
        #for item in items:
            #print("{0} ({1})".format(item['name'], item['id']))

    return company_list

def get_filenames_in_folder(folder_id):
    """Gets a list of file names from a folder from Google Drive.
    
    Only records .xls, .xlsx, .csv, and .pdf files.
    
    Args:
        folder_id: ID of the folder to download from.
        
    Returns:
        file_list: list of file names.
    """    
    drive_service = initialize_drive_service()
    parent_id = folder_id
    folder_name = find_file_name(parent_id)
    page_token = None
    files_exist = False
    file_list = []
    
    #Searches within folder for non-Google files with xls, xlsx, csv, or pdf extensions
    while True:
        response = drive_service.files().list(q='"%s" in parents and (not mimeType contains "google-apps")' % (parent_id), 
                                        spaces='drive', fields='nextPageToken, files(id, name)', pageToken=page_token).execute()
    
        for file in response.get('files', []):
            # Process change
            file_id = file.get('id')
            file_name_full = file.get('name')
            file_name = (".").join(file.get('name').split(".")[:-1])
            extension = file.get('name').split(".")[-1]
            
            if (extension == 'xls') or (extension == 'xlsx') or (extension == 'csv') or (extension == 'pdf'):
                files_exist = True
                logging.info("Found file: %s (ID: %s)" % (file_name_full, file_id))
                print ("Found file: %s (ID: %s)" % (file_name_full, file_id))
                file_list.append(file_name_full.encode('utf-8'))
                    
        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break;
        
    items = response.get('files', [])
    if (not items or files_exist == False):
        logging.info("No files found in " + folder_name + " (" + parent_id + ").")
        print("No files found in " + folder_name + " (" + parent_id + ").")
    else:
        logging.info("Finished retrieving file names from " + folder_name + " (" + parent_id + ").")
        print("Finished retrieving file names from " + folder_name + " (" + parent_id + ").")
        #print(file_list)
            
    return file_list
   
def find_file_id(filename):
    """Gets file ID of a file in Google Drive using file name.
    
    Assumes there is only one file or folder with that name in the entire Drive.
    
    Args:
        filename: name of the file.

    Returns:
        file_id, file ID of the file.
    """    
    drive_service = initialize_drive_service()   
    file_id = None
    page_token = None
    
    #Search for file by name to retrieve ID
    while True:
        response = drive_service.files().list(q= 'name = "%s"' % filename, spaces='drive', 
                                        fields='nextPageToken, files(id, name)', 
                                        pageToken=page_token).execute()
        for file in response.get('files', []):
            # Process change
            file_id = file.get('id')
            page_token = response.get('nextPageToken', None)
        if page_token is None:
            break;    
    
    #If no matching files found
    if file_id == None:
        logging.warning("File \"%s\" not found." % filename)
        print("File \"%s\" not found." % filename)
    return file_id

def find_file_id_using_parent(filename, parent_id):
    """Gets file ID of a file in Google Drive using file name and parent ID.
    
    Assumes there is only one file with that name in the specified folder.
    
    Args:
        filename: name of the file.
        parent_id: ID of the parent.

    Returns:
        file_id, file ID of the file.
    """    
    drive_service = initialize_drive_service()   
    file_id = None
    page_token = None
    
    #Search for file by name to retrieve ID
    while True:
        response = drive_service.files().list(q= '"{0}" in parents and (name = "{1}")'.format(parent_id, filename), spaces='drive', 
                                        fields='nextPageToken, files(id, name)', 
                                        pageToken=page_token).execute()
        for file in response.get('files', []):
            # Process change
            file_id = file.get('id')
            page_token = response.get('nextPageToken', None)
        if page_token is None:
            break;    
    
    #If no matching files found
    if file_id == None:
        logging.warning("File \"%s\" not found." % filename)        
        print("File \"%s\" not found." % filename)
    return file_id

def find_file_name(file_id):
    """Gets name of a file in Google Drive using its ID.
    
    Args:
        file_id: ID of the file.

    Returns:
        file_name, name of the file.
    """    
    drive_service = initialize_drive_service()   
    file_name = None
    page_token = None
    
    #Search for file by name to retrieve ID
    while True:
        response = drive_service.files().list(spaces='drive', 
                                        fields='nextPageToken, files(id, name)', 
                                        pageToken=page_token).execute()
        for file in response.get('files', []):
            # Process change
            if (file.get('id') == file_id):
                file_name = file.get('name')
            page_token = response.get('nextPageToken', None)
        if page_token is None:
            break;    
    
    #If no matching files found
    if file_name == None:
        if file_id == None:
            logging.warning("No file ID given.")
            print("No file ID given.")
        else:
            logging.warning("File with ID \"" + file_id + "\" not found.")
            print(": File with ID \"" + file_id + "\" not found.")
        
    return file_name

def move_to_folder(file_id, folder_id):
    """Moves a file to a folder in Google Drive using file ID.
    
    Args:
        file_id: ID of the file to move.
        folder_id: ID of the destination folder.
    """    
    drive_service = initialize_drive_service()      
    filename = find_file_name(file_id)
    parent_id = folder_id
    folder_name = find_file_name(parent_id)
    
    try:
        # Retrieve the existing parents to remove
        file = drive_service.files().get(fileId=file_id,
                                     fields='parents').execute();
        previous_parents = ",".join(file.get('parents'))
        # Move the file to the new folder
        file = drive_service.files().update(fileId=file_id,
                                        addParents=parent_id,
                                        removeParents=previous_parents,
                                        fields='id, parents').execute()
        logging.info("Moved \"" + filename + "\" (ID: %s) to " % file_id + folder_name + " (ID: %s)" % parent_id)
        print("Moved \"" + filename + "\" (ID: %s) to " % file_id + folder_name + " (ID: %s)" % parent_id)
    except TypeError:
        logging.error("Could not find file to move.")
        print("Could not find file to move.")    
    except googleapiclient.errors.HttpError:
        logging.error("Invalid folder ID.")
        print("Invalid folder ID.")

def move_to_folder_using_name(filename, folder_id):
    """Moves a file to a folder in Google Drive using file name.
    
    Assumes there is only one file with that name in the entire Drive.
    
    Args:
        filename: name of the file to move.
        folder_id: ID of the destination folder.
    """    
    drive_service = initialize_drive_service()      
    file_id = find_file_id(filename)
    parent_id = folder_id
    folder_name = find_file_name(parent_id)
    
    try:
        # Retrieve the existing parents to remove
        file = drive_service.files().get(fileId=file_id,
                                     fields='parents').execute();
        previous_parents = ",".join(file.get('parents'))
        # Move the file to the new folder
        file = drive_service.files().update(fileId=file_id,
                                        addParents=parent_id,
                                        removeParents=previous_parents,
                                        fields='id, parents').execute()
        logging.info("Moved \"" + filename + "\" (ID: %s) to " % file_id + folder_name + " (ID: %s)" % parent_id)
        print("Moved \"" + filename + "\" (ID: %s) to " % file_id + folder_name + " (ID: %s)" % parent_id)
    except TypeError:
        logging.error("Could not find file to move.")
        print("Could not find file to move.")    
    except googleapiclient.errors.HttpError:
        logging.error("Invalid folder ID.")
        print("Invalid folder ID.")
   
def upload_as_gsheet(file_to_upload, filename):
    """Uploads a file as a Google Sheet to Google Drive.
    
    Default path is working directory. Takes Excel files (.xls, .xlsx) only.
    
    Args:
        file_to_upload: full path of the file to upload, including extension.
        filename: name of the file to be displayed on Google Drive.
    """    
    drive_service = initialize_drive_service()
    file_found = False
    
    try:
        extension = file_to_upload.split(".")[-1]
        
        file_metadata = {
            'name' : filename,
          'mimeType' : 'application/vnd.google-apps.spreadsheet'
        }
        
        if (extension == 'xls'):
            media = MediaFileUpload('%s' % (file_to_upload),
                                mimetype='application/vnd.ms-excel',
                                resumable=True)
            file_found = True
        elif (extension == 'xlsx'):
            media = MediaFileUpload('%s' % (file_to_upload),
                                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                resumable=True)
            file_found = True
        else:
            pass        
    except IndexError:
        pass
    
    if (file_found == True):
        file = drive_service.files().create(body=file_metadata,
                                        media_body=media,
                                        fields='id').execute()
        logging.info("File \"{0}\" uploaded as: {1} (ID: {2}).".format(file_to_upload, filename, file.get('id')))
        print("File \"{0}\" uploaded as: {1} (ID: {2}).".format(file_to_upload, filename, file.get('id')))    
    else:
        logging.error("Invalid file name or extension. Provide full file name with .xls or .xlsx extensions.")
        print("Invalid file name or extension. Provide full file name with .xls or .xlsx extensions.")
        
def get_email(ind):
    """Gets an email message resource from the Gmail inbox.
    It only looks at emails in the inbox with the label "New".
    
    Args:
        ind: index of the email.
        
    Returns:
        mail, an email resource.
    """
    gmail_service = initialize_gmail_service()
    label_ids = ["INBOX", "Label_2"]
    
    results = gmail_service.users().messages().list(userId='me',labelIds=label_ids).execute()
    #messages is the list of messages 
    messages = results['messages']
    mail = gmail_service.users().messages().get(userId='me', id=messages[ind]['id'], format='full').execute()
    return mail

def get_email_date(ind):
    """Gets the date of an email message from the Gmail inbox.
    It only looks at emails in the inbox with the label "New".
    
    Args:
        ind: index of the email.
        
    Returns:
        date, the date of the email.
    """
    gmail_service = initialize_gmail_service()
    label_ids = ["INBOX", "Label_2"]    
    
    results = gmail_service.users().messages().list(userId='me',labelIds=label_ids).execute()
    messages = results['messages']
    mail = gmail_service.users().messages().get(userId='me', id=messages[ind]['id'], format='metadata').execute()
    value = mail['payload']['headers']
    for dict_val in value:
        if dict_val['name'] == 'Date':
            date = dict_val['value']
            
    date = convert_date_email(date.encode('utf-8'))
    return date

def get_email_sender(ind):
    """Gets the sender of an email message from the Gmail inbox.
    It only looks at emails in the inbox with the label "New".
    
    Args:
        ind: index of the email.
        
    Returns:
        sender, the source of an email and the email address.
    """         
    gmail_service = initialize_gmail_service()
    label_ids = ["INBOX", "Label_2"]
    sender = []
    info = ""
    
    results = gmail_service.users().messages().list(userId='me',labelIds=label_ids).execute()
    messages = results['messages']
    mail = gmail_service.users().messages().get(userId='me', id=messages[ind]['id'], format='metadata').execute()
    value = mail['payload']['headers']
    for dict_val in value:
        if dict_val['name'] == 'From':
            info = dict_val['value']
    
    ind = info.rfind('<', 0)
    address = info[ind+1:-1].encode('utf-8')
    source = find_source_from_email(address)
    sender.append(source)
    sender.append(address)
    
    return sender

def get_email_attachment(ind):
    """Gets the name of an attachment of an email message from the Gmail inbox.
    It only looks at emails in the inbox with the label "New".
    
    Args:
        ind: index of the email.
        
    Returns:
        file, a list containing the file name of the attachment of an email.
    """
    gmail_service = initialize_gmail_service()
    label_ids = ["INBOX", "Label_2"]
    file = []
    
    results = gmail_service.users().messages().list(userId='me',labelIds=label_ids).execute()
    messages = results['messages']
    mail = gmail_service.users().messages().get(userId='me', id=messages[ind]['id']).execute()
    for part in mail['payload']['parts']:        
        filename = part_id(part)
        if filename != None:
            file.append(filename.encode('utf-8'))
                
        filename2 = part_find(part)
        if filename2 != None:
            if filename2 not in file:
                file.append(filename2.encode('utf-8'))

    return file

def part_id(part):
    '''Finds file name of an attachment using partId.
    
    Args:
        part: dictionary "part" of the email.
        
    Returns:
        filename, the name of the file.
    '''
    if 'partId' in part.keys() and part['partId'] > 0:
        if part['filename'] != "":
            filename = part['filename']
            return filename

def part_find(part):
    '''Recursive method that finds file name of an attachment using partId, looking through parts.
    
    Args:
        part: dictionary "part" of the email.
        
    Returns:
        filename, the name of the file.
    
    '''    
    filename = part_id(part)
    
    if 'parts' in part.keys():
        for part2 in part['parts']:
            filename = part_id(part2)
            if filename == None:
                part_find(part2)
                return filename
            else:
                return filename
    return filename

def get_email_attachment_list(dl_list):
    """Gets a list of attachments from Google Drive. 
    It only looks at emails in the inbox with the label "New".
    
    Args:
        dl_list: list of downloaded files.
        
    Returns:
        attach_list, a list of lists with each sublist containing the file name, source, email address, and date sent.
    """
    gmail_service = initialize_gmail_service()
    label_ids = ["INBOX", "Label_2"]
    results = gmail_service.users().messages().list(userId='me',labelIds=label_ids).execute() 
    if results['resultSizeEstimate']!=0:
        messages = results['messages']
    else:
        messages=[]
    
    ind = 0
    last_ind = len(messages)
    file_attach = []
    attach_list = []
    remove = []    
    loop_break = True
    is_thread = False
    
    while (ind < last_ind):
        #if ind != last_ind - 1:
            #if not is_thread and messages[ind]['threadId'] == messages[ind + 1]['threadId']:
                #dl_list.reverse()
                #is_thread = True
                #pass
            #if is_thread and messages[ind]['threadId'] != messages[ind + 1]['threadId']:
                #dl_list.reverse()
                #is_thread = False
                #pass
        file_attach = get_email_attachment(ind)
        sender = get_email_sender(ind)
        date = get_email_date(ind)
        for attachment in file_attach:
            sublist = []
            sublist.append(attachment)
            sublist.append(sender[0])
            sublist.append(sender[1])
            sublist.append(date)
            for i in range(len(dl_list)):
                temp = dl_list[i]
                index = temp.rfind(".")
                file_short = temp[:index]
                if file_short[len(file_short) - 3] == '(' and file_short[len(file_short) - 1] == ')':
                    sublist[0] = dl_list[i]
                    attach_list.append(sublist)
                    dl_list.pop(i)
                    break                    
                elif attachment == dl_list[i]:
                    attach_list.append(sublist)
                    dl_list.pop(i)
                    break
            if attachment == "":
                del sublist
            else:
                if ind == last_ind:
                    break
                #loop_break = False
                #break
        if ind == last_ind:
            break
        #if loop_break == False:
            #break
        remove.append(ind)
        ind = ind + 1
    for index in remove:
        remove_label(remove[0])

    return attach_list

def find_source_from_email(email_string):
    """Finds the source of an email by using the email address.
    
    Uses the following spreadsheet to determine the source:
    https://docs.google.com/spreadsheets/d/1rJlhCxJIy1DyYlzp8G9aVap505QBwxcTmiH9zleZzG4
    
    Args:
        email_string: string containing the sender name and email address in the following format:
        "Sender Name" <email@email.com>
        
    Returns:
        source, the name of the source.
    """    
    source_exists = False    

    if not os.path.isfile('Aggregator Source Sheet.xlsx'):
        export_sheet('1rJlhCxJIy1DyYlzp8G9aVap505QBwxcTmiH9zleZzG4')
    
    for i in range(1):  
        book = xlrd.open_workbook('Aggregator Source Sheet.xlsx')
        sheet = book.sheet_by_index(0)
        rownum = sheet.nrows
    
        for x in range(rownum):          
            if email_string.lower().encode("utf-8") == str(sheet.cell(x,0).value).encode("utf-8"):
                source = str(sheet.cell(x,1).value).encode("utf-8")
                source_exists = True
                return source

        if source_exists == False:
            export_sheet('1rJlhCxJIy1DyYlzp8G9aVap505QBwxcTmiH9zleZzG4')

def convert_date_email(date):
    '''Returns a date-time object using a string. Used for the email methods.
    
    Args:
        date: string of the date to be converted. Format example: 20 Jun 2017
        
    Returns:
        date_obj, the date-time object.    
    '''
    #Reg-ex for 4 digit numbers
    form = r"[0-9]{4}"
    find = re.findall(form, date)[0]
    start=0
    if "," in date:
        start = date.index(",")+1
    end = date.index(find)
    date = date[start:end+len(find)].strip()
    date_format = "%d %b %Y"
    date_obj = datetime.datetime.strptime(date, date_format)
    return date_obj.date()

def convert_date(date):
    '''Returns a date-time object using a string.
    
    Args:
        date: string of the date to be converted. Format example: 2017-06-20
        
    Returns:
        date_obj, the date-time object.    
    '''
    date_format = "%Y-%m-%d"
    date_obj = datetime.datetime.strptime(date, date_format)
    return date_obj.date()

def move_to_day_folder(file_id, datetime_obj, parent_id):
    """Organizes the files according to their respective "day" folders and 
    places the day folder into the parent folder. If a day folder does not 
    exist, it will be created.
    
    Only looks at files in the "Files" folder.

    Args:
        file_id: string of the ID of the file to move.
        datetime_obj: date of the email as date-time object.
        parent_id: string of the ID of the parent folder.

    """
    file_name = find_file_name(file_id)
    
    if parent_id == '0BzlU44AWMToxeFhld1pfNWxDTWs':
        folder_name = str(datetime_obj) + ' NR'
    elif parent_id == '0BzlU44AWMToxOGtyYWZzSVAyNkE':
        folder_name = str(datetime_obj) + ' NP'
    elif parent_id == '0BzlU44AWMToxVU8ySkNBQzJQeFE':
        folder_name = str(datetime_obj) + ' P'
    else:
        folder_name = str(datetime_obj) + ' N/A'
    
    folder_id = find_file_id_using_parent(folder_name, parent_id)
    if folder_id == None:
        create_folder(folder_name)
        folder_id = find_file_id(folder_name)
    move_to_folder(file_id, folder_id)
    move_to_folder(folder_id, parent_id)

def move_to_day_folder_using_names(filename, datetime_obj, parent_name):
    """Organizes the files according to their respective "day" folders and 
    places the day folder into the parent folder. If a day folder does not 
    exist, it will be created.
    
    Only looks at files in the "Files" folder.

    Args:
        filename: string of the name of the file to move.
        datetime_obj: date of the email as date-time object.
        parent_name: string of the name of the parent folder.

    """
    file_id = find_file_id_using_parent(filename, '0BzlU44AWMToxZnh5ekJaVUJUc2c')
    parent_id = find_file_id(parent_name)    
    
    folder_name = str(datetime_obj)
    folder_id = find_file_id_using_parent(folder_name, parent_id)
    if folder_id == None:
        create_folder(folder_name)
        folder_id = find_file_id_using_parent(folder_name, 'my-drive')
    move_to_folder(file_id, folder_id)
    move_to_folder_using_name(folder_name, parent_id)

def remove_label(ind):
    """Removes the "New" label and adds the "Processed" label.

    Args:
        ind: Index of the message from which the attachments have been pulled.
    """
    gmail_service = initialize_gmail_service()
    label_id = ["INBOX", "Label_2"]
    
    results = gmail_service.users().messages().list(userId='me', labelIds=label_id).execute()
    messages = results['messages']
    mail = gmail_service.users().messages().modify(userId='me', id=messages[ind]['id'],body={'removeLabelIds': ["Label_2"]}).execute()
    mail = gmail_service.users().messages().modify(userId='me', id=messages[ind]['id'],body={'addLabelIds': ["Label_3"]}).execute()    

def format_cell_alignment(sheet_id):
    """Formats the alignment of a spreadsheet cell.

    Args:
        sheet_id: ID of the sheet.
    """    
    sheets_service = initialize_sheets_service()
    
    spreadsheet_id = sheet_id
    
    col_list = [5, 7, 10]
    
    for col in col_list:
        if col == 10:
            alignment = 'CENTER'
        else:
            alignment = 'RIGHT'

        batch_update_spreadsheet_request_body = {
            'requests': [
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": 0,
                            "startRowIndex": 0,
                            "startColumnIndex": col,
                            "endColumnIndex": col + 1
                            },
                        "cell": {
                            "userEnteredFormat": {
                                "horizontalAlignment" : alignment
                            }
                        },
                    "fields": "userEnteredFormat(horizontalAlignment)"
                    }
                },
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": 0,
                            "startRowIndex": 0,
                            "endRowIndex": 1,
                            "startColumnIndex": 0
                            },
                        "cell": {
                            "userEnteredFormat": {
                                "horizontalAlignment" : 'CENTER'
                            }
                        },
                    "fields": "userEnteredFormat(horizontalAlignment)"
                    }
                },                    
            ],
        }
    
        request = sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=batch_update_spreadsheet_request_body)   
        response = request.execute()

def conditional_format(spreadsheet_id):
    """Conditionally formats the colors of spreadsheet rows based on a value in a cell.

    Args:
        spreadsheet_id: ID of the sheet.
    """       
    sheets_service = initialize_sheets_service()
        
    myRange = {
        "sheetId": 0,
      "startRowIndex": 0,
      "startColumnIndex": 0,
    }
    
    requests = []
    
    requests.append({
        "addConditionalFormatRule": {
          "rule": {
              "ranges": [ myRange ],
            "booleanRule": {
              "condition": {
                "type": "CUSTOM_FORMULA",
              "values": [ { "userEnteredValue": "=EXACT(\"Increase\", $K1)" } ]
              },
            "format": {
                "backgroundColor": { "red": 1.0, "green": 0.3, "blue": 0.3 }
            }
          }
          },
        "index": 0
      }
    })
    
    requests.append({
        "addConditionalFormatRule": {
          "rule": {
              "ranges": [ myRange ],
            "booleanRule": {
              "condition": {
                "type": "CUSTOM_FORMULA",
              "values": [ { "userEnteredValue": "=EXACT(\"Decrease\", $K1)" } ]
              },
            "format": {
                "backgroundColor": { "red": 0.4, "green": 0.65, "blue": 0.3 }
            }
          }
          },
        "index": 0
      }
    })

    requests.append({
        "autoResizeDimensions": {
          "dimensions": {
            "sheetId": 0,
            "dimension": "COLUMNS",
            "startIndex": 0,
            "endIndex": 2
          }
        }
    })    
    
    body = {
        'requests': requests
    }
    result = sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id,
                                                body=body).execute()    

def freeze_first_row(spreadsheet_id, rowCount):
    """Freezes the first row of the spreadsheet.

    Args:
        spreadsheet_id: ID of the sheet.
        rowCount: number of rows.
    """     
    sheets_service = initialize_sheets_service()
    
    requests = []
    
    requests.append({
        "updateSheetProperties": {
              "fields": "*",
              "properties": {
                  "title": 'Sheet1',
                  "gridProperties": {
                      "frozenRowCount": 1,
                      "rowCount": rowCount,
                      "columnCount": 11,
                  }
              }
        }
    })
    
    body = {
        'requests': requests
    }
    
    result = sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id,
                                                       body=body).execute()  

def unfreeze_first_row(spreadsheet_id, rowCount):
    """Unreezes the first row of the spreadsheet.

    Args:
        spreadsheet_id: ID of the sheet.
        rowCount: number of rows.
    """     
    sheets_service = initialize_sheets_service()
    
    requests = []
    
    requests.append({
        "updateSheetProperties": {
              "fields": "*",
              "properties": {
                  "title": 'Sheet1',
                  "gridProperties": {
                      "frozenRowCount": 0,
                      "rowCount": rowCount,
                      "columnCount": 11,
                  }
              }
        }
    })
    
    body = {
        'requests': requests
    }
    
    result = sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id,
                                                       body=body).execute()
    
def upload_excel(file_name):
    """Uploads an Excel (.xlsx) file to Drive.
    
    Default path is working directory. Takes Excel files (.xlsx) only.

    Args:
        file_name: full name of the file including extension.
    """    
    drive_service = initialize_drive_service()
    file_metadata = { 'name' : file_name }
    try:
        media = MediaFileUpload(file_name,
                                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        file = drive_service.files().create(body=file_metadata,
                                            media_body=media,
                                            fields='id').execute()
        logging.info("Uploaded " + file_name + "(ID: %s) to Drive" % file.get('id'))
        print("Uploaded " + file_name + "(ID: %s) to Drive" % file.get('id'))
    except IOError:
        logging.error("File not found.")
        print("File not found.")
        
def upload_log(file_name):
    """Uploads a log (.log) file to Drive.
    
    Default path is working directory. Takes log files (.log) only.

    Args:
        file_name: full name of the file including extension.
    """    
    drive_service = initialize_drive_service()
    file_metadata = { 'name' : file_name }
    try:
        media = MediaFileUpload(file_name,
                                mimetype='text/plain')
        file = drive_service.files().create(body=file_metadata,
                                            media_body=media,
                                            fields='id').execute()
        
        print("Uploaded " + file_name + "(ID: %s) to Drive" % file.get('id'))
    except IOError:
        print("File not found.")
    
def main():
    drive_service = initialize_drive_service()
    gmail_service = initialize_gmail_service()
    sheets_service = initialize_sheets_service()