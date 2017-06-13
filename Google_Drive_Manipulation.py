from apiclient import discovery
from apiclient.http import MediaFileUpload
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import googleapiclient
import httplib2
import os
import io

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at sheets.googleapis.com-python.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Rates'

def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    dir_path = os.path.dirname(os.path.realpath(__file__))
    credential_path = os.path.join(dir_path,
                                   'sheets.googleapis.com-python.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print("Storing credentials to " + credential_path)
    return credentials

def initialize_service():
    """Initializes service.
    
    Returns:
        service, a service object.
    """    
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())    
    service = discovery.build('drive', 'v3', http=http)    
    
    return service

def delete_file(file_id):
    """Permanently delete a file from Google Drive, skipping the trash.
    
      Args:
        file_id: ID of the file to delete.
      """       
    service = initialize_service()
    
    try:
        service.files().delete(fileId=file_id).execute()
    except errors.HttpError, error:
        print("An error occurred: %s" % error)    

def clean_folder(folder_id):
    """Permanently deletes files from a Google Drive folder, skipping the trash.
    
    Only leaves .xls, .xlsx, .csv, and .pdf files, deleting all other file types, including Google app files (sheets, docs, slides, etc.).

      Args:
        folder_id: ID of the folder to delete from.
      """      
    service = initialize_service()  
    parent_id = folder_id
    folder_name = find_file_name(parent_id)
    page_token = None
    delete = False
    
    #Delete Google files
    while True:
        response = service.files().list(q='"%s" in parents and (mimeType contains "google-apps")' % (parent_id), 
                                        spaces='drive', fields='nextPageToken, files(id, name)', pageToken=page_token).execute()
    
        for file in response.get('files', []):
            # Process change
            file_id = file.get('id')
            file_name_full = file.get('name')
            
            print("Deleted Google file: %s (%s)" % (file_name_full, file_id))
            delete_file(file_id)
            delete = True
                
        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break;    
    
    #Delete files that are not xls, xlsx, csv, or pdf format
    while True:
        response = service.files().list(q='"%s" in parents' % (parent_id), spaces='drive', 
                                        fields='nextPageToken, files(id, name)', pageToken=page_token).execute()
    
        for file in response.get('files', []):
            # Process change
            file_id = file.get('id')
            file_name_full = file.get('name')
            file_name = (".").join(file.get('name').split(".")[:-1])
            extension = file.get('name').split(".")[-1]
            
            if (extension != 'xls') and (extension != 'xlsx') and (extension != 'csv') and (extension != 'pdf'):
                print("Deleted file: %s (%s)" % (file_name_full, file_id))
                delete_file(file_id)
                delete = True
                
        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break;
    
    if (delete == False):
        print("No files found to remove from " + folder_name + " (" + parent_id + ").")
    else:
        print("Finished cleaning folder from " + folder_name + "(" + parent_id + ").")     
   
def rename_file(filename, newname):
    """Renames a file in Google Drive.
    
    Args:
        filename: name of the file to rename.
        newname: new name of the file.
    """
    service = initialize_service()
    
    file_id = find_file_id(filename)

    file_metadata = {
        'name' : newname
        }
    
    try:
        file = service.files().update(fileId=file_id, body=file_metadata, fields='id').execute()
        print("File \"{0}\" renamed as: {1} ({2}).".format(filename, newname, file.get('id')))
    except TypeError:
        pass     
        
def dl_file(file_id, filename, extension):
    """Downloads a file from Google Drive.
    
    Args:
        file_id: ID of the file to download.
        filename: name of the file to download without extension.
        extension: extension of the file.
    """    
    service = initialize_service()
    request = service.files().get_media(fileId=file_id)
    #fh = io.BytesIO()
    fh = io.FileIO("{0}.{1}".format(filename, extension), 'wb')
    downloader = googleapiclient.http.MediaIoBaseDownload(fh, request)
    done = False
    
    while done is False:
        status, done = downloader.next_chunk()
        print("Download %d%%." % int(status.progress() * 100))
        
def dl_folder(folder_id):
    """Downloads files from a folder from Google Drive.
    
    Only downloads .xls, .xlsx, .csv, and .pdf files.
    
    Args:
        folder_id: ID of the folder to download from.
    """    
    service = initialize_service()
    parent_id = folder_id
    folder_name = find_file_name(parent_id)
    page_token = None
    
    #Searches within folder for non-Google files with xls, xlsx, csv, or pdf extensions
    while True:
        response = service.files().list(q='"%s" in parents and (not mimeType contains "google-apps")' % (parent_id), 
                                        spaces='drive', fields='nextPageToken, files(id, name)', pageToken=page_token).execute()
    
        for file in response.get('files', []):
            # Process change
            file_id = file.get('id')
            file_name_full = file.get('name')
            file_name = (".").join(file.get('name').split(".")[:-1])
            extension = file.get('name').split(".")[-1]
            
            if (extension == 'xls') or (extension == 'xlsx') or (extension == 'csv') or (extension == 'pdf'):
                print ("Found file: %s (%s)" % (file_name_full, file_id))
                dl_file(file_id, file_name, extension)
                    
        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break;
        
    items = response.get('files', [])
    if not items:
        print("No files found to download from " + folder_name + " (" + parent_id + ").")
    else:
        print("Finished downloading files from " + folder_name + " (" + parent_id + ").")
        #print("List of Files:")
        #for item in items:
            #print("{0} ({1})".format(item['name'], item['id']))
   
def find_file_id(filename):
    """Gets file ID of a file in Google Drive using file name.
    
    Assumes there is only one file or folder with that name in the entire Drive.
    
    Args:
        filename: name of the file.

    Returns:
        file_id, file ID of the file.
    """    
    service = initialize_service()   
    file_id = None
    page_token = None
    
    #Search for file by name to retrieve ID
    while True:
        response = service.files().list(q= 'name = "%s"' % filename, spaces='drive', 
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
        print("File not found.")
        
    return file_id

def find_file_name(file_id):
    """Gets name of a file in Google Drive using its ID.
    
    Args:
        file_id: ID of the file.

    Returns:
        file_name, name of the file.
    """    
    service = initialize_service()   
    file_name = None
    page_token = None
    
    #Search for file by name to retrieve ID
    while True:
        response = service.files().list(spaces='drive', 
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
        print("File not found.")
        
    return file_name

def move_to_folder(filename, folder_id):
    """Moves a file to a folder in Google Drive.
    
    Assumes there is only one file with that name in the entire Drive.
    
    Args:
        filename: name of the file to move.
        folder_id: ID of the destination folder.
    """    
    service = initialize_service()      
    file_id = find_file_id(filename)
    parent_id = folder_id
    folder_name = find_file_name(parent_id)
    
    try:
        # Retrieve the existing parents to remove
        file = service.files().get(fileId=file_id,
                                     fields='parents').execute();
        previous_parents = ",".join(file.get('parents'))
        # Move the file to the new folder
        file = service.files().update(fileId=file_id,
                                        addParents=parent_id,
                                        removeParents=previous_parents,
                                        fields='id, parents').execute()
        print("Moved \"" + filename + "\" to " + folder_name + " (%s)" % parent_id)
    except TypeError:
        print("Could not find file to move.")    
    except googleapiclient.errors.HttpError:
        print("Invalid folder ID.")

def move_to_noRates(filename):
    """Moves a file to the noRates folder in Google Drive.
    
    Args:
        filename: name of the file to move.
    """    
    move_to_folder(filename, '0BzlU44AWMToxeFhld1pfNWxDTWs')
                 
def move_to_processed(filename):
    """Moves a file to the Processed folder in Google Drive.
    
    Args:
        filename: name of the file to move.
    """        
    move_to_folder(filename, '0BzlU44AWMToxVU8ySkNBQzJQeFE')
    
def move_to_notProcessed(filename):
    """Moves a file to the NotProcessed folder in Google Drive.
    
    Args:
        filename: name of the file to move.
    """        
    move_to_folder(filename, '0BzlU44AWMToxOGtyYWZzSVAyNkE')    
    
def upload_as_gsheet(file_to_upload, filename):
    """Uploads a file as a Google Sheet to Google Drive.
    
    Default path is working directory. Takes Excel files (.xls, .xlsx) only.
    
    Args:
        file_to_upload: full path of the file to upload, including extension.
        filename: name of the file to be displayed on Google Drive.
    """    
    service = initialize_service()
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
        file = service.files().create(body=file_metadata,
                                        media_body=media,
                                        fields='id').execute()
        print("File \"{0}\" uploaded as: {1} ({2}).".format(file_to_upload, filename, file.get('id')))    
    else:
        print("Invalid file name or extension. Provide full file name with .xls or .xlsx extensions.")

def main():
    service = initialize_service()