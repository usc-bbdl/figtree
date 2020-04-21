import os

from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

from upload_to_gdrive import ValeroLabMeetings_folder_id, drive, tempDict


def etl_build_details():
    tempDict = {}
    with open('temp_output.txt', 'r') as f:
        for line in f:
            split_line = line.split()
            if len(split_line) == 2:
                tempDict[split_line[0]] = int(split_line[1])
            else:
                tempDict[split_line[0]] = " ".join(split_line[1:])
    return tempDict


def authenticate(gauth):
    if gauth.credentials is None:
        # Authenticate if they're not there
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        # Refresh them if expired
        gauth.Refresh()
    else:
        # Initialize the saved creds
        gauth.Authorize()
    return gauth


def agenda_and_figures_are_empty(tempDict):
    return tempDict['agendaItemsCount'] == 0 and tempDict['figureCount'] == 0

## @param folder_name name for the folder without slashes
## @return labMeetingFolder_id the google drive ID for the new folder
def mkdir_for_new_lab_meeting(folder_name):
    labMeetingFolder_metadata = {
        'title': folder_name,
        # Define the file type as folder
        'mimeType': 'application/vnd.google-apps.folder',
        # ID of the parent folder
        'parents': [{
            "kind": "drive#fileLink",
            "id": ValeroLabMeetings_folder_id
        }]
    }
    labMeetingFolder = drive.CreateFile(labMeetingFolder_metadata)
    labMeetingFolder.Upload()
    labMeetingFolder_id = drive.ListFile(
        {'q': "title='{}' and '{}' in parents and trashed=false".format(
            tempDict['labMeetingFolderName'][:-1],
            ValeroLabMeetings_folder_id
        )}
    ).GetList()[0]['id']
    return labMeetingFolder_id

## @param folder_name string target folder name without slashes.
## @param parent_name string the parent directory. If it's in the root directory, use 'root'.
## @note have not tested spaces and unescaped inputs
def compose_folder_find_query(folder_name, parent_name):
    query_text = "title='%s' and '%s' in parents and trashed=false" % (folder_name,parent_name)
    return query_text


def find_and_id_drive_folder(folder_name,parent_name):
    query_text = compose_folder_find_query(folder_name, parent_name)
    results = drive.ListFile({'q': query_text}).GetList()
    if len(results) == 1:
        return results[0]['id']
    elif len(results) > 1:
        raise NameError('Multiple folders named %s/ found; unclear which one should be chosen' % folder_name)
    else:
        raise NameError('Folder %s/ not found' % folder_name)


def compose_mkdir_query():
    return {
        'title': foldername,
        # Define the file type as folder
        'mimeType': 'application/vnd.google-apps.folder',
        # ID of the parent folder
        'parents': [{
            "kind": "drive#fileLink",
            "id": parent_id
        }]
    }

## @param dirname name of directory. try to avoid spaces and weird symbols
## @param parent_id drive-id that will house the new folder.
def drive_mkdir(dirname, parentid):
    newFolder = drive.CreateFile(compose_mkdir_query(dirname, parentid))
    newFolder.Upload() #sync new change to drive
    return newFolder.get('id')


def authenticate_drive():
    gauth = GoogleAuth()
    # Try to load saved client credentials
    if os.path.exists("mycreds.txt"):
        gauth.LoadCredentialsFile("mycreds.txt")
    gauth = authenticate(gauth)
    # Save the current credentials to a file
    gauth.SaveCredentialsFile("mycreds.txt")
    return GoogleDrive(gauth)

## @param filepath string, a single file to upload.
## @param folder_id string folder target drive id
## @return drive_file a drive-file with an ID and weblink
def drive_newfile(filepath, folder_id):
    file = drive.CreateFile(
        {
            "parents": [{
                "kind": "drive#fileLink",
                "id": folder_id
            }]
        }
    )
    file.SetContentFile(
        filepath
    )
    file.Upload()
    return file