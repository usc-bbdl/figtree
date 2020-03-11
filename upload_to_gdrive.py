import os.path
from os import listdir
from datetime import date

from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

# from slack_functions import *

### Import relevant values from temp_output.txt
###############################################################################

assert os.path.exists('temp_output.txt'), "'temp_output.txt' does not exist. Need to run 'build_weekly_ppt.py' before " \
                                          "'upload_to_gdrive.py'. "

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


tempDict = etl_build_details()


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


if tempDict['agendaItemsCount'] == 0 and tempDict['figureCount'] == 0:
    print("No Agenda or Figures this week... :(")
    # TODO: still send a slack message to Brian
else:
    ### Connect to Google Drive API
    ###########################################################################

    gauth = GoogleAuth()

    # Try to load saved client credentials
    if os.path.exists("mycreds.txt"):
        gauth.LoadCredentialsFile("mycreds.txt")

    gauth = authenticate(gauth)

    # Save the current credentials to a file
    gauth.SaveCredentialsFile("mycreds.txt")
    drive = GoogleDrive(gauth)

    ### Find 'ValeroLabMeetings' folder id
    ###########################################################################
    assert len(
        drive.ListFile({'q': "title='ValeroLabMeetings' and 'root' in parents and trashed=false"}).GetList()
    ) == 1, \
        "Error. 'ValeroLabMeetings/' not found in Google Drive."

    ValeroLabMeetings_folder_id = \
    drive.ListFile({'q': "title='ValeroLabMeetings' and 'root' in parents and trashed=false"}).GetList()[0]['id']

    ### Create Folder for Latest Lab Meeting
    ###########################################################################

    labMeetingFolder_metadata = {
        'title': tempDict['labMeetingFolderName'][:-1],
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

    ### Upload powerpoint presenatation
    ###########################################################################

    labMeetingPresentation_title = (
            tempDict['labMeetingFolderName'][:-1].replace(" ", "_")
            + ".pptx"
    )
    labMeetingPresentation = drive.CreateFile(
        {
            "parents": [{
                "kind": "drive#fileLink",
                "id": labMeetingFolder_id
            }]
        }
    )
    labMeetingPresentation.SetContentFile(
        labMeetingPresentation_title
    )
    labMeetingPresentation.Upload()

    ### Distribute link to lab
    #######################################################################

    download_link = labMeetingPresentation.metadata['webContentLink']
    # today = date.today()
    # if today.weekday()==0: # Monday
    #     distribute_link_to_lab(download_link)

    ### Upload figures
    ###########################################################################

    FigureQueue_folder_id = drive.ListFile(
        {'q': "title='Figure Queue' and '{}' in parents and trashed=false".format(ValeroLabMeetings_folder_id)}
    ).GetList()[0]['id']

    FigureQueue_item_list = drive.ListFile(
        {'q': "title!='README.md' and '{}' in parents and trashed=false".format(FigureQueue_folder_id)}).GetList()

    if len(FigureQueue_item_list) != 0:
        ### Create new figures subfolder
        #######################################################################
        figuresFolder_metadata = {
            'title': 'Figures',
            # Define the file type as folder
            'mimeType': 'application/vnd.google-apps.folder',
            # ID of the parent folder
            'parents': [{
                "kind": "drive#fileLink",
                "id": labMeetingFolder_id
            }]
        }

        figuresFolder = drive.CreateFile(figuresFolder_metadata)
        figuresFolder.Upload()
        figuresFolder_id = drive.ListFile(
            {'q': "title='Figures' and '{}' in parents and trashed=false".format(labMeetingFolder_id)}
        ).GetList()[0]['id']

        ### Move files from Figure Queue to new subfolder
        #######################################################################
        for item in sorted(FigureQueue_item_list, key=lambda x: x['title']):
            drive.auth.service.files().update(
                fileId=item['id'],
                addParents=figuresFolder_id,
                removeParents=FigureQueue_folder_id,
                fields='id, parents'
            ).execute()
