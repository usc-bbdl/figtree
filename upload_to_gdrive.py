from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from os import remove, mkdir, listdir
import os.path
import sys

### Import relevant values from temp_output.txt
###############################################################################

assert os.path.exists('temp_output.txt'), "'temp_output.txt' does not exist. Need to run 'build_weekly_ppt.py' before 'upload_to_gdrive.py'."

tempDict = {}
with open('temp_output.txt', 'r') as f:
    for line in f:
        splitLine = line.split()
        if len(splitLine)==2:
            tempDict[splitLine[0]] = int(splitLine[1])
        else:
            tempDict[splitLine[0]] = " ".join(splitLine[1:])

if tempDict['agendaItemsCount']==0 and tempDict['figureCount']==0:
    print("No Agenda or Figures this week... :(")
else:
    ### Connect to Google Drive API
    ###########################################################################

    gauth = GoogleAuth()

    # Try to load saved client credentials
    if os.path.exists("mycreds.txt"):gauth.LoadCredentialsFile("mycreds.txt")

    if gauth.credentials is None:
        # Authenticate if they're not there
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        # Refresh them if expired
        gauth.Refresh()
    else:
        # Initialize the saved creds
        gauth.Authorize()

    # Save the current credentials to a file
    gauth.SaveCredentialsFile("mycreds.txt")
    drive = GoogleDrive(gauth)

    ### Find 'ValeroLabMeetings' folder id
    ###########################################################################
    assert len(
                drive.ListFile({'q': "title='ValeroLabMeetings' and 'root' in parents and trashed=false"}).GetList()
            )==1,\
        "Error. 'ValeroLabMeetings/' not found in Google Drive."

    ValeroLabMeetings_folder_id = drive.ListFile({'q': "title='ValeroLabMeetings' and 'root' in parents and trashed=false"}).GetList()[0]['id']

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
        tempDict['labMeetingFolderName'][:-1].replace(" ","_")
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
        tempDict['labMeetingFolderName']
        + labMeetingPresentation_title
    )
    labMeetingPresentation.Upload()

    ### Upload figures
    ###########################################################################

    if tempDict['figureCount']!=0:
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

        for item in listdir(tempDict['labMeetingFolderName']+'Figures/'):
            figure = drive.CreateFile(
                {
                    "parents": [{
                        "kind": "drive#fileLink",
                        "id": figuresFolder_id
                    }]
                }
            )
            figure.SetContentFile(
                tempDict['labMeetingFolderName']
                + "Figures/"
                + item)
            figure.Upload()