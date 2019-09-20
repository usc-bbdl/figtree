import os.path
from os import mkdir

from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

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

### Connect to Google Drive API
###############################################################################

gauth = GoogleAuth()

# Try to load saved client credentials
if os.path.exists("mycreds.txt"):
    gauth.LoadCredentialsFile("mycreds.txt")

gauth = authenticate(gauth)

# Save the current credentials to a file
gauth.SaveCredentialsFile("mycreds.txt")
drive = GoogleDrive(gauth)

### Make TEMP_INPUT_FOLDER with 'Figure Queue' subfolder
###############################################################################

TEMP_INPUT_FOLDER = 'TEMP_INPUT_FOLDER/'
mkdir(TEMP_INPUT_FOLDER)
mkdir(TEMP_INPUT_FOLDER + 'Figure Queue/')

### Find 'ValeroLabMeetings' folder id
###############################################################################
assert len(
    drive.ListFile({'q': "title='ValeroLabMeetings' and 'root' in parents and trashed=false"}).GetList()
) == 1, \
    "Error. 'ValeroLabMeetings/' not found in Google Drive."

ValeroLabMeetings_folder_id = \
drive.ListFile({'q': "title='ValeroLabMeetings' and 'root' in parents and trashed=false"}).GetList()[0]['id']

### Find 'ValeroLabMeetings/Meeting Agenda.docx' folder id and save it locally.
###############################################################################
assert len(
    drive.ListFile(
        {'q': "title='Meeting Agenda.docx' and '{}' in parents and trashed=false".format(ValeroLabMeetings_folder_id)}
    ).GetList()
) == 1, \
    "Error. 'ValeroLabMeetings/Meeting Agenda.docx' not found in Google Drive."

MeetingAgenda_file = drive.ListFile(
    {'q': "title='Meeting Agenda.docx' and '{}' in parents and trashed=false".format(ValeroLabMeetings_folder_id)}
).GetList()[0]

MeetingAgenda_file.GetContentFile(
    TEMP_INPUT_FOLDER
    + MeetingAgenda_file['title']
)

### Find 'ValeroLabMeetings/Figure Queue' folder id and save its contents
### locally.
###############################################################################
assert len(
    drive.ListFile(
        {'q': "title='Figure Queue' and '{}' in parents and trashed=false".format(ValeroLabMeetings_folder_id)}
    ).GetList()
) == 1, \
    "Error. 'ValeroLabMeetings/Figure Queue' not found in Google Drive."

FigureQueue_folder_id = drive.ListFile(
    {'q': "title='Figure Queue' and '{}' in parents and trashed=false".format(ValeroLabMeetings_folder_id)}
).GetList()[0]['id']

FigureQueue_item_list = drive.ListFile(
    {'q': "title!='README.md' and '{}' in parents and trashed=false".format(FigureQueue_folder_id)}).GetList()

i = 1
if len(FigureQueue_item_list) != 0:
    for item in sorted(FigureQueue_item_list, key=lambda x: x['title']):
        print('Downloading {} from GDrive ({}/{})'.format(item['title'], i, len(FigureQueue_item_list)))
        item.GetContentFile(
            TEMP_INPUT_FOLDER
            + 'Figure Queue/'
            + item['title']
        )
        i += 1
