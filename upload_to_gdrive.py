from datetime import date

from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

from gdrive_functions import etl_build_details, authenticate, agenda_and_figures_are_empty
from slack_functions import *

### Import relevant values from temp_output.txt
###############################################################################

assert os.path.exists('temp_output.txt'), "'temp_output.txt' does not exist. Need to run 'build_weekly_ppt.py' before " \
                                          "'upload_to_gdrive.py'. "

tempDict = etl_build_details()

## @param folder_name name for the folder without slashes
## @return labMeetingFolder_id the google drive ID for the new folder
def gen_folder_for_new_lab_meeting(folder_name):
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

if agenda_and_figures_are_empty():
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
    meetings_foldername = "ValeroLabMeetings"
    assert foldername_exists_in_drive(meetings_foldername), \
        "Error. '%s/' not found in Google Drive."%meetings_foldername
    ValeroLabMeetings_folder_id = find_and_id_drive_folder(meetings_foldername)

    ### Create Folder for Latest Lab Meeting
    ###########################################################################

    gen_folder_for_new_lab_meeting(tempDict['labMeetingFolderName'][:-1])

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
    today = date.today()
    if today.weekday() == 0:  # Monday
        distribute_link_to_lab(download_link)

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
        figuresFolder_id = figuresFolder.get('id')

        ### Move files from Figure Queue to new subfolder
        #######################################################################
        for item in sorted(FigureQueue_item_list, key=lambda x: x['title']):
            drive.auth.service.files().update(
                fileId=item['id'],
                addParents=figuresFolder_id,
                removeParents=FigureQueue_folder_id,
                fields='id, parents'
            ).execute()
