from datetime import date

from gdrive_functions import etl_build_details, agenda_and_figures_are_empty, \
    gen_folder_for_new_lab_meeting, find_and_id_drive_folder, drive_mkdir, authenticate_drive, drive_newfile
from slack_functions import *

### Import relevant values from temp_output.txt
###############################################################################

assert os.path.exists('temp_output.txt'), "'temp_output.txt' does not exist. Need to run 'build_weekly_ppt.py' before " \
                                          "'upload_to_gdrive.py'. "

tempDict = etl_build_details()

## @param folder_name name for the folder without slashes
## @return labMeetingFolder_id the google drive ID for the new folder

## @param folder_name string target folder name without slashes.
## @param parent_name string the parent directory. If it's in the root directory, use 'root'.
## @note have not tested spaces and unescaped inputs


## @param dirname name of directory. try to avoid spaces and weird symbols
## @param parent_id drive-id that will house the new folder.


## @param filepath string, a single file to upload.
## @param folder_id string folder target drive id
## @return drive_file a drive-file with an ID and weblink


if agenda_and_figures_are_empty(tempDict):
    print("No Agenda or Figures this week... :(")
    # TODO: still send a slack message to Brian

else:
    drive = authenticate_drive()

    meetings_foldername = "ValeroLabMeetings"
    ValeroLabMeetings_folder_id = find_and_id_drive_folder(meetings_foldername)
    labMeetingFolder_id = gen_folder_for_new_lab_meeting(tempDict['labMeetingFolderName'][:-1])

    ### Upload powerpoint presenatation
    ###########################################################################
    ppt_filepath = (
            tempDict['labMeetingFolderName'][:-1].replace(" ", "_")
            + ".pptx"
    )

    download_line = id_to_weblink(id=drive_newfile(filepath=ppt_filepath, folder_id=labMeetingFolder_id))


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




        figuresFolder_id = drive_mkdir('Figures', labMeetingFolder_id)


        ### Move files from Figure Queue to new subfolder
        #######################################################################
        for item in sorted(FigureQueue_item_list, key=lambda x: x['title']):
            drive.auth.service.files().update(
                fileId=item['id'],
                addParents=figuresFolder_id,
                removeParents=FigureQueue_folder_id,
                fields='id, parents'
            ).execute()
