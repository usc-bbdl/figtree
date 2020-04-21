from datetime import date

from gdrive_functions import etl_build_details, agenda_and_figures_are_empty, \
    mkdir_for_new_lab_meeting, find_and_id_drive_folder, drive_mkdir, authenticate_drive, drive_newfile
from slack_functions import *

### Import relevant values from temp_output.txt
###############################################################################

assert os.path.exists('temp_output.txt'), "'temp_output.txt' does not exist. Need to run 'build_weekly_ppt.py' before " \
                                          "'upload_to_gdrive.py'. "

tempDict = etl_build_details()

## @param item_list a GoogleDriveFileList (acquired via drive.ListFile)
def drive_mv_bulk(item_list, target_parent_id, prior_parent_id_to_remove):
    for item in sorted(item_list, key=lambda x: x['title']):
        drive.auth.service.files().update(
            fileId=item['id'],
            addParents=target_parent_id,
            removeParents=prior_parent_id_to_remove,
            fields='id, parents'
        ).execute()


if agenda_and_figures_are_empty(tempDict):
    print("No Agenda or Figures this week... :(")
    # TODO: still send a slack message to Brian

else:
    drive = authenticate_drive()

    meetings_foldername = "ValeroLabMeetings"
    ValeroLabMeetings_folder_id = find_and_id_drive_folder(meetings_foldername)
    labMeetingFolder_id = mkdir_for_new_lab_meeting(tempDict['labMeetingFolderName'][:-1])

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

    figure_files_in_the_figure_queue = drive.ListFile(
        {'q': "title!='README.md' and '{}' in parents and trashed=false".format(FigureQueue_folder_id)}).GetList()

    # Clear figure Queue by filling a Figures folder in the meeting folder.
    if len(figure_files_in_the_figure_queue) != 0:
        figuresFolder_id = drive_mkdir('Figures', labMeetingFolder_id)
        drive_mv_bulk(figure_files_in_the_figure_queue, figuresFolder_id, FigureQueue_folder_id)
