# import requests
import gdown
def return_google_files():
    url = 'https://drive.google.com/drive/u/0/folders/0AMJzoQplWDs9Uk9PVA'
    agenda_docx = 'TEMP_INPUT_FOLDER/Meeting Agenda.docx'
    figure_queue = 'TEMP_INPUT_FOLDER/Figure Queue/'
    gdown.download(url, agenda_docx, quiet=False)
    gdown.download(url, figure_queue, quiet=False)
# def return_google_files():
#     headers = {
#         'Authorization': 'Bearer [YOUR_ACCESS_TOKEN]',
#         'Accept': 'application/json',
#     }
#
#     params = (
#         ('key', '[YOUR_API_KEY]'),
#     )
#
#     response = requests.get('https://www.googleapis.com/drive/v3/files', headers=headers, params=params)
#
#     #NB. Original query string below. It seems impossible to parse and
#     #reproduce query strings 100% accurately so the one below is given
#     #in case the reproduced version is not "correct".
#     # response = requests.get('https://www.googleapis.com/drive/v3/files?key=[YOUR_API_KEY]', headers=headers)
