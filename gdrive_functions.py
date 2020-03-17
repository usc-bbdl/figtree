from upload_to_gdrive import tempDict


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


def agenda_and_figures_are_empty():
    return tempDict['agendaItemsCount'] == 0 and tempDict['figureCount'] == 0