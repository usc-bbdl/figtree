import requests
import datetime
import os
DRIVE_FOLDER=os.environ['DRIVE_FOLDER']
DRIVE_QUEUE=os.environ['DRIVE_QUEUE']
SLACK_API=os.environ['SLACK_API']
headers = {
    'Content-type': 'application/json',
}

def remind_lab_to_upload():
    data = '{"text":"Please upload your weekly update figure here: \n %s"}'%(DRIVE_QUEUE)
    response = requests.post(SLACK_API, headers=headers, data=data)

def distribute_link_to_lab():
    data = '{"text":"Lab Meeting PPT for Week  %s: \n %s"}'%(week_number(),DRIVE_FOLDER)
    response = requests.post(SLACK_API, headers=headers, data=data)

def week_number():
    return datetime.date(2010, 6, 16).isocalendar()[1]