import datetime
import os

import requests

DRIVE_FOLDER = os.environ['DRIVE_FOLDER']
DRIVE_QUEUE = os.environ['DRIVE_QUEUE']
SLACK_API = os.environ['SLACK_API']
headers = {
    'Content-type': 'application/json',
}


def slack_confirmed_receipt(response):
    return (response.status_code == 200)

def remind_lab_to_upload():
    data = '{"text":"Please upload your weekly update figure here: \n %s"}' % (DRIVE_QUEUE)
    response = requests.post(SLACK_API, headers=headers, data=data)
    print("WebHook Successful: " + str(slack_confirmed_receipt(response)))

def distribute_link_to_lab(link_string):
    data = '{"text":"Download ValeroLab Meeting PPT, Week %s: \n %s"}' % (week_number(), link_string)
    response = requests.post(SLACK_API, headers=headers, data=data)
    print("WebHook Successful: " + str(slack_confirmed_receipt(response)))

def week_number():
    return datetime.datetime.now().isocalendar()[1]
