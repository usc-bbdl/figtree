# figtree
##

## Challenge
How do we get a large lab of >8 students in sync about their research on a weekly basis, and how do we establish a culture of discipline with producing visible progress in our Ph.D.'s?

## Approach
Social pressure can drive accountability for projects, so we decided to host a weekly *figure review* with the lab, where we can show one figure or image of our last week's progress, creating a visual topic for discussion.

## UX
1. See this reminder on Slack on Monday, and upload your figure to the google drive folder.
image of reminder
2. On the day of the lab meeting, lab members get auto-generated PPT that combines the draft agenda and all of the figures
image of valerolab slack message

## Result
PPT pic

### Technical Overview
A Zapier CRON job pings a Slack channel webhook to post the upload link as a reminder 1x/wk. Travis-CI is set up with a 1x/wk CRON job that authenticates & downloads the contents of the Figures Drive folder, computes a PPT, uploads the results to the drive, clears the queue, and pings Slack via a webhook with the direct link to the PPT for download (link is mobile friendly too). Credentials and drive links are set as private environment variables within the Travis-CI GUI.

### Languages
Python, Bash

### Install
Put this one in Travis-CI
```bash
export DRIVE_FOLDER="https://drive.google.com/..."
export DRIVE_QUEUE="https://drive.google.com/..."
export SLACK_API="https://hooks.slack.com/services/..."
CREDENTIALS='{"access_token": ...'
echo $CREDENTIALS >> mycreds1.txt 
```
Where the credentials are the access token for the Google Drive Python API.

```bash
python3 download_from_gdrive.py
python3 build_weekly_ppt.py >> temp_output.txt
python3 upload_to_gdrive.py
```
