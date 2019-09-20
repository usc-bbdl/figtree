# FigTree
[![Build Status](https://travis-ci.org/usc-bbdl/figtree.svg?branch=master)](https://travis-ci.org/usc-bbdl/figtree)
## Continuous integration service for building a PPT from all of the lab members' weekly progress figures.

Designed by [@danhagen](https://www.github.com/danhagen) and [@bc](https://www.github.com/bc) for [ValeroLab](https://valerolab.org)

## Challenge
How do we get a large lab of >8 students in sync about their research on a weekly basis, and how do we establish a culture of discipline with producing visible progress in our Ph.D.'s?

## Approach
Social pressure can drive accountability for projects, so we decided to host a weekly *figure review* with the lab, where we can show one figure or image of our last week's progress, creating a visual topic for discussion.

## UX
1. See this reminder on Slack on Monday, and upload your figure to the google drive folder.  
![Pasted_Image_9_20_19__2_59_PM](https://user-images.githubusercontent.com/13772726/65361109-31a66f80-dbb7-11e9-8448-13ff03223125.png)

2. On the day of the lab meeting, lab members get auto-generated PPT that combines the draft agenda and all of the figures  
![Pasted_Image_9_20_19__2_59_PM](https://user-images.githubusercontent.com/13772726/65361120-38cd7d80-dbb7-11e9-96a3-1637b6a4f8b9.png)

## Result  
![image](https://user-images.githubusercontent.com/13772726/65361205-91047f80-dbb7-11e9-85a7-20901d868750.png)

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
Scripts ran successfully if they return with `0`.
![image](https://user-images.githubusercontent.com/13772726/65361427-53ecbd00-dbb8-11e9-9e23-b6d459157006.png)

