from pptx import Presentation
from pptx.util import Inches,Emu
from datetime import datetime,timedelta #
import os.path
from os import listdir, remove
from shutil import rmtree,copy2
import tkinter as tk
from tkinter import filedialog
from PIL import ImageTk,Image
import subprocess
from win32com.client import Dispatch
import requests
### Assumes files will be downloaded to a folder name TEMP_INPUT_FOLDER
def make_weekly_ppt(TEMP_INPUT_FOLDER):
    """
    Takes in the name of the folder that contains the 'Figure Queue' as well as the 'Meeting Agenda' (TEMP_INPUT_FOLDER), builds a powerpoint presentation for that week's lab meeting, and moves the presentation and corresponding figures to a new folder that will be uploaded to Google Drive (TEMP_OUTPUT_FOLDER)
    """

    if TEMP_INPUT_FOLDER[-1]!="/": TEMP_INPUT_FOLDER=TEMP_INPUT_FOLDER+"/"
    assert os.path.exists(TEMP_INPUT_FOLDER)

    TEMP_OUTPUT_FOLDER = "TEMP_OUTPUT_FOLDER"
    os.path.mkdir(TEMP_OUTPUT_FOLDER)

    todaysDate = datetime.now()
    daysToLabMeeting = 0
    date = todaysDate
    for i in range(7):
        if date.weekday()==1: #Tuesday
            daysToLabMeeting = i
            break
        else:
            date = datetime.now() + timedelta(days=i+1)
    labMeetingDate = datetime.now() + timedelta(days=daysToLabMeeting)
    labMeetingDateStr = labMeetingDate.strftime("%A, %B %d, %Y")
    labMeetingFolderName = labMeetingDate.strftime("%Y_%m_%d Lab Meeting/")
    labMeetingPresentationName = labMeetingDate.strftime(
        "%Y_%m_%d_Lab_Meeting.pptx"
    )

    # Find all figures in Figure Queue
    fileNames = [
        f
        for f in listdir(TEMP_INPUT_FOLDER+"Figure Queue/")
        if os.path.isfile(os.path.join(TEMP_INPUT_FOLDER+"Figure Queue/", f))
    ]

    # Create presentation from BBDL format
    prs = Presentation('graphics_and_templates/template.pptx')
    title_slide_layout = prs.slide_layouts[0]
    agenda_slide_layout = prs.slide_layouts[1]
    content_slide_layout = prs.slide_layouts[2]

    # Add Title slide
    title_slide = prs.slides.add_slide(title_slide_layout)
    title_slide_title = title_slide.placeholders[10]
    title_slide_subtitle = title_slide.placeholders[11]
    title_slide_title.text = "Valero Lab Weekly Meeting"
    title_slide_subtitle.text = labMeetingDateStr

    # Add Agenda slide
    agenda_slide = prs.slides.add_slide(agenda_slide_layout)
    agenda_slide_title = agenda_slide.placeholders[10]
    agenda_slide_subtitle = agenda_slide.placeholders[11]
    agenda_slide_title.text = "Today's Meeting Agenda"
    thisWeeksAgenda = open("agenda.txt", "r").read()
    if (thisWeeksAgenda[:66]
            ==
            '[Add Agenda Items Here. Each on a New Line, Starting with " - "]\n\n'
            ):
        thisWeeksAgenda = thisWeeksAgenda[66:]
    agenda_slide_subtitle.text = thisWeeksAgenda

    # Add all appropriate files as Content slides
    for file in fileNames:
        content_slide = prs.slides.add_slide(content_slide_layout)
        content_slide_title = content_slide.placeholders[10]
        titleStr = file.replace("_"," ")[:file.find(".")]
        presenterName = titleStr[(titleStr.rfind(" ")+1):]
        titleStr = titleStr[:titleStr.rfind(" ")]
        content_slide_title.text = titleStr + " (" + presenterName + ")"
        figure = content_slide.shapes.add_picture(
            "Figure Queue/" +file,
            Inches(0.5),
            Inches(1.75)
        )
        ratio = figure.height/figure.width
        if ratio>4/5:
            figure.height = Emu(Inches(4))
            figure.width = Emu(Inches(4/ratio))
        else:
            figure.width = Emu(Inches(5))
            figure.height = Emu(Inches(5*ratio))
        figure.left = Emu((Inches(13.3333)- figure.width)/2)
        figure.top = Emu((Inches(7.5)- figure.height)/2)
        remove("Figure Queue/"+file)

    # Remove template slides
    rIdToDrop = ['rId1','rId2', 'rId3', 'rId4']
    for i in reversed(range(len(prs.slides._sldIdLst))):
       if prs.slides._sldIdLst[i].rId in rIdToDrop:
           prs.part.drop_rel(prs.slides._sldIdLst[i].rId)
           del prs.slides._sldIdLst[i]

    prs.save(fileName)

    #########################################################


    # Delete contents of agenda.txt and recreate a blank document
    remove("agenda.txt")
    nextWeeksAgenda = open('agenda.txt','w')
    nextWeeksAgenda.write(
        '[Add Agenda Items Here. Each on a New Line, Starting with " - "]\n\n'
        + ' - '
    )
    nextWeeksAgenda.close()

    # Open pptx
    filePath = os.path.abspath(fileName)
    ppt = Dispatch('PowerPoint.Application')
    wb = ppt.Presentations.Open(filePath)
    ppt.Visible = True

    self.master.destroy()

class UploadFileWidget:
    def __init__(self,parent):
        self.master = tk.Tk()
        self.parent = parent
        # self.folderName = folderName
        self.master.configure(background="#991D20")
        self.master.resizable(width=False, height=False)
        self.canvas = tk.Canvas(
            self.master,
            width=500,
            height=200,
            bg="#991D20")
        self.master.title("Upload File to Lab Meeting Folder")
        self.canvas.pack()

        self.widgetTitleText = tk.Text(
            self.master,
            height=1.5,
            width=45,
            bg="#991D20",
            fg="#F5BF33",
            relief=tk.FLAT,
            font=("Helvetica", 14)

        )
        self.widgetTitleText.tag_configure('tag-center', justify='center')
        self.widgetTitleText.insert(tk.INSERT, "Add a File to This Week's Lab Meeting Presentation",'tag-center')
        self.widgetTitleText.place(relx=0.5, rely=0.25, anchor=tk.CENTER)

        self.fileNameLabel = tk.Text(
            self.master,
            height=1.5,
            width=int(len("File Name: ")-2),
            bg="#991D20",
            fg="#F5BF33",
            relief=tk.FLAT,
            font=("Helvetica", 11)

        )
        self.fileNameLabel.insert(tk.INSERT, "File Name: ")
        self.fileNameLabel.place(
            relx=(
                0.40
                - int(len("File Name: "))*11/2/500
            ),
            rely=0.50+(1.5*11/2/200),
            anchor=tk.CENTER
        )

        self.fileNameEntry = tk.Entry(self.master)
        self.fileNameEntry.place(
            relx=0.40 + 20*11/2/500,
            rely=0.50,
            anchor=tk.CENTER
        )
        self.fileNameEntry.configure(
            width=30,
            bg = "white",
            fg="black",
            font=("Helvetica", 11)
        )

        self.uploadedFileNameLabel = tk.Text(
            self.master,
            height=1.5,
            width=int(len("File to be Uploaded: ")-5),
            bg="#991D20",
            fg="#F5BF33",
            relief=tk.FLAT,
            font=("Helvetica", 11)

        )
        self.uploadedFileNameLabel.insert(tk.INSERT, "File to be Uploaded: ")
        self.uploadedFileNameLabel.place(
            relx=(
                0.40
                - int(len("File to be Uploaded: ")-5)*11/2/500
            ),
            rely=0.65+(1.5*11/2/200),
            anchor=tk.CENTER
        )

        self.uploadedFileName = "[None]"
        self.uploadedFileNameText = tk.Text(
            self.master,
            height=1,
            width=34,
            bg="#991D20",
            fg="#F5BF33",
            relief=tk.FLAT,
            font=("Helvetica", 9)
        )
        self.uploadedFileNameText.tag_configure('tag-center', justify='center')
        self.uploadedFileNameText.insert(
            tk.INSERT,
            self.uploadedFileName,
            'tag-center'
        )
        self.uploadedFileNameText.place(
            relx=0.62,
            rely=0.65,
            anchor=tk.CENTER
        )

        self.chooseButton = tk.Button(
            self.master,
            text='Choose',
            font=("Helvetica", 11),
            command=lambda: self.choose_file()
        )
        self.canvas.create_window(
            125,
            0.85*200,
            width=100,
            height=30,
            window=self.chooseButton
        )

        self.saveButton = tk.Button(
            self.master,
            text='Save',
            font=("Helvetica", 11),
            command=lambda: self.save_file()
        )
        self.canvas.create_window(
            250,
            0.85*200,
            width=100,
            height=30,
            window=self.saveButton
        )

        self.exitButton = tk.Button(
            self.master,
            text='Exit',
            font=("Helvetica", 11),
            command=lambda: self.master.destroy()
        )
        self.canvas.create_window(
            375,
            0.85*200,
            width=100,
            height=30,
            window=self.exitButton
        )
        self.master.mainloop()

    def choose_file(self):
        self.uploadedFileName = filedialog.askopenfilename(
            initialdir = os.path.dirname(os.path.abspath(__file__)),
            title = "Add File to Weekly Presentation",
            filetypes = (
                ("PNG files", "*.png"),
                ("all files","*.*")
            )
        )
        self.uploadedFileNameText = tk.Text(
            self.master,
            height=1,
            width=34,
            bg="#991D20",
            fg="#F5BF33",
            relief=tk.FLAT,
            font=("Helvetica", 9)
        )
        self.uploadedFileNameText.tag_configure('tag-center', justify='center')
        self.uploadedFileNameText.insert(
            tk.INSERT,
            self.uploadedFileName,
            'tag-center'
        )
        self.uploadedFileNameText.place(
            relx=0.62,
            rely=0.65,
            anchor=tk.CENTER
        )
    def save_file(self):
        if "[None]" in self.uploadedFileName:
            self.warningText = tk.Text(
                self.master,
                height=1,
                width=20,
                bg="#991D20",
                fg="white",
                relief=tk.FLAT,
                font=("Helvetica", 10)

            )
            self.warningText.tag_configure('tag-center', justify='center')
            self.warningText.insert(tk.INSERT, "Choose File to Upload",'tag-center')
            self.warningText.place(relx=0.5, rely=0.375, anchor=tk.CENTER)
        else:
            if " " in self.fileNameEntry.get():
                self.warningText = tk.Text(
                    self.master,
                    height=1,
                    width=20,
                    bg="#991D20",
                    fg="white",
                    relief=tk.FLAT,
                    font=("Helvetica", 10)

                )
                self.warningText.tag_configure('tag-center', justify='center')
                self.warningText.insert(tk.INSERT, "No Spaces!",'tag-center')
                self.warningText.place(relx=0.5, rely=0.375, anchor=tk.CENTER)
            else:
                copy2(
                    self.uploadedFileName,
                    (
                        os.path.dirname(os.path.abspath(__file__))
                        +"/Figure Queue/"
                        + self.fileNameEntry.get()
                    )
                )
                self.master.destroy()
class WeeklyMeetingGUI:
    def __init__(self):
        self.master = tk.Tk()
        self.master.configure(background="#991D20")
        self.master.resizable(width=False, height=False)
        self.canvas = tk.Canvas(self.master, width = 400, height = 400)
        self.master.title("Valero Lab Meetings")
        background_img_png = Image.open("graphics_and_templates/Valero_lab_graphics_sans_yellow.png")
        background_img_png = background_img_png.resize((400, 400), Image.ANTIALIAS)
        self.BTC_img = ImageTk.PhotoImage(
            background_img_png,
            master=self.canvas
        )
        self.BTC_img_label = tk.Label(self.canvas, image=self.BTC_img)
        self.BTC_img_label.image = self.BTC_img
        self.BTC_img_label.grid(row=2, column=0)
        self.canvas.pack()
        self.deleteFolder = tk.IntVar()

        todaysDate = datetime.now()
        daysToLabMeeting = 0
        date = todaysDate
        for i in range(7):
            if date.weekday()==1: #Tuesday
                daysToLabMeeting = i
                break
            else:
                date = datetime.now() + timedelta(days=i+1)
        self.labMeetingDate = datetime.now() + timedelta(days=daysToLabMeeting)
        self.labMeetingDateStr = self.labMeetingDate.strftime("%A, %B %d, %Y")

        self.labMeetingDateText = tk.Text(
            self.master,
            height=1.5,
            width=25,
            bg="#991D20",
            fg="#F5BF33",
            relief=tk.FLAT,
            font=("Helvetica", 18)

        )
        self.labMeetingDateText.tag_configure('tag-center', justify='center')
        self.labMeetingDateText.insert(tk.INSERT, "Next Lab Meeting:\n" + self.labMeetingDateStr,'tag-center')
        self.labMeetingDateText.place(
            relx=0.5,
            rely=0.575,
            anchor=tk.CENTER
        )

        ### COL 1

        # row 1
        self.openWklyReportButton = tk.Button(
            self.master,
            text='Open Weekly Report',
            font=("Helvetica", 11),
            command= lambda: self.open_weekly_presentation()
        )
        self.canvas.create_window(
            110,
            0.775*400,
            width=160,
            height=30,
            window=self.openWklyReportButton
        )

        # row 2
        self.mkWklyReportButton = tk.Button(
            self.master,
            text='Make Weekly Report',
            font=("Helvetica", 11),
            command=lambda: self.make_weekly_ppt()
        )
        self.canvas.create_window(
            110,
            0.875*400,
            width=160,
            height=30,
            window=self.mkWklyReportButton
        )

        ### COL 2

        # row 1
        self.addFileButton = tk.Button(
            self.master,
            text='Add File',
            font=("Helvetica", 11),
            command= lambda: self.add_figure_to_queue()
        )
        self.canvas.create_window(
            290,
            0.775*400,
            width=160,
            height=30,
            window=self.addFileButton
        )

        # row 2
        self.exitButton = tk.Button(
            self.master,
            text='Exit',
            font=("Helvetica", 11),
            command=lambda: self.master.destroy()
        )
        self.canvas.create_window(
            290,
            0.875*400,
            width=160,
            height=30,
            window=self.exitButton
        )

        self.master.mainloop()

    def open_weekly_presentation(self):
        presentationFileName = filedialog.askopenfilename(
            initialdir = os.path.dirname(os.path.abspath(__file__)),
            title = "Select Weekly Presentation",
            filetypes = (
                ("Microsoft PowerPoint Presentation", "*.pptx"),
                ("all files","*.*")
            )
        )
        ppt = Dispatch('PowerPoint.Application')
        wb = ppt.Presentations.Open(presentationFileName)
        ppt.Visible = True
        self.master.destroy()
    def add_figure_to_queue(self):
        UploadFileWidget(self.master)
    def make_weekly_ppt(self):
        labMeetingDateStr = self.labMeetingDate.strftime("%A, %B %d, %Y")
        fileName = self.labMeetingDate.strftime("%Y_%m_%d_LabMeeting.pptx")

        if os.path.exists(fileName):
            # Find all files in this week's lab folder (excluding agenda)
            fileNames = [
                f
                for f in listdir("Figure Queue/")
                if os.path.isfile(os.path.join("Figure Queue/", f))
            ]

            # Create presentation from BBDL format
            prs = Presentation(fileName)
            title_slide_layout = prs.slide_layouts[0]
            agenda_slide_layout = prs.slide_layouts[1]
            content_slide_layout = prs.slide_layouts[2]

            # Add Agenda slide
            agenda_slide = prs.slides[1]
            agenda_slide_subtitle = agenda_slide.placeholders[11]
            previouslyAddedAgenda = agenda_slide_subtitle.text
            thisWeeksAgenda = open("agenda.txt", "r").read()
            if (thisWeeksAgenda != '[Add Agenda Items Here. Each on a New Line, Starting with " - "]\n\n - '):
                if (thisWeeksAgenda[:66]
                        ==
                        '[Add Agenda Items Here. Each on a New Line, Starting with " - "]\n\n'
                        ):
                    thisWeeksAgenda = thisWeeksAgenda[66:]
            else:
                thisWeeksAgenda = ""
            agenda_slide_subtitle.text = (
                previouslyAddedAgenda
                + "\n"
                + thisWeeksAgenda
            )

            # Add all appropriate files as Content slides
            for file in fileNames:
                content_slide = prs.slides.add_slide(content_slide_layout)
                content_slide_title = content_slide.placeholders[10]
                titleStr = file.replace("_"," ")[:file.find(".")]
                presenterName = titleStr[(titleStr.rfind(" ")+1):]
                titleStr = titleStr[:titleStr.rfind(" ")]
                content_slide_title.text = titleStr + " (" + presenterName + ")"
                figure = content_slide.shapes.add_picture(
                    "Figure Queue/" +file,
                    Inches(0.5),
                    Inches(1.75)
                )
                ratio = figure.height/figure.width
                if ratio>4/5:
                    figure.height = Emu(Inches(4))
                    figure.width = Emu(Inches(4/ratio))
                else:
                    figure.width = Emu(Inches(5))
                    figure.height = Emu(Inches(5*ratio))
                figure.left = Emu((Inches(13.3333)- figure.width)/2)
                figure.top = Emu((Inches(7.5)- figure.height)/2)
                remove("Figure Queue/"+file)

            prs.save(fileName)

            #########################################################


            # Delete contents of agenda.txt and recreate a blank document
            remove("agenda.txt")
            nextWeeksAgenda = open('agenda.txt','w')
            nextWeeksAgenda.write(
                '[Add Agenda Items Here. Each on a New Line, Starting with " - "]\n\n'
                + ' - '
            )
            nextWeeksAgenda.close()

            # Open Presenation
            filePath = os.path.abspath(fileName)
            ppt = Dispatch('PowerPoint.Application')
            wb = ppt.Presentations.Open(filePath)
            ppt.Visible = True

            self.master.destroy()
        else:
            # Find all files in this week's lab folder (excluding agenda)
            fileNames = [
                f
                for f in listdir("Figure Queue/")
                if os.path.isfile(os.path.join("Figure Queue/", f))
            ]

            # Create presentation from BBDL format
            prs = Presentation('graphics_and_templates/bbdl_template.pptx')
            title_slide_layout = prs.slide_layouts[0]
            agenda_slide_layout = prs.slide_layouts[1]
            content_slide_layout = prs.slide_layouts[2]

            # Add Title slide
            title_slide = prs.slides.add_slide(title_slide_layout)
            title_slide_title = title_slide.placeholders[10]
            title_slide_subtitle = title_slide.placeholders[11]
            title_slide_title.text = "Valero Lab Weekly Meeting"
            title_slide_subtitle.text = labMeetingDateStr

            # Add Agenda slide
            agenda_slide = prs.slides.add_slide(agenda_slide_layout)
            agenda_slide_title = agenda_slide.placeholders[10]
            agenda_slide_subtitle = agenda_slide.placeholders[11]
            agenda_slide_title.text = "Today's Meeting Agenda"
            thisWeeksAgenda = open("agenda.txt", "r").read()
            if (thisWeeksAgenda[:66]
                    ==
                    '[Add Agenda Items Here. Each on a New Line, Starting with " - "]\n\n'
                    ):
                thisWeeksAgenda = thisWeeksAgenda[66:]
            agenda_slide_subtitle.text = thisWeeksAgenda

            # Add all appropriate files as Content slides
            for file in fileNames:
                content_slide = prs.slides.add_slide(content_slide_layout)
                content_slide_title = content_slide.placeholders[10]
                titleStr = file.replace("_"," ")[:file.find(".")]
                presenterName = titleStr[(titleStr.rfind(" ")+1):]
                titleStr = titleStr[:titleStr.rfind(" ")]
                content_slide_title.text = titleStr + " (" + presenterName + ")"
                figure = content_slide.shapes.add_picture(
                    "Figure Queue/" +file,
                    Inches(0.5),
                    Inches(1.75)
                )
                ratio = figure.height/figure.width
                if ratio>4/5:
                    figure.height = Emu(Inches(4))
                    figure.width = Emu(Inches(4/ratio))
                else:
                    figure.width = Emu(Inches(5))
                    figure.height = Emu(Inches(5*ratio))
                figure.left = Emu((Inches(13.3333)- figure.width)/2)
                figure.top = Emu((Inches(7.5)- figure.height)/2)
                remove("Figure Queue/"+file)

            # Remove template slides
            rIdToDrop = ['rId1','rId2', 'rId3', 'rId4']
            for i in reversed(range(len(prs.slides._sldIdLst))):
               if prs.slides._sldIdLst[i].rId in rIdToDrop:
                   prs.part.drop_rel(prs.slides._sldIdLst[i].rId)
                   del prs.slides._sldIdLst[i]

            prs.save(fileName)

            #########################################################


            # Delete contents of agenda.txt and recreate a blank document
            remove("agenda.txt")
            nextWeeksAgenda = open('agenda.txt','w')
            nextWeeksAgenda.write(
                '[Add Agenda Items Here. Each on a New Line, Starting with " - "]\n\n'
                + ' - '
            )
            nextWeeksAgenda.close()

            # Open pptx
            filePath = os.path.abspath(fileName)
            ppt = Dispatch('PowerPoint.Application')
            wb = ppt.Presentations.Open(filePath)
            ppt.Visible = True

            self.master.destroy()

my_gui = WeeklyMeetingGUI()
