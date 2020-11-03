# ******************************************************************************************************
#  DESCRIPTION:
#         This program is used to parse the CPW leftover list and notify recipients of available tags
#
#  CHANGES:
#         WHO         REV    DATE        DETAIL
#         megeep      1.5    11/1/2020   Added dictionary indexing of tag data
#         ctstewar    1.4    10/28/2020  Added texting
#         ctstewar    1.3    09/29/2020  Updated sender to tag.holler@gmail.com, created email function
#         foxyt       1.2    09/15/2020  Created dictionary and keys for searching leftover contents
#         megeep      1.1    07/01/2019  Improved handling of hunt codes and recipients
#         megeep      1.0    07/01/2018  Initial creation?
#
#  NOTES:
#          Originally written for Python2. Python version inconsistencies remain.
#
# Copyright (2018) Megee Software Limited
#
# ******************************************************************************************************
# ******************************************************************************************************
#                                       IMPORT LIBRARIES
# ******************************************************************************************************
# Used to parse PDF
import PyPDF2

# regex
import re

# Import requests (to download the page)
import requests

# Import BeautifulSoup (to parse what we download)
from bs4 import BeautifulSoup  # not used

# Import Time (to add a delay between the times the scape runs)
import time

# Used to build SMTP oject for email interface
import smtplib

# Used to convert text to standard MIME output
from email.mime.text import MIMEText

# Used to generate current time for email
from datetime import datetime


# ******************************************************************************************************
#                                       DEFINE FUNCTIONS
# ******************************************************************************************************
def send_email(recipient, subject_line, email_body):
    # Define SMTP settings
    SMTP_SERVER = "smtp.gmail.com"
    SMTP_PORT = 587
    SMTP_USERNAME = "tag.holler@gmail.com"
    SMTP_PASSWORD = "020hunts4life"

    # Pull the current time as dd/mm/YY H:M:S
    now = datetime.now()
    dt_string = now.strftime("%m/%d/%Y %H:%M:%S")

    # Reassign variables
    EMAIL_SPACE = ", "
    EMAIL_SUBJECT = subject_line
    EMAIL_FROM = SMTP_USERNAME
    DATA = email_body + '\n\n' + 'Notification compiled at ' + dt_string

    # Allow single recipient by ensuring array
    if isinstance(recipient, list):
        EMAIL_TO = recipient
    else:
        EMAIL_TO = [recipient]

    # Build email
    msg = MIMEText(DATA)
    msg['Subject'] = EMAIL_SUBJECT
    msg['To'] = EMAIL_SPACE.join(EMAIL_TO)
    msg['From'] = EMAIL_FROM

    try:
        # Create SMTP object mail and send email
        mail = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        mail.starttls()
        mail.login(SMTP_USERNAME, SMTP_PASSWORD)
        mail.sendmail(EMAIL_FROM, EMAIL_TO, msg.as_string())
        # Send text message through SMS gateway of destination number
        mail.quit()
    except:
        print('ERROR: Unspecified email delivery failure.')


# ******************************************************************************************************
#                                            MAIN
# ******************************************************************************************************
# Initialize variables to allow the code to enter correctly into the time check loop
CPW_time = '00:00:01'
old_CPW_time = '00:00:00'
Text = 'Run Date and Time: and then 00:00:01'
CPW_website = 'cpwshop.com'
email_footer = CPW_website + '\n\n' + 'Thank you for using the Tag Holler system. For issues or concerns contact the tag holler team by responding to this message.'
while True:
    try:
        # Pull PDF from CPW website and parse using PyPDF2
        file_url = "https://cpw.state.co.us/Documents/Leftover.pdf"
        r = requests.get(file_url, stream=True)
        with open("python.pdf", "wb") as pdf:
            for chunk in r.iter_content(chunk_size=1024):
                # writing one chunk at a time to pdf file
                if chunk:
                    pdf.write(chunk)
        object = PyPDF2.PdfFileReader("python.pdf")

        # Determine number of pages
        NumPages = object.getNumPages()

        # Define email lists
        # thirstyBois = ['megeep@gmail.com', 'colinstewrat@gmail.com', 'a.lodolce@gmail.com']
        # goatBois = thirstyBois + ['mlanderson4645@gmail.com', 'cswatk2@gmail.com', 'tyler.fox09@gmail.com']
        # elkBois = thirstyBois + ['rileygelatt@gmail.com']
        # deerBois = ['megeep@gmail.com', 'shco6527@gmail.com', 'a.lodolce@gmail.com', 'colinstewrat@gmail.com',
        #                     '7209365042@vtext.com', '3034374578@messaging.sprintpcs.com']
        testBois = ['megeep@gmail.com', 'colinstewrat@gmail.com', 'rileygelatt@gmail.com',
                    'a.lodolce@gmail.com', '7209365042@vtext.com', '3034374578@messaging.sprintpcs.com',
                    '9707788390@messaging.sprintpcs.com']
        # Makes a dictionary object. The tag code is the key, while the email list is the value. Use the key to retrieve the value
        tagDict = {"DF012O3R", "DF015O3R", "DF025O3R", "DF035O3R", "DF044O3R", "DF049O3R", "DF444O3R", "DF501O3R",
                   "DF511O3R", "DM006O3R", "DM012O3R", "DM015O3R", "DM025O3R", "DM031O3R", "DM035O3R", "DM043O3R",
                   "DM044O3R", "DM049O3R", "DM050O3R", "DM051O3R", "DM055O3R", "DM059O3R", "DM068O3R", "DM069O3R",
                   "DM082O3R", "DM085O3R", "DM140O3R", "DM161O3R", "DM171O3R", "DM444O3R", "DM551O3R", "DE131P3R",
                   "DM131O3R", "EF012O3R", "EF014O3R", "EF016O3R", "EF032O3R", "EF042O3R", "EF161O3R", "EF231O3R",
                   "EF471O3R", "EF500O3R","EF007O3R", "EF020O3R", "EF048O3R", "EF049O3R", "EF049S3R", "EF055O3R",
                   "EF056O3R", "EF551O3R", "EM020O3R", "EM039O3R", "EM048O3R", "EM049O3R", "EM050O3R", "EM056O3R",
                   "EM057O3R", "EM069O3R", "EM500O3R", "EM501O3R"}

        #TestTag , "EF131P3R"
        # EF131P3R = {"List": "TEST TAG", "Drawn At": "Leftover", "%Draw": "", "Success": "", "3 yr Avg%": "", "OTC ELK": "Y",
        #             "%Public": "0.11", "ANY BULL": "N", "Driving (hr)": "3.1", "#Tags": "6", "Public Per Tag": "",
        #             "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "", "Sq Mi Public": "",
        #             "Bull to Cow": "23", }

        DF012O3R = {"List": "A", "Drawn At": "2", "%Draw": "0.14", "Success": "1", "3 yr Avg%": "0.69", "OTC ELK": "Y",
                    "%Public": "0.55", "ANY BULL": "N", "Driving (hr)": "3.5", "#Tags": "10", "Public Per Tag": "",
                    "Leftover Day": "0", "3 year leftover day avg": "0.333333333333333", "Valid Units": "",
                    "Sq Mi Public": "", "Bull to Cow": "23", }
        DF015O3R = {"List": "B", "Drawn At": "0", "%Draw": "0.44", "Success": "0.39", "3 yr Avg%": "0.47", "OTC ELK": "Y",
                    "%Public": "0.68", "ANY BULL": "N", "Driving (hr)": "2.5", "#Tags": "10", "Public Per Tag": "",
                    "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "", "Sq Mi Public": "",
                    "Bull to Cow": "30", }
        DF025O3R = {"List": "B", "Drawn At": "Choice 2", "%Draw": "0.39", "Success": "0.43", "3 yr Avg%": "0.4",
                    "OTC ELK": "Y", "%Public": "0.83", "ANY BULL": "N", "Driving (hr)": "2.4", "#Tags": "70",
                    "Public Per Tag": "", "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "",
                    "Sq Mi Public": "", "Bull to Cow": "23", }
        DF035O3R = {"List": "B", "Drawn At": "0", "%Draw": "0.46", "Success": "0.26", "3 yr Avg%": "0.41", "OTC ELK": "Y",
                    "%Public": "0.7", "ANY BULL": "N", "Driving (hr)": "2.1", "#Tags": "10", "Public Per Tag": "",
                    "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "", "Sq Mi Public": "",
                    "Bull to Cow": "25", }
        DF044O3R = {"List": "B", "Drawn At": "0", "%Draw": "0.14", "Success": "0.43", "3 yr Avg%": "0.55", "OTC ELK": "Y",
                    "%Public": "0.78", "ANY BULL": "N", "Driving (hr)": "2.1", "#Tags": "15", "Public Per Tag": "",
                    "Leftover Day": "0", "3 year leftover day avg": "0.333333333333333", "Valid Units": "",
                    "Sq Mi Public": "", "Bull to Cow": "22", }
        DF049O3R = {"List": "A", "Drawn At": "5", "%Draw": "0.71", "Success": "1", "3 yr Avg%": "0.94", "OTC ELK": "N",
                    "%Public": "0.73", "ANY BULL": "Y", "Driving (hr)": "1.9", "#Tags": "10", "Public Per Tag": "",
                    "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "", "Sq Mi Public": "",
                    "Bull to Cow": "31", }
        DF444O3R = {"List": "B", "Drawn At": "0", "%Draw": "0.17", "Success": "0.75", "3 yr Avg%": "0.53", "OTC ELK": "Y",
                    "%Public": "0.67", "ANY BULL": "N", "Driving (hr)": "2.9", "#Tags": "40", "Public Per Tag": "",
                    "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "", "Sq Mi Public": "",
                    "Bull to Cow": "22", }
        DF501O3R = {"List": "A", "Drawn At": "1", "%Draw": "0.81", "Success": "0.45", "3 yr Avg%": "0.48", "OTC ELK": "N",
                    "%Public": "0.9", "ANY BULL": "Y", "Driving (hr)": "1.5", "#Tags": "60", "Public Per Tag": "",
                    "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "", "Sq Mi Public": "",
                    "Bull to Cow": "36", }
        DF511O3R = {"List": "A", "Drawn At": "2", "%Draw": "0.5", "Success": "0.62", "3 yr Avg%": "0.87", "OTC ELK": "Y",
                    "%Public": "0.62", "ANY BULL": "N", "Driving (hr)": "1.3", "#Tags": "20", "Public Per Tag": "",
                    "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "", "Sq Mi Public": "",
                    "Bull to Cow": "25", }
        DM006O3R = {"List": "A", "Drawn At": "5", "%Draw": "0.55", "Success": "0.68", "3 yr Avg%": "0.81", "OTC ELK": "Y",
                    "%Public": "0.68", "ANY BULL": "N", "Driving (hr)": "2.5", "#Tags": "25", "Public Per Tag": "",
                    "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "", "Sq Mi Public": "",
                    "Bull to Cow": "34", }
        DM012O3R = {"List": "A", "Drawn At": "Choice 2", "%Draw": "0.58", "Success": "0.31", "3 yr Avg%": "0.44",
                    "OTC ELK": "Y", "%Public": "0.55", "ANY BULL": "N", "Driving (hr)": "3.5", "#Tags": "1090",
                    "Public Per Tag": "", "Leftover Day": "11", "3 year leftover day avg": "7.33333333333333",
                    "Valid Units": "", "Sq Mi Public": "", "Bull to Cow": "23", }
        DM015O3R = {"List": "A", "Drawn At": "Choice 3", "%Draw": "0.06", "Success": "0.3", "3 yr Avg%": "0.35",
                    "OTC ELK": "Y", "%Public": "0.68", "ANY BULL": "N", "Driving (hr)": "2.5", "#Tags": "370",
                    "Public Per Tag": "", "Leftover Day": "6", "3 year leftover day avg": "5.33333333333333",
                    "Valid Units": "", "Sq Mi Public": "", "Bull to Cow": "30", }
        DM025O3R = {"List": "A", "Drawn At": "0", "%Draw": "1", "Success": "0.25", "3 yr Avg%": "0.36", "OTC ELK": "Y",
                    "%Public": "0.83", "ANY BULL": "N", "Driving (hr)": "2.4", "#Tags": "285", "Public Per Tag": "",
                    "Leftover Day": "8", "3 year leftover day avg": "6", "Valid Units": "", "Sq Mi Public": "",
                    "Bull to Cow": "23", }
        DM031O3R = {"List": "A", "Drawn At": "1", "%Draw": "0.77", "Success": "0.62", "3 yr Avg%": "0.66", "OTC ELK": "Y",
                    "%Public": "0.58", "ANY BULL": "N", "Driving (hr)": "3.5", "#Tags": "230", "Public Per Tag": "",
                    "Leftover Day": "3", "3 year leftover day avg": "1.66666666666667", "Valid Units": "",
                    "Sq Mi Public": "", "Bull to Cow": "", }
        DM035O3R = {"List": "A", "Drawn At": "0", "%Draw": "0.96", "Success": "0.28", "3 yr Avg%": "0.37", "OTC ELK": "Y",
                    "%Public": "0.7", "ANY BULL": "N", "Driving (hr)": "2.1", "#Tags": "640", "Public Per Tag": "",
                    "Leftover Day": "15", "3 year leftover day avg": "11.6666666666667", "Valid Units": "",
                    "Sq Mi Public": "", "Bull to Cow": "25", }
        DM043O3R = {"List": "A", "Drawn At": "0", "%Draw": "0.64", "Success": "0.45", "3 yr Avg%": "0.52", "OTC ELK": "Y",
                    "%Public": "0.73", "ANY BULL": "N", "Driving (hr)": "2.8", "#Tags": "175", "Public Per Tag": "",
                    "Leftover Day": "7", "3 year leftover day avg": "4.66666666666667", "Valid Units": "",
                    "Sq Mi Public": "", "Bull to Cow": "24", }
        DM044O3R = {"List": "A", "Drawn At": "14", "%Draw": "0.41", "Success": "0.53", "3 yr Avg%": "0.7", "OTC ELK": "Y",
                    "%Public": "0.78", "ANY BULL": "N", "Driving (hr)": "2.1", "#Tags": "20", "Public Per Tag": "",
                    "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "", "Sq Mi Public": "",
                    "Bull to Cow": "22", }
        DM049O3R = {"List": "A", "Drawn At": "1", "%Draw": "0.31", "Success": "0.46", "3 yr Avg%": "0.47", "OTC ELK": "N",
                    "%Public": "0.73", "ANY BULL": "Y", "Driving (hr)": "1.9", "#Tags": "600", "Public Per Tag": "",
                    "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "", "Sq Mi Public": "",
                    "Bull to Cow": "31", }
        DM050O3R = {"List": "A", "Drawn At": "1", "%Draw": "0.1", "Success": "0.4", "3 yr Avg%": "0.36", "OTC ELK": "N",
                    "%Public": "0.49", "ANY BULL": "Y", "Driving (hr)": "1.7", "#Tags": "300", "Public Per Tag": "",
                    "Leftover Day": "1", "3 year leftover day avg": "1.66666666666667", "Valid Units": "",
                    "Sq Mi Public": "", "Bull to Cow": "36", }
        DM051O3R = {"List": "A", "Drawn At": "1", "%Draw": "0.51", "Success": "0.52", "3 yr Avg%": "0.61", "OTC ELK": "N",
                    "%Public": "0.45", "ANY BULL": "Y", "Driving (hr)": "1", "#Tags": "90", "Public Per Tag": "",
                    "Leftover Day": "5", "3 year leftover day avg": "2", "Valid Units": "", "Sq Mi Public": "",
                    "Bull to Cow": "15", }
        DM055O3R = {"List": "A", "Drawn At": "5", "%Draw": "0.25", "Success": "0.44", "3 yr Avg%": "0.65", "OTC ELK": "Y",
                    "%Public": "0.89", "ANY BULL": "N", "Driving (hr)": "3", "#Tags": "80", "Public Per Tag": "",
                    "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "", "Sq Mi Public": "",
                    "Bull to Cow": "24", }
        DM059O3R = {"List": "A", "Drawn At": "2", "%Draw": "0.81", "Success": "0.4", "3 yr Avg%": "0.63", "OTC ELK": "Y",
                    "%Public": "0.35", "ANY BULL": "N", "Driving (hr)": "1.6", "#Tags": "75", "Public Per Tag": "",
                    "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "", "Sq Mi Public": "",
                    "Bull to Cow": "25", }
        DM068O3R = {"List": "A", "Drawn At": "2", "%Draw": "0.01", "Success": "0.61", "3 yr Avg%": "0.55", "OTC ELK": "Y",
                    "%Public": "0.85", "ANY BULL": "N", "Driving (hr)": "3.3", "#Tags": "205", "Public Per Tag": "",
                    "Leftover Day": "5", "3 year leftover day avg": "3.33333333333333", "Valid Units": "",
                    "Sq Mi Public": "", "Bull to Cow": "18", }
        DM069O3R = {"List": "A", "Drawn At": "0", "%Draw": "0.14", "Success": "0.41", "3 yr Avg%": "0.52", "OTC ELK": "N",
                    "%Public": "0.35", "ANY BULL": "Y", "Driving (hr)": "2.3", "#Tags": "490", "Public Per Tag": "",
                    "Leftover Day": "2", "3 year leftover day avg": "3", "Valid Units": "", "Sq Mi Public": "",
                    "Bull to Cow": "39", }
        DM082O3R = {"List": "A", "Drawn At": "1", "%Draw": "0.83", "Success": "0.45", "3 yr Avg%": "0.45", "OTC ELK": "Y",
                    "%Public": "0.63", "ANY BULL": "N", "Driving (hr)": "2.9", "#Tags": "125", "Public Per Tag": "",
                    "Leftover Day": "1", "3 year leftover day avg": "1", "Valid Units": "", "Sq Mi Public": "",
                    "Bull to Cow": "52", }
        DM085O3R = {"List": "A", "Drawn At": "0", "%Draw": "0.31", "Success": "0.42", "3 yr Avg%": "0.54", "OTC ELK": "Y",
                    "%Public": "0.16", "ANY BULL": "N", "Driving (hr)": "2.7", "#Tags": "300", "Public Per Tag": "",
                    "Leftover Day": "2", "3 year leftover day avg": "5.33333333333333", "Valid Units": "",
                    "Sq Mi Public": "", "Bull to Cow": "22", }
        DM140O3R = {"List": "A", "Drawn At": "Leftover", "%Draw": "", "Success": "0.63", "3 yr Avg%": "0.65",
                    "OTC ELK": "N", "%Public": "0.1", "ANY BULL": "N", "Driving (hr)": "3.2", "#Tags": "100",
                    "Public Per Tag": "", "Leftover Day": "1", "3 year leftover day avg": "0.333333333333333",
                    "Valid Units": "", "Sq Mi Public": "", "Bull to Cow": "", }
        DM161O3R = {"List": "A", "Drawn At": "3", "%Draw": "1", "Success": "0.7", "3 yr Avg%": "0.74", "OTC ELK": "Y",
                    "%Public": "0.75", "ANY BULL": "N", "Driving (hr)": "3", "#Tags": "25", "Public Per Tag": "",
                    "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "", "Sq Mi Public": "",
                    "Bull to Cow": "34", }
        DM171O3R = {"List": "A", "Drawn At": "0", "%Draw": "0.59", "Success": "0.2", "3 yr Avg%": "0.38", "OTC ELK": "Y",
                    "%Public": "0.69", "ANY BULL": "N", "Driving (hr)": "3", "#Tags": "25", "Public Per Tag": "",
                    "Leftover Day": "0", "3 year leftover day avg": "0.333333333333333", "Valid Units": "",
                    "Sq Mi Public": "", "Bull to Cow": "34", }
        DM444O3R = {"List": "A", "Drawn At": "0", "%Draw": "0.53", "Success": "0.57", "3 yr Avg%": "0.53", "OTC ELK": "Y",
                    "%Public": "0.67", "ANY BULL": "N", "Driving (hr)": "2.9", "#Tags": "100", "Public Per Tag": "",
                    "Leftover Day": "1", "3 year leftover day avg": "1.66666666666667", "Valid Units": "",
                    "Sq Mi Public": "", "Bull to Cow": "22", }
        DM551O3R = {"List": "A", "Drawn At": "5", "%Draw": "0.79", "Success": "0.62", "3 yr Avg%": "0.72", "OTC ELK": "Y",
                    "%Public": "0.86", "ANY BULL": "N", "Driving (hr)": "3.2", "#Tags": "35", "Public Per Tag": "",
                    "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "", "Sq Mi Public": "",
                    "Bull to Cow": "24", }
        DE131P3R = {"List": "A", "Drawn At": "Leftover", "%Draw": "", "Success": "", "3 yr Avg%": "", "OTC ELK": "Y",
                    "%Public": "0.11", "ANY BULL": "N", "Driving (hr)": "3.1", "#Tags": "25", "Public Per Tag": "",
                    "Leftover Day": "10", "3 year leftover day avg": "4.33333333333333", "Valid Units": "",
                    "Sq Mi Public": "", "Bull to Cow": "23", }
        DM131O3R = {"List": "A", "Drawn At": "0", "%Draw": "0.88", "Success": "0", "3 yr Avg%": "0.06", "OTC ELK": "Y",
                    "%Public": "0.11", "ANY BULL": "N", "Driving (hr)": "3.1", "#Tags": "35", "Public Per Tag": "",
                    "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "", "Sq Mi Public": "",
                    "Bull to Cow": "23", }
        EF012O3R = {"List": "B", "Drawn At": "Leftover", "%Draw": "", "Success": "0.44", "3 yr Avg%": "0.27",
                    "OTC ELK": "Y", "%Public": "0.55", "ANY BULL": "N", "Driving (hr)": "3.5", "#Tags": "846",
                    "Public Per Tag": "", "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "",
                    "Sq Mi Public": "", "Bull to Cow": "23", }
        EF014O3R = {"List": "B", "Drawn At": "0", "%Draw": "0.89", "Success": "0.28", "3 yr Avg%": "0.26", "OTC ELK": "Y",
                    "%Public": "0.85", "ANY BULL": "N", "Driving (hr)": "2.5", "#Tags": "75", "Public Per Tag": "",
                    "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "", "Sq Mi Public": "",
                    "Bull to Cow": "16", }
        EF016O3R = {"List": "B", "Drawn At": "Leftover", "%Draw": "", "Success": "0.28", "3 yr Avg%": "0.24",
                    "OTC ELK": "Y", "%Public": "0.52", "ANY BULL": "N", "Driving (hr)": "3", "#Tags": "51",
                    "Public Per Tag": "", "Leftover Day": "109", "3 year leftover day avg": "114", "Valid Units": "",
                    "Sq Mi Public": "", "Bull to Cow": "34", }
        EF032O3R = {"List": "B", "Drawn At": "Leftover", "%Draw": "", "Success": "0.38", "3 yr Avg%": "0.28",
                    "OTC ELK": "Y", "%Public": "0.38", "ANY BULL": "N", "Driving (hr)": "3.2", "#Tags": "13",
                    "Public Per Tag": "", "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "",
                    "Sq Mi Public": "", "Bull to Cow": "", }
        EF042O3R = {"List": "B", "Drawn At": "Leftover", "%Draw": "", "Success": "0.3", "3 yr Avg%": "0.24", "OTC ELK": "Y",
                    "%Public": "0.59", "ANY BULL": "N", "Driving (hr)": "2.8", "#Tags": "189", "Public Per Tag": "",
                    "Leftover Day": "447", "3 year leftover day avg": "510", "Valid Units": "", "Sq Mi Public": "",
                    "Bull to Cow": "", }
        EF161O3R = {"List": "B", "Drawn At": "Leftover", "%Draw": "", "Success": "0.15", "3 yr Avg%": "0.22",
                    "OTC ELK": "Y", "%Public": "0.75", "ANY BULL": "N", "Driving (hr)": "3", "#Tags": "260",
                    "Public Per Tag": "", "Leftover Day": "66", "3 year leftover day avg": "108.666666666667",
                    "Valid Units": "", "Sq Mi Public": "", "Bull to Cow": "34", }
        EF231O3R = {"List": "B", "Drawn At": "Leftover", "%Draw": "", "Success": "0.29", "3 yr Avg%": "0.27",
                    "OTC ELK": "Y", "%Public": "0.64", "ANY BULL": "N", "Driving (hr)": "3", "#Tags": "66",
                    "Public Per Tag": "", "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "",
                    "Sq Mi Public": "", "Bull to Cow": "23", }
        EF471O3R = {"List": "B", "Drawn At": "Leftover", "%Draw": "", "Success": "0", "3 yr Avg%": "0.22", "OTC ELK": "Y",
                    "%Public": "0.88", "ANY BULL": "N", "Driving (hr)": "2.9", "#Tags": "3", "Public Per Tag": "",
                    "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "", "Sq Mi Public": "",
                    "Bull to Cow": "24", }
        EF500O3R = {"List": "B", "Drawn At": "Choice 2", "%Draw": "0.07", "Success": "0.43", "3 yr Avg%": "0.27",
                    "OTC ELK": "N", "%Public": "0.73", "ANY BULL": "Y", "Driving (hr)": "1.3", "#Tags": "83",
                    "Public Per Tag": "", "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "",
                    "Sq Mi Public": "", "Bull to Cow": "36", }
        EF007O3R = {"List": "A", "Drawn At": "0", "%Draw": "0.05", "Success": "0.28", "3 yr Avg%": "0.3",
                    "OTC ELK": "N", "%Public": "0.87", "ANY BULL": "N", "Driving (hr)": "2.5", "#Tags": "150",
                    "Public Per Tag": "", "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "",
                    "Sq Mi Public": "", "Bull to Cow": "24", }
        EF020O3R = {"List": "A", "Drawn At": "0", "%Draw": "0.33", "Success": "0.4", "3 yr Avg%": "0.24",
                    "OTC ELK": "N", "%Public": "0.48", "ANY BULL": "Y", "Driving (hr)": "1", "#Tags": "20",
                    "Public Per Tag": "", "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "",
                    "Sq Mi Public": "", "Bull to Cow": "47", }
        EF048O3R = {"List": "A", "Drawn At": "0", "%Draw": "0.16", "Success": "0.65", "3 yr Avg%": "0.46",
                    "OTC ELK": "N", "%Public": "0.89", "ANY BULL": "Y", "Driving (hr)": "2", "#Tags": "17",
                    "Public Per Tag": "", "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "",
                    "Sq Mi Public": "", "Bull to Cow": "26", }
        EF049O3R = {"List": "A", "Drawn At": "1", "%Draw": "0.28", "Success": "0.33", "3 yr Avg%": "0.27",
                    "OTC ELK": "N", "%Public": "0.73", "ANY BULL": "Y", "Driving (hr)": "1.9", "#Tags": "110",
                    "Public Per Tag": "", "Leftover Day": "0", "3 year leftover day avg": "1", "Valid Units": "",
                    "Sq Mi Public": "", "Bull to Cow": "31", }
        EF049S3R = {"List": "A", "Drawn At": "Choice 2", "%Draw": "0.12", "Success": "0.38", "3 yr Avg%": "0.38",
                    "OTC ELK": "N", "%Public": "0.73", "ANY BULL": "Y", "Driving (hr)": "1.9", "#Tags": "50",
                    "Public Per Tag": "", "Leftover Day": "1", "3 year leftover day avg": "0.333333333333333",
                    "Valid Units": "", "Sq Mi Public": "", "Bull to Cow": "31", }
        EF055O3R = {"List": "A", "Drawn At": "1", "%Draw": "0.83", "Success": "0.08", "3 yr Avg%": "0.26",
                    "OTC ELK": "Y", "%Public": "0.89", "ANY BULL": "N", "Driving (hr)": "3", "#Tags": "100",
                    "Public Per Tag": "", "Leftover Day": "4", "3 year leftover day avg": "3.66666666666667",
                    "Valid Units": "", "Sq Mi Public": "", "Bull to Cow": "24", }
        EF056O3R = {"List": "A", "Drawn At": "Choice 4", "%Draw": "0.17", "Success": "0.11", "3 yr Avg%": "0.26",
                    "OTC ELK": "N", "%Public": "0.68", "ANY BULL": "Y", "Driving (hr)": "2.5", "#Tags": "55",
                    "Public Per Tag": "", "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "",
                    "Sq Mi Public": "", "Bull to Cow": "26", }
        EF551O3R = {"List": "A", "Drawn At": "1", "%Draw": "0.94", "Success": "0.6", "3 yr Avg%": "0.41",
                    "OTC ELK": "Y", "%Public": "0.86", "ANY BULL": "N", "Driving (hr)": "3.2", "#Tags": "70",
                    "Public Per Tag": "", "Leftover Day": "0", "3 year leftover day avg": "2.33333333333333",
                    "Valid Units": "", "Sq Mi Public": "", "Bull to Cow": "24", }
        EM020O3R = {"List": "A", "Drawn At": "4", "%Draw": "1", "Success": "0.44", "3 yr Avg%": "0.55", "OTC ELK": "N",
                    "%Public": "0.48", "ANY BULL": "Y", "Driving (hr)": "1", "#Tags": "25", "Public Per Tag": "23.136",
                    "Leftover Day": "0", "3 year leftover day avg": "0", "Valid Units": "20", "Sq Mi Public": "578.4",
                    "Bull to Cow": "47", }
        EM039O3R = {"List": "A", "Drawn At": "Choice 3", "%Draw": "0.03", "Success": "0.28", "3 yr Avg%": "0.26",
                    "OTC ELK": "N", "%Public": "0.65", "ANY BULL": "Y", "Driving (hr)": "0.9", "#Tags": "50",
                    "Public Per Tag": "4.797", "Leftover Day": "3", "3 year leftover day avg": "1.33333333333333",
                    "Valid Units": "39", "Sq Mi Public": "239.85", "Bull to Cow": "33", }
        EM048O3R = {"List": "A", "Drawn At": "1", "%Draw": "0.84", "Success": "0.3", "3 yr Avg%": "0.29",
                    "OTC ELK": "N", "%Public": "0.89", "ANY BULL": "Y", "Driving (hr)": "2", "#Tags": "16",
                    "Public Per Tag": "16.6875", "Leftover Day": "0", "3 year leftover day avg": "0",
                    "Valid Units": "48", "Sq Mi Public": "267", "Bull to Cow": "26", }
        EM049O3R = {"List": "A", "Drawn At": "4", "%Draw": "0.41", "Success": "0.3", "3 yr Avg%": "0.36",
                    "OTC ELK": "N", "%Public": "0.73", "ANY BULL": "Y", "Driving (hr)": "1.9", "#Tags": "34",
                    "Public Per Tag": "11.5511764705882", "Leftover Day": "0", "3 year leftover day avg": "0",
                    "Valid Units": "49", "Sq Mi Public": "392.74", "Bull to Cow": "31", }
        EM050O3R = {"List": "A", "Drawn At": "1", "%Draw": "0.89", "Success": "0.34", "3 yr Avg%": "0.28",
                    "OTC ELK": "N", "%Public": "0.49", "ANY BULL": "Y", "Driving (hr)": "1.7", "#Tags": "64",
                    "Public Per Tag": "3.828125", "Leftover Day": "0", "3 year leftover day avg": "0",
                    "Valid Units": "50", "Sq Mi Public": "245", "Bull to Cow": "36", }
        EM056O3R = {"List": "A", "Drawn At": "0", "%Draw": "0.62", "Success": "0.13", "3 yr Avg%": "0.24",
                    "OTC ELK": "N", "%Public": "0.68", "ANY BULL": "Y", "Driving (hr)": "2.5", "#Tags": "45",
                    "Public Per Tag": "3.64177777777778", "Leftover Day": "1",
                    "3 year leftover day avg": "0.333333333333333", "Valid Units": "56", "Sq Mi Public": "163.88",
                    "Bull to Cow": "26", }
        EM057O3R = {"List": "A", "Drawn At": "2", "%Draw": "0.04", "Success": "0.34", "3 yr Avg%": "0.25",
                    "OTC ELK": "N", "%Public": "0.74", "ANY BULL": "Y", "Driving (hr)": "2.2", "#Tags": "80",
                    "Public Per Tag": "8.50975", "Leftover Day": "0", "3 year leftover day avg": "0.666666666666667",
                    "Valid Units": "58,57", "Sq Mi Public": "680.78", "Bull to Cow": "31", }
        EM069O3R = {"List": "A", "Drawn At": "3", "%Draw": "0.56", "Success": "0.57", "3 yr Avg%": "0.48",
                    "OTC ELK": "N", "%Public": "0.35", "ANY BULL": "Y", "Driving (hr)": "2.3", "#Tags": "40",
                    "Public Per Tag": "14.955", "Leftover Day": "0", "3 year leftover day avg": "0",
                    "Valid Units": "84,69", "Sq Mi Public": "598.2", "Bull to Cow": "39", }
        EM500O3R = {"List": "A", "Drawn At": "0", "%Draw": "0.21", "Success": "0.3", "3 yr Avg%": "0.31",
                    "OTC ELK": "N", "%Public": "0.73", "ANY BULL": "Y", "Driving (hr)": "1.3", "#Tags": "9",
                    "Public Per Tag": "13.14", "Leftover Day": "0", "3 year leftover day avg": "0",
                    "Valid Units": "500", "Sq Mi Public": "118.26", "Bull to Cow": "36", }
        EM501O3R = {"List": "A", "Drawn At": "2", "%Draw": "0.76", "Success": "0.47", "3 yr Avg%": "0.4",
                    "OTC ELK": "N", "%Public": "0.9", "ANY BULL": "Y", "Driving (hr)": "1.5", "#Tags": "25",
                    "Public Per Tag": "18.036", "Leftover Day": "0", "3 year leftover day avg": "0",
                    "Valid Units": "501", "Sq Mi Public": "450.9", "Bull to Cow": "36", }

        # For each page extract text and do the search for tag codes
        for i in range(0, NumPages):
            PageObj = object.getPage(i)
            Text = PageObj.extractText()
            if re.search("Run Date and Time:", Text):
                # Store leftover list run date and time
                # Expected format is Sep 29 2020 1:08:00 PM MDT
                leftover_list_time = re.search(r'.\d{1}:\d{2}:\d{2}', Text)
                if leftover_list_time:
                    CPW_time = leftover_list_time.group()
            # If the CPW_time has changed, it's a new leftover list. Loop through the pages and check for codes
            if CPW_time != old_CPW_time:
                page_number = i + 1
                if page_number == 1:
                    print("\n-------------------------------------------------------------------------")
                print("\nNew leftover list found - reading Page " + str(page_number))
                print("CPW time is " + CPW_time)
                now = datetime.now()
                dt_string = now.strftime("%m/%d/%Y %H:%M:%S")
                print("Actual time is " + dt_string)
                print("Old CPW time is " + old_CPW_time)
                # Loops through the dict, grabbing the next key and value
                for tagCode in tagDict:
                    if re.search(tagCode, Text):
                        # Send email
                        msg = 'List:'+ locals()[tagCode].get('List') + '\n' \
                        +  'Drawn At:'+ locals()[tagCode].get('Drawn At') + '\n'\
                        +  '3 yr %:'+ locals()[tagCode].get('3 yr Avg%') + '\n'\
                        +  'OTC Elk:'+ locals()[tagCode].get('OTC ELK') + '\n'\
                        +  '%Public:' + locals()[tagCode].get('%Public') + '\n'\
                        +  'Any Bull:'+ locals()[tagCode].get('ANY BULL') +  '\n\n' + email_footer
                        subject = tagCode + ' @ ' + CPW_time
                        send_email(testBois, subject, msg)
                        now = datetime.now()
                        dt_string = now.strftime("%H:%M:%S")
                        print("  Email sent for " + subject + ' @ ' + dt_string)
        old_CPW_time = CPW_time
    except:
        print("ERROR: failed to complete")
        time.sleep(10)
        continue
    else:
        time.sleep(0.05)