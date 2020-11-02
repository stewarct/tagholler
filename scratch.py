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
    # try:
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
    testBois = ['megeep@gmail.com', 'colinstewrat@gmail.com',
                '7209365042@vtext.com', '3034374578@messaging.sprintpcs.com']
    # Makes a dictionary object. The tag code is the key, while the email list is the value. Use the key to retrieve the value
    tagDict = {"DF015O3R", "DF025O3R", "DF027O3R", "DF035O3R", "DF044O3R", "DF049O3R", "DF058O3R", "DF501O3R", "DF511O3R",
               "DM006O3R", "DM015O3R", "DM025O3R", "DM027O3R", "DM035O3R", "DM044O3R", "DM048O3R", "DM049O3R", "DM050O3R",
               "DM051O3R", "DM059O3R", "DM069O3R", "EF006O3R", "EF015O3R", "EF025O3R", "EF027O3R", "EF050O3R", "EF501O3R",
               "EF131P3R", "DE131P3R", "EF002L1R", "EF131O3R"}
    DF015O3R = {"List": "B", "Drawn": "0", "Success": "0.39", "OTC": "Y", "%Public": "0.68", "Any Bull": "N",
                "Drive Time (HR)": "2.5", "#Tags": "10", "Bull2Cow": "30"}
    DF025O3R = {"List": "B", "Drawn": "Choice 2", "Success": "0.43", "OTC": "Y", "%Public": "0.83", "Any Bull": "N",
                "Drive Time (HR)": "2.4", "#Tags": "70", "Bull2Cow": "23"}
    DF027O3R = {"List": "B", "Drawn": "Leftover", "Success": "0.16", "OTC": "Y", "%Public": "0.53", "Any Bull": "N",
                "Drive Time (HR)": "2", "#Tags": "255", "Bull2Cow": "30"}
    DF035O3R = {"List": "B", "Drawn": "0", "Success": "0.26", "OTC": "Y", "%Public": "0.7", "Any Bull": "N",
                "Drive Time (HR)": "2.1", "#Tags": "10", "Bull2Cow": "25"}
    DF044O3R = {"List": "B", "Drawn": "0", "Success": "0.43", "OTC": "Y", "%Public": "0.78", "Any Bull": "N",
                "Drive Time (HR)": "2.1", "#Tags": "15", "Bull2Cow": "22"}
    DF049O3R = {"List": "A", "Drawn": "5", "Success": "1", "OTC": "N", "%Public": "0.73", "Any Bull": "Y",
                "Drive Time (HR)": "1.9", "#Tags": "10", "Bull2Cow": "31"}
    DF058O3R = {"List": "A", "Drawn": "5", "Success": "0", "OTC": "N", "%Public": "0.47", "Any Bull": "Y",
                "Drive Time (HR)": "2.2", "#Tags": "10", "Bull2Cow": "NA"}
    DF501O3R = {"List": "A", "Drawn": "1", "Success": "0.45", "OTC": "N", "%Public": "0.9", "Any Bull": "Y",
                "Drive Time (HR)": "1.5", "#Tags": "60", "Bull2Cow": "36"}
    DF511O3R = {"List": "A", "Drawn": "2", "Success": "0.62", "OTC": "Y", "%Public": "0.62", "Any Bull": "N",
                "Drive Time (HR)": "1.3", "#Tags": "20", "Bull2Cow": "25"}
    DM006O3R = {"List": "A", "Drawn": "5", "Success": "0.68", "OTC": "Y", "%Public": "0.68", "Any Bull": "N",
                "Drive Time (HR)": "2.5", "#Tags": "25", "Bull2Cow": "34"}
    DM015O3R = {"List": "A", "Drawn": "Choice 3", "Success": "0.3", "OTC": "Y", "%Public": "0.68", "Any Bull": "N",
                "Drive Time (HR)": "2.5", "#Tags": "370", "Bull2Cow": "30"}
    DM025O3R = {"List": "A", "Drawn": "0", "Success": "0.25", "OTC": "Y", "%Public": "0.83", "Any Bull": "N",
                "Drive Time (HR)": "2.4", "#Tags": "285", "Bull2Cow": "23"}
    DM027O3R = {"List": "A", "Drawn": "Choice 2", "Success": "0.25", "OTC": "Y", "%Public": "0.53", "Any Bull": "N",
                "Drive Time (HR)": "2", "#Tags": "445", "Bull2Cow": "30"}
    DM035O3R = {"List": "A", "Drawn": "0", "Success": "0.28", "OTC": "Y", "%Public": "0.7", "Any Bull": "N",
                "Drive Time (HR)": "2.1", "#Tags": "640", "Bull2Cow": "25"}
    DM044O3R = {"List": "A", "Drawn": "14", "Success": "0.53", "OTC": "Y", "%Public": "0.78", "Any Bull": "N",
                "Drive Time (HR)": "2.1", "#Tags": "20", "Bull2Cow": "22"}
    DM048O3R = {"List": "A", "Drawn": "0", "Success": "0", "OTC": "N", "%Public": "0.89", "Any Bull": "Y",
                "Drive Time (HR)": "2", "#Tags": "250", "Bull2Cow": "26"}
    DM049O3R = {"List": "A", "Drawn": "1", "Success": "0.46", "OTC": "N", "%Public": "0.73", "Any Bull": "Y",
                "Drive Time (HR)": "1.9", "#Tags": "600", "Bull2Cow": "31"}
    DM050O3R = {"List": "A", "Drawn": "1", "Success": "0.4", "OTC": "N", "%Public": "0.49", "Any Bull": "Y",
                "Drive Time (HR)": "1.7", "#Tags": "300", "Bull2Cow": "36"}
    DM051O3R = {"List": "A", "Drawn": "1", "Success": "0.52", "OTC": "N", "%Public": "0.45", "Any Bull": "Y",
                "Drive Time (HR)": "1", "#Tags": "90", "Bull2Cow": "15"}
    DM059O3R = {"List": "A", "Drawn": "2", "Success": "0.4", "OTC": "Y", "%Public": "0.35", "Any Bull": "N",
                "Drive Time (HR)": "1.6", "#Tags": "75", "Bull2Cow": "25"}
    DM069O3R = {"List": "A", "Drawn": "0", "Success": "0.41", "OTC": "N", "%Public": "0.35", "Any Bull": "Y",
                "Drive Time (HR)": "2.3", "#Tags": "490", "Bull2Cow": "39"}
    EF006O3R = {"List": "B", "Drawn": "Leftover", "Success": "0.17", "OTC": "Y", "%Public": "0.68", "Any Bull": "N",
                "Drive Time (HR)": "2.5", "#Tags": "134", "Bull2Cow": "34"}
    EF015O3R = {"List": "B", "Drawn": "Leftover", "Success": "0.08", "OTC": "Y", "%Public": "0.68", "Any Bull": "N",
                "Drive Time (HR)": "2.5", "#Tags": "228", "Bull2Cow": "30"}
    EF025O3R = {"List": "B", "Drawn": "0", "Success": "0.24", "OTC": "Y", "%Public": "0.83", "Any Bull": "N",
                "Drive Time (HR)": "2.4", "#Tags": "84", "Bull2Cow": "23"}
    EF027O3R = {"List": "B", "Drawn": "Leftover", "Success": "0.11", "OTC": "Y", "%Public": "0.53", "Any Bull": "N",
                "Drive Time (HR)": "2", "#Tags": "60", "Bull2Cow": "30"}
    EF050O3R = {"List": "B", "Drawn": "0", "Success": "0.17", "OTC": "N", "%Public": "0.49", "Any Bull": "Y",
                "Drive Time (HR)": "1.7", "#Tags": "140", "Bull2Cow": "36"}
    EF501O3R = {"List": "B", "Drawn": "1", "Success": "0.22", "OTC": "N", "%Public": "0.9", "Any Bull": "Y",
                "Drive Time (HR)": "1.5", "#Tags": "52", "Bull2Cow": "36"}
    EF131P3R = {"List": "B", "Drawn": "Leftover", "Success": "NA", "OTC": "Y", "%Public": "0.11", "Any Bull": "N",
                "Drive Time (HR)": "3.1", "#Tags": "6", "Bull2Cow": "23"}
    DE131P3R = {"List": "A", "Drawn": "Leftover", "Success": "NA", "OTC": "Y", "%Public": "0.11", "Any Bull": "N",
                "Drive Time (HR)": "3.1", "#Tags": "25", "Bull2Cow": "23"}
    EF002L1R = {"List": "B", "Drawn": "2", "Success": "0.76", "OTC": "N", "%Public": "0.92", "Any Bull": "Y",
                "Drive Time (HR)": "4.5", "#Tags": "25", "Bull2Cow": "22"}
    EF131O3R = {"List": "B", "Drawn": "Leftover", "Success": "0.24", "OTC": "Y", "%Public": "0.11", "Any Bull": "N",
                "Drive Time (HR)": "3.1", "#Tags": "15", "Bull2Cow": "23"}


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
                    +  'Drawn At:'+ locals()[tagCode].get('Drawn') + '\n' +  'Success Rate:'+ locals()[tagCode].get('Success') +\
                    '\n' +  'OTC Elk:'+ locals()[tagCode].get('OTC') + '\n' +  '%Public:'+\
                    locals()[tagCode].get('%Public') + '\n' +  'Any Bull:'+ locals()[tagCode].get('Any Bull') +  '\n\n' + email_footer
                    subject = tagCode + ' @ ' + CPW_time
                    send_email(testBois, subject, msg)
                    now = datetime.now()
                    dt_string = now.strftime("%H:%M:%S")
                    print("  Email sent for " + subject + ' @ ' + dt_string)
    old_CPW_time = CPW_time
    # except:
    #     print("ERROR: failed to complete")
    #     time.sleep(10)
    #     continue
    # else:
    #     time.sleep(0.05)