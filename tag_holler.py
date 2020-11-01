# ******************************************************************************************************
#  DESCRIPTION:
#         This program is used to parse the CPW leftover list and notify recipients of available tags
#
#  CHANGES:
#         WHO         REV    DATE        DETAIL
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
from bs4 import BeautifulSoup # not used

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
CPW_website = 'https://www.cpwshop.com/licensing.page'
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
        thirstyBois = ['megeep@gmail.com', 'shco6527@gmail.com', 'colinstewrat@gmail.com', 'a.lodolce@gmail.com', 'Wagnercorym@gmail.com']
        goatBois = thirstyBois + ['mlanderson4645@gmail.com', 'cswatk2@gmail.com', 'tyler.fox09@gmail.com']
        elkBois = thirstyBois + ['rileygelatt@gmail.com']
        deerBois = thirstyBois + ['william.dockins@nblenergy.com']
        testBois = ['megeep@gmail.com', 'shco6527@gmail.com', 'a.lodolce@gmail.com', 'colinstewrat@gmail.com','7209365042@vtext.com', '3034374578@messaging.sprintpcs.com']
        # Makes a dictionary object. The tag code is the key, while the email list is the value. Use the key to retrieve the value
        tagDict= {}
        tagDict["AF119O1R"] = goatBois
        tagDict["AM119O1R"] = goatBois
        tagDict["EE069O1A"] = thirstyBois
        tagDict["DF069P3R"] = thirstyBois
        tagDict["AM120O1R"] = thirstyBois
        tagDict["EF069P5R"] = thirstyBois
        tagDict["DF020O3R"] = deerBois
        tagDict["DM020O3R"] = deerBois
        tagDict["EE020O1A"] = elkBois
        tagDict["EF020O1A"] = elkBois
        tagDict["EM020O1A"] = elkBois
        tagDict["EM054O1R"] = elkBois
        tagDict["EF054O1R"] = elkBois
        tagDict["EM020O1R"] = elkBois
        tagDict["EF020O2R"] = elkBois
        tagDict["EM020O2R"] = elkBois
        tagDict["EF020O3R"] = elkBois
        tagDict["EM020O3R"] = elkBois
        tagDict["EF020O4R"] = elkBois
        tagDict["EM020O4R"] = elkBois
        tagDict["EF020L1R"] = elkBois
        tagDict["EM020L1R"] = elkBois
        tagDict["EM020L2R"] = elkBois
        tagDict["DF020O2R"] = elkBois
        tagDict["DF020O4R"] = elkBois
        tagDict["EF020L3R"] = ['megeep@gmail.com', 'colinstewrat@gmail.com', 'dcmaes@gmail.com']
        # Test tag: 
        # tagDict["EF131O2R"] = 
        tagDict["DE131P3R"] = testBois
        tagDict["DM131O3R"] = testBois
        tagDict["DF029O2R"] = testBois
        tagDict["DF029O3R"] = testBois
        tagDict["DF029O4R"] = testBois
        tagDict["EF029O1R"] = testBois
        tagDict["EF029O2R"] = testBois
        tagDict["EF029O3R"] = testBois
        tagDict["EF029O4R"] = testBois
        tagDict["EM029O1R"] = testBois
        tagDict["EM029O2R"] = testBois
        tagDict["EM029O3R"] = testBois
        tagDict["EM029O4R"] = testBois
        tagDict["EF002O2R"] = testBois
        tagDict["EF201O2R"] = testBois
        tagDict["EF040O2R"] = testBois
        tagDict["EF049O2R"] = testBois
        tagDict["EF500O2R"] = testBois
        tagDict["EF056O2R"] = testBois
        tagDict["EF061O2R"] = testBois
        tagDict["EF067O2R"] = testBois
        tagDict["EF001O3R"] = testBois
        tagDict["EF002O3R"] = testBois
        tagDict["EF010O3R"] = testBois
        tagDict["EF040O3R"] = testBois
        tagDict["EF048O3R"] = testBois
        tagDict["EF049O3R"] = testBois
        tagDict["EF500O3R"] = testBois
        tagDict["EF561O3R"] = testBois
        tagDict["EF061O3R"] = testBois
        tagDict["EF001O4R"] = testBois
        tagDict["EF002O4R"] = testBois
        tagDict["EF201O4R"] = testBois
        tagDict["EF010O4R"] = testBois
        tagDict["EF048O4R"] = testBois
        tagDict["EF057O4R"] = testBois
        tagDict["EF058O4R"] = testBois
        tagDict["EF061O4R"] = testBois
        tagDict["EM039O2R"] = testBois
        tagDict["EM046O2R"] = testBois
        tagDict["EM048O2R"] = testBois
        tagDict["EM481O2R"] = testBois
        tagDict["EM049O2R"] = testBois
        tagDict["EM500O2R"] = testBois
        tagDict["EM056O2R"] = testBois
        tagDict["EM057O2R"] = testBois
        tagDict["EM058O2R"] = testBois
        tagDict["EM061O2R"] = testBois
        tagDict["EM066O2R"] = testBois
        tagDict["EM076O2R"] = testBois
        tagDict["EM039O3R"] = testBois
        tagDict["EM048O3R"] = testBois
        tagDict["EM481O3R"] = testBois
        tagDict["EM049O3R"] = testBois
        tagDict["EM050O3R"] = testBois
        tagDict["EM500O3R"] = testBois
        tagDict["EM501O3R"] = testBois
        tagDict["EM051O3R"] = testBois
        tagDict["EM057O3R"] = testBois
        tagDict["EM058O3R"] = testBois
        tagDict["EM061O3R"] = testBois
        tagDict["EM066O3R"] = testBois
        tagDict["EM076O3R"] = testBois
        tagDict["EM039O4R"] = testBois
        tagDict["EM049O4R"] = testBois
        tagDict["EM050O4R"] = testBois
        tagDict["EM500O4R"] = testBois
        tagDict["EM058O4R"] = testBois
        tagDict["EM061O4R"] = testBois
        tagDict["EM501O4R"] = testBois
        tagDict["EE040O2R"] = testBois
        tagDict["EE040O3R"] = testBois
        tagDict["EE040O4R"] = testBois
        tagDict["EF002L1R"] = testBois
        tagDict["EF011L1R"] = testBois
        tagDict["EF002O4R"] = testBois
        tagDict["EF001L1R"] = testBois
        tagDict["EF025O4R"] = testBois
        tagDict["EF005O3R"] = testBois
        tagDict["EF001O3R"] = testBois
        tagDict["EF010O3R"] = testBois
        tagDict["EF002O3R"] = testBois
        tagDict["EF010O4R"] = testBois
        tagDict["EF022L1R"] = testBois
        tagDict["EF010L1R"] = testBois
        tagDict["EF026L1R"] = testBois
        tagDict["EF040O3R"] = testBois
        tagDict["EF022O4R"] = testBois

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
                page_number = i+1
                if page_number == 1:
                    print("\n-------------------------------------------------------------------------")
                print("\nNew leftover list found - reading Page "+str(page_number))
                print("CPW time is "+CPW_time)
                now = datetime.now()
                dt_string = now.strftime("%m/%d/%Y %H:%M:%S")
                print("Actual time is " + dt_string)
                print("Old CPW time is "+old_CPW_time)
                # Loops through the dict, grabbing the next key and value
                for tagCode, emailList in tagDict.items():                
                    if re.search(tagCode, Text):
                        # Send email
                        msg = 'Subject: ' + tagCode + ' is available as of ' + CPW_time + '\n\n' + email_footer
                        subject = tagCode + ' @ ' + CPW_time
                        send_email(emailList, subject, msg)
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
