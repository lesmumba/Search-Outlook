"""
This script uses the win32com module to search Microsoft Outlook folders
for emails which match a set criteria.
Matching emails are saved as HTML documents to an appropriate folder
in the file system for easy retrival labelled as:
    sender_name - message.Subject.html

A HTML page is also genearted which serves as an index file
with links to emails in their respective locations.

Some links worth remembering:
- http://www.boddie.org.uk/python/COM.html
- https://msdn.microsoft.com/en-us/library/office/dn467914.aspx
- https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.aspx
- https://msdn.microsoft.com/en-us/library/office/aa221870%28v=office.11%29.aspx

Known Issues:
- UTF-8 encoding of the message body causes text not to be displayed properly.
"""
import calendar
import datetime
import logging
import os
import re
import shutil
import sys
from win32com import client


BASE_FOLDER = "Extracted Emails"

try:
    shutil.rmtree(BASE_FOLDER)
except:
    print "{0} does not exist".format(BASE_FOLDER)
    pass

try:
    os.makedirs(BASE_FOLDER)
except:
    print "Could not create folder {0}".format(BASE_FOLDER)
    pass


TOTAL_MESSAGES = 0
RESULT_COUNTER = 0

# Open the index.html page in write mode.
HTML_INDEX = open("{0}\index.html".format(BASE_FOLDER), "w")
HTML_INDEX.write("""
<!DOCTYPE html>
<html>
    <head>
        <title>Extracted Emails</title>
    </head>
    <body>
        <h1>All Extracted Emails</h1>
        <ul>
""")
HTML_INDEX.close()

"""
    Sanitize a provided string by encoding it in UTF-8
    and removing reserved characters.
    UTF-8 encoding should limit issues with moving files between OS's.
"""
def sanitize_fields(input_field):
    encoded_field = input_field.encode(encoding="utf-8", errors="ignore")
    sanitized_field = re.sub('[<>"/\|?*!@#$:]', '', encoded_field)
    return sanitized_field

# Add the message we just created to the HTML index.
def add_message_to_index(subject, sender, location):
    HTML_INDEX = open("{0}\index.html".format(BASE_FOLDER), 'a')
    content = '<li><a href="file:.\{location}\{sender} - {subject}.html" \
               target="_blank">{sender} - {subject}</a></li>'.format(
                    location=location, sender=sender, subject=subject)
    HTML_INDEX.write("{0}".format(content))
    return True

# Assume that we'll always want to save the file to
# BASE_FOLDER/year/month/day.
def send_message_to_folder(message):
    try:
        msg_year = message.ReceivedTime.year
        msg_month = message.ReceivedTime.month
        msg_day = message.ReceivedTime.day
        folder_structure = "{0}\{1}\{2}".format(str(msg_year), str(msg_month),
                                                str(msg_day))
        new_dir = "{0}\{1}\{2}\{3}".format(BASE_FOLDER, str(msg_year),
                                           str(msg_month), str(msg_day))
        message_subject = sanitize_fields(message.Subject)
        sender_name = sanitize_fields(message.SenderName)
        location = "{0}\{1} - {2}".format(
            new_dir, sender_name, message_subject)
        if not os.path.exists(new_dir):
            os.makedirs(new_dir)
        try:
            message_file = open("{0}.html".format(location), "w")
            message_body = ""
            try:
                message_body = message.HTMLbody.encode(encoding="utf-8",
                                                       errors="ignore")
            except:
                message_body = message.body.encode(encoding="utf-8",
                                                   errors="ignore")
            # print message_body
            message_file.write(message_body)
            message_file.close()
            if message.Attachments:
                for attachment in message.Attachments:
                    current_directory = os.getcwd()
                    attachment.SaveAsFile("{0}\{1}\{2}".format(
                        current_directory, new_dir, attachment.FileName))
            print ".",
        except Exception, e:
            print "*",
            pass
    except Exception, e:
        print message.SenderName,
        pass


def matching_message(message):
    if ("birthday" in message.SenderEmailAddress.lower() or
        "birthday" in message.Subject.lower() or
            "birthday" in message.Subject.lower()):
        return True
    else:
        return False


# Find the root Outlook folder which holds folders for all email accounts.
outlook = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")
folders = outlook.Folders

for folder in folders:
    # "My Email" is the name of my Outlook email address profile.
    # Substitute with whatever yours is called.
    if folder.name == "My Email":
        print 'Inbox "My Email" found'
        print "---------------------------------------"
        # Here I'm looking in my Inbox.
        # You can change this to any folder you like.
        inbox = folder.Folders('Inbox')
        messages = inbox.Items
        message = messages.GetNext()
        sender_name = ""
        while message:
            TOTAL_MESSAGES += 1
            try:
                sender_email_address = message.SenderEmailAddress
                if matching_message(message):
                    RESULT_COUNTER += 1
                    send_message_to_folder(message)
            except:
                print "!"
                pass
            message = messages.GetNext()

        print "\n%d TOTAL MESSAGES." % TOTAL_MESSAGES
        print "%d MESSAGES MATCHING THE CRITERIA." % RESULT_COUNTER

HTML_INDEX = open("{0}\index.html".format(BASE_FOLDER), 'a')
# Walk through BASE_FOLDER and create the HTML index.
for current_root, sub_dirs, files in os.walk(BASE_FOLDER, topdown=True):
    folder_root, folder_name = os.path.split(current_root)
    if sub_dirs and folder_name != "Emails":
        if "Emails\\" not in folder_root:
            HTML_INDEX.write("<h1>{0}</h1>\n".format(folder_name))
        else:
            HTML_INDEX.write("<h2>{0}</h2>\n".format(
                calendar.month_name[int(folder_name)]))
    if files and folder_name != "Emails":
        HTML_INDEX.write("<h3>{0}</h3>\n<ul>".format(folder_name))
        for file_name in files:
            message_title, message_extention = os.path.splitext(file_name)
            if message_extention == ".html":
                HTML_INDEX.write(
                    '<li><a href="..\{0}\{1}" target="_blank">{2}</a></li>\n'.format(
                        current_root, file_name, message_title))
        HTML_INDEX.write("</ul>\n".format(folder_name))

# Finish up our HTML file by adding closing tags.
HTML_INDEX = open("{0}\index.html".format(BASE_FOLDER), 'a')
HTML_INDEX.write("""
    </ul>
</body>
</html>
""")
print "FINISHED."
