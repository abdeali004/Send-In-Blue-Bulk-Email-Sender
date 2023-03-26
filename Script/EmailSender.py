import json
import csv
import os
import time
from uuid import uuid4
import sib_api_v3_sdk as sib_api
from sib_api_v3_sdk.rest import ApiException

# Globals
global API_INSTANCE


API_INSTANCE = ""
CONFIG_FILE = "config.json"
SIB_API_KEY = "abc"
HTML_FILE_PATH = ""
EMAILS_CSV_DATA_FILE = ""
EMAILS_STATUS_DATA_FILE = ""
HTML_CONTENT = ""
HTML_TITLE = "Important Notification"

print("Script Started.")

# Functions
def configureAPI():
    global API_INSTANCE, sib_api, SIB_API_KEY
    
    # Setting up API
    configuration = sib_api.Configuration()
    configuration.api_key['api-key'] = SIB_API_KEY
    API_INSTANCE = sib_api.TransactionalEmailsApi(sib_api.ApiClient(configuration))

def sendMail(data):
    global sib_api, uuid4

    # Send Mail
    if not data[1] or not data[3]:
        raise Exception("Senders email or To email is missing.")
    
    sender = {"name":f"{data[0]}","email":f"{data[1]}"}
    to = [{"name":f"{data[2]}","email":f"{data[3]}"}]
    cc = [{"name":f"{data[4]}","email":f"{data[5]}"}] if data[5] else to
    bcc = [{"name":f"{data[6]}","email":f"{data[7]}"}] if data[7] else to

    headers = {"Some-Custom-Name":f"{uuid4()}"}
    # params = {"parameter":"My param value","subject":"New Subject"}
    send_smtp_email = sib_api.SendSmtpEmail(sender=sender, to=to, bcc=bcc, cc=cc, headers=headers, html_content=HTML_CONTENT, subject=HTML_TITLE)
    return API_INSTANCE.send_transac_email(send_smtp_email)


status_headers = ["Emails", "Status", "Remarks"]
status_headers_len = len(status_headers)

try:
    with open(CONFIG_FILE, "r", encoding="utf-8") as file:
        json_data = json.load(file)
        SIB_API_KEY = json_data["api_key"]
        HTML_FILE_PATH = json_data["html_template_path"]
        EMAILS_CSV_DATA_FILE = json_data["emails_file_path"]
        EMAILS_STATUS_DATA_FILE = json_data["email_status_path"]
        # print(HTML_FILE_PATH, EMAILS_CSV_DATA_FILE)

    with open(HTML_FILE_PATH,"r",encoding="utf-8") as file:
        HTML_CONTENT = file.read()

    # Parsing Title
    with open(HTML_FILE_PATH,"r",encoding="utf-8") as file:
        # print(file.readline())
        for line in file.readlines():
            if "title" in line:
                HTML_TITLE = line.split(">")[1].split("<")[0].strip()
    # print(HTML_TITLE)
    # reading csv
    with open(EMAILS_CSV_DATA_FILE, "r", encoding="utf-8") as file:
        csv_data = list(csv.reader(file))
        # print(csv_data)
        csv_data = csv_data[1:]
    
    status_results = []
    for data in csv_data:
        status = ["" for _ in range(status_headers_len)]
        status[0] = data[3]
        try:
            check = all(elem == "" for elem in data)
            if not data and not check:
                continue
            # print(data)
            configureAPI()
            api_response = sendMail(data)
            status[1] = "Success"
            status[2] = f'{api_response}'
        except Exception as e:
            exception = ("Exception when calling SMTPApi->%s\n" % e)
            status[1] = "Failed"
            status[2] = f'{exception}'

        status_results.append(status)
    
    # Appending Results
    EMAILS_STATUS_DATA_FILE = os.path.join(EMAILS_STATUS_DATA_FILE, "Email Status")
    if not os.path.exists(EMAILS_STATUS_DATA_FILE):
        os.makedirs(EMAILS_STATUS_DATA_FILE)

    EMAILS_STATUS_DATA_FILE = os.path.join(EMAILS_STATUS_DATA_FILE, f'{uuid4()}.csv')
    with open(EMAILS_STATUS_DATA_FILE, "w", encoding="utf-8", newline="\n") as file:
        # creating a csv writer object 
        csvwriter = csv.writer(file) 
        # writing the fields 
        csvwriter.writerow(status_headers) 
        # writing the data rows 
        csvwriter.writerows(status_results)


except Exception as ex:
    exception = f'This is outer exception - {ex}'
    status_results.append(["", "Failed", f'{exception}'])

print("Script Completed.")
time.sleep(5)