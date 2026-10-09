# notification for completed file processing
from win11toast import toast
from pathlib import Path
import os
import contextlib
from dotenv import load_dotenv
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import json
import sys

def resource_path(filename):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.abspath("."), filename)

with open(resource_path("LivePaths.json")) as json_paths:
    PATH = json.load(json_paths)
errors = format(PATH["Error_path"])

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(dotenv_path=os.path.join(BASE_DIR, "Environment.env"))

SMTP_SERVER = "10.35.80.40"
SMTP_PORT = 25
SENDER = "tmoloi@mba-group.com"

cnt = 0
todaysCount = 0


def run(type, pdfname, file_path=None, latest_file=None): #latest file = original file 
    return
    try:
        message = f"Processed: {pdfname} \n{type}"
        if latest_file:
            message += f"\nOriginal: {os.path.basename(latest_file)}"
        button = None
        if file_path:
            file_uri = Path(file_path).resolve().as_uri()
            button = {
                "activationType": "protocol",
                "arguments": file_uri,
                "content": "Open PDF File",
            }
        with contextlib.redirect_stdout(open(os.devnull, "w")), contextlib.redirect_stderr(open(os.devnull, "w")):  # suppress console output for notifications
            toast("Task Completed",
                message,
                duration="short",
                icon="C:\\Users\\tmoloi\\Desktop\\Python\\Client\\Script\\dist\\649467.ico",
                button=button)
    except Exception as e:
        print(f"Error showing notification: {e}")
    return

# file received notification

def startNotification():
    try:
        message = f"Processing Files"
        with contextlib.redirect_stdout(open(os.devnull, "w")), contextlib.redirect_stderr(open(os.devnull, "w")):  # suppress console output for notifications
            toast("Auto UU Files Recieved",
                message,
                duration="short",
                icon= "C:\\Users\\tmoloi\\Desktop\\Python\\Client\\Script\\dist\\whitet3p2334.ico")
    except Exception as e:
        print(f"Error displaying notification: {e}")
    return

# observer running notification

def observerRunning():
    try:
        message = f"Observer Running"
        with contextlib.redirect_stdout(open(os.devnull, "w")), contextlib.redirect_stderr(open(os.devnull, "w")): # suppress console output for notifications
            toast("Auto UU Observer",
                message,
                duration="short",
                icon= "C:\\Users\\tmoloi\\Desktop\\Python\\Client\\Script\\dist\\whitet3p2334.ico")
    except Exception as e:
        print(f"Error displaying notification: {e}")
    return


# error notification for no address match
def errorNotification(doc_path, message1):
    try:
        message = f"{os.path.basename(doc_path)} Error  ||  File Attached \n\nError: {message1} \nFile moved here - {errors} \n\n"
        with contextlib.redirect_stdout(open(os.devnull, "w")), contextlib.redirect_stderr(open(os.devnull, "w")): # suppress console output for notifications
            toast("Auto UU File Error",
                message,
                duration="short",
                # button=button,
                icon= "C:\\Users\\tmoloi\\Desktop\\Python\\Client\\Script\\dist\\sign-red-error-icon-1.ico")
    except Exception as e:
        print(f"Error showing notification: {e}")
        
    
    # doc_path is normally already in the errors folder - fall back to errors\<name> for older callers
    docfile = doc_path if os.path.exists(doc_path) else os.path.join(errors, os.path.basename(doc_path))
    file = None
    try:
        with open(docfile, 'rb') as f:
            file = MIMEApplication(f.read(),
            name=os.path.basename(docfile))
        file.add_header('Content-Disposition', 'attachment', filename=os.path.basename(docfile))
    except OSError as e:
        print(f"Could not attach {docfile}: {e}")
        message += "Original file could not be attached"

    send_email(file, recipient="tmoloi@mba-group.com", subject="Auto UU File Error", body=message)
    return

def CompletedNotificaiton(counter, pcsed_count, errdfiles_count, file_path, proofs_path, processed, wsheet):
    # email and notification for completed file processing
    
    global cnt, todaysCount
    cnt += 1
    todaysCount += counter

    #get filenames from file_path list
    files = []
    if isinstance(file_path, list):
        for f in file_path:
            files.append(os.path.basename(f))
    else:
        files.append(os.path.basename(file_path))
    files_str = "\n".join(files)

    file = None    
    if wsheet is not None:
        wsheet = wsheet[0]
        with open(wsheet, 'rb') as f:
            file = MIMEApplication(f.read(),
            name=os.path.basename(wsheet))
        file.add_header('Content-Disposition', 'attachment', filename=os.path.basename(wsheet))

    # notification         
    message = f"UUDM Files  ||  {pcsed_count} Passed and {errdfiles_count} Failed\n\nFiles: \n{files_str} \n\n{proofs_path}\n\n"
    if file is None:
        message = message + "No wavesheet file created"
    
    try:
        with contextlib.redirect_stdout(open(os.devnull, "w")), contextlib.redirect_stderr(open(os.devnull, "w")): # suppress console output for notifications
            toast("Auto UU Files",
                message,
                duration="short",
                icon= "C:\\Users\\tmoloi\\Desktop\\Python\\Client\\Script\\dist\\whitet3p2334.ico")
    except Exception as e:
            print(f"Error showing notification: {e}")

    #grab wavesheet
    # if all files errored there would be no wavesheet file created. to prevent error passing "file" set it as None.
    # if a wavesheet file exists "file" gets set as the created wavesheet file to pass through
   

    send_email(file, recipient="tmoloi@mba-group.com", subject="Auto UU Files Processed", body=message)
    return



def send_email(file, recipient: str, subject: str, body: str) -> bool:
    """Send a plain-text notification email via the internal relay.

    Returns True on success, False on failure (logs the error instead of crashing).
    """
    message = MIMEMultipart()
    message["From"] = SENDER
    message["To"] = recipient
    message["Subject"] = subject
    message.attach(MIMEText(body, "plain"))
    if file is not None:
        message.attach(file)

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=10) as server:
            server.sendmail(SENDER, recipient, message.as_string())
            print("Email sent successfully.")
        return True

    except Exception as e:
        print(f"[sendemail] Failed to send email: {e}")
        return False
