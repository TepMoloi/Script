# notification for completed file processing
from win11toast import toast
from pathlib import Path
import os
import contextlib
import win32com.client
import time

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
    return
    try:
        message = f"Error: {message1}"
        button = None
        file_uri = Path(doc_path).resolve().as_uri()
        button = {
            "activationType": "protocol",
            "arguments": file_uri,
            "content": "Open File",
        }
        with contextlib.redirect_stdout(open(os.devnull, "w")), contextlib.redirect_stderr(open(os.devnull, "w")): # suppress console output for notifications
            toast("Auto UU File Error",
                message,
                duration="short",
                button=button,
                icon= "C:\\Users\\tmoloi\\Desktop\\Python\\Client\\Script\\dist\\sign-red-error-icon-1.ico")
    except Exception as e:
        print(f"Error showing notification: {e}")
    return

def CompletedNotificaiton(counter, pcsed_count, errdfiles_count):

    global cnt, todaysCount
    cnt += 1
    todaysCount += counter

    try:
        message = f"{todaysCount} file(s) have been processed today \n {counter} just processed \n{pcsed_count} Passed \n {errdfiles_count} Failed"
        with contextlib.redirect_stdout(open(os.devnull, "w")), contextlib.redirect_stderr(open(os.devnull, "w")): # suppress console output for notifications
            toast("Auto UU Files Processed",
                message,
                duration="short",
                icon= "C:\\Users\\tmoloi\\Desktop\\Python\\Client\\Script\\dist\\whitet3p2334.ico")
    except Exception as e:
            print(f"Error showing notification: {e}")
    return

