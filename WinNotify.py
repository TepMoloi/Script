# notification for completed file processing
from win11toast import toast
from pathlib import Path
import os
import contextlib

def run(type, pdfname, file_path=None, latest_file=None): #latest file = original file 
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
                duration="short")
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
                duration="short")
    except Exception as e:
        print(f"Error displaying notification: {e}")
    return


# error notification for no address match
def errorNotification(doc_path):
    try:
        message = f"Error - No Address Match for file {os.path.basename(doc_path)}"
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
                button=button)
    except Exception as e:
        print(f"Error showing notification: {e}")
    return
