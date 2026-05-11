# word doc script for UU dm jobs

# looking through the words found on line 1, 2, 3, and 4.
# the UU address on every record - this is what it should match with from defined line numbers
# those line numbers are chosen as at the top of every UU DM the Unitied Utilities address is always from first line to the 4th. 
# if those words are used anywhere else in the document they will NOT be compared. 
# if there are different words on the line they WILL be compared and the program will NOT continue

import subprocess
from docx import Document
import win32com.client as win32 # this for accurate page counting
import json
import WinNotify
import os
import shutil
import time
import pandas as pd
import sys
import threading
import DataLoad
import uuid


# function to get resource path for json file - non hard coded file paths
def resource_path(filename):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.abspath("."), filename)


# load live paths from json file
with open(resource_path("TestPaths.json")) as json_paths:
    PATH = json.load(json_paths)
Srv_path = format(PATH["Server_path"])
Client_path = format(PATH["Client_path"])
old = format(PATH["old_path"])
errors = format(PATH["Error_path"])
input_path = format(PATH["Input_path"])
script_path = format(PATH["Script_path"])
dashboard_path = format(PATH["Dashboard_path"])

json_filename = "AutoUU_Processed_Files.json"
json_path = os.path.join(dashboard_path, json_filename)

# Limit concurrent Word instances (tune for your machine)
MAX_WORD_INSTANCES = 5
word_open_semaphore = threading.Semaphore(MAX_WORD_INSTANCES)

# -------------------------------------------------------------------------------
# FIX: Removed global start_time, global doc, global WordID_selected entirely.
# These are now local variables created per-call and passed between functions.
# This means each thread has its own isolated copy and cannot overwrite another
# thread's values, which was the root cause of the race condition.
# -------------------------------------------------------------------------------

def main(latest_file=None, file_index=0):
    reference_id = f"REF-{uuid.uuid4().hex[:12].upper()}" # generate a unique reference ID for this file processing run

    # Each thread that calls main() gets its own start_time and Unique Reference ID
    start_time = time.strftime("%H:%M:%S", time.localtime())
    print(f"Start Time - {start_time}")

    file = shutil.move(latest_file, Srv_path)  # move file to Server path

    # docx library word open - address check
    with word_open_semaphore:
        doc_path = file
        Docfile = Document(doc_path)

        record = (
            Docfile.paragraphs[0].text.strip(),
            Docfile.paragraphs[1].text.strip(),
            Docfile.paragraphs[2].text.strip(),
            Docfile.paragraphs[3].text.strip()
        )

        lookfor = ("United Utilities", "PO Box 50", "Warrington", "WA55 1AQ")
        lookfor2 = ("Sample 2", "Address", "Address2", "Post Code")  # test data
        Match = (record == lookfor)

        if not Match:
            print(f"No address match up. \nFile moved to errors, waiting for next file...")
            shutil.move(doc_path, errors)
            DataLoad.ErrorStage(doc_path, dashboard_path, reference_id)
            WinNotify.errorNotification(doc_path)
            return

    DataLoad.JLog(doc_path, dashboard_path, "Address Match Passed")

    # FIX: Pass start_time into word_process so it stays local to this thread
    word_process(doc_path, start_time, file_index, reference_id)


def word_process(doc_path, start_time, file_index=0, reference_id=None):
    # FIX: doc and WordID_selected are now local variables in this function.
    # They are passed forward to production() explicitly instead of via globals.
    # This means two threads can each have their own doc and WordID_selected
    # without either one overwriting the other.
    record_pages = set()
    seen_pages = set()

    print("Data1")
    DataLoad.Queue(doc_path, dashboard_path, reference_id)
    print("Data2")

    doc = None
    word_id_selected = None  # FIX: renamed and made local
    word = None

    try:
        word = win32.DispatchEx("Word.Application")

        # FIX: Get the PID immediately after opening THIS specific Word instance
        # before any other thread can open another one, by querying right away.
        # We also now store it as a local variable (word_id_selected) not a global.
        pid = subprocess.Popen(
            ["powershell", "-Command", "Get-Process -Name WINWORD | Select-Object -ExpandProperty Id"],
            stdout=subprocess.PIPE
        ).communicate()[0]
        WordID = pid.strip().decode("utf-8")
        x = WordID.split()
        if file_index < len(x):
            word_id_selected = x[file_index]
        else:
            word_id_selected = x[0] if x else None
        print(f"Word Process ID for file index {file_index}: {word_id_selected}")


# ------ Cache error delete cache and retry word open ---- #
    except Exception as e:
        print(f"Removing cache from %TEMP%\\gen_py \n{e}")
        cachepATH = os.path.join(os.environ["TEMP"], "gen_py")
        shutil.rmtree(cachepATH)
        print("Cache removed, trying again...")
        word = win32.DispatchEx("Word.Application")

# wrap process in pipeline try except to catch any craches and log them to dashbaord #
    try:
        print("Data3")
        with word_open_semaphore:
            print("Opening Word Document")
            word.Visible = False
            # FIX: doc is now a local variable assigned here and passed forward
            doc = word.Documents.Open(doc_path)
            print("Open")
            DataLoad.JLog(doc_path, dashboard_path, "Logic running")
            DataLoad.ProcessingStage(doc_path, dashboard_path, reference_id)
            page_count = doc.ComputeStatistics(2)

            for i in range(1, doc.Paragraphs.Count + 1):
                para = doc.Paragraphs(i)
                page_num = para.Range.Information(3)
                if page_num not in seen_pages:
                    word.Selection.GoTo(What=1, Which=2)
                    seen_pages.add(page_num)
                    if para.Range.Text.strip() == "United Utilities":
                        record_pages.add(page_num)
                        word.Selection.GoTo(What=1, Which=2)
                        print(len(record_pages))
                        if len(record_pages) == 3:
                            break

        DataLoad.JLog(doc_path, dashboard_path, "Logic Check Passed")
        print("Data4")
        gap(record_pages, doc_path, word, doc, word_id_selected, start_time, page_count, reference_id)

    except Exception as e:
        # Any unexpected crash anywhere in the pipeline lands here
        print(f"Pipeline error for {os.path.basename(doc_path)}: {e}")
        DataLoad.JLog(doc_path, dashboard_path, f"Pipeline Error - {e}")
        # if e == f"- Destination path {doc_path} already exists":
        #         print("File with same name already exists in destination. Moving to errors.")
        # print(doc_path)
        shutil.move(doc_path, errors)
        DataLoad.ErrorStage(doc_path, dashboard_path, reference_id)
        WinNotify.errorNotification(doc_path)

        # Clean up Word if it was opened before the crash
        try:
            if doc is not None:
                doc.Close(SaveChanges=0)
            if word is not None:
                word.Quit()
            if word_id_selected is not None:
                os.system(f"taskkill /pid {word_id_selected} /f")
        except Exception as cleanup_error:
            print(f"Cleanup error: {cleanup_error}")


def gap(record_pages, latest_file, word, doc, word_id_selected, start_time, page_count, reference_id=None):
    """Calculates the gap between record page numbers to determine Simplex/Duplex"""

    print(f"{page_count} pages in the document")

    if len(record_pages) == 1:
        print("1 Record Found")
        gapnum = 1
        full_record_amount = len(record_pages)
        output(gapnum, latest_file, word, doc, word_id_selected, start_time, full_record_amount, page_count, reference_id)
        return  

    elif len(record_pages) == 2:
        print("2 Records Found")
        sorted_numbers = sorted(record_pages)
        gapnum = sorted_numbers[1] - sorted_numbers[0]
        if gapnum == 0:
            gapnum = 1
        full_record_amount = 2
        output(gapnum, latest_file, word, doc, word_id_selected, start_time, full_record_amount, page_count, reference_id)
        return  

    print(f"First 3 records page numbers: {sorted(record_pages)}")

    # For 3 or more records
    sorted_numbers = sorted(record_pages)
    gapz = [sorted_numbers[i+1] - sorted_numbers[i] for i in range(len(sorted_numbers) - 1)]
    gapnum = int(gapz[0])
    opr_amount = page_count / gapnum
    full_record_amount = int(opr_amount)
    print(f"Records in File: {full_record_amount}")
    output(gapnum, latest_file, word, doc, word_id_selected, start_time, full_record_amount, page_count, reference_id)


def output(gapnum, latest_file, word, doc, word_id_selected, start_time, full_record_amount, page_count, reference_id=None):

    if page_count > gapnum:
        print("More")
    if gapnum < 2:
        print("Less")

    print(f"Gap between numbers: {gapnum}")

    if gapnum == 1 and page_count > 1 and page_count < 5:
        print("Duplex 1 record but more than 1 page found")
        match page_count:
            case 1:
                type = "Simplex~#1"
            case 2:
                type = "Duplex~#2"
            case 3:
                type = "Duplex~#3"
            case 4:
                type = "Duplex~#4"
            case _:
                print("Error Incorrect format")
                return
        print(type)
        production(latest_file, word, doc, word_id_selected, start_time, full_record_amount, type, page_count, reference_id)
        return

    match gapnum:
        case 1:
            type = "Simplex~#1"
        case 2:
            type = "Duplex~#2"
        case 3:
            type = "Duplex~#3"
        case 4:
            type = "Duplex~#4"
        case _:
            print("Error Incorrect format")
            return
    print(type)
    DataLoad.JLog(latest_file, dashboard_path, "Type Identified")
    production(latest_file, word, doc, word_id_selected, start_time, full_record_amount, type, page_count, reference_id)
    return


def production(latest_file, word, doc, word_id_selected, start_time, full_record_amount, type, page_count, reference_id=None):

    processing_json_path = os.path.join(dashboard_path, "ProccessingQueue.json")
    timestamp = time.strftime("%H:%M:%S", time.localtime())
    processing_frame = pd.DataFrame(
        DataLoad.datadict(os.path.basename(latest_file), "PDF Output", timestamp, reference_id)
    )
    try:
        DataLoad.push2(processing_json_path, processing_frame)
        DataLoad.JLog(latest_file, dashboard_path, "PDF Output running")
    except Exception as e:
        DataLoad.JLog(latest_file, dashboard_path, f"Error during PDF output stage: {e}")
        shutil.move(latest_file, errors)

        DataLoad.ErrorStage(latest_file, dashboard_path, reference_id)
        print(f"Error during PDF output processing: {e}")

    pdf_path = latest_file.replace('.docx', '')
    basename = os.path.basename(pdf_path)
    pdfname = (f"{basename}~#{full_record_amount}~#{type}")
    print(f"PDF Filename - {pdfname}")
    pathtosave = (f"{Srv_path}\\{pdfname}.pdf")

    try:
        doc.ExportAsFixedFormat(pathtosave, 17)
    except Exception as e:
        print(f"Error saving document {e}")

    try:
        doc.Close(SaveChanges=0)
        word.Quit()
        print("Closed Word Document 1")
    except Exception as e:
        print(f"Error closing Word application: {e}")

    try:
        os.system(f"taskkill /pid {word_id_selected} /f")  # FIX: uses local word_id_selected
    except Exception as e:
        print(f"Error killing Word process: {e}")
    print("Killed Word Process")

    time.sleep(5)

    linkpath = shutil.move(pathtosave, input_path)
    shutil.move(latest_file, old)

    End_time = time.strftime("%H:%M:%S", time.localtime())
    print(f"End Time - {End_time}")

    status = "Completed"

    dataframe = {
        "Filename": [pdfname],
        "Reference": [reference_id or ""],
        "Start_Time": [start_time],   # FIX: uses local start_time passed from main()
        "End_Time": [End_time],
        "Page_Amount": [page_count],
        "Record_Amount": [full_record_amount],
        "Type": [type],
        "Date": [time.strftime("%Y-%m-%d", time.localtime())],
        "Status": [status]
    }
    df = pd.DataFrame(dataframe)

    try:
        DataLoad.push2(processing_json_path, df)
        DataLoad.JobDisplay(df, json_path)
        WinNotify.run(type, pdfname, file_path=linkpath, latest_file=latest_file)
        DataLoad.JLog(latest_file, dashboard_path, "Completed")
    except Exception as e:
        DataLoad.JLog(latest_file, dashboard_path, f"Error during final processing steps: {e}")
        shutil.move(latest_file, errors)
        DataLoad.ErrorStage(latest_file, dashboard_path, reference_id)
        print(f"Error during final processing steps: {e}")

    return




# program starts in pointer
if __name__ == "__main__":
    main()