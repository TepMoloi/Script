# word doc script for UU dm jobs

# looking through the words found on line 1, 2, 3, and 4.
# the UU address on every record - this is what it should match with from defined line numbers
# those line numbers are chosen as at the top of every UU DM the Unitied Utilities address is always from first line to the 4th. 
# if those words are used anywhere else in the document they will NOT be compared. 
# if there are different words on the line they WILL be compared and the program will NOT continue

from pydoc import doc
import subprocess
from docx import Document
import win32com.client as win32 # this for accurate page counting
import json
import Pointer
import WinNotify
import os
import shutil
import time
import pandas as pd
import sys
import threading
import DataLoad
import uuid
from docx.oxml.ns import qn

# function to get resource path for json file - non hard coded file paths
def resource_path(filename):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.abspath("."), filename)


# load live paths from json file
with open(resource_path("LivePaths.json")) as json_paths:
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
MAX_WORD_INSTANCES = 10
word_open_semaphore = threading.Semaphore(MAX_WORD_INSTANCES)

errdfiles = set() # track files that errored
pcsed = set() # track files that processed successfully

# -------------------------------------------------------------------------------
# each thread has its own isolated copy and cannot overwrite another
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
            print(f"No address match up. File moved to errors")
            shutil.move(doc_path, errors)
            DataLoad.ErrorStage(doc_path, dashboard_path, reference_id)
            WinNotify.errorNotification(doc_path, "No address match - Check formatting")
            errdfiles.add(doc_path)
            return
        
        
    # print("Address Match Passed") - reducing console clutter
    DataLoad.JLog(doc_path, dashboard_path, "Address Match Passed")
    
    word_process(doc_path, start_time, file_index, reference_id)
    # return

def word_process(doc_path, start_time, file_index=0, reference_id=None):

    record_pages = set()
    seen_pages = set()

    # print("Data1") - reducing console clutter
    DataLoad.Queue(doc_path, dashboard_path, reference_id)
    # print("Data2") - reducing console clutter

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
        # print(f"Word Process ID for file index {file_index}: {word_id_selected}") - reducing console clutter

# ------ Cache error delete cache and retry word open ---- #
    except Exception as e:
        print(f"Removing cache from %TEMP%\\gen_py \n{e}")
        cachepATH = os.path.join(os.environ["TEMP"], "gen_py")
        shutil.rmtree(cachepATH)
        print("Cache removed, trying again...")
        word = win32.DispatchEx("Word.Application")

# wrap process in pipeline try except to catch any craches and log them to dashbaord #
    try:
        # print("Data3") - reducing console clutter
        with word_open_semaphore:
            print("Opening Word Document")
            word.Visible = False
            # FIX: doc is now a local variable assigned here and passed forward
            doc = word.Documents.Open(doc_path)
            # print("Open") - reducing console clutter
            DataLoad.JLog(doc_path, dashboard_path, "Logic running")
            DataLoad.ProcessingStage(doc_path, dashboard_path, reference_id)
            page_count = doc.ComputeStatistics(2)

            print("secStarting")   

            if len(doc.sections) >= 1:
                para = doc.Paragraphs(1)  # Check the first paragraph for the address
                for i in range(1, doc.Sections.Count + 1):
                    section = doc.Sections(i)
                    if para.Range.Text.strip() == "United Utilities": # for test data use - "Sample 2"
                        page_number = section.Range.Information(3)
                        record_pages.add(page_number)
                        print(f"Section break page number: {page_number}")
                        print(record_pages)
                        if len(record_pages) == 3:
                            break
            print("secComplete")   

    except Exception as e:
        # Any unexpected crash anywhere in the pipeline lands here
        print(f"Pipeline error for {os.path.basename(doc_path)}: {e}")
        DataLoad.JLog(doc_path, dashboard_path, f"Pipeline Error - {e}")
        
        DataLoad.ErrorStage(doc_path, dashboard_path, reference_id)
        errdfiles.add(doc_path)
        WinNotify.errorNotification(doc_path, "Pipeline Error - Check formatting")

        # Clean up Word if it was opened before the crash
        try:
            if doc is not None:
                doc.Close(SaveChanges=0)
            if word is not None:
                word.Quit()
                word.Quit()  # Ensure full cleanup of Word instance
            shutil.move(doc_path, errors)
        except Exception as cleanup_error:
            print(f"Cleanup error: {cleanup_error}")


    DataLoad.JLog(doc_path, dashboard_path, "Logic Check Passed")
    print("Data4")
    gap(record_pages, doc_path, word, doc, word_id_selected, start_time, page_count, reference_id)

    # return


def gap(record_pages, latest_file, word, doc, word_id_selected, start_time, page_count, reference_id=None):
    """Calculates the gap between record page numbers to determine Simplex/Duplex"""

    print(f"{page_count} pages in the document")
    print(len(record_pages))

    if len(record_pages) == 1:
        print("1 Record Found")
        gapnum = 1
        full_record_amount = len(record_pages)
        output(gapnum, latest_file, word, doc, word_id_selected, start_time, full_record_amount, page_count, reference_id)
        return
        
    if len(record_pages) == 2:
        print("2 Records Found")
        sorted_numbers = sorted(record_pages)
        gapnum = sorted_numbers[1] - sorted_numbers[0]
        if gapnum == 0:
            gapnum = 1
        full_record_amount = 2
        output(gapnum, latest_file, word, doc, word_id_selected, start_time, full_record_amount, page_count, reference_id)
        return
        

    
    print(f"First 3 record page numbers: {sorted(record_pages)}")

    # For 3 or more records
    sorted_numbers = sorted(record_pages)
    gapz = [sorted_numbers[i+1] - sorted_numbers[i] for i in range(len(sorted_numbers) - 1)]
    print(gapz)
    if gapz is None or len(gapz) == 0:
        gapz = [1]  # Default to 1 if no gaps found, to prevent errors in output function
    else:
        gapnum = int(gapz[0])
    print(gapz)
    opr_amount = page_count / gapnum
    full_record_amount = int(opr_amount)
    print(f"Records in File: {full_record_amount}")
    output(gapnum, latest_file, word, doc, word_id_selected, start_time, full_record_amount, page_count, reference_id)
    # return


def output(gapnum, latest_file, word, doc, word_id_selected, start_time, full_record_amount, page_count, reference_id=None):

    # failsafe if the records are not divisible by the page count
    if page_count % gapnum != 0:
        print("Page count is not divisible by gap between records. Check document formatting.")
        DataLoad.JLog(latest_file, dashboard_path, "Page count not divisible by gap - [Check formatting]")
        print("error notification")
        WinNotify.errorNotification(latest_file, "Page count not divisible by gap - [Check formatting]")       
        doc.Close()
        word.Quit()    # KEEP BOTH IN
        word.Quit() 
        # THE FIRST QUIT DOES NOT CLOSE IT, IT JUST ENDS THE WORD INSTANCE, 
        # THE SECOND QUIT ENSURES IT FULLY CLOSES AND RELEASES THE FILE LOCK
        shutil.move(latest_file, errors)
        errdfiles.add(latest_file)
        DataLoad.ErrorStage(latest_file, dashboard_path, reference_id)
        return

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
    production(latest_file, word, doc, start_time, full_record_amount, type, page_count, reference_id)
    # return


def production(latest_file, word, doc, start_time, full_record_amount, type, page_count, reference_id=None):

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
        errdfiles.add(latest_file)
        DataLoad.ErrorStage(latest_file, dashboard_path, reference_id)
        print(f"Error during PDF output processing: {e}")

    pdf_path = latest_file.replace('.docx', '')
    basename = os.path.basename(pdf_path)
    pdfname = (f"{basename}~#{full_record_amount}~#{type}")
    print(f"PDF Filename - {pdfname}")
    pathtosave = (f"{Srv_path}\\{pdfname}.pdf")
    existing_file = (f"{input_path}\\{pdfname}.pdf")

    try:
        doc.ExportAsFixedFormat(pathtosave, 17)
    except Exception as e:
        print(f"Error saving document {e}")

    try:
        doc.Close(SaveChanges=0)
        word.Quit()
        print("Closed Word Document")
    except Exception as e:
        print(f"Error closing Word application: {e}")


    print("moving")
    if os.path.exists(existing_file):
        print(f"File {pdfname}.pdf already exists in input")
    else:
        try:
            linkpath = shutil.move(pathtosave, input_path)
        except Exception as e:
            print(f"Error moving file: {e}")

    old_file = os.path.join(old, os.path.basename(latest_file))
    if os.path.exists(old_file):
        print(f"File {os.path.basename(latest_file)} already exists in old")
    else:
        try:     
            shutil.move(latest_file, old)
        except Exception as e:
            print(f"Error moving file to old: {e}")
                
    

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

    print("finalizing")

    try:
        DataLoad.push2(processing_json_path, df)
        print("push2")
        DataLoad.JobDisplay(df, json_path)
        print("Job Display")
        WinNotify.run(type, pdfname, file_path=linkpath, latest_file=latest_file)
        print("notification")
        DataLoad.JLog(latest_file, dashboard_path, "Completed")
        print("JLog")
    except Exception as e:
        DataLoad.JLog(latest_file, dashboard_path, f"Error during final processing steps: {e}")
        shutil.move(latest_file, errors)
        errdfiles.add(latest_file)
        DataLoad.ErrorStage(latest_file, dashboard_path, reference_id)
        print(f"Error during final processing steps: {e}")

    # pcsed.add(latest_file)
    # return
    print("hellossfworld")
