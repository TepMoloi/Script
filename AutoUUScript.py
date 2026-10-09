# word doc script for UU dm jobs

# looking through the words found on line 1, 2, 3, and 4.
# the UU address on every record - this is what it should match with from defined line numbers
# those line numbers are chosen as at the top of every UU DM the Unitied Utilities address is always from first line to the 4th. 
# if those words are used anywhere else in the document they will NOT be compared. 
# if there are different words on the line they WILL be compared and the program will NOT continue

from docx import Document
import win32com.client as win32
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
import glob 
import pdfplumber
import re

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
output_path = format(PATH["Output_path"])
proofing_path = format(PATH["Proofs_path"])
mrdf_path = format(PATH["MRDF_path"])
wavesheet_path = format(PATH["Wavesheet_path"])

json_filename = "AutoUU_Processed_Files.json"
json_path = os.path.join(dashboard_path, json_filename)

# Limit concurrent Word instances (tune for your machine)
MAX_WORD_INSTANCES = 10
word_open_semaphore = threading.Semaphore(MAX_WORD_INSTANCES)

pcsed = set() # track files that processed successfully
errdfiles = set() # track files that errored
active_word_instances = []
active_word_lock = threading.Lock()
# each thread has its own isolated copy and cannot overwrite another
copy_lock = threading.Lock()


def safe_move(src, dest_dir, retries=10, delay=1):
    """Moves src into dest_dir without overwriting an existing file.
    Retries while Word releases its file lock. Returns the new path."""
    name, ext = os.path.splitext(os.path.basename(src))
    dest = os.path.join(dest_dir, name + ext)
    if os.path.exists(dest):
        dest = os.path.join(dest_dir, f"{name}_{time.strftime('%Y%m%d%H%M%S')}_{uuid.uuid4().hex[:4]}{ext}")
    for attempt in range(retries):
        try:
            return shutil.move(src, dest)
        except PermissionError:
            if attempt == retries - 1:
                raise
            time.sleep(delay)


def close_word(doc, word):
    """Closes the document and quits Word - safe to call if either is already closed"""
    try:
        if doc is not None:
            doc.Close(SaveChanges=0)
    except Exception as e:
        print(f"Error closing document: {e}")
    try:
        if word is not None:
            word.Quit(SaveChanges=0)
    except Exception as e:
        print(f"Error quitting Word: {e}")


def fail_file(path, message, reference_id=None, doc=None, word=None, move_path=None):
    """Error route for a file: close Word -> move to errors -> log -> notify (notify needs the file already in errors).
    move_path is used when the file has already moved on from path (e.g. into old)."""
    close_word(doc, word)
    moved = move_path or path
    try:
        moved = safe_move(moved, errors)
    except Exception as e:
        print(f"Could not move {moved} to errors: {e}")
    DataLoad.JLog(path, dashboard_path, message)
    DataLoad.ErrorStage(path, dashboard_path, reference_id)
    errdfiles.add(path)
    WinNotify.errorNotification(moved, message)


def main(latest_file, proofs_path):
    reference_id = f"REF-{uuid.uuid4().hex[:12].upper()}" # generate a unique reference ID for this file processing run
    # Each thread that calls main() gets its own start_time and Unique Reference ID
    
    start_time = time.strftime("%H:%M:%S", time.localtime())
    print(f"Start Time - {start_time}")
    # a file with the same name already in Precomposed would make the move fail (or clash with a running job)
    if os.path.exists(os.path.join(Srv_path, os.path.basename(latest_file))):
        print(f"Duplicate filename in Precomposed. File moved to errors")
        fail_file(latest_file, "Duplicate filename - a file with this name is already in Precomposed", reference_id)
        return
    file = shutil.move(latest_file, Srv_path)  # move file to Server path

    # docx library word open - address check
    with word_open_semaphore:
        doc_path = file
        try:
            Docfile = Document(doc_path)
        except Exception as e:
            print(f"Could not open {doc_path}: {e}")
            fail_file(doc_path, "Could not open document - file may be corrupt", reference_id)
            return
        # slice instead of indexing so a document with fewer than 4 paragraphs is a mismatch, not a crash
        record = tuple(p.text.strip() for p in Docfile.paragraphs[:4])
        lookfor = ("United Utilities", "PO Box 50", "Warrington", "WA55 1AQ")
        lookfor2 = ("Sample 2", "Address", "Address2", "Post Code")  # test data
        Match = (record == lookfor)

        if not Match:
            print(f"No address match up. File moved to errors")
            fail_file(doc_path, "No address match - Check formatting", reference_id)
            return
        
    Docfile = None  # Ensure the Document object is released before proceedin
    DataLoad.JLog(doc_path, dashboard_path, "Address Match Passed")
    word_process(
        doc_path,
        start_time,
        proofs_path=proofs_path,
        reference_id=reference_id,
    )

def word_process(doc_path, start_time, proofs_path, reference_id=None):

    record_pages = set()
    DataLoad.Queue(doc_path, dashboard_path, reference_id)
    doc = None
    word_id_selected = None  # FIX: renamed and made local
    word = None
    
# wrap process in pipeline try except to catch any craches and log them to dashbaord #
    try:
        with word_open_semaphore:
            # print("Opening Word Document") muting cluster
            word = win32.DispatchEx("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(doc_path)
            DataLoad.JLog(doc_path, dashboard_path, "Logic running")
            DataLoad.ProcessingStage(doc_path, dashboard_path, reference_id)
            page_count = doc.ComputeStatistics(2)


            if len(doc.sections) >= 1:
                para = doc.Paragraphs(1)  # Check the first paragraph for the address
                for i in range(1, doc.Sections.Count + 1):
                    section = doc.Sections(i)
                    if para.Range.Text.strip() == "United Utilities": # for test data use - "Sample 2"
                        page_number = section.Range.Information(3)
                        record_pages.add(page_number)
                        # print(f"Section break page number: {page_number}")  muting cluster
                        print(record_pages)
                        if len(record_pages) == 3:
                            break

# ----------------------------------------------------------------------------------------
    # First error to come back is a typeError in the production module.
    # The production is sending 9 variables when it only accepts 8. I have added a reference_id variable to the production module to fix this issue.
# ----------------------------------------------------------------------------------------

        DataLoad.JLog(doc_path, dashboard_path, "Logic Check Passed")
        # gap -> output -> production now run inside the try so a crash there still closes Word and routes the file
        gap(
            record_pages,
            doc_path,
            word,
            doc,
            word_id_selected,
            start_time,
            page_count,
            proofs_path=proofs_path,
            reference_id=reference_id,
        )

    except Exception as e:
        # Any unexpected crash anywhere in the pipeline lands here
        print(f"Pipeline error for {os.path.basename(doc_path)}: {e}")
        DataLoad.JLog(doc_path, dashboard_path, f"Pipeline Error - {e}")
        # close Word -> move to errors -> notify (files already moved on by production are skipped with a print)
        fail_file(doc_path, "Pipeline Error - Check formatting", reference_id, doc, word)
        return

def gap(record_pages, latest_file, word, doc, word_id_selected, start_time, page_count,proofs_path, reference_id=None):
    """Calculates the gap between record page numbers to determine Simplex/Duplex"""

    # print(f"{page_count} pages in the document with {len(record_pages)} record pages found") muting cluster

    if len(record_pages) == 0:
        raise ValueError("No record pages found in document")

    if len(record_pages) == 1:
        gapnum = 1
        full_record_amount = len(record_pages)
        output(gapnum, latest_file, word, doc, word_id_selected, start_time, full_record_amount, page_count, proofs_path=proofs_path, reference_id=reference_id)
        return
        
    if len(record_pages) == 2:
        sorted_numbers = sorted(record_pages)
        gapnum = sorted_numbers[1] - sorted_numbers[0]
        if gapnum == 0:
            gapnum = 1
        full_record_amount = 2
        output(gapnum, latest_file, word, doc, word_id_selected, start_time, full_record_amount, page_count, proofs_path=proofs_path, reference_id=reference_id)
        return
        
    # print(f"First 3 record page numbers: {sorted(record_pages)}") muting cluster

    # For 3 or more records
    sorted_numbers = sorted(record_pages)
    gapz = [sorted_numbers[i+1] - sorted_numbers[i] for i in range(len(sorted_numbers) - 1)]
    gapnum = int(gapz[0]) if gapz else 1  # Default to 1 if no gaps found, to prevent errors in output function
    # print(gapz)  muting cluster
    opr_amount = page_count / gapnum
    full_record_amount = int(opr_amount)
    # print(f"Records in File: {full_record_amount}") muting cluster
    output(gapnum, latest_file, word, doc, word_id_selected, start_time, full_record_amount, page_count, proofs_path=proofs_path, reference_id=reference_id)
    # return


def output(gapnum, latest_file, word, doc, word_id_selected, start_time, full_record_amount, page_count, proofs_path, reference_id=None):

    # failsafe if the records are not divisible by the page count
    if page_count % gapnum != 0:
        print("Page count is not divisible by gap between records. Check document formatting.")
        # single Quit - safe_move inside fail_file retries until Word has released the file lock
        fail_file(latest_file, "Page count not divisible by gap - [Check formatting]", reference_id, doc, word)
        return

    # print(f"Gap between numbers: {gapnum}") muting cluster

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
                fail_file(latest_file, f"Unsupported page count ({page_count}) - [Check formatting]", reference_id, doc, word)
                return

        production(latest_file, word, doc, start_time, full_record_amount, type, page_count, proofs_path=proofs_path, reference_id=reference_id)
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
            # previously returned with Word still open and the file locked in Precomposed
            print("Error Incorrect format")
            fail_file(latest_file, f"Unsupported record gap ({gapnum}) - [Check formatting]", reference_id, doc, word)
            return
    # print(type) muting cluster
    DataLoad.JLog(latest_file, dashboard_path, "Type Identified")
    production(latest_file, word, doc, start_time, full_record_amount, type, page_count, proofs_path=proofs_path, reference_id=reference_id)
    # return


def production(latest_file, word, doc, start_time, full_record_amount, type, page_count, proofs_path, reference_id=None):

    processing_json_path = os.path.join(dashboard_path, "ProccessingQueue.json")
    timestamp = time.strftime("%H:%M:%S", time.localtime())
    processing_frame = pd.DataFrame(
        DataLoad.datadict(os.path.basename(latest_file), "PDF Output", timestamp, reference_id)
    )
    try:
        DataLoad.push2(processing_json_path, processing_frame)
        DataLoad.JLog(latest_file, dashboard_path, "PDF Output running")
    except Exception as e:
        print(f"Error during PDF output processing: {e}")
        fail_file(latest_file, f"Error during PDF output stage: {e}", reference_id, doc, word)
        return  # previously carried on with a file that had already been moved to errors

    # pdf_path = latest_file.replace('.docx', '')
    basename = os.path.splitext(os.path.basename(latest_file))[0]  # splitext handles .DOCX as well as .docx
    pdfname = (f"{basename}~#{full_record_amount}~#{type}")
    print(f"PDF Filename - {pdfname}")
    pathtosave = (f"{Srv_path}\\{pdfname}.pdf")
    existing_file = (f"{input_path}\\{pdfname}.pdf")

    try:
        doc.ExportAsFixedFormat(pathtosave, 17)
    except Exception as e:
        # no PDF means no downstream output - fail now instead of waiting 15 minutes for it
        print(f"Error saving document {e}")
        fail_file(latest_file, f"PDF export failed - {e}", reference_id, doc, word)
        return

    close_word(doc, word)
    doc = None  # Ensure the Document object is released
    word = None  # Ensure the Word application object is released

    linkpath = existing_file
    if os.path.exists(existing_file):
        print(f"File {pdfname}.pdf already exists in input")
    else:
        try:
            linkpath = shutil.move(pathtosave, input_path)  #pdf input folder
        except Exception as e:
            print(f"Error moving file: {e}")

    # safe_move adds a timestamp suffix if the name already exists in old, so the file never gets left behind
    try:
        old_file = safe_move(latest_file, old)
    except Exception as e:
        print(f"Error moving file to old: {e}")
        old_file = latest_file

#-------------------------- section --------------- finds and picks up file from output folder, copies to proofs folder                 
    output_pattern = f'{basename}_OUTPUT_'
    print(f"Looking for output file with pattern: {output_pattern}") # debugging line to check the output pattern
    try:
        ready_file = output_file(basename)
        if not ready_file:
            raise TypeError(f"ready file should not be empty check script log for file {latest_file}")
        shutil.copy2(ready_file, proofs_path)
        DataLoad.JLog(latest_file, dashboard_path, "Output file copied to proofs")
    except (TimeoutError, OSError, TypeError) as e:
        print(f"Error: {e}")
        DataLoad.JLog(latest_file, dashboard_path, f"Output file timeout - {e}")
        fail_file(latest_file, "Output PDF never landed - check downstream system", reference_id, move_path=old_file)
        return

    #grab mrdf - a missing/unreadable MRDF is logged but doesn't fail the job
    PDF_Output = os.path.join(proofs_path, os.path.basename(ready_file))
    try:
        MRDF(PDF_Output, proofs_path, type)
    except Exception as e:
        print(f"MRDF error: {e}")
        DataLoad.JLog(latest_file, dashboard_path, f"MRDF lookup failed - {e}")

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
        pcsed.add(latest_file)
    except Exception as e:
        print(f"Error during final processing steps: {e}")
        # latest_file is already in old at this point - move that copy, not the original path
        fail_file(latest_file, f"Error during final processing steps: {e}", reference_id, move_path=old_file)

    print("Returning")
    return

def output_file(basename, timeout=900, check_interval=1, stable_checks=3):
    """
    Waits for a file matching the basename (glob pattern, e.g. '...OUTPUT_*.pdf')
    """
    start = time.time()
    last_size = -1
    stable_count = 0

    while True:
        if time.time() - start > timeout:
            raise TimeoutError(f"No stable file matching {basename} within {timeout}s")


        matches = glob.glob(output_path + "/" + basename + "_OUTPUT_??????????.PDF")
        if not matches:
            time.sleep(check_interval)
            continue  # nothing there yet — don't count this as "stable"

        matched_file = matches[0]
        current_size = os.path.getsize(matched_file)

        if current_size == last_size and current_size > 0:
            stable_count += 1
            if stable_count >= stable_checks:
                return matched_file
        else:
            stable_count = 0

        last_size = current_size
        time.sleep(check_interval)
    
def MRDF(PDF_Output, proofs_path, type):
    print("inside mrdf")
    with pdfplumber.open(PDF_Output) as pdf:  # closes the PDF so it isn't left locked in the proofs folder
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    # print(text) # debug
    if "Simplex" in type: 
        match = re.search(r'10000000\d{10}', text)
    else:
        match = re.search(r'10100000\d{10}', text)
    if match:
        value = match.group() # clean regex output 
        mrdf = value[8:][::-1] #remove first 8 digits and reverse
        print(mrdf) #debug

        mrdf_file = glob.glob(mrdf_path + "\\" + mrdf + ".mrd")
        print(mrdf_file)
        if not mrdf_file:
            print("No mrdf found")
            return

        shutil.copy2(mrdf_file[0], proofs_path) #glob returns a list so i need to index 0 for the only match
        print("MRDF file moved to proofs") # debug
        return
            
    print("No MRDF in file found")
    # ------------------- ------------------ --------------- need to build wavesheet grab next --------------- - - - - - -- 

def Wavesheet(proofs_path):
    filename = time.strftime('%d%m%Y%H', time.localtime())
    Wsheet = glob.glob(wavesheet_path+ "\\" + filename + ".csv")
    print(Wsheet) #debug
    if not Wsheet:
        print("No wavesheet found")
        return

    shutil.copy2(Wsheet[0], proofs_path) # glob returns list - index 0 to get first result (only result)
    print("Wavesheet moved to proofs")
    return Wsheet