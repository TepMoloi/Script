# Point
import time
import logging 
import threading
from watchdog.events import LoggingEventHandler
import watchdog.observers 
import watchdog.events
import os
import json
import pandas as pd
import AutoUUScript
import WinNotify
import sys
import DataLoad

def resource_path(filename):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.abspath("."), filename)

moved_event_flag = threading.Event()
latest_file = None
filelist = []  # Global list to accumulate files
processing_files = set()  # Track files currently being processed
start_time = None
counter = 0

class Handler(watchdog.events.PatternMatchingEventHandler, LoggingEventHandler):
    def __init__(self):
        watchdog.events.PatternMatchingEventHandler.__init__(self, 
        patterns=["*.docx"],
        ignore_patterns=["~$*"],
        ignore_directories=True,
        case_sensitive=False)
 
        logging.basicConfig(filename="UUlog.log",filemode="a",
                        level=logging.INFO,
                        format="%(asctime)s - %(message)s",
                        datefmt="%Y-%m-%d %H:%M:%S")

    def on_created(self, event):
        global latest_file, filelist, counter, filename, processing_files
        counter += 1
        filename = os.path.basename(event.src_path)
        print("File Detected")
        latest_file = event.src_path
        # Only add file if not already in list or being processed
        if latest_file not in filelist and latest_file not in processing_files:
            filelist.append(latest_file)
            moved_event_flag.set()
        print(f"Queue size: {len(filelist)}")

################################################################################

def datastore(filename, type):
    dataframe = {
        "Filename": [filename],
        "Type": [type],
        "Timestamp": [time.ctime()],
    }

    # Ensure download/backup path exists
    try:
        os.makedirs(src_path2, exist_ok=True)
    except Exception:
        pass

    df = pd.DataFrame(dataframe)
    filenamez = "AutoUU_Data.csv"
    csv_path = os.path.join(src_path2, filenamez)
    if not os.path.exists(csv_path):
        df.to_csv(csv_path, mode="w", header=True, index=False)
    else:
        df.to_csv(csv_path, mode="a", header=False, index=False)
    print("Data saved")
    return

################################################################################

if __name__ == "__main__":
# load paths from json file
    with open(resource_path("TestPaths.json")) as json_paths:
        PATH = json.load(json_paths)
    src_path = format(PATH["Server_path"])
    src_path2 = format(PATH["Client_path"])
    Dashboard_path = format(PATH["Dashboard_path"])

# initialize observer
# loop through both folders
    observer = watchdog.observers.Observer()
    event_handler = LoggingEventHandler()
    event_handler = Handler()
    observer.schedule(event_handler, src_path2, recursive=True)
    file = "UUlog.log"

# start observer
    observer.start()
    print("Observer running..")
    WinNotify.observerRunning()
    
# process files concurrently as they arrive

    def process_file_thread(file_path):
        global processing_files
        # pcdpath = DataLoad.datadict(os.path.basename(file_path), "Queued", times)
        # procframe = pd.DataFrame(pcdpath)
        # procjson_path = os.path.join(Dashboard_path, "ProccessingQueue.json")

        DataLoad.ProcessingStage(file_path, Dashboard_path)
        try:
            processing_files.add(file_path)
            print(f"Processing: {os.path.basename(file_path)}")
            AutoUUScript.main(file_path)
        except Exception as e:
            print(f"Error processing {file_path}: {e}")
        finally:
            processing_files.discard(file_path)
        print("Waiting..")

    try:
        while True:
            moved_event_flag.wait()  # blocks until file is detected
            moved_event_flag.clear()  # reset flag
            # Process all files in the queue concurrently, only if not already being processed
            if filelist:
                WinNotify.startNotification()
                to_process = [f for f in filelist if f not in processing_files]
                for f in to_process:
                    filelist.remove(f)  # remove file before processing
                    processing_files.add(f)
                    t = threading.Thread(target=process_file_thread, args=(f,))
                    # t2 = threading.Thread(target=process_file_thread, args=(f,))
                    t.start()
                    # t2.start()
            time.sleep(0.5)

    except KeyboardInterrupt:
        observer.stop()
        observer.join()
        logging.shutdown()
        print(f"Shutting down")
        exit()
    
