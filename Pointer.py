# Pointer
import time
import logging
import threading
import watchdog.observers
import watchdog.events
import os
import json
import AutoUUScript
import WinNotify
import sys
import DataLoad

def resource_path(filename):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.abspath("."), filename)

moved_event_flag = threading.Event()
queue_lock = threading.Lock()  # guards filelist, processing_files and last_event_time across observer/worker threads
filelist = []  # Global list to accumulate files
processing_files = set()  # Track files currently being processed
last_event_time = 0.0

QUIET_PERIOD = 5        # seconds with no new files before a batch starts - files dropped together go out as one batch
STABLE_CHECKS = 3       # identical size reads in a row before a file counts as fully copied
STABLE_TIMEOUT = 300    # give up on a file that is still copying after this many seconds


class Batch:
    """Counters for one group of files - each batch has its own so overlapping batches can't reset each other"""
    def __init__(self, files, proofs_path):
        self.files = list(files)
        self.names = {os.path.basename(f) for f in files}
        self.proofs_path = proofs_path
        self.total = len(files)
        self.done = 0
        self.passed = 0
        self.lock = threading.Lock()

    def finish_one(self, passed):
        """Returns True only for the thread that finishes the last file in the batch"""
        with self.lock:
            self.done += 1
            if passed:
                self.passed += 1
            return self.done == self.total


class Handler(watchdog.events.PatternMatchingEventHandler):
    def __init__(self):
        watchdog.events.PatternMatchingEventHandler.__init__(self,
        patterns=["*.docx"],
        ignore_patterns=["~$*"],
        ignore_directories=True,
        case_sensitive=False)

    def on_created(self, event):
        self.queue_file(event.src_path)

    def on_moved(self, event):
        # copy tools (MoveIT) often write a temp name then rename it to .docx - that is a move, not a create
        name = os.path.basename(event.dest_path)
        if name.lower().endswith(".docx") and not name.startswith("~$"):
            self.queue_file(event.dest_path)

    def queue_file(self, path):
        global last_event_time
        print("File Detected")
        with queue_lock:
            last_event_time = time.time()
            # Only add file if not already in list or being processed
            if path not in filelist and path not in processing_files:
                filelist.append(path)
                moved_event_flag.set()

################################################################################

################################################################################

def wait_until_copied(file_path):
    # return #not in use
    """Waits until the file stops growing and can be opened, so a half-copied file is never moved"""
    start = time.time()
    last_size = -1
    stable = 0
    while time.time() - start < STABLE_TIMEOUT:
        try:
            size = os.path.getsize(file_path)
            if size == last_size and size > 0:
                stable += 1
                if stable >= STABLE_CHECKS:
                    with open(file_path, "rb+"):  # fails while the copier still has the file open
                        return True
            else:
                stable = 0
            last_size = size
        except FileNotFoundError:
            return False  # file was removed before we got to it
        except OSError:
            stable = 0  # still locked by the copier
        time.sleep(1)
    return False

# process files concurrently as they arrive
def process_file_thread(file_path, batch):     # seperate thread for each file
    name = os.path.basename(file_path)
    passed = False
    try:
        DataLoad.ProcessingStage(file_path, Dashboard_path)
        print(f"Processing: {name}")
        if not wait_until_copied(file_path):
            raise TimeoutError(f"{name} never finished copying or was removed")
        AutoUUScript.main(file_path, proofs_path=batch.proofs_path)  # call main function from AutoUUScript.py
        passed = name in {os.path.basename(p) for p in list(AutoUUScript.pcsed)}
    except Exception as e:
        # anything AutoUUScript didn't handle - logged so the batch still completes
        logging.exception(f"Unhandled error processing {file_path}")
        print(f"Error processing {file_path}: {e}")
        DataLoad.JLog(file_path, Dashboard_path, f"Unhandled processing error - {e}")
        AutoUUScript.errdfiles.add(file_path)
    finally:
        with queue_lock:
            processing_files.discard(file_path)
        print(f"{name} finished ({batch.done + 1}/{batch.total})")
        if batch.finish_one(passed):
            finish_batch(batch)

def finish_batch(batch):
    failed = batch.total - batch.passed  # anything that didn't pass counts as failed, including unhandled crashes
    try:
        wsheet = None
        if batch.passed > 0:
            wsheet = AutoUUScript.Wavesheet(batch.proofs_path) # only grab wavesheet if successful jobs is not 0
        sendInfo(batch.total, batch.files, batch.proofs_path, batch.passed, failed, wsheet) # send count numbers for notifications/emails
    except Exception:
        logging.exception("Error sending batch notification")
    finally:
        # only clear this batch's entries - another batch may still be running
        for results in (AutoUUScript.pcsed, AutoUUScript.errdfiles):
            for p in list(results):
                if os.path.basename(p) in batch.names:
                    results.discard(p)

def startz():
    while True:
        moved_event_flag.wait()  # blocks until file is detected
        # wait for a quiet period so a drop of many files becomes one batch, not one batch per file
        while time.time() - last_event_time < QUIET_PERIOD:
            time.sleep(0.5)
        moved_event_flag.clear()  # reset flag

        with queue_lock:
            to_process = [f for f in filelist if f not in processing_files]
        if not to_process:
            continue

        try:
            proof_path = create_proofs_folder()  # create proofs folder before processing
        except OSError:
            # network share unavailable - keep the files queued and retry instead of killing the watcher
            logging.exception("Could not create proofs folder - retrying in 30s")
            time.sleep(30)
            moved_event_flag.set()
            continue

        with queue_lock:
            for f in to_process:
                filelist.remove(f)  # remove file before processing
                processing_files.add(f)

        batch = Batch(to_process, proof_path)
        WinNotify.startNotification()
        print(f"Batch to process: {batch.total}")
        for f in to_process:
            t = threading.Thread(target=process_file_thread, args=(f, batch))
            t.start()

def sendInfo(total, files, proofs_path, pcsed_count, errdfiles_count, wsheet):

    print(f"{pcsed_count} files processed successfully.")
    print(f"{errdfiles_count} files errored.")

    if pcsed_count == 0:
        proofs_path = f"DELETE FOLDER IS EMPTY - {proofs_path}"

    print("All files in batch processed.")
    WinNotify.CompletedNotificaiton(total, pcsed_count, errdfiles_count, files, proofs_path, total, wsheet)


def create_proofs_folder():
    ext = 0
    folderName = f"UUDM_{time.strftime('%d%m%Y')}"
    taken = True
    while taken:
        proofs_path = os.path.join(proofing_path, folderName)
        if not os.path.exists(proofs_path):
            os.makedirs(proofs_path)
            taken = False
        else:
            ext += 1
            folderName = f"UUDM_{time.strftime('%d%m%Y')}_{ext}"

    print(f"Created Proofs folder at {proofs_path}")
    return proofs_path

if __name__ == "__main__":

    logging.basicConfig(filename="UUlog.log",filemode="a",
                    level=logging.INFO,
                    format="%(asctime)s - %(message)s",
                    datefmt="%Y-%m-%d %H:%M:%S")
    observer = None
    exit_code = 0

    try:
    # load paths from json file
        with open(resource_path("LivePaths.json")) as json_paths:
            PATH = json.load(json_paths)
        src_path2 = format(PATH["Client_path"])
        Dashboard_path = format(PATH["Dashboard_path"])
        proofing_path = format(PATH["Proofs_path"])

    # initialize observer
        observer = watchdog.observers.Observer()
        observer.schedule(Handler(), src_path2, recursive=True)

    # start observer
        observer.start()
        print("Observer running..")
        logging.info("Observer running")
        WinNotify.observerRunning()
        startz()

    except KeyboardInterrupt:
        print(f"Shutting down")
    except Exception:
        # no console in the exe - without this a crash leaves no trace
        logging.exception("Pointer stopped unexpectedly")
        exit_code = 1
    finally:
        if observer is not None:
            observer.stop()
            observer.join()
        logging.shutdown()
    sys.exit(exit_code)
