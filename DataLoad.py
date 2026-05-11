import json
import os
import time
import pandas as pd
import threading

# -------------------------------------------------------------------------------
# FIX: A single module-level lock is used for all JSON file operations.
# This prevents two threads from reading and writing the same file at the same
# time, which was causing one thread to silently overwrite the other's changes.
# -------------------------------------------------------------------------------
_json_lock = threading.Lock()


def datadict(filename, status, timestamp, reference_id=None):
    dataframe = {
        "Filename": [filename],
        "Status": [status],
        "Last_Updated": [timestamp],
        "Reference": [reference_id or ""],
    }
    return dataframe


def ErrorStage(doc_path, dashboard_path, reference_id=None):
    timestamp = time.strftime("%H:%M:%S", time.localtime())
    filename = os.path.basename(doc_path)
    dataframe = datadict(filename, "Error - No Address Match", timestamp, reference_id)
    frame = pd.DataFrame(dataframe)
    json_path = os.path.join(dashboard_path, "ErrorFiles.json")
    push(json_path, frame)
    return


def JobDisplay(df, json_path):
    push(json_path, df)
    return


def Queue(doc_path, dashboard_path, reference_id=None):
    timestamp = time.strftime("%H:%M:%S", time.localtime())
    filename = os.path.basename(doc_path)
    dataframe = datadict(filename, "Queued", timestamp, reference_id)
    frame = pd.DataFrame(dataframe)
    json_path = os.path.join(dashboard_path, "ProccessingQueue.json")
    push(json_path, frame)
    return


def ProcessingStage(doc_path, dashboard_path, reference_id=None):
    timestamp = time.strftime("%H:%M:%S", time.localtime())
    filename = os.path.basename(doc_path)
    dataframe = datadict(filename, "Processing", timestamp, reference_id)
    frame = pd.DataFrame(dataframe)
    json_path = os.path.join(dashboard_path, "ProccessingQueue.json")
    push(json_path, frame)
    return


def PDFStage(doc_path, dashboard_path, reference_id=None):
    timestamp = time.strftime("%H:%M:%S", time.localtime())
    filename = os.path.basename(doc_path)
    dataframe = datadict(filename, "PDF Output", timestamp, reference_id)
    frame = pd.DataFrame(dataframe)
    json_path = os.path.join(dashboard_path, "ProccessingQueue.json")
    push(json_path, frame)
    return


def CompletedStage(latest_file, dashboard_path, dataframe, reference_id=None):
    timestamp = time.strftime("%H:%M:%S", time.localtime())
    filename = os.path.basename(latest_file)
    dataframe = datadict(filename, "Completed", timestamp, reference_id)
    frame = pd.DataFrame(dataframe)
    json_path = os.path.join(dashboard_path, "ProccessingQueue.json")
    push(json_path, frame)
    return


def push2(json_path, frame):
    """
    Updates an existing record in the JSON file matched by Filename.
    If no matching record is found, appends a new one.

    FIX 1: Now matches by Filename instead of by list position (procfiles[-1]).
            Previously, two concurrent threads would both update the last record
            in the list, meaning the wrong file's entry would be overwritten.

    FIX 2: The entire read-modify-write is wrapped in _json_lock so only one
            thread can access the file at a time. Without this, two threads
            could read the same list simultaneously and one would silently
            discard the other's write.
    """
    record = frame.to_dict(orient="records")[0]
    target_filename = record.get("Filename")

    with _json_lock:
        procfiles = _read_json(json_path)

        # FIX: Find the record by Filename, not by position
        matched = False
        for existing in procfiles:
            if existing.get("Filename") == target_filename:
                existing.update(record)  # Update the correct record in place
                matched = True
                break

        if not matched:
            # No existing record for this file — append as new
            procfiles.append(record)

        _write_json(json_path, procfiles)
    return


def push(json_path, frame):
    """
    Appends a new record to the JSON file.

    FIX: Wrapped in _json_lock so concurrent threads cannot read and write
         the file simultaneously, which was causing records to be lost.
    """
    with _json_lock:
        procfiles = _read_json(json_path)
        procfiles.append(frame.to_dict(orient="records")[0])
        _write_json(json_path, procfiles)
    return


def JLog(file, dashboard_path, type):
    filename = os.path.basename(file)
    timestamp = time.strftime("%H:%M:%S", time.localtime())
    log_entry = f"{timestamp} - {filename} - {type}\n"
    json_path = os.path.join(dashboard_path, "ActivityLog.json")
    # FIX: Log appends are also locked to prevent interleaved writes
    with _json_lock:
        try:
            with open(json_path, "a", encoding="utf-8") as f:
                f.write(log_entry)
        except Exception as e:
            print(f"Error writing to log: {e}")
    return


# -------------------------------------------------------------------------------
# Private helpers — shared read/write logic used by push and push2
# These should only ever be called from inside a _json_lock block
# -------------------------------------------------------------------------------

def _read_json(json_path):
    """Reads and returns the JSON list from a file. Returns [] on any error."""
    if not os.path.exists(json_path):
        return []
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            content = f.read().strip()
            if not content:
                return []
            data = json.loads(content)
            if not isinstance(data, list):
                return [data]
            return data
    except (json.JSONDecodeError, ValueError) as e:
        print(f"JSON read error ({json_path}): {e} — starting fresh")
        return []


def _write_json(json_path, data):
    """Writes a list back to a JSON file."""
    try:
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"Error saving to JSON ({json_path}): {e}")


if __name__ == "__main__":
    pass