import logging
import os
import sys
import traceback
from googleapiclient.http import MediaIoBaseUpload
from googleapiclient.discovery import build
from io import BytesIO, StringIO
from datetime import datetime

# global cache of handlers (so all services reuse them)
_shared_log_stream = None
_shared_file_handler = None

def attach_drive_client(logger, gs_client, constants=None, log_stream=None):
    logger.gs_client = gs_client
    logger.log_stream = log_stream
    if constants:
        logger.constants = constants
    logger.info("Google Drive client attached to logger.")

def log_traceback(logger, exception):
    """
    Log full traceback of an exception only to FileHandler logs.
    """
    traceback_str = ''.join(traceback.format_exception(type(exception), exception, exception.__traceback__))
    
    wrote = False
    for handler in logger.handlers:
        if isinstance(handler, logging.FileHandler):
            handler.acquire()
            try:
                handler.stream.write(traceback_str + "\n")
                handler.flush()
                wrote = True
            finally:
                handler.release()

    if not wrote:
        # fallback: log at ERROR so you at least see it in console
        logger.error(traceback_str)


def setup_in_memory_logger(service_name: str):
    global _shared_log_stream, _shared_file_handler

    logger = logging.getLogger(service_name)

    if logger.hasHandlers():
        logger.handlers.clear()

    logger.setLevel(logging.INFO)

    formatter = logging.Formatter(
        fmt="[%(asctime)s,%(msecs)03d]: [%(name)s] : [%(levelname)s]:"
            "[%(filename)s:%(lineno)d - %(funcName)s()]: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    # Create shared log_stream and file handler if not already made
    if _shared_log_stream is None:
        _shared_log_stream = StringIO()

    if _shared_file_handler is None:
        now = datetime.now()
        folder_path = os.path.join(
            "local_logs",
            now.strftime("%Y"),
            now.strftime("%B"),
            now.strftime("%d"),
            now.strftime("%H")
        )
        os.makedirs(folder_path, exist_ok=True)

        file_name = f"{now.strftime('%Y-%m-%d_%H-%M-%S')}.log"
        _shared_file_handler = logging.FileHandler(os.path.join(folder_path, file_name), mode="w", encoding="utf-8")
        _shared_file_handler.setFormatter(formatter)

    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # Shared memory handler
    stream_handler = logging.StreamHandler(_shared_log_stream)
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    # Shared file handler (same file for all services)
    logger.addHandler(_shared_file_handler)

    return logger, _shared_log_stream

def create_drive_folder(service, name: str, parent_id: str = None) -> str:
    """Create a folder in Drive (or return existing one)."""
    query = f"mimeType='application/vnd.google-apps.folder' and trashed=false and name='{name}'"
    if parent_id:
        query += f" and '{parent_id}' in parents"
    results = service.files().list(
        q=query,
        fields="files(id, name)",
        supportsAllDrives=True,                    
        includeItemsFromAllDrives=True             
    ).execute()
    items = results.get('files', [])
    if items:
        return items[0]['id']
    file_metadata = {
        'name': name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_id] if parent_id else []
    }
    folder = service.files().create(body=file_metadata,supportsAllDrives=True,fields='id').execute()
    return folder['id']


def create_nested_drive_path(service, base_folder_id: str, path_parts: list[str]) -> str:
    """Create nested folders like ['2025', '06', '20', '14'] under a base folder."""
    current_folder_id = base_folder_id
    for part in path_parts:
        current_folder_id = create_drive_folder(service, part, current_folder_id)
    return current_folder_id



def upload_log_to_drive(service, content: str, filename: str, folder_id: str):
    # Encode to UTF-8 bytes
    byte_stream = BytesIO(content.encode("utf-8"))
    media = MediaIoBaseUpload(byte_stream, mimetype="text/plain", resumable=False)
    
    file_metadata = {
        "name": filename,
        "parents": [folder_id]
    }
    uploaded_file = service.files().create(
        body=file_metadata,
        media_body=media,
        supportsAllDrives=True,
        fields="id, name"
    ).execute()
    print(f"Uploaded log: {uploaded_file['name']} to Google Drive")

def finalize_log_upload(logger,root_folder):
    #Uploads the in-memory logs to Drive after process ends.

    if hasattr(logger, "gs_client") and hasattr(logger, "log_stream") and hasattr(logger, "constants"):
        
        drive = logger.gs_client.authenticate_google_drive()
        now = datetime.now()
        filename = f"{logger.name}_{now.strftime('%Y-%m-%d_%H-%M-%S')}.log"
        subfolders = [now.strftime("%Y"), now.strftime("%B"), now.strftime("%d"), now.strftime("%H")]
        target_folder_id = create_nested_drive_path(drive, root_folder, subfolders)

        upload_log_to_drive(drive, logger.log_stream.getvalue(), filename, target_folder_id)
        logger.info("Logs uploaded to Google Drive.")