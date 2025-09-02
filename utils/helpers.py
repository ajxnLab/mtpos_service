import time
from datetime import datetime
from utils.logger import setup_in_memory_logger

def wait(seconds):
    time.sleep(seconds)

logger, log_stream = setup_in_memory_logger(service_name="helpers")

def get_datetime(format: str = "iso", tz=None) -> str:
    """
    Returns the current date/time in the requested format.
    
    Args:
        format: str - The format type ("iso", "date", "time", "custom")
        tz: timezone - Optional timezone (use from pytz or zoneinfo)
        
    Returns:
        str - formatted datetime string
    """
    now = datetime.now(tz)
    
    if format == "iso":
        return now.isoformat()
    elif format == "date":
        return now.strftime("%Y-%m-%d")
    elif format == "time":
        return now.strftime("%H:%M:%S")
    elif format == "full":
        return now.strftime("%Y-%m-%d %H:%M:%S")
    elif isinstance(format, str):
        return now.strftime(format)
    else:
        raise ValueError("Unsupported format")

def duration_time(start_time, end_time):
    try:
        start_time_duration = datetime.strptime(start_time, "%Y-%m-%d %H:%M:%S")
        end_time_duration = datetime.strptime(end_time, "%Y-%m-%d %H:%M:%S")
        duration = end_time_duration - start_time_duration
        return duration
    except ValueError as ve:
        # Raised when strptime() fails to parse the datetime
        logger.error(f"Invalid date format: {ve}")
        return None
    
    except TypeError as te:
        # Raised if input is not a string
        logger.error(f"Type error in duration_time(): {te}")
        return None
    except Exception as e:
        # Catch-all for unexpected issues
        logger.error(f"Unexpected error in duration_time(): {e}")
        return None

def data_strip(data):
    data = [
        {k.strip(): str(v).strip() for k, v in row.items()}
        for row in data
    ]
    return data