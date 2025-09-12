import time
import utils.helpers
import concurrent.futures
from pywinauto import Application, Desktop
from config.env_config import get_env_variable, load_environment
from utils.google_sheet import GSheetClient
from utils.logger import setup_in_memory_logger
from matcode_mtpos.mtpos_constant import MTPOS_Constants
from pywinauto.keyboard import send_keys
from utils.helpers import wait
from pywinauto.controls.uiawrapper import UIAWrapper
from comtypes.gen import UIAutomationClient
from comtypes import POINTER
from pywinauto.base_wrapper import BaseWrapper
from pywinauto.application import WindowSpecification

UIA_ValuePatternId = UIAutomationClient.UIA_ValuePatternId
UIA_LegacyIAccessiblePatternId = UIAutomationClient.UIA_LegacyIAccessiblePatternId

logger,log_stream = setup_in_memory_logger(__name__)


# Instantiate constants
mtpos = MTPOS_Constants()

PATTERNS = {
    UIAutomationClient.UIA_ValuePatternId: "ValuePattern",
    UIAutomationClient.UIA_LegacyIAccessiblePatternId: "LegacyIAccessiblePattern",
    UIAutomationClient.UIA_InvokePatternId: "InvokePattern",
    UIAutomationClient.UIA_SelectionItemPatternId: "SelectionItemPattern",
    UIAutomationClient.UIA_ExpandCollapsePatternId: "ExpandCollapsePattern",
    UIAutomationClient.UIA_TextPatternId: "TextPattern",
    UIAutomationClient.UIA_ScrollItemPatternId: "ScrollItemPattern",
    # add more if needed
}

def find_element_in_parent(
        child_control_type,
        parent_control_type=None,
        parent_name=None,
        parent_automation_id=None,
        child_name=None,
        child_automation_id=None,
        variable=None,
        action=None,
        visible_only=False,
        search_descendants=False,
        element=None,
        timeout=180,
        interval=1,
        main_window=None 
):
    """
        Find a child (or descendant) element inside a parent element, using optional filters.
        Executes action if specified.

        Args:
            search_descendants (bool): If True, searches all nested controls with .descendants().
            element: Optional parent element already found (skip searching main_window).
        """
    try:
        # 1. Locate parent
        if element:
            parent = element
        else:
            parent = main_window.child_window(
                control_type=parent_control_type,
                title=parent_name if parent_name else None,
                auto_id=parent_automation_id if parent_automation_id else None
            )
            parent.wait("exists ready", timeout=timeout)
            parent = parent.wrapper_object()

        # 2. Poll for child within timeout
        end_time = time.time() + timeout
        while time.time() < end_time:
            # Get elements
            if search_descendants:
                elements = [UIAWrapper(e) for e in parent.element_info.descendants()]
            else:
                elements = [UIAWrapper(e) for e in parent.element_info.children()]

            for elem in elements:
                try:
                    # Filter by control_type, name, automation_id, visibility
                    if child_control_type and getattr(elem.element_info, "control_type", None) != child_control_type:
                        continue
                    if child_name and elem.window_text() != child_name:
                        continue
                    if child_automation_id and getattr(elem.element_info, "automation_id", None) != child_automation_id:
                        continue
                    if visible_only and getattr(elem.element_info, "is_offscreen", True):
                        continue

                    logger.info(f"Found child element: {elem.window_text()} ({getattr(elem.element_info, 'control_type', None)})")
                    return perform_action(element=elem, variable=variable, action=action, main_window=main_window)
                except Exception as e:
                    logger.warning(f"Skipping element due to error: {e}")
                    continue

            time.sleep(interval)

        logger.error(f"No child element found within {timeout} seconds (type={child_control_type}).")
        return None

    except Exception as e:
        logger.error(f"Error in find_element_in_parent: {e}")
        return None
    
def find_element_with_index(
        main_window,
        control_type: str,
        automation_id: str = None,
        name: str = None,
        found_index: int = None,
        timeout: int = 60,
        retry_interval: int = 1,
        variable=None,
        action=None,
        search_descendants: bool = True,
        element=None
    ):
    """
    Finds a UI element by control_type and optional automation_id / name.
    If multiple matches exist, select one by found_index (default = 0).
    """
    try:
        # 1. Pick search parent
        parent = element if element else main_window
        # Collect matching elements
        candidates = parent.descendants(control_type=control_type) if search_descendants \
            else parent.children(control_type=control_type)

        # Filter by automation_id
        if automation_id:
            candidates = [c for c in candidates if getattr(c.element_info, "automation_id", "") == automation_id]

        # Filter by name
        if name:
            candidates = [c for c in candidates if c.window_text() == name]

        if not candidates:
            raise RuntimeError(f"No element found (type={control_type}, id={automation_id}, name={name})")

        # Log multiple matches
        if len(candidates) > 1:
            logger.warning(f"Multiple elements found ({len(candidates)}) for control_type={control_type}")
            for i, c in enumerate(candidates):
                logger.warning(f" Index {i}: type={c.element_info.control_type}, "
                               f"name='{c.window_text()}', "
                               f"auto_id='{getattr(c.element_info, 'automation_id', '')}'")

        # Pick by index
        index = found_index if found_index is not None and 0 <= found_index < len(candidates) else 0
        element = candidates[index]

        if found_index is not None:
            logger.info(f"Using element at index {index}")
        elif len(candidates) > 1:
            logger.info("No found_index specified, defaulting to index 0")

        #Only call wait() if it's a WindowSpecification
        if isinstance(element, WindowSpecification):
            element.wait("exists ready", timeout=timeout, retry_interval=retry_interval)
        else:
            # Already a UIAWrapper -> just check it's alive
            if not element.is_enabled() and not element.is_visible():
                raise RuntimeError(f"Element at index {index} is not enabled/visible")

        # Perform optional action
        result = perform_action(element=element, variable=variable, action=action, main_window=main_window)
        return result if result is not None else element

    except Exception as e:
        logger.error(f"Error in find_element_with_index: {e}")
        raise
        
def find_element(
        control_type: str,
        automation_id: str = None,
        name: str = None,
        variable=None,
        action=None,
        timeout: int = 60,
        retry_interval: int = 1,
        main_window=None 
    ):
        """
        Finds a UI element by control_type and optional automation_id / name.
        Then performs optional action (e.g., click, sendkeys, find).
        """
        criteria = {"control_type": control_type}
        if automation_id:
            criteria["auto_id"] = automation_id
        if name:
            criteria["title"] = name

        try:
            element_spec = main_window.child_window(**criteria)
            if element_spec.exists(timeout=timeout, retry_interval=retry_interval):
                
                # Perform optional action
                result = perform_action(element=element_spec, variable=variable, action=action,main_window=main_window)

                # Return what perform_action returned if any, else the element itself
                return result if result is not None else element_spec
            else:
                logger.error(f"Element not found within timeout. Criteria: {criteria}")
                raise RuntimeError(f"Element not found: {criteria}")

        except Exception as e:
            logger.error(f"Error finding element with criteria {criteria}: {e}")
            raise RuntimeError(f"Failed to find element: {e}") from e
        
def perform_action(element, action,main_window=None, variable=None):
    if not element:
        logger.error("Element not found")
        raise RuntimeError("Element not found")

    try:
        if action == "click":
            element.click_input()
        elif action == "double_click":
            element.double_click_input()
        elif action == "sendkeys":
                element.set_edit_text("")   # clear old text
                if variable is not None:
                    element.set_edit_text(variable)  # enter new text
                else:
                    logger.warning("No variable provided to sendkeys")
        elif action == "send_type":
            element.set_focus()
            # clear old text by selecting all and deleting
            element.type_keys("^a{BACKSPACE}", set_foreground=True)
            if variable is not None:
                element.type_keys(variable, with_spaces=True, set_foreground=True)
            else:
                logger.warning("No variable provided to sendkeys")
        elif action == "find":
            return element
        elif action == "control_keys":
            send_keys_to(element=element ,keys=variable)
        else:
            logger.error(f"Unknown action: {action}")
            raise ValueError(f"Unknown action: {action}")
    except Exception as e:
        logger.error(f"Error performing action '{action}' on element: {e}")
        raise
def wait_until_element_present(control_type: str, automation_id: str = None, name: str = None,
                                retries: int = 3, single_attempt_timeout: int = 1, retry_interval: int = 1, main_window=None):
        """
        Try up to `retries` times to find a visible element matching the criteria.
        Each attempt can wait up to `single_attempt_timeout` seconds.
        """
        criteria = {"control_type": control_type}
        if automation_id:
            criteria["auto_id"] = automation_id
        if name:
            criteria["title"] = name

        for attempt in range(1, retries + 1):
            try:
                element = main_window.child_window(**criteria)
                element.wait('visible', timeout=single_attempt_timeout)
                logger.info(f"Found element on attempt {attempt}")
                return element
            except Exception as e:
                logger.debug(f"Attempt {attempt} failed: {e}")
                if attempt < retries:
                    time.sleep(retry_interval)

        logger.error(f"Element not found after {retries} retries × {single_attempt_timeout}s each")
        return None

def send_keys_to(keys: str, main_window=None,element=None):
    """
    Send keys to a specific element if provided, otherwise to the main window.
    """
    if element:
        try:
            element.set_focus()
            send_keys(keys, pause=0.05)
            logger.info(f"Sent keys '{keys}' to element")
        except Exception as e:
            logger.error(f"Failed to send keys to element: {e}")
            raise
    elif main_window:
        try:
            main_window.set_focus()
            send_keys(keys, pause=0.05)
            logger.info(f"Sent keys '{keys}' to main window")
        except Exception as e:
            logger.error(f"Failed to send keys to main window: {e}")
            raise
    else:
        logger.error("No element or main window to send keys to")
        raise RuntimeError("No target for send_keys")   
    
def safe_get_value(uia_elem):
    """Try ValuePattern -> LegacyIAccessiblePattern -> fallback Name"""
    try:
        vp_raw = uia_elem.GetCurrentPattern(UIA_ValuePatternId)
        vp = vp_raw.QueryInterface(UIAutomationClient.IUIAutomationValuePattern)
        return vp.CurrentValue
    except Exception:
        pass

    try:
        legacy_raw = uia_elem.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        legacy = legacy_raw.QueryInterface(UIAutomationClient.IUIAutomationLegacyIAccessiblePattern)
        return legacy.CurrentValue
    except Exception:
        pass

    try:
        return uia_elem.CurrentName
    except Exception:
        return None


    
def debug_element(uia_elem):

    print("Supported Patterns for:", uia_elem.CurrentName)
    for pid, pname in PATTERNS.items():
        try:
            pattern = uia_elem.GetCurrentPattern(pid)
            if pattern:
                print(f"  ✔ {pname}")
        except Exception:
            pass


def clear_all(child_name):
    
    find_element_in_parent(
    parent_control_type="Custom",
    child_control_type="DataItem", 
    parent_name="Filter Row",
    child_name=child_name, 
    action="double_click"
    )
    wait(0.3)
    send_keys('^+a{BACKSPACE}')

def extract_text(row):
    uia_elem = row.element_info.element
    raw_value = safe_get_value(uia_elem)
    if raw_value:
        return raw_value.split(";")  # split into columns
    return []

def extract_all_rows(data_panel, scrollbar, max_scrolls=100):
    """
    Extract all rows from a virtualized table by scrolling until all unique rows are collected.
    
    Args:
        data_panel: pywinauto wrapper for the Custom Data Panel (row container)
        scrollbar: pywinauto wrapper for the vertical scrollbar
        max_scrolls: safety limit to prevent infinite loops
    
    Returns:
        list of row strings
    """
    seen = set()
    all_rows = []

    for i in range(max_scrolls):
        # --- read visible rows ---
        rows = data_panel.children()
        for row in rows:
            try:
                value = safe_get_value(row.element_info.element)
                if value and value not in seen:
                    seen.add(value)
                    all_rows.append(value)
            except Exception:
                continue

        # --- try to scroll down ---
        try:
            page_down = scrollbar.child_window(title="Page Down", control_type="Button")
            page_down.click_input()
        except Exception:
            break  # no more scrolling possible

        time.sleep(0.3)  # give UI time to refresh

    return all_rows

try:
    load_environment("dev")
    # Google Sheets client
    gsheet_credential = get_env_variable("GOOGLE_SERVICE_ACCOUNT")
    gsheet_id = get_env_variable("GSHEET")

    gs = GSheetClient(gsheet_credential, gsheet_id)
    sheet_tab_sor = mtpos.WORKSHEET_TAB_SOR    
    data = gs.get_sheet_data(sheet_tab_sor) 
    identifier_column = mtpos.MATERIAL_CODE

    today_date = utils.helpers.get_datetime("date") 

    filtered_data = []
    for row in data:
        deployment_date = str(row.get("Deployment Date", "")).strip()

        if deployment_date == today_date:
            filtered_data.append(row)

    print(len(filtered_data))
    pending_updates = [] 
    for row in filtered_data:

         
        material_code = row.get("Material Code")

        result = gs.find_row_index(data, identifier_column, material_code)
        row_index = result["row_index"]

        pending_updates.append((
        mtpos.WORKSHEET_TAB_SOR,   # sheet/tab name
        row_index,
        "RPA Definition Remarks - GT",
        "Success"
        ))

        pending_updates.append((
        mtpos.WORKSHEET_TAB_SOR,   # sheet/tab name
        row_index,
        "RPA Deployment Remarks - GT",
        "Published"
        ))

        
    print(len(pending_updates))
    gs.flush_updates(mtpos.WORKSHEET_TAB_SOR, pending_updates)



except Exception as e:
    print(f"Failed: {e}") 


