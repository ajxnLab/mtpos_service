import time
import win32gui
from pywinauto import Application, Desktop
from utils.logger import setup_in_memory_logger
from matcode_mtpos.mtpos_constant import MTPOS_Constants
from pywinauto.keyboard import send_keys
from utils.helpers import wait
from pywinauto.controls.uiawrapper import UIAWrapper
from comtypes.gen import UIAutomationClient
from comtypes import POINTER

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
            send_keys(keys)
            logger.info(f"Sent keys '{keys}' to element")
        except Exception as e:
            logger.error(f"Failed to send keys to element: {e}")
            raise
    elif main_window:
        try:
            main_window.set_focus()
            send_keys(keys)
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

try:
    
    # Connect to the top-level Coupon List window
    app = Application(backend="uia").connect(title_re=r"^Coupon list - Ver.*$")

    coupon_win = app.window(title_re=r"^Coupon list - Ver.*$")
    coupon_win.set_focus()

    # Navigate down to Promotion Wizard
    promo_wizard = coupon_win.child_window(title_re=r"^Promotion Wizard -.*$", control_type="Window")

    # Navigate down further to the Promotion window
    promo_selected = promo_wizard.child_window(title_re=r"^Promotion \[.*\] Selected Cats$", control_type="Window")

    if promo_selected.exists():
        promo_selected.set_focus()

        custom_data_panel = wait_until_element_present(automation_id="GroupControl1", control_type="Pane", retries=3, single_attempt_timeout=60, retry_interval=0,main_window=promo_selected)

        ctrl=find_element_in_parent(
                            child_control_type="TreeItem",
                            action="find",
                            element = custom_data_panel,
                            search_descendants = True,
                            main_window=promo_selected
                        )

        #ctrl = promo_selected.child_window(title="Node0", control_type="TreeItem").wrapper_object()
        uia_elem = ctrl.element_info.element

        value = safe_get_value(uia_elem)
        debug_element(uia_elem)
        print("Value:", value)
        print("Found:", promo_selected.window_text())
    else:
        print("Promotion window not found")

except Exception as e:
    print(f"Failed: {e}")