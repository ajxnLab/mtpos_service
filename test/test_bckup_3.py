import time
from pywinauto import Application
from utils.logger import setup_in_memory_logger
from matcode_mtpos.mtpos_constant import MTPOS_Constants
from pywinauto.keyboard import send_keys
from utils.helpers import wait

logger,log_stream = setup_in_memory_logger(__name__)

# Instantiate constants
mtpos = MTPOS_Constants()

def find_element_in_parent(
    parent_control_type,
    child_control_type,
    parent_name=None,
    parent_automation_id=None,
    child_name=None,
    child_automation_id=None,
    variable=None,
    action=None,
    visible_only=False,
    main_window=None 
):
    """
    Find a child element inside a parent element, using optional name and/or automation_id filters.
    If visible_only=True, only considers visible (not offscreen) elements.
    """
    try:
        # Find parent
        parent = main_window.child_window(
            control_type=parent_control_type,
            title=parent_name if parent_name else None,
            auto_id=parent_automation_id if parent_automation_id else None
        )
        parent.wait('exists ready', timeout=5)

        # Get all children
        all_children = parent.children()

        # Filter children by criteria
        filtered_children = []
        for c in all_children:
            if child_control_type and c.element_info.control_type != child_control_type:
                continue
            if child_name and c.window_text() != child_name:
                continue
            if child_automation_id and c.element_info.automation_id != child_automation_id:
                continue
            if visible_only and c.is_offscreen():
                continue
            filtered_children.append(c)

        if not filtered_children:
            logger.warning("No matching child elements found (after visible_only filter)" if visible_only else "No matching child elements found")
            return None

        child = filtered_children[0]

        logger.info(f"Found child: {child.window_text()}")
        return perform_action(element=child, variable=variable, action=action, main_window=main_window)

    except Exception as e:
        logger.error(f"Error finding child element inside parent: {e}")
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
                                retries: int = 3, single_attempt_timeout: int = 5, retry_interval: int = 1, main_window=None):
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
        
try:
    app = Application(backend="uia").connect(title_re=r"^MT-POS ENT.*")
    main_window = app.window(title_re=r"^MT-POS ENT.*")

    if app:
        main_window.set_focus()

        find_element_in_parent(
            parent_control_type="ToolBar",
            child_control_type="Button", 
            parent_name="Quick Access Toolbar",
            child_name="Logout", 
            action="click",
            main_window=main_window )

       
    else:
        print("Failed")

except Exception as e:
    print(f"Failed: {e}")