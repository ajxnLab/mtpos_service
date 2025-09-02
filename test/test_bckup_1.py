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
            main_window,
            parent_control_type,
            child_control_type,
            parent_name=None,
            parent_automation_id=None,
            child_name=None,
            child_automation_id=None,
            variable=None,
            action=None
        ):
        """
        Find a child element inside a parent element, using optional name and/or automation_id filters.
        """
        try:
            # Step 1: find parent(s)
            parent = main_window.child_window(
                control_type=parent_control_type,
                title=parent_name if parent_name else None,
                auto_id=parent_automation_id if parent_automation_id else None
            )
            parent.wait('exists ready', timeout=5)

            # Step 2: find child inside parent
            child = parent.child_window(
                control_type=child_control_type,
                title=child_name if child_name else None,
                auto_id=child_automation_id if child_automation_id else None
            )
            child.wait('exists ready', timeout=5)

            logger.info(f"Found child: {child.window_text()}")
            return perform_action(element=child, variable=variable, action=action,main_window=main_window)

        except Exception as e:
            logger.error(f"Error finding child element inside parent: {e}")
            return None
        
def find_element(control_type: str, automation_id: str = None, name: str = None, variable=None, action=None,main_window=None ):
        """
        Finds a UI element using control_type, and optionally automation_id and/or name.
        """
        try:
            criteria = {"control_type": control_type}
            if automation_id:
                criteria["auto_id"] = automation_id
            if name:
                criteria["title"] = name

            element_spec = main_window.child_window(**criteria)
            element_spec.wait('visible', timeout=10)

            # Capture return value if any
            result = perform_action(element=element_spec,main_window=main_window, variable=variable, action=action)

            # If perform_action returned something (e.g., for 'find'), return it
            return result if result is not None else element_spec

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
            send_keys_to(element=element ,keys=variable,main_window =main_window,)
        else:
            logger.error(f"Unknown action: {action}")
            raise ValueError(f"Unknown action: {action}")
    except Exception as e:
        logger.error(f"Error performing action '{action}' on element: {e}")
        raise

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
    app = Application(backend="uia").connect(title_re="^Inventory -.*$")
    main_window = app.window(title_re="^Inventory -.*$")

    """ app = Application(backend="uia").connect(title_re="^Update Stores Inventory -.*$")
    main_window = app.window(title_re="^Update Stores Inventory -.*$") """
   
    
    """ wait(3)
    app_1 = Application(backend="uia").connect(title_re="^Update Stores Inventory -.*$")
    main_window_1 = app_1.window(title_re="^Update Stores Inventory -.*$")
    main_window_1.print_control_identifiers() """

    if app:
        main_window.set_focus()
        send_keys('^m')

        # Wait for the new window to appear
        found = False
        for attempt in range(30):  # check up to ~30 seconds
            try:
                app_1 = Application(backend="uia").connect(title_re="^Update Stores Inventory -.*$")
                main_window_1 = app_1.window(title_re="^Update Stores Inventory -.*$")
                if main_window_1.exists(timeout=1):
                    found = True
                    break
            except Exception:
                pass
            wait(1)

        if not found:
            logger.error("Update Stores Inventory window did not appear in time.")
            raise RuntimeError("Update Stores Inventory window not found.")
    
        main_window_1.set_focus()

        # Now wait for inner element
        element = main_window_1.child_window(auto_id="SplitContainer1", control_type="Pane")
        if element.exists(timeout=60, retry_interval=1):
            print("Success")

            for auto_id, control_type in [
                ("ckSelectAllLocations", "CheckBox")
            ]:
                find_element(automation_id=auto_id, control_type=control_type, action="click", main_window=main_window_1)
            """ for auto_id in ("dtFromDate", "dtToDate"):
                found = False
                for attempt in range(1, 4):  # try 3 times
                    try:
                        t0 = time.time()
                        pane = find_element(automation_id=auto_id, control_type="Pane",action="find", main_window=main_window_1,)
                        pane.wait('exists ready', timeout=1)  # short timeout
                        pane.set_focus()
                        send_keys('{SPACE}')
                        logger.info(f"Clicked {auto_id} in {time.time() - t0:.2f}s on attempt {attempt}")
                        found = True
                        break
                    except Exception as e:
                        logger.warning(f"Attempt {attempt}: failed to find {auto_id}: {e}")
                        time.sleep(0.5)  # wait before retry
                if not found:
                    raise RuntimeError(f"Could not find date picker with auto_id={auto_id}")
            logger.info("Update Stores Inventory window found and ready!") """
        else:
            logger.error("PanelControl1 not found inside new window.")

        
        """ main_window.set_focus() 
        send_keys('^m')
        wait(60)
        app_1 = Application(backend="uia").connect(title_re="^Update Stores Inventory -.*$")
        main_window_1 = app_1.window(title_re="^Update Stores Inventory -.*$")
        main_window_1.print_control_identifiers() """

        """ find_element(
            main_window=main_window,
            variable = "ACCESSORIES",
            automation_id="cboCatDX0", 
            control_type="Edit", 
            action="sendkeys"
            ) """
    else:
        print("Failed")

except Exception as e:
    print(f"Failed: {e}")