from pywinauto import Application
import re
import time
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto.keyboard import send_keys
from utils.logger import setup_in_memory_logger,log_traceback
from pywinauto.application import WindowSpecification


#Setup logger for the MTPOS process
logger,log_stream = setup_in_memory_logger(__name__)

class AppAutomation:
    def __init__(self, exe_path: str):
        self.exe_path = exe_path
        self.app = None
        self.main_window = None

    def start_app(self):
        self.app = Application(backend="uia").start(self.exe_path)
        self.main_window = self.app.window(title_re=".*")
        self.main_window.wait('visible')

    def get_window_by_title(self, title=None, title_re=None, auto_id=None):
        try:
            if auto_id:
                window = self.app.window(auto_id=auto_id)
            elif title:
                window = self.app.window(title=title)
            elif title_re:
                window = self.app.window(title_re=title_re)
            else:
                raise ValueError("Must provide title, title_re, or auto_id")

            logger.info(f"Found window: title={title or title_re}, auto_id={auto_id}")
            return window

        except Exception as e:
            logger.error(f"Error finding window {title or title_re or auto_id}: {e}")
            return None

        
    def find_element(
        self,
        control_type: str,
        automation_id: str = None,
        name: str = None,
        variable=None,
        action=None,
        timeout: int = 60,
        retry_interval: int = 1,
        element = None
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
            parent = element if element else self.main_window

            if isinstance(parent, UIAWrapper):
                parent = self.app.window(handle=parent.handle)

            element_spec = parent.child_window(**criteria)

            if element_spec.exists(timeout=timeout, retry_interval=retry_interval):
                
                # Perform optional action
                result = self.perform_action(element=element_spec, variable=variable, action=action)
                logger.info(f"Found element {result}")
                # Return what perform_action returned if any, else the element itself
                return result if result is not None else element_spec
            else:
                logger.error(f"Element not found within timeout. Criteria: {criteria}")
                raise RuntimeError(f"Element not found: {criteria}")

        except Exception as e:
            logger.error(f"Error finding element with criteria {criteria}: {e}")
            raise RuntimeError(f"Failed to find element: {e}") from e


    def wait_until_element_present(self, control_type: str, automation_id: str = None, name: str = None,
                                retries: int = 3, single_attempt_timeout: int = 1, retry_interval: int = 1):
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
                element = self.main_window.child_window(**criteria)
                element.wait('visible', timeout=single_attempt_timeout)
                logger.info(f"Found element on attempt {attempt}")
                return element
            except Exception as e:
                logger.debug(f"Attempt {attempt} failed: {e}")
                if attempt < retries:
                    time.sleep(retry_interval)

        logger.error(f"Element not found after {retries} retries × {single_attempt_timeout}s each")
        return None


    def find_partial_element(self, partial_name, control_type, timeout=180, interval=1):

        end_time = time.time() + timeout

        while time.time() < end_time:
            try:
                elements = self.main_window.descendants(control_type=control_type)
                for elem in elements:
                    name = elem.window_text()
                    if name.startswith(partial_name):
                        logger.info(f"Found element starting with '{partial_name}': {name}")
                        return elem
            except Exception as e:
                logger.error(f"Error while searching for element: {e}")

            time.sleep(interval)

        logger.error(f"No element starting with '{partial_name}' found within {timeout} seconds.")
        return None
    def find_element_with_index(
        self,
        control_type: str,
        automation_id: str = None,
        name: str = None,
        found_index: int = None,
        timeout: int = 60,
        retry_interval: int = 1,
        variable=None,
        action=None,
        search_descendants: bool = True,
    ):
        """
        Finds a UI element by control_type and optional automation_id / name.
        If multiple matches exist, select one by found_index (default = 0).
        """
        try:
            # Collect matching elements
            if search_descendants:
                candidates = self.main_window.descendants(control_type=control_type)
            else:
                candidates = self.main_window.children(control_type=control_type)

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
            result = self.perform_action(element=element, variable=variable, action=action)
            return result if result is not None else element

        except Exception as e:
            logger.error(f"Error in find_element_with_index: {e}")
            raise

        
    def find_element_in_parent(
        self,
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
        interval=1
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
                parent = self.main_window.child_window(
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
                        return self.perform_action(element=elem, variable=variable, action=action)
                    except Exception as e:
                        logger.warning(f"Skipping element due to error: {e}")
                        continue

                time.sleep(interval)

            logger.error(f"No child element found within {timeout} seconds (type={child_control_type}).")
            return None

        except Exception as e:
            logger.error(f"Error in find_element_in_parent: {e}")
            return None

    def perform_action(self, element, action, variable=None):
        if not element:
            logger.error("Element not found")
            raise RuntimeError("Element not found")

        try:
            if isinstance(action, (list, tuple)):
                for act in action:
                    time.sleep(0.5)
                    self.perform_action(element, act, variable)
                return element
            if action == "click":
                element.click_input()
            elif action == "double_click":
                element.double_click_input()
            elif action == "right_click":
                element.right_click_input()
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
                    logger.warning("No variable provided to send_type")
            elif action == "find":
                return element
            elif action == "control_keys":
                self.send_keys_to(element=element ,keys=variable)
            else:
                logger.error(f"Unknown action: {action}")
                raise ValueError(f"Unknown action: {action}")
        except Exception as e:
            logger.error(f"Error performing action '{action}' on element: {e}")
            raise

    def send_keys_to(self, keys: str, element=None):
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
        elif self.main_window:
            try:
                self.main_window.set_focus()
                send_keys(keys)
                logger.info(f"Sent keys '{keys}' to main window")
            except Exception as e:
                logger.error(f"Failed to send keys to main window: {e}")
                raise
        else:
            logger.error("No element or main window to send keys to")
            raise RuntimeError("No target for send_keys")   
    
    def wait_until_ready(element, timeout=180, interval=1):
        start = time.time()
        while time.time() - start < timeout:
            try:
                element.wait("exists ready", timeout=interval)
                return True
            except Exception:
                time.sleep(interval)
        return False
        
    def close_app(self):
        self.app.kill()
