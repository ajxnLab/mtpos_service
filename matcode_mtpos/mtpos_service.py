from config.env_config import get_env_variable
from utils.app_controler import AppAutomation
import re
from utils.helpers import wait
from utils.logger import log_traceback,finalize_log_upload,attach_drive_client,setup_in_memory_logger
import utils.helpers
from matcode_mtpos.mtpos_constant import MTPOS_Constants
from matcode_mtpos.mtpos_inventory import MtposInventory
from utils.google_sheet import GSheetClient



# Instantiate constants
mtpos = MTPOS_Constants()

#Setup logger for the MTPOS process
logger,log_stream = setup_in_memory_logger(service_name=mtpos.SERVICE_NAME)

class Mtpos_Service:
    def __init__(self):

        # Google Sheets client
        gsheet_credential = get_env_variable("GOOGLE_SERVICE_ACCOUNT")
        gsheet_id = get_env_variable("GSHEET")

        self.gs = GSheetClient(gsheet_credential, gsheet_id)
        self.sheet_tab_sor = mtpos.WORKSHEET_TAB_SOR 
        self.sheet_tab = mtpos.WORKSHEET_TAB_CREDENTIAL
        self.creds_data = self.gs.get_sheet_data(self.sheet_tab)
        self.data = self.gs.get_sheet_data(self.sheet_tab_sor)
        self.data = self.data_strip(self.data)
        self.creds_row = None

        # Start datetime
        self.start_time = utils.helpers.get_datetime("full")

        # App paths from env
        self.APP_PATH_GT = get_env_variable("APP_PATH_MTPOS_GT")
        self.APP_PATH_PD = get_env_variable("APP_PATH_MTPOS_PD")

        self.proceed_to_publish = None

      

    def run(self):
        logger.info(">>> Starting MTPOS process sequence")
        today_date = utils.helpers.get_datetime("date")

        

        filtered_app_type = sorted(
            mtpos.app_type,
            key=lambda x: x.lower()
        )

        filtered_data = []
        for row in self.data:
            deployment_date = str(row.get("Deployment Date", "")).strip()

            if deployment_date == today_date:
                filtered_data.append(row)

        if filtered_data:

            for row_procedure in filtered_app_type:
                
                

                logger.info(f"Retrieved {len(self.data)} rows from sheet '{self.sheet_tab_sor}'")

                filtered_data_process = []
                success_def_only = []
                filtered_sorted_data = []

                for row in filtered_data:
                    def_remarks = str(row.get(f"RPA Definition Remarks - {row_procedure}", "")).strip().lower()
                    dep_remarks = str(row.get(f"RPA Deployment Remarks - {row_procedure}", "")).strip().lower()

                    if def_remarks == "success" and dep_remarks == "published":
                        # Both success
                        continue
                    elif def_remarks != "success" and dep_remarks != "published":
                        # Both blank
                        filtered_data_process.append(row)
                    elif def_remarks == "success" and dep_remarks != "published":
                        # definition success, deployment blank
                        success_def_only.append(row)
                    else:
                        # Other cases
                        continue

                app_path = get_env_variable(f"APP_PATH_MTPOS_{row_procedure}")
                # After loop, decide what to do
                if filtered_data_process:
                    logger.info(f"Found {len(filtered_data_process)} rows to process (both remarks blank)")
                            # sort alphabetically
                    filtered_sorted_data = sorted(
                        filtered_data_process,
                        key=lambda row: row.get("Procedure", "").lower()  
                    )
                    self.run_app(app_path, row_procedure, filtered_sorted_data,success_def_only)
                elif success_def_only:
                    logger.info(f"Found {len(success_def_only)} rows ready to publish")
                    self.run_app(app_path, row_procedure,filtered_sorted_data, success_def_only)
                    self.proceed_to_publish = 1
                else:
                    # Other cases
                    continue
        else:
            logger.warning("No matching data found to process or publish today.")

    def run_app(self, app_path, app_name, filtered_data,success_def_only):
        """
        Run the given app using AppAutomation and update Google Sheet status.
        """
        try:
            logger.info(f">>> Starting {app_name} process sequence")
            self.bot = AppAutomation(app_path)
            self.bot.start_app()
            wait(3)

            # Login automation (same logic for GT and PD; adjust if needed)
            self.login(app_name)

            self.mtpos_inventory(filtered_data,success_def_only, app_name)

            self.logout()

            logger.info(f">>> {app_name} process succeeded")
            #wait(3)
            #self.bot.close_app()

            # End datetime and log duration
            end_time = utils.helpers.get_datetime("full")
            duration = utils.helpers.duration_time(self.start_time, end_time)
            logger.info(f"Start Time: {self.start_time}")
            logger.info(f"End Time: {end_time}")
            logger.info(f"Duration: {duration}")
           
        except Exception as e:
            logger.error(f">>> {app_name} process failed: {e}")
            log_traceback(logger, e)
        finally:
            attach_drive_client(logger, self.gs ,mtpos, log_stream)
            finalize_log_upload(logger)
            

    def login(self, app_name):

        # Find row where App Type == app_name
        self.creds_row = next(
            (row for row in self.creds_data if row["App Type"] == app_name),
            None
        )

        if self.creds_row:
            username = self.creds_row["Username"]
            password = self.creds_row["Password"]
            
            self.bot.find_element(automation_id="SecurtyWarningOkBut", control_type="Button", action="click")
            element = self.bot.wait_until_element_present(automation_id="mtclink", control_type="Pane")
            if element:
                if app_name == "PD" :
                    self.bot.find_element(variable=get_env_variable("PD"), automation_id = "txtAgentID", control_type="Edit", action="sendkeys")
                else:
                    self.bot.find_element(variable=get_env_variable("GT"), automation_id = "txtAgentID", control_type="Edit", action="sendkeys") 

                self.bot.find_element(variable=username, automation_id = "UserTxt", control_type="Edit", action="sendkeys")

                self.bot.find_element(variable=password, automation_id = "PassTxt", control_type="Edit", action="send_type")

     
                self.bot.find_element(automation_id="cmdOK", control_type="Button", action="click")
            else:
                logger.error("Element not found!")
                raise RuntimeError("Element not found!")
              
            wait(10)
            element = self.bot.find_partial_element(
                partial_name="User last logged-in from terminal",
                control_type="Text"
            )
  
            if element:
                name_text = element.window_text()
                dynamic_parts = re.findall(r'\[(.*?)\]', name_text)
                logger.info(f"User/terminal: {dynamic_parts}")
                self.bot.find_element(name="OK", control_type="Button", action="click")
                wait(3)
                element = self.bot.wait_until_element_present(name="No",control_type="Button")

                if element:
                    logger.info("Password Expiration Notice appeared; clicking 'No'")
                    self.bot.find_element(name="No", control_type="Button", action="click")
                else:
                    logger.info("Password Expiration Notice not shown; proceeding")
            else:
                logger.error("Login failed: could not find confirmation element.")
                raise RuntimeError("Not successfully logged in — please check if your credentials or app path are correct.")
            

            logger.info(">>> Login sequence completed")

        else:
            logger.error(f"[MTPOS] No credentials found  '{app_name}'") 

    def mtpos_inventory(self, filtered_data,success_def_only, app_name):

        # Extract only the 'Name' values
        if isinstance(filtered_data, dict):
            filtered_data = [filtered_data]
        names = [row.get("Material Description", "Unnamed") for row in filtered_data]
        logger.info(f"Material Description: {names}")


        logger.info(f"Redirecting to Inventory Tab") 

        self.bot.wait_until_element_present(name="Inventory", control_type="TabItem")

        self.bot.find_element(name="Inventory", control_type="TabItem", action = "click") 

        self.bot.find_element_in_parent(
            parent_name="Catalog",
            child_name="Inventory",
            child_control_type="Button",
            parent_control_type = "ToolBar",
            action = "click"
        )
        self.bot.wait_until_element_present(automation_id="frmInventoryNew", control_type="Window", retries=5, single_attempt_timeout=60, retry_interval=0)
        # NEW: switch to the new window that pops up
        new_window = self.bot.get_window_by_title(title_re="^Inventory -.*$")

        if new_window:
            self.bot.main_window = new_window
            logger.info("Switched to Inventory window")
            proc = MtposInventory(self.bot, self.gs, self.data)

            if self.proceed_to_publish:
 
                proc.run_publish_to_all(app_name,success_def_only, self.creds_row)

            else:
                for row in filtered_data:
                    procedure = row.get("Procedure", None)
                    material_description = row.get("Material Description", None)

                    if procedure:
                            try:
                                if procedure == "create":
                                    logger.info(f"Processing MTPOS Material {material_description} create procedure") 
                                    proc.run_create(row, app_name)
                                elif procedure == "update-srp":
                                    logger.info(f"Processing MTPOS Material {material_description} {procedure} procedure") 
                                    proc.run_update_srp(row, app_name,procedure)
                                elif procedure == "update-description":
                                    logger.info(f"Processing MTPOS Material {material_description} {procedure} procedure") 
                                    proc.run_update_srp(row, app_name,procedure)
                                else:
                                    logger.warning(f"Unknown MTPOS Material {material_description} procedure: {procedure}")
                            except Exception as e:
                                logger.error(f"Error processing material code {row.get('Material Code')}: {e}")
                                logger.info("Continuing to next material code...")
                    else:
                        logger.warning("No 'Procedure' value found in row")
                 
                filtered_data_publish = (filtered_data or []) + (success_def_only or [])
                logger.info(filtered_data_publish)       
                proc.run_publish_to_all(app_name,filtered_data_publish, self.creds_row)
       
        else:
            logger.error("Inventory window not found")
            raise RuntimeError("Inventory window not found")

    def data_strip(self, data):
        cleaned_data = []
        for row in data:
            cleaned_row = {}
            for k, v in row.items():
                # Handle None keys
                key = k.strip() if isinstance(k, str) else str(k or "").strip()
                # Handle None values
                val = v.strip() if isinstance(v, str) else str(v or "").strip()
                cleaned_row[key] = val
            cleaned_data.append(cleaned_row)
        return cleaned_data    

    def logout(self):
        logger.info(">>> Starting logout sequence")

        # Close Update Stores Inventory window if present
        try:
            new_window = self.bot.get_window_by_title(title_re=r"^Update Stores Inventory -.*$")
            if new_window:
                logger.info("Found 'Update Stores Inventory' window; closing it")
                new_window.close()
                wait(1)
            else:
                logger.info("'Update Stores Inventory' window not found; skipping")
        except Exception as e:
            logger.warning(f"Could not close 'Update Stores Inventory' window: {e}")

        # Close Inventory window if present
        try:
            inventory_window = self.bot.get_window_by_title(title_re=r"^Inventory -.*$")
            if inventory_window:
                logger.info("Found 'Inventory' window; closing it")
                inventory_window.close()
                wait(1)
            else:
                logger.info("'Inventory' window not found; skipping")
        except Exception as e:
            logger.warning(f"Could not close 'Inventory' window: {e}")

        # Switch back to main MT-POS ENT window
        try:
            main_window = self.bot.get_window_by_title(title_re=r"^MT-POS ENT.*")
            if main_window:
                logger.info("Switched back to MT-POS ENT main window")
                self.bot.main_window = main_window

                try:
                    self.bot.find_element_in_parent(
                            parent_control_type="ToolBar",
                            child_control_type="Button", 
                            parent_name="Quick Access Toolbar",
                            child_name="Logout", 
                            action="click",
                            visible_only=True)
                    self.bot.wait_until_element_present(name="Logout", control_type="Window")
                    self.bot.find_element(name="Yes", control_type="Button", action="click")
                    self.bot.wait_until_element_present(automation_id="frmLogin", control_type="Window")

                    # Handle post-logout security/cancel buttons gracefully
                    sec_ok = self.bot.find_element(automation_id="SecurtyWarningOkBut", control_type="Button", action="click")
                    if sec_ok:
                        logger.info("Clicked 'Security Warning OK' after logout")
                    else:
                        logger.info("'Security Warning OK' button not found; skipping")

                    cmd_cancel = self.bot.find_element(automation_id="cmdCancel", control_type="Button", action="click")
                    if cmd_cancel:
                        logger.info("Clicked 'Exit' on login screen after logout")
                    else:
                        logger.info("'Exit' button not found; skipping")
                        raise

                except Exception as e_inner:
                    logger.warning(f"Partial failure during main window logout actions: {e_inner}")
                    raise

            else:
                logger.warning("Could not find MT-POS ENT main window; skipping logout actions")

        except Exception as e:
            logger.error(f"Unexpected error during logout: {e}")
            raise

        logger.info(">>> Logout sequence completed")




            





        

       


        