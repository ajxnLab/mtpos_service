import time
from config.env_config import get_env_variable
from utils.app_controler import AppAutomation
import re
from utils.helpers import wait
from utils.logger import log_traceback,finalize_log_upload,attach_drive_client,setup_in_memory_logger
import utils.helpers
from promo_code.promocode_constant import PromoCode_Constants
from promo_code.promocode_process import PromoCode_Process
from utils.google_sheet import GSheetClient



# Instantiate constants
promo_code = PromoCode_Constants()

#Setup logger for the MTPOS process
logger,log_stream = setup_in_memory_logger(service_name=promo_code.SERVICE_NAME)

class PromoCode:
    def __init__(self):

        # Google Sheets client
        gsheet_credential = get_env_variable("GOOGLE_SERVICE_ACCOUNT")
        gsheet_id = get_env_variable("GSHEET")

        self.gs = GSheetClient(gsheet_credential, gsheet_id)

        self.sheet_tab_sor = promo_code.WORKSHEET_TAB_SOR 
        self.sheet_tab = promo_code.WORKSHEET_TAB_CREDENTIAL
        self.creds_data = self.gs.get_sheet_data(self.sheet_tab)
        self.creds_row = None

        # Start datetime
        self.start_time = utils.helpers.get_datetime("full")

        # App paths from env
        self.APP_PATH_GT = get_env_variable("APP_PATH_MTPOS_GT")
        self.APP_PATH_PD = get_env_variable("APP_PATH_MTPOS_PD")

        self.proceed_to_publish = None

        self.data = [
        {k.strip(): v for k, v in row.items()}
        for row in self.gs.get_sheet_data(self.sheet_tab_sor)
        ]       

    def run(self):
        logger.info(">>> Starting MTPOS process sequence")
        logger.info(">>> Promo Code Definition")

        
        today_date = utils.helpers.get_datetime("date")
        #logger.info(f"Data : {self.data} ")
        logger.info(f"Retrieved {len(self.data)} rows from sheet '{self.sheet_tab_sor}'")

        filtered_data = []
        gt_rows = []
        pd_rows = []

        for row in self.data:
            deployment_date = str(row.get("Deployment Date", "")).strip()
            if deployment_date != today_date:
                continue  # just skip, don't log here

            filtered_data.append(row)


            app_type = self.normalize_app_type(row.get("App Type"))
            raw_parts = re.split(r"[,/\\]", app_type)
            parts = {p.strip() for p in raw_parts if p.strip()}


            if "GT" in parts:
                gt_rows.append(row)
            if "PD" in parts:
                pd_rows.append(row)


        # Process GT first (GT-only and GT/PD)
        logger.info(">>> Processing GT")
        gt_processed = False
        for filtered_data_gt in gt_rows:
            def_remarks_gt = str(filtered_data_gt.get("RPA Remarks - GT", "")).strip().lower()

            if def_remarks_gt == "success":
                # Both success
                continue
            elif def_remarks_gt != "success":
                # Both blank
                app_path = get_env_variable(f"APP_PATH_MTPOS_GT")
                logger.info(f"Found {len(gt_rows)} rows to process ")
                self.run_app(app_path, promo_code.GT_PROCEDURE, filtered_data_gt)
                gt_processed = True

            else:
                # Other cases
                logger.warning("No matching data found to process today.")
                continue

        
        if not gt_processed:
            logger.info("No matching data found to process today for GT.")
  
        # Then process PD (PD-only and GT/PD again)
        logger.info(">>> Processing PD")
        pd_processed = False
        for filtered_data_pd in pd_rows:
            def_remarks_pd = str(filtered_data_pd.get("RPA Remarks - PD", "")).strip().lower()

            if def_remarks_pd == "success":
                # Both success
                continue
            elif def_remarks_pd == "" and def_remarks_pd != "success":
                # Both blank
                app_path = get_env_variable(f"APP_PATH_MTPOS_PD")
                logger.info(f"Found {len(pd_rows)} rows to process ")
                self.run_app(app_path, promo_code.PD_PROCEDURE, filtered_data_pd)
                pd_processed = True

            else:
                # Other cases
                logger.warning("No matching data found to process today.")
                continue
        
        if not pd_processed:
            logger.info("No matching data found to process today for PD.")

    def normalize_app_type(self, app_type):
        if app_type is None:
            return ""
        return str(app_type).strip().upper()

    def run_app(self, app_path, app_name, filtered_data):
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
            
            self.promocode_process(filtered_data, app_name)

            self.logout()

            logger.info(f">>> {app_name} process done")
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
            attach_drive_client(logger, self.gs ,promo_code, log_stream)
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

    def promocode_process(self, filtered_data, app_name):
    
        t_start = time.time()
        main_window = self.bot.get_window_by_title(title_re=r"^MT-POS ENT.*")
        if main_window:
            self.bot.main_window = main_window
            # Extract only the 'Promo Code' values
            if isinstance(filtered_data, dict):
                filtered_data = [filtered_data]

            promo = [row.get("Promo Code", "Unnamed") for row in filtered_data]

            logger.info(f"Promo Code: {promo}")

            filtered_sorted_data = sorted(
                        filtered_data,
                        key=lambda row: row.get("Procedure", "").lower()  
                    )
            logger.info(f"Redirecting to Promotions Tab") 

            self.bot.wait_until_element_present(name="Management", control_type="TabItem")

            self.bot.find_element(name="Management", control_type="TabItem", action = "click") 
            self.bot.find_element(name="Promotions", control_type="MenuItem", action = "click")
            self.bot.find_element(name="Coupons / Discounts", control_type="Button", action = "click")

            #self.bot.find_element(name="Coupons / Discounts", control_type="Button", action = "click")
            wait(2)
            found = False
            for attempt in range(60): 
                try:
                    new_window = self.bot.get_window_by_title(title_re="^Coupon list -.*$")
                    if new_window.exists(timeout=1):
                        self.bot.main_window = new_window
                        found = True
                        break
                except Exception:
                    raise
                wait(1)

            if not found:
                logger.error("Coupon list window did not appear in time.")
                raise RuntimeError("Coupon list window not found.")
            
            logger.info(f"Switched to Coupon list window in {time.time() - t_start:.2f}s")

            proc = PromoCode_Process(self.bot, self.gs , self.data)

            self.bot.find_element_in_parent(
                parent_control_type="Custom",
                child_control_type="DataItem", 
                parent_name="Filter Row",
                child_name="Promotion row -2147483646", 
                action="right_click"
            )
            
            for row in filtered_sorted_data:
                procedure = row.get("Procedure", None)
                details = row.get("Details", None)

                if procedure:
                        try:
                            if procedure == "create":

                                self.bot.find_element_in_parent(
                                parent_control_type="Menu",
                                child_control_type="MenuItem", 
                                parent_name="DropDown",
                                child_name="New Coupon", 
                                action="click"
                                )

                                logger.info(f"Processing Promo Details {details} {procedure} procedure") 
                                proc.run_create(row, app_name)
                            elif procedure == "update":
                                logger.info(f"Processing Promo Details {details} {procedure} procedure") 
                                proc.run_update_srp(row, app_name)
                            else:
                                logger.warning(f"Unknown Promo Details {details} procedure: {procedure}")
                        except Exception as e:
                            logger.error(f"Error processing Promo Details {details}: {e}")
                            logger.info("Continuing to next Promo Details...")
                else:
                    logger.warning("No 'Procedure' value found in row")

        else:
            logger.error("Main window not found")
            raise RuntimeError("Main window not found")
        
    def logout(self):
        logger.info(">>> Starting logout sequence")

        # Close Coupon list window if present
        try:
            new_window = self.bot.get_window_by_title(title_re="^Coupon list -.*$")
            if new_window:
                logger.info("Found 'Coupon list' window; closing it")
                new_window.close()
                wait(1)
            else:
                logger.info("'Coupon list' window not found; skipping")
        except Exception as e:
            logger.warning(f"Could not close 'Coupon list' window: {e}")

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
                            action="click"
                            )
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




            





        

       


        