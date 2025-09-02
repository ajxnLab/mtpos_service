import time
from datetime import datetime
from pywinauto.keyboard import send_keys
from utils.helpers import wait, data_strip
from utils.logger import log_traceback,finalize_log_upload,attach_drive_client,setup_in_memory_logger
from promo_code.promocode_constant import PromoCode_Constants

# Instantiate constants
promo_code = PromoCode_Constants()

#Setup logger for the MTPOS process
logger,log_stream = setup_in_memory_logger(service_name=promo_code.SERVICE_NAME)

class PromoCode_Process():
    def __init__(self, bot, gs, data):

        self.bot = bot
        self.gs = gs
        self.data = data
        self.productlistmatrix_data = self.gs.get_sheet_data(promo_code.WORKSHEET_TAB_PRODUCTLISTMATRIX)

    def run_create(self,row, app_name):
        try:
            self.productlistmatrix_data_strip = data_strip(self.productlistmatrix_data)
            self.app_name = app_name
            self.row= row
            
            self.promotion_step_1()

            self.promotion_step_2()

            self.promotion_step_3()

            self.promotion_step_4()

            self.update_gsheet("SOR")


        except Exception as inner_e:
            logger.error(f"Error processing {self.promo_code_value}: {inner_e}")

            # Try marking as Failed if not already done
            try:
                result = self.gs.find_row_index(self.data, promo_code.PROMO_CODE, self.promo_code_value)
                if result:
                    row_index = result["row_index"]
                    self.gs.update_cell(promo_code.WORKSHEET_TAB_SOR, row_index, f"RPA Remarks - {self.app_name}", "Failed")
                    logger.info(f"Updated GSheet row {row_index} with Failed status")
            except Exception as sheet_update_error:
                logger.error(f"Failed to update GSheet for {self.promo_code_value}: {sheet_update_error}")

            # Finally re-raise so outer loop can catch and move to next
            raise
            
    def promotion_step_1(self):
        self.promotion_window()

        self.bot.wait_until_element_present(automation_id="listCoupon", control_type="List", retries=10, single_attempt_timeout=60, retry_interval=0)

        self.promo_code_value = self.row.get("Promo Code")
        self.details = self.row.get("Details")
        self.promo_type = self.row.get("Promo Type")
        self.discount_php = self.row.get("Discount (Php)")
        self.discount_percent = self.row.get("Discount (Percentage)")
        self.date_start = self.row.get("Effective Date: Start")
        self.date_end = self.row.get("Effective Date: End")
        self.participating_stores= self.row.get("Participating Stores")

        discount = None

        promo = self.bot.find_element(name=self.promo_type, control_type="ListItem", action = "click") 

        conditions = {
            "Promo Type": self.promo_type,
            "Details": self.details,
            "Promo Code": self.promo_code_value,
        }

        if self.discount_php:
            discount = self.discount_php
        
        if self.discount_percent:
            discount = self.discount_percent

        self.result_index = self.gs.find_row_index_multi(self.data, conditions)

        if not promo:
            logger.warning(f"Warning: Promo type does not match expected value.")
            raise ValueError("Promo type element not found")

        self.bot.find_element(automation_id="txtZ", control_type="Edit", action = "sendkeys", variable = discount ) 
        self.bot.find_element(name="&Next >", control_type="Button", action = "click") 

    def promotion_step_2(self):
            
            step_2=self.bot.wait_until_element_present(automation_id="pgGeneral", control_type="Pane", retries=2, single_attempt_timeout=2, retry_interval=0)

            if not step_2:
                logger.warning("Unable to proceed to next step 2, verify your input.")
                raise ValueError("Unable to proceed to next step")
            
            for auto_id, variable in [
                 ("txtCode" , self.promo_code_value),
                 ("txtDesc" , self.details),
                 ("cmbExclusivity" , promo_code.NON_EXCLUSIVE)                                                   

            ]:
                self.bot.find_element(automation_id=auto_id, control_type="Edit", action = "sendkeys", variable = variable )  

            duplicate_promo_code = self.bot.wait_until_element_present(automation_id="picCodeChecker", control_type="Pane", retries=2, single_attempt_timeout=1, retry_interval=0) 
            if duplicate_promo_code:
                logger.warning(f"Duplicate coupon code. Proceeding to next...")
                self.bot.find_element_in_parent(
                        parent_control_type="Pane",
                        child_control_type="Button", 
                        parent_automation_id="WizardControl1",
                        child_name="Cancel", 
                        action="click"
                        )
                raise ValueError("Duplicate coupon code") 

            fmt = "%#m/%#d/%Y %I:%M %p"
            formatted_am = datetime.strptime(self.date_start, "%m/%d/%Y").replace(hour=0, minute=0).strftime(fmt)
            formatted_pm = datetime.strptime(self.date_end, "%m/%d/%Y").replace(hour=23, minute=59).strftime(fmt)

            self.bot.find_element(automation_id="txtEffStartDate", control_type="Pane", action="send_type", variable=formatted_am)
            self.bot.find_element(automation_id="txtEffEndDate", control_type="Pane", action="send_type", variable=formatted_pm)
            self.bot.find_element(name="&Next >", control_type="Button", action = "click") 

    def promotion_step_3(self):
            
            step_3=self.bot.wait_until_element_present(automation_id="pgEligibility", control_type="Pane", retries=2, single_attempt_timeout=2, retry_interval=0)

            if not step_3:
                logger.warning("Unable to proceed to next step 3, verify your input.")
                raise ValueError("Unable to proceed to next step")
            
            
            param_matrix = self.filter_data(self.productlistmatrix_data_strip, "PROMO CODE", self.promo_code_value)
            
            logger.info(f"param_matrix: {len(param_matrix)}")

            if not param_matrix:
                logger.warning("No matching Promo Code in Product List Matrix, verify your input.")
                raise ValueError("Unable to proceed to next step")            
            
            product_matrix_sku = self.filter_data(param_matrix, "SKU/Category", "SKU")
            logger.info(f"product_matrix_sku: {len(product_matrix_sku)}")

            product_matrix_sku_include = self.filter_data(product_matrix_sku, "Include/Exclude", "Include")
            product_matrix_sku_exclude = self.filter_data(product_matrix_sku, "Include/Exclude", "Exclude")

            if product_matrix_sku_include:
                # Handle include SKUs
                for auto_child_id, auto_parent_id, control_type_child, control_type_parent in [
                    ("btnSkus", "pnlSkus", "Button", "Group"),
                    ("btnModify", "PnlBtm", "Button", "Pane")
                ]:
                    self.bot.find_element_in_parent(
                        parent_control_type=control_type_parent,
                        child_control_type=control_type_child,
                        parent_automation_id=auto_parent_id,
                        child_automation_id=auto_child_id,
                        action="click"
                    )

                include = "Included"
                self.add_sku_items(product_matrix_sku_include, include)
            else:
                logger.warning("No Include SKUs found to process.")

            if product_matrix_sku_exclude:
                # Handle exclude SKUs
                for auto_child_id, auto_parent_id, control_type_child, control_type_parent in [
                    ("btnExceptionalSkus", "pnlSkus", "Button", "Group"),
                    ("btnModify", "PnlBtm", "Button", "Pane")
                ]:
                    self.bot.find_element_in_parent(
                        parent_control_type=control_type_parent,
                        child_control_type=control_type_child,
                        parent_automation_id=auto_parent_id,
                        child_automation_id=auto_child_id,
                        action="click"
                    )
                exclude = "Excluded"
                self.add_sku_items(product_matrix_sku_exclude, exclude)

            else:
                logger.warning("No Exclude SKUs found to process.")

            
            self.promotion_window()

            self.bot.find_element_in_parent(
                parent_control_type="Pane",
                parent_automation_id="WizardControl1",
                child_control_type="Button",
                child_name="&Next >",
                action="click"
            )
    def promotion_step_4(self):
        promotion_window=self.bot.wait_until_element_present(automation_id="WizardControl1", control_type="Pane", retries=3, single_attempt_timeout=60, retry_interval=0)
        result_store=self.parse_store_data(self.participating_stores)

        if result_store:
            if self.participating_stores == "All Stores":
                    
                    self.bot.find_element_in_parent(
                    child_control_type="Button",
                    child_automation_id="btnSelectAll",
                    action="click",
                    element = promotion_window,
                    search_descendants = True
                    )
        
            else:

                custom_data_panel = self.bot.find_element_in_parent(
                child_control_type="Custom",
                child_name="Data Panel",
                action="find",
                element = promotion_window,
                search_descendants = True
                )
                for store_id, store_name in result_store:

                    self.bot.find_element_in_parent(
                    child_control_type="DataItem",
                    child_name="Store ID row -2147483646",
                    action=["click", "send_type"],
                    element = custom_data_panel,
                    search_descendants = True,
                    variable = store_id
                    )

                    element = self.bot.find_element_in_parent(
                        parent_control_type="Custom",
                        child_control_type="DataItem", 
                        parent_name="Row 1",
                        child_name="checkbox row 0", 
                        action="find"
                    )
                    
                    if element:
                        self.bot.find_element_in_parent(
                        child_control_type="DataItem",
                        child_name="checkbox row 0",
                        action="click",
                        element = custom_data_panel,
                        search_descendants = True
                        )
                    
                    else:
                        self.clear_all("Store ID row -2147483646")

                        self.bot.find_element_in_parent(
                        child_control_type="DataItem",
                        child_name="Store Name row -2147483646",
                        action=["click", "send_type"],
                        element = custom_data_panel,
                        search_descendants = True,
                        variable = store_name
                        )


            self.bot.find_element_in_parent(
            child_control_type="Button",
            child_name="&Finish",
            action="click",
            element = promotion_window
            )

        else:
            logger.error("Data processing failed: please check the input")
            raise RuntimeError("Data processing failed: please check the input")

    def filter_data(self, data, column_name, value):
        return [row for row in data if row.get(column_name) == value]
        
    def add_sku_items(self,sku_include_exclude, process):
        promotion_window=self.bot.wait_until_element_present(automation_id="frmSkuSelector", control_type="Window", retries=3, single_attempt_timeout=60, retry_interval=0)
        custom_data_panel = self.bot.wait_until_element_present(automation_id="GroupControl1", control_type="Pane", retries=3, single_attempt_timeout=60, retry_interval=0)
        if promotion_window:
            if custom_data_panel:

                custom_filter_row = self.bot.find_element_in_parent(
                            child_control_type="Custom",
                            child_name="Filter Row",
                            action="find",
                            element = custom_data_panel,
                            search_descendants = True
                        )
            
            logger.info(f"Processing SKUs {process} Items ")
            for row in sku_include_exclude:
                matcode = row.get("Matcodes")
                desc = row.get("Product Description")
                promo_code_matrix = row.get("PROMO CODE") 

                conditions_matrix = {
                "PROMO CODE": promo_code_matrix,
                "Product Description": desc,
                "Matcodes": matcode,
                }

                logger.info(f"Conditions Matrix : {conditions_matrix}")

                self.result_index_matrix = self.gs.find_row_index_multi(self.productlistmatrix_data_strip, conditions_matrix)
                logger.info(f"Conditions : {self.productlistmatrix_data_strip}")
                logger.info(f"Conditions result : {self.result_index_matrix}")

                self.bot.find_element_in_parent(
                    child_control_type="DataItem",
                    child_name="Item ID row -2147483646",
                    action=["click", "send_type"],
                    element = custom_filter_row,
                    search_descendants = True,
                    variable = matcode
                )


                target_item_value = self.bot.find_element_in_parent(
                    child_control_type="Custom",
                    child_name="Row 1",
                    action="find",
                    element = custom_data_panel,
                    search_descendants = True
                )
                        
                target_item_value_click = self.bot.find_element_in_parent(
                        child_control_type="DataItem",
                        child_name="Item ID row 0",
                        action="find",
                        element = target_item_value
       
                    )
                
                self.bot.perform_action(target_item_value_click,"click")
                self.bot.find_element(automation_id="btnadd", control_type="Button", action = "click") 

                logger.info(f"SKUs {process} Item code {matcode} successfully added")
                matrix = "MATRIX"
                self.update_gsheet(matrix)


            self.bot.find_element_in_parent(
                child_control_type="Button",
                child_name="Close",
                action="click",
                element = promotion_window,
                search_descendants = True
            )

    def add_cat_items(self,sku_include_exclude, process):
        pass
           
    def promotion_window(self):
            t_start = time.time()
            found = False
            for attempt in range(60): 
                try:
                    new_window = self.bot.get_window_by_title(title_re="^Promotion Wizard -.*$")
                    if new_window.exists(timeout=1):
                        self.bot.main_window = new_window
                        found = True
                        break
                except Exception:
                    pass
                wait(1)

            if not found:
                    logger.error("Promotion Wizard window did not appear in time.")
                    raise RuntimeError("Promotion Wizard window not found.")
                
            logger.info(f"Switched to Promotion Wizard window in {time.time() - t_start:.2f}s")
            
    def parse_store_data(self,data: str):

        # Remove any trailing commas
        data = data.rstrip(',')

        # Split by comma to get individual stores
        stores = data.split(',')

        result = []
        for store in stores:
            # Split only on the first " - " to separate store ID from name
            parts = store.split(' - ', 1)
            if len(parts) == 2:
                store_id = parts[0].strip()
                store_name = parts[1].strip()
                result.append((store_id, store_name))

        return result

    def clear_all(self,child_name):
        self.bot.find_element_in_parent(
        parent_control_type="Custom",
        child_control_type="DataItem", 
        parent_name="Filter Row",
        child_name=child_name, 
        action="double_click"
        )
        wait(0.3)
        send_keys('^+a{BACKSPACE}')

    def update_gsheet(self, worksheet):
        logger.info(f"SKUs  Item code successfully added")
        if worksheet == "SOR":
            if self.result_index:
                self.gs.update_cell(promo_code.WORKSHEET_TAB_SOR, self.result_index, f"RPA Remarks - {self.app_name}" , "Success")
                logger.info(f"Updated GSheet row {self.result_index} with Success status")
            else:
                logger.warning(f"Promo Code {self.promo_code_value} not found in filtered data for GSheet update")
                return
        else:
            logger.info(f"SKUs  Item code successfully added")
            if self.result_index_matrix:
                logger.info(f"SKUs  Item code successfully added")
                self.gs.update_cell(promo_code.WORKSHEET_TAB_PRODUCTLISTMATRIX, self.result_index_matrix, f"RPA Remarks - {self.app_name}", "Success")
                logger.info(f"Updated GSheet row {self.result_index_matrix} with Success status")
            else:
                logger.warning(f"Promo Code {self.promo_code_value} not found in filtered data for GSheet update")
                return

            
          