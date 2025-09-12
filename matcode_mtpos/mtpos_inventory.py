import time
from pywinauto.keyboard import send_keys
from utils.logger import setup_in_memory_logger
from utils.helpers import wait,data_strip
import utils.helpers
from matcode_mtpos.mtpos_constant import MTPOS_Constants


# Instantiate constants
mtpos = MTPOS_Constants()

#Setup logger for the MTPOS process
logger,log_stream = setup_in_memory_logger("MTPOS")

class MtposInventory:

    def __init__(self, bot, gs, data, all_items,pending_updates):

        self.bot = bot
        self.gs = gs
        self.data = data
        self.all_items = all_items
        self.pending_updates = pending_updates


    def run_create(self,row, app_name):

        reqSNEntry = row.get("ReqSNEntry")
        material_code = row.get("Material Code")
        category = row.get("Category")
        subcategory = row.get("SubCategory")
        retail_price = row.get("Retail Price")
        mat_desc=row.get("Material Description")
        deploymen_date=row.get("Deployment Date")
        procedure=row.get("Procedure")
        

        try:

            conditions = {
                "Material Code": material_code,
                "Material Description": mat_desc,
                "Deployment Date": deploymen_date,
                "Procedure": procedure,
            }
            self.result_index = self.gs.find_row_index_multi(self.data, conditions)

            logger.info(f"Index {self.result_index} material code {mat_desc} procedure {procedure} delopyment date {deploymen_date}")

            logger.info(f"Adding MTPOS Material code {material_code}")
            self.bot.find_element(name="Catalog Def", control_type="TabItem", action="click")
            parent_element_add = self.bot.wait_until_element_present(automation_id="PanelRight", control_type="Pane", retries=6, single_attempt_timeout=30, retry_interval=1)
            #Add item
            #self.bot.find_element(automation_id="cmdAdd", control_type="Button", action="click")
            self.bot.find_element_in_parent(
                child_control_type="Button",
                child_automation_id="cmdAdd",
                action="click",
                element = parent_element_add,
                search_descendants = True 
            )
            #element = self.bot.wait_until_element_present(name="Microtelecom", control_type="Window", retries=1, single_attempt_timeout=2, retry_interval=0)
            wait(1)
            send_keys('{ESC}')
            
            element_add =  self.bot.find_element_in_parent(
                child_control_type="Tab",
                child_automation_id="xtbCtrlRight",
                action="find",
                element = parent_element_add,
                search_descendants = True 
            )
            
            for variable, found_index in [
            (category, 21),
            (subcategory, 19),
            (mtpos.SUPPLIER, 23)
               ]:        
                element_1 = self.bot.find_element_with_index(
                    control_type="Edit",
                    action="find",
                    found_index=found_index,
                    search_descendants=True 
                )

                self.bot.find_element_in_parent(
                    child_control_type="Button",
                    child_name="Open", 
                    element=element_1,
                    search_descendants=True,
                    action="click"
                )

                wait(2)
                element = self.bot.find_element_with_index(
                    control_type="Window",
                    action="find",
                    found_index=0,
                    search_descendants=True 
                )
                self.bot.find_element_in_parent(
                    child_control_type="Custom",
                    child_name="Filter Row", 
                    element=element,
                    search_descendants=True,
                    action="send_type",
                    variable=variable
                )
                self.bot.send_keys_to("{ENTER}", element=element)

            for auto_id, variable in [
                ("itemID", material_code),
                ("FaceValue", retail_price),
                ("Retail", retail_price),
                ("ItemDesc", mat_desc)
            ]:
                   # self.bot.find_element(variable = variable,automation_id=auto_id, control_type="Edit", action="sendkeys")

                    self.bot.find_element_in_parent(
                    child_control_type="Edit",
                    child_automation_id=auto_id, 
                    element=element_add,
                    search_descendants=True,
                    action="sendkeys",
                    variable=variable
                )

            if reqSNEntry == "Y":
                
                self.bot.find_element_in_parent(
                    child_control_type="CheckBox",
                    child_automation_id="ckReqSNEntry", 
                    element=element_add,
                    action="click",
                    search_descendants=True,
                    timeout=2,
                    interval=1
                )
                
                #self.bot.find_element(automation_id="ckReqSNEntry", control_type="CheckBox", action="click")

            #self.bot.find_element(variable = "0",automation_id="SNLength", control_type="Edit", action="sendkeys")
            self.bot.find_element_in_parent(
                    child_control_type="Edit",
                    child_automation_id="SNLength", 
                    element=element_add,
                    action="sendkeys",
                    search_descendants=True,
                    variable = "0"
                )
            self.bot.find_element_in_parent(
                    child_control_type="Button",
                    child_automation_id="cmdUpdate", 
                    element=parent_element_add,
                    search_descendants=True,
                    action="click"
                )
            #self.bot.find_element(automation_id="cmdUpdate", control_type="Button", action="click")
            
            element = self.bot.wait_until_element_present(name="MT.Main.v5", control_type="Window",single_attempt_timeout = 1,retries= 1)
            if element:
                logger.warning(f"Item {material_code} is already defined in location 60001 inventory catalog.")

                # Step 2: Update Google Sheet with "Failed"
                #result = self.gs.find_row_index(self.data, identifier_column, material_code)
                if self.result_index:
                    #row_index = result["row_index"]
                    self.gs.update_cell(mtpos.WORKSHEET_TAB_SOR, self.result_index, f"RPA Definition Remarks - {app_name}", "Failed")
                    self.gs.update_cell(mtpos.WORKSHEET_TAB_SOR, self.result_index, f"RPA Deployment Remarks - {app_name}", "Publish Failed")
                    logger.info(f"Updated GSheet row {self.result_index} with Failed status")
                else:
                    logger.warning(f"Material code {material_code} not found in filtered data for GSheet update")

                #Mark failure and raise to proceed to next
                raise Exception(f"Material code {material_code} is already defined; skipping to next")
            logger.info(f"Successfully saved MTPOS material code {material_code}. Moving to Options tab.")
            #self.bot.find_element(name="Options", control_type="TabItem", action="click")
            self.bot.find_element_in_parent(
                    child_control_type="TabItem",
                    child_name="Options", 
                    element=element_add,
                    action="click",
                    search_descendants=True
                )
            
            self.bot.find_element_in_parent(
                    child_control_type="Button",
                    child_automation_id="cmdEdit", 
                    element=parent_element_add,
                    action="click",
                    search_descendants=True
                )

            #self.bot.find_element(automation_id="cmdEdit", control_type="Button", action="click")

            list_element = self.bot.find_element_in_parent(
                child_control_type="List",
                child_automation_id="OptionLV", 
                element=element_add,
                action="find",
                search_descendants=True
            )

            for name in [
                ("Right Trim S/N to specified length"),
                ("Retail Price Include the Tax"),
                ("Require S/N on Sale"),
            ]:
                    pane = self.bot.find_element(
                        control_type="ListItem",
                        name = name,
                        action="click",
                        element=list_element 
                    )
                    #pane = self.bot.find_element(name=name, control_type="ListItem", action="click")
                    pane.set_focus()
                    send_keys('{SPACE}')

                    wait(1)
                    send_keys('{ESC}')

            self.bot.find_element(automation_id="cmdUpdate", control_type="Button", action="click")
            logger.info(f"MTPOS material code {material_code} added successfully. Proceeding to next matcode.")

            #Update Google Sheet with "Success"
            #result = self.gs.find_row_index(self.data, identifier_column, material_code)
            if self.result_index:
                #row_index = result["row_index"]
                self.gs.update_cell(mtpos.WORKSHEET_TAB_SOR, self.result_index, f"RPA Definition Remarks - {app_name}", "Success")
                logger.info(f"Updated GSheet row {self.result_index} with Success status")
            else:
                logger.warning(f"Material code {material_code} not found in filtered data for GSheet update")

        except Exception as inner_e:
            logger.error(f"Error processing {material_code}: {inner_e}")

            # Try marking as Failed if not already done
            try:
                #result = self.gs.find_row_index(self.data, identifier_column, material_code)
                if self.result_index:
                    #row_index = result["row_index"]
                    self.gs.update_cell(mtpos.WORKSHEET_TAB_SOR, self.result_index, f"RPA Definition Remarks - {app_name}", "Failed")
                    self.gs.update_cell(mtpos.WORKSHEET_TAB_SOR, self.result_index, f"RPA Deployment Remarks - {app_name}", "Publish Failed")
                    logger.info(f"Updated GSheet row {self.result_index} with Failed status")
            except Exception as sheet_update_error:
                logger.error(f"Failed to update GSheet for {material_code}: {sheet_update_error}")

            # Finally re-raise so outer loop can catch and move to next
            raise
    

    def run_update_srp(self, row, app_name,procedure):
        material_code = row.get("Material Code")
        retail_price = row.get("Retail Price")
        description = row.get("Material Description")
        deploymen_date=row.get("Deployment Date")
        procedure_column=row.get("Procedure")

        

        try:
            logger.info(f"Searching MTPOS Material code {material_code}")

            
            conditions = {
                "Material Code": material_code,
                "Material Description": description,
                "Deployment Date": deploymen_date,
                "Procedure": procedure_column,
            }
            self.result_index_update = self.gs.find_row_index_multi(self.data, conditions)

            logger.info(f"Index {self.result_index_update} material code {material_code} procedure {procedure_column} delopyment date {deploymen_date}")

             #Add item
            parent_element = self.bot.wait_until_element_present(automation_id="Frame1", control_type="Pane", retries=6, single_attempt_timeout=30, retry_interval=1)
            self.bot.find_element(name="Catalog Def", control_type="TabItem", action="click")
            send_keys('{ESC}')
            
            # Search for item
            if not self.all_items:
                self.bot.find_element_in_parent(
                                child_control_type="Edit",
                                child_automation_id="cboSearchIn2",
                                action="send_type",
                                element = parent_element,
                                search_descendants = True,
                                variable="All Items"
                            )
                self.all_items = True

            #self.bot.find_element(variable="All Items", control_type="Edit", automation_id="cboSearchIn2", action="send_type")
            self.bot.find_element_in_parent(
                            child_control_type="Edit",
                            child_automation_id="teFind",
                            action="sendkeys",
                            element = parent_element,
                            search_descendants = True,
                            variable=material_code
                        )
            self.bot.find_element_in_parent(
                            child_control_type="Button",
                            child_automation_id="cmdRefreshInv",
                            action="click",
                            element = parent_element
                        )
            #self.bot.perform_action(parent_element, "send_type", "All Items")
            #self.bot.find_element(variable=material_code, control_type="Edit", automation_id="teFind", action="sendkeys")
            #self.bot.find_element(automation_id="cmdRefreshInv", control_type="Button", action="click")

            element = self.bot.find_element_in_parent(
                            child_control_type="Custom",
                            child_name="Row 1",
                            action="find",
                            element = parent_element,
                            search_descendants = True,
                            timeout = 5,
                            interval = 1 
                        )

            #element = self.bot.wait_until_element_present(name="Row 1", control_type="Custom", retries=3, single_attempt_timeout=60, retry_interval=0)
            if not element:
                logger.warning(f"MTPOS Material code {material_code} not found in UI")

                # Step 2: Update Google Sheet with "Failed"
                #result = self.gs.find_row_index(self.data, identifier_column, material_code)
                if self.result_index_update:
                    #row_index = result["row_index"]

                    self.gs.update_cell(mtpos.WORKSHEET_TAB_SOR, self.result_index_update, f"RPA Definition Remarks - {app_name}", "Failed")
                    self.gs.update_cell(mtpos.WORKSHEET_TAB_SOR, self.result_index_update, f"RPA Deployment Remarks - {app_name}", "Publish Failed")
                    logger.info(f"Updated GSheet row {self.result_index_update} with Failed status")
                else:
                    logger.warning(f"Material code {material_code} not found in filtered data for GSheet update")

                #Mark failure and raise to proceed to next
                raise Exception(f"Material code {material_code} not found; skipping to next")

            logger.info(f"MTPOS Material code {material_code} found; editing MSRP and Retail Price")

            #Edit item
            parent_element_edit = self.bot.wait_until_element_present(automation_id="PanelRight", control_type="Pane", retries=6, single_attempt_timeout=30, retry_interval=1)

            self.bot.find_element_in_parent(
                            child_control_type="Button",
                            child_automation_id="cmdEdit",
                            action="click",
                            element = parent_element_edit,
                            search_descendants = True 
                        )
            send_keys('{ESC}')
            #self.bot.find_element(automation_id="cmdEdit", control_type="Button", action="click")
            parent_element_update=self.bot.find_element_in_parent(
                child_control_type="Pane",
                child_automation_id="_DataFrame_0",
                action="find",
                element = parent_element_edit,
                search_descendants = True 
            )
            if procedure == "update-srp":

                #self.bot.wait_until_element_present(automation_id="FaceValue", control_type="Edit")

                self.bot.find_element(variable=retail_price, control_type="Edit", element=parent_element_update, automation_id="FaceValue", action="sendkeys")
                self.bot.find_element(variable=retail_price, control_type="Edit",  element=parent_element_update,automation_id="Retail", action="sendkeys")

                self.bot.find_element_in_parent(
                    child_control_type="Button",
                    child_automation_id="cmdUpdate",
                    action="click",
                    element = parent_element_edit,
                    search_descendants = True 
                )

                #self.bot.find_element(automation_id="cmdUpdate", control_type="Button", element=parent_element_update, action="click")
            else:

                self.bot.find_element(variable=description, control_type="Edit", automation_id="ItemDesc", element=parent_element_update, action="sendkeys")

                self.bot.find_element_in_parent(
                    child_control_type="Button",
                    child_automation_id="cmdUpdate",
                    action="click",
                    element = parent_element_edit,
                    search_descendants = True 
                )

            logger.info(f"MTPOS Material code {material_code} successfully updated")

            #Update Google Sheet with "Success"
            #result = self.gs.find_row_index(self.data, identifier_column, material_code)
            if self.result_index_update:
               # row_index = result["row_index"]

                self.gs.update_cell(mtpos.WORKSHEET_TAB_SOR, self.result_index_update, f"RPA Definition Remarks - {app_name}", "Success")
                logger.info(f"Updated GSheet row {self.result_index_update} with Success status")
            else:
                logger.warning(f"Material code {material_code} not found in filtered data for GSheet update")

        except Exception as inner_e:
            logger.error(f"Error processing {material_code}: {inner_e}")

            # Try marking as Failed if not already done
            try:
                #result = self.gs.find_row_index(self.data, identifier_column, material_code)
                if self.result_index_update:
                    #row_index = result["row_index"]

                    self.gs.update_cell(mtpos.WORKSHEET_TAB_SOR, self.result_index_update, f"RPA Definition Remarks - {app_name}", "Failed")
                    self.gs.update_cell(mtpos.WORKSHEET_TAB_SOR, self.result_index_update, f"RPA Deployment Remarks - {app_name}", "Publish Failed")
                    logger.info(f"Updated GSheet row {self.result_index_update} with Failed status")
            except Exception as sheet_update_error:
                logger.error(f"Failed to update GSheet for {material_code}: {sheet_update_error}")

            # Finally re-raise so outer loop can catch and move to next
            raise
    
    def run_publish_to_all(self, app_name, filtered_sorted_data, creds_row):
        
        username = creds_row["Username"]
        password = creds_row["Password"]
        identifier_column = mtpos.MATERIAL_CODE

        sheet_tab_sor = mtpos.WORKSHEET_TAB_SOR 
        data = self.gs.get_sheet_data(sheet_tab_sor)
        data = data_strip(data)
        today_date = utils.helpers.get_datetime("date")

        data_success = []

        for row in data:
            def_remarks = str(row.get(f"RPA Definition Remarks - {app_name}", "")).strip().lower()
            deployment_date = str(row.get("Deployment Date", "")).strip()

            if def_remarks == "success" and deployment_date == today_date:
                data_success.append(row)

        logger.info(f"Total successful definitions: {len(data_success)}")

        try:
            logger.info("Starting publish_to_all process")
            t_start = time.time()

            self.bot.main_window.set_focus()
            send_keys('^m')

            found = False
            for attempt in range(60): 
                try:
                    new_window = self.bot.get_window_by_title(title_re="^Update Stores Inventory -.*$")
                    if new_window.exists(timeout=1):
                        self.bot.main_window = new_window
                        found = True
                        break
                except Exception:
                    pass
                wait(1)

            if not found:
                logger.error("Update Stores Inventory window did not appear in time.")
                raise RuntimeError("Update Stores Inventory window not found.")
            
            logger.info(f"Switched to new window in {time.time() - t_start:.2f}s")
            

            parent_element_publish = self.bot.wait_until_element_present(automation_id="SplitContainer1", control_type="Pane", retries=6, single_attempt_timeout=30, retry_interval=1)
            #self.bot.find_element(automation_id="SplitContainer1", control_type="Pane", action="find")

            for auto_id in ("dtFromDate", "dtToDate"):
                found = False
                for attempt in range(1, 4):  # try 3 times
                    try:
                        t0 = time.time()
                        pane=self.bot.find_element_in_parent(
                            child_control_type="Pane",
                            child_automation_id=auto_id,
                            action="find",
                            element = parent_element_publish,
                            search_descendants = True
                        )
                        #pane = self.bot.find_element(automation_id=auto_id, control_type="Pane",action="find")
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
          
            # Click Search 
            self.bot.find_element_in_parent(
                child_control_type="Button",
                child_automation_id="cmdSearch",
                action="click",
                element = parent_element_publish,
                search_descendants = True
            )
            #self.bot.find_element(automation_id="cmdSearch", control_type="Button", action="click")

            wait(3)
            table= self.bot.find_element_in_parent(
                child_control_type="Table",
                child_automation_id="GCInv",
                action="find",
                element = parent_element_publish,
                search_descendants = True
            )
            logger.info("Proceeding to select only defined matcodes")
            for row in data_success:
                matcode = row.get("Material Code")
                desc = row.get("Material Description").strip()

                custom_filter_row = self.bot.find_element_in_parent(
                    child_control_type="Custom",
                    child_name="Filter Row",
                    action="find",
                    element = table,
                    search_descendants = True
                )

                self.bot.find_element_in_parent(
                        child_control_type="DataItem",
                        child_name="Item Code row -2147483646",
                        action=["click", "send_type"],
                        element = custom_filter_row,
                        search_descendants = True,
                        variable = matcode
                )

                target_item_value = self.bot.find_element_in_parent(
                    child_control_type="Custom",
                    child_name="Row 1",
                    action="find",
                    element = table,
                    search_descendants = True,
                    timeout=5,
                    interval = 1
                )

                if target_item_value:
                        
                    self.bot.find_element_in_parent(
                            child_control_type="DataItem",
                            child_name="Select row 0",
                            action="click",
                            element = target_item_value
        
                        )
                    
                    self.clear_all("Item Code row -2147483646",custom_filter_row)
                    continue
                    
                self.clear_all("Item Code row -2147483646", custom_filter_row)

                self.bot.find_element_in_parent(
                        child_control_type="DataItem",
                        child_name="Description row -2147483646",
                        action=["click", "send_type"],
                        element = custom_filter_row,
                        search_descendants = True,
                        variable = desc
                )

                target_item_value_des = self.bot.find_element_in_parent(
                    child_control_type="Custom",
                    child_name="Row 1",
                    action="find",
                    element = table,
                    search_descendants = True,
                    timeout=5,
                    interval = 1
                )

                if target_item_value_des:
                        
                    self.bot.find_element_in_parent(
                            child_control_type="DataItem",
                            child_name="Select row 0",
                            action="click",
                            element = target_item_value_des
        
                        )
                    
                    self.clear_all("Description row -2147483646",custom_filter_row)
                    continue
                    
                self.clear_all("Description row -2147483646",custom_filter_row)
                #self.bot.perform_action(target_item_value_click,"click")

            right_panel= self.bot.find_element_in_parent(
                child_control_type="Pane",
                child_automation_id="Picture2",
                action="find",
                element = parent_element_publish,
                search_descendants = True
            )  

            # Click Open button and ALL ListItem inside ComboBox
            for name, child_control_type in [
                ("Open", "Button"),
                ("ALL", "ListItem")
            ]:
                self.bot.find_element_in_parent(
                    child_control_type=child_control_type,
                    child_name=name,
                    action="click",
                    element = right_panel,
                    search_descendants = True
                )

            logger.info("Selected 'ALL' from Market dropdown")

            # Click Select All Locations and Copy Inventory
            for auto_id, control_type in [
                ("ckSelectAllLocations", "CheckBox"),
                ("cmdCopyInventory", "Button")
            ]:
                self.bot.find_element_in_parent(
                    child_control_type=control_type,
                    child_automation_id=auto_id,
                    action="click",
                    element = right_panel,
                    search_descendants = True
                )
                #self.bot.find_element(automation_id=auto_id, control_type=control_type, action="click")
            logger.info("Clicked Select All Locations & Copy Inventory")

            # Handle user verification popup quickly
            pass_check = self.bot.wait_until_element_present(automation_id="frmPassCheck", control_type="Window", retries=6, single_attempt_timeout=30, retry_interval=1)
            #pass_check = self.bot.find_element(automation_id="frmPassCheck", control_type="Window", action="find")
            if pass_check:
                logger.info("Found user verification prompt")

                for variable , auto_id in [
                    (username,"txtUID"),
                    (password,"PassTxt")

                ]:
                    self.bot.find_element_in_parent(
                        child_control_type="Edit",
                        child_automation_id=auto_id,
                        action="sendkeys",
                        element = pass_check,
                        variable=variable,
                        search_descendants = True
                    )
                #self.bot.find_element(variable=username, automation_id="txtUID", control_type="Edit", action="sendkeys")
                #self.bot.find_element(variable=password, automation_id="PassTxt", control_type="Edit", action="send_type")
                self.bot.find_element_in_parent(
                        child_control_type="Button",
                        child_name="C&ontinue",
                        action="click",
                        element = pass_check,
                        search_descendants = True
                )
                #self.bot.find_element(name="C&ontinue", control_type="Button", action="click")
            else:
                logger.error("User verification window not found. Failed to publish.")
                raise RuntimeError("User verification window not found")

            # Confirm No Items Selected → Yes
            confirmation = self.bot.wait_until_element_present(name="No Items Selected", control_type="Window", retries=6, single_attempt_timeout=30, retry_interval=1)

            if confirmation:
                confirmation.set_focus()
                send_keys('{ENTER}')
                #self.bot.find_element(name="Yes", control_type="Button", action="click",element = confirmation)
                logger.info("Clicked confirmation Yes")
            
            else:
                logger.error("Confirmation window not found. Failed to publish.")
                raise RuntimeError("Confirmation not found")

            # Wait for export result
            export_res= self.bot.wait_until_element_present(name="Microtelecom", control_type="Window", retries=20, single_attempt_timeout=30, retry_interval=1)
            if export_res:
                export_res.set_focus()
                send_keys('{ENTER}')
                #self.bot.find_element(name="OK", control_type="Button", action="click",element = export_res)
                logger.info("Export completed successfully")

                for row in data_success:
                    material_code = row.get("Material Code")
                    try:
                        result = self.gs.find_row_index(self.data, identifier_column, material_code)
                        if result:
                            row_index = result["row_index"]

                            self.pending_updates.append((
                                mtpos.WORKSHEET_TAB_SOR,   
                                row_index,
                                f"RPA Deployment Remarks - {app_name}",
                                "Published"
                            ))
                            #self.gs.update_cell(mtpos.WORKSHEET_TAB_SOR, row_index, f"RPA Deployment Remarks - {app_name}", "Published")
                            #logger.info(f"Marked Material Code {material_code} as Published")
                          
                        else:
                            logger.warning(f"Material Code {material_code} not found in GSheet")
                    except Exception as e:
                        logger.error(f"GSheet update failed for {material_code}: {e}")

            else:
                logger.error("Export window did not appear (fail)")
                raise RuntimeError("Export window not found")

            logger.info(f"run_publish_to_all finished in {time.time() - t_start:.2f}s")

        except Exception as e:
            logger.error(f"Exception during publish: {e}")
            # Mark items as Publish Failed
            for row in filtered_sorted_data:
                material_code = row.get("Material Code")
                try:
                    result = self.gs.find_row_index(self.data, identifier_column, material_code)
                    if result:
                        row_index = result["row_index"]
                        
                        self.pending_updates.append((
                            mtpos.WORKSHEET_TAB_SOR,   
                            row_index,
                            f"RPA Deployment Remarks - {app_name}",
                            "Publish Failed"
                        ))
                        #self.gs.update_cell(mtpos.WORKSHEET_TAB_SOR, row_index, f"RPA Deployment Remarks - {app_name}", "Publish Failed")
                        logger.info(f"Marked Material Code {material_code} as Publish Failed")
                      
                    else:
                        logger.warning(f"Material Code {material_code} not found in GSheet")
                except Exception as e2:
                    logger.error(f"GSheet update failed for {material_code}: {e2}")

            raise

    def clear_all(self,child_name,custom_filter_row):
        self.bot.find_element_in_parent(
        child_control_type="DataItem", 
        child_name=child_name, 
        action="double_click",
        element = custom_filter_row
        )
        wait(0.3)
        send_keys('^+a{BACKSPACE}')
