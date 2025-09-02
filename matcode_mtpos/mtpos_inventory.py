import time
from pywinauto.keyboard import send_keys
from utils.logger import setup_in_memory_logger
from utils.helpers import wait
from matcode_mtpos.mtpos_constant import MTPOS_Constants


# Instantiate constants
mtpos = MTPOS_Constants()

#Setup logger for the MTPOS process
logger,log_stream = setup_in_memory_logger("MTPOS")

class MtposInventory:

    def __init__(self, bot, gs, data):

        self.bot = bot
        self.gs = gs
        self.data = data


    def run_create(self,row, app_name):

        reqSNEntry = row.get("ReqSNEntry")
        material_code = row.get("Material Code")
        category = row.get("Category")
        subcategory = row.get("SubCategory")
        retail_price = row.get("Retail Price")
        mat_desc=row.get("Material Description")
        identifier_column = mtpos.MATERIAL_CODE

        try:
            logger.info(f"Adding MTPOS Material code {material_code}")
            self.bot.find_element(name="Catalog Def", control_type="TabItem", action="click")
            #Add item
            self.bot.find_element(automation_id="cmdAdd", control_type="Button", action="click")
            element = self.bot.wait_until_element_present(name="Microtelecom", control_type="Window",single_attempt_timeout = 1,retries= 2)

            if element:
                self.bot.find_element_in_parent(
                    parent_control_type="Window",
                    child_control_type="Button", 
                    parent_name="Microtelecom",
                    child_name="OK", 
                    action="click")
            else:
                send_keys('{ESC}')
            
            for parent_id, variable in [
            ("cboCatDX0", category),
            ("cboCatDX1", subcategory),
            ("txtSearch", mtpos.SUPPLIER)
               ]:
                self.bot.find_element_in_parent(
                    parent_control_type="Edit",
                    child_control_type="Button", 
                    parent_automation_id=parent_id,
                    child_name="Open", 
                    action="click"
                )
                wait(2)
                element = self.bot.find_element_with_index(
                    control_type="Window",
                    action="find",
                    found_index=0,
                    search_descendants=True 
                )
                print("done")
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
                    self.bot.find_element(variable = variable,automation_id=auto_id, control_type="Edit", action="sendkeys")

            if reqSNEntry == "Y":
                self.bot.find_element(automation_id="ckReqSNEntry", control_type="CheckBox", action="click")

            self.bot.find_element(variable = "0",automation_id="SNLength", control_type="Edit", action="sendkeys")

            self.bot.find_element(automation_id="cmdUpdate", control_type="Button", action="click")

            element = self.bot.wait_until_element_present(name="MT.Main.v5", control_type="Window",single_attempt_timeout = 1,retries= 1)
            if element:
                logger.warning(f"Item {material_code} is already defined in location 60001 inventory catalog.")

                # Step 2: Update Google Sheet with "Failed"
                result = self.gs.find_row_index(self.data, identifier_column, material_code)
                if result:
                    row_index = result["row_index"]
                    self.gs.update_cell(mtpos.WORKSHEET_TAB_SOR, row_index, f"RPA Definition Remarks - {app_name}", "Failed")
                    self.gs.update_cell(mtpos.WORKSHEET_TAB_SOR, row_index, f"RPA Deployment Remarks - {app_name}", "Publish Failed")
                    logger.info(f"Updated GSheet row {row_index} with Failed status")
                else:
                    logger.warning(f"Material code {material_code} not found in filtered data for GSheet update")

                #Mark failure and raise to proceed to next
                raise Exception(f"Material code {material_code} is already defined; skipping to next")

            self.bot.find_element(name="Options", control_type="TabItem", action="click")

            self.bot.find_element(automation_id="cmdEdit", control_type="Button", action="click")

            for name in [
                    ("Right Trim S/N to specified length").strip(),
                    ("Retail Price Include the Tax").strip(),
                    ("Require S/N on Sale").strip()
                ]:
                    
                    pane = self.bot.find_element(name=name, control_type="ListItem", action="click")
                    pane.set_focus()
                    send_keys('{SPACE}')
                    element = self.bot.wait_until_element_present(name="Microtelecom", control_type="Window",single_attempt_timeout = 1,retries= 2)

                    if element:
                        self.bot.find_element_in_parent(
                            parent_control_type="Window",
                            child_control_type="Button", 
                            parent_name="Microtelecom",
                            child_name="OK", 
                            action="click")
                    else:
                        send_keys('{ESC}')

            self.bot.find_element(automation_id="cmdUpdate", control_type="Button", action="click")

            #Update Google Sheet with "Success"
            result = self.gs.find_row_index(self.data, identifier_column, material_code)
            if result:
                row_index = result["row_index"]
                self.gs.update_cell(mtpos.WORKSHEET_TAB_SOR, row_index, f"RPA Definition Remarks - {app_name}", "Success")
                logger.info(f"Updated GSheet row {row_index} with Success status")
            else:
                logger.warning(f"Material code {material_code} not found in filtered data for GSheet update")

        except Exception as inner_e:
            logger.error(f"Error processing {material_code}: {inner_e}")

            # Try marking as Failed if not already done
            try:
                result = self.gs.find_row_index(self.data, identifier_column, material_code)
                if result:
                    row_index = result["row_index"]
                    self.gs.update_cell(mtpos.WORKSHEET_TAB_SOR, row_index, f"RPA Definition Remarks - {app_name}", "Failed")
                    self.gs.update_cell(mtpos.WORKSHEET_TAB_SOR, row_index, f"RPA Deployment Remarks - {app_name}", "Publish Failed")
                    logger.info(f"Updated GSheet row {row_index} with Failed status")
            except Exception as sheet_update_error:
                logger.error(f"Failed to update GSheet for {material_code}: {sheet_update_error}")

            # Finally re-raise so outer loop can catch and move to next
            raise
    

    def run_update_srp(self, row, app_name,procedure):
        material_code = row.get("Material Code")
        retail_price = row.get("Retail Price")
        description = row.get("Material Description")
        identifier_column = mtpos.MATERIAL_CODE

        try:
            logger.info(f"Searching MTPOS Material code {material_code}")

             #Add item
            self.bot.find_element(name="Catalog Def", control_type="TabItem", action="click")
            element = self.bot.wait_until_element_present(name="Microtelecom", control_type="Window",single_attempt_timeout = 1,retries= 2)

            if element:
                self.bot.find_element_in_parent(
                    parent_control_type="Window",
                    child_control_type="Button", 
                    parent_name="Microtelecom",
                    child_name="OK", 
                    action="click")
            else:
                send_keys('{ESC}')

            # Search for item
            element_1 = self.bot.wait_until_element_present(name="All Items", control_type="Edit",single_attempt_timeout = 1,retries= 2)
            if not element_1:
                self.bot.find_element(variable="All Items", control_type="Edit", automation_id="cboSearchIn2", action="send_type")
            self.bot.find_element(variable=material_code, control_type="Edit", automation_id="teFind", action="sendkeys")
            self.bot.find_element(automation_id="cmdRefreshInv", control_type="Button", action="click")

            element = self.bot.wait_until_element_present(name="Row 1", control_type="Custom")
            if not element:
                logger.warning(f"MTPOS Material code {material_code} not found in UI")

                # Step 2: Update Google Sheet with "Failed"
                result = self.gs.find_row_index(self.data, identifier_column, material_code)
                if result:
                    row_index = result["row_index"]
                    self.gs.update_cell(mtpos.WORKSHEET_TAB_SOR, row_index, f"RPA Definition Remarks - {app_name}", "Failed")
                    self.gs.update_cell(mtpos.WORKSHEET_TAB_SOR, row_index, f"RPA Deployment Remarks - {app_name}", "Publish Failed")
                    logger.info(f"Updated GSheet row {row_index} with Failed status")
                else:
                    logger.warning(f"Material code {material_code} not found in filtered data for GSheet update")

                #Mark failure and raise to proceed to next
                raise Exception(f"Material code {material_code} not found; skipping to next")

            logger.info(f"MTPOS Material code {material_code} found; editing MSRP and Retail Price")

            #Edit item
            self.bot.find_element(automation_id="cmdEdit", control_type="Button", action="click")
            if procedure == "update-srp":
                self.bot.wait_until_element_present(automation_id="FaceValue", control_type="Edit")

                self.bot.find_element(variable=retail_price, control_type="Edit", automation_id="FaceValue", action="sendkeys")
                self.bot.find_element(variable=retail_price, control_type="Edit", automation_id="Retail", action="sendkeys")

                self.bot.find_element(automation_id="cmdUpdate", control_type="Button", action="click")
            else:

                self.bot.find_element(variable=description, control_type="Edit", automation_id="ItemDesc", action="sendkeys")
                self.bot.find_element(automation_id="cmdUpdate", control_type="Button", action="click")

            logger.info(f"MTPOS Material code {material_code} successfully updated")

            #Update Google Sheet with "Success"
            result = self.gs.find_row_index(self.data, identifier_column, material_code)
            if result:
                row_index = result["row_index"]
                self.gs.update_cell(mtpos.WORKSHEET_TAB_SOR, row_index, f"RPA Definition Remarks - {app_name}", "Success")
                logger.info(f"Updated GSheet row {row_index} with Success status")
            else:
                logger.warning(f"Material code {material_code} not found in filtered data for GSheet update")

        except Exception as inner_e:
            logger.error(f"Error processing {material_code}: {inner_e}")

            # Try marking as Failed if not already done
            try:
                result = self.gs.find_row_index(self.data, identifier_column, material_code)
                if result:
                    row_index = result["row_index"]
                    self.gs.update_cell(mtpos.WORKSHEET_TAB_SOR, row_index, f"RPA Definition Remarks - {app_name}", "Failed")
                    self.gs.update_cell(mtpos.WORKSHEET_TAB_SOR, row_index, f"RPA Deployment Remarks - {app_name}", "Publish Failed")
                    logger.info(f"Updated GSheet row {row_index} with Failed status")
            except Exception as sheet_update_error:
                logger.error(f"Failed to update GSheet for {material_code}: {sheet_update_error}")

            # Finally re-raise so outer loop can catch and move to next
            raise
    
    def run_publish_to_all(self, app_name, filtered_sorted_data, creds_row):
        
        username = creds_row["Username"]
        password = creds_row["Password"]
        identifier_column = mtpos.MATERIAL_CODE

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

           
            self.bot.find_element(automation_id="SplitContainer1", control_type="Pane", action="find")
            
            for auto_id in ("dtFromDate", "dtToDate"):
                found = False
                for attempt in range(1, 4):  # try 3 times
                    try:
                        t0 = time.time()
                        pane = self.bot.find_element(automation_id=auto_id, control_type="Pane",action="find")
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
            logger.info("Update Stores Inventory window found and ready!")
          
            # Click Search 
            self.bot.find_element(automation_id="cmdSearch", control_type="Button", action="click")

            wait(3)

            for row in filtered_sorted_data:
                matcode = row.get("Material Code")
                desc = row.get("Material Description").strip()

                # Step 1: type item code in filter
                for action in [
                    ("click"),
                    ("send_type"),
                ]:
                    self.bot.find_element_in_parent(
                    parent_control_type="Custom",
                    child_control_type="DataItem", 
                    parent_name="Filter Row",
                    child_name="Item Code row -2147483646", 
                    action=action,
                    variable=matcode
                    )

                # Step 2: try find Select row 0
                element = self.bot.find_element_in_parent(
                    parent_control_type="Custom",
                    child_control_type="DataItem", 
                    parent_name="Row 1",
                    child_name="Select row 0", 
                    action="find"
                )

                if element:
                    self.bot.find_element_in_parent(
                    parent_control_type="Custom",
                    child_control_type="DataItem", 
                    parent_name="Row 1",
                    child_name="Select row 0", 
                    action="click"
                    )

                    # Item code exists, no need to type description
                    # Clear filter row
                    self.clear_all("Item Code row -2147483646")

                    continue  # skip to next row

                # Not found by Item Code → now add description filter to help find
                # Double-click to clear item code filter
                self.clear_all("Item Code row -2147483646")

                # Type description
                for action in [
                    ("click"),
                    ("send_type"),
                ]:
                    self.bot.find_element_in_parent(
                    parent_control_type="Custom",
                    child_control_type="DataItem", 
                    parent_name="Filter Row",
                    child_name="Description row -2147483646", 
                    action=action,
                    variable=desc
                    )

                # Step 3: click to confirm row
                self.bot.find_element_in_parent(
                    parent_control_type="Custom",
                    child_control_type="DataItem", 
                    parent_name="Row 1",
                    child_name="Select row 0", 
                    action="click"
                )

                # Finally clear description filter
                self.clear_all("Description row -2147483646")

            # Find ComboBox parent once
            combo_start = time.time()
            combo_parent = new_window.child_window(control_type="ComboBox", auto_id="cboHouse")
            combo_parent.wait('exists ready', timeout=1)
            logger.info(f"Found ComboBox parent in {time.time() - combo_start:.2f}s")

            # Click Open button and ALL ListItem inside ComboBox
            for name, child_control_type in [
                ("Open", "Button"),
                ("ALL", "ListItem")
            ]:
                c_start = time.time()
                child = combo_parent.child_window(control_type=child_control_type, title=name)
                child.wait('exists ready', timeout=1)
                child.click_input()
                logger.info(f"Clicked '{name}' ({child_control_type}) in {time.time() - c_start:.2f}s")

            # Click Select All Locations and Copy Inventory
            for auto_id, control_type in [
                ("ckSelectAllLocations", "CheckBox"),
                ("cmdCopyInventory", "Button")
            ]:
                self.bot.find_element(automation_id=auto_id, control_type=control_type, action="click")
            logger.info("Clicked Select All Locations & Copy Inventory")

            # Handle user verification popup quickly
            pass_check = self.bot.find_element(automation_id="frmPassCheck", control_type="Window", action="find")
            if pass_check:
                logger.info("Found user verification prompt")
                self.bot.find_element(variable=username, automation_id="txtUID", control_type="Edit", action="sendkeys")
                self.bot.find_element(variable=password, automation_id="PassTxt", control_type="Edit", action="send_type")
                self.bot.find_element(name="C&ontinue", control_type="Button", action="click")
            else:
                logger.error("User verification window not found. Failed to publish.")
                raise RuntimeError("User verification window not found")

            # Confirm No Items Selected → Yes
            self.bot.find_element(name="Yes", control_type="Button", action="click")
            logger.info("Clicked confirmation Yes")

            # Wait for export result
            if self.bot.wait_until_element_present(name="Microtelecom", control_type="Window", retries=30, single_attempt_timeout=60, retry_interval=0):
                self.bot.find_element(name="OK", control_type="Button", action="click")
                logger.info("Export completed successfully")

                # Update GSheet: mark Published for items with empty remark
                for row in filtered_sorted_data:
                    material_code = row.get("Material Code")
                    try:
                        result = self.gs.find_row_index(self.data, identifier_column, material_code)
                        if result:
                            row_index = result["row_index"]
                            row_data = self.gs.get_row(mtpos.WORKSHEET_TAB_SOR, row_index)
                            latest_remark = row_data.get(f"RPA Deployment Remarks - {app_name}")

                            if not latest_remark:
                                self.gs.update_cell(mtpos.WORKSHEET_TAB_SOR, row_index,
                                    f"RPA Deployment Remarks - {app_name}", "Published")
                                logger.info(f"Marked Material Code {material_code} as Published")
                            else:
                                logger.info(f"Skipped {material_code}, already has remark '{latest_remark}'")
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
                        row_data = self.gs.get_row(mtpos.WORKSHEET_TAB_SOR, row_index)
                        latest_remark = row_data.get(f"RPA Deployment Remarks - {app_name}")

                        if not latest_remark:
                            self.gs.update_cell(mtpos.WORKSHEET_TAB_SOR, row_index,
                                f"RPA Deployment Remarks - {app_name}", "Publish Failed")
                            logger.info(f"Marked Material Code {material_code} as Publish Failed")
                        else:
                            logger.info(f"Skipped {material_code}, already has remark '{latest_remark}'")
                    else:
                        logger.warning(f"Material Code {material_code} not found in GSheet")
                except Exception as e2:
                    logger.error(f"GSheet update failed for {material_code}: {e2}")

            raise

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
