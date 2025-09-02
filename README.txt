# Setup
Google Sheet Link: https://docs.google.com/spreadsheets/d/14kmKYQh9FG0GHJVMxbdlxoa10Juw0aXCMjp0L1KnNK8/edit?gid=2145120567#gid=2145120567
## Install Prerequisites and Service Dependencies

## For Windows

### Python Installation
### NOTE: If the PC/Laptop has already a python installed, ignore this step
1. Go to this link https://www.python.org/downloads/ to download .exe installer for windows
2. Once download is done, proceed to install python then follow the instruction

### Dependencies Installation
1. Run the dependency loader script in your machine. Select win_dependency_loader.bat and run

## Dependencies Manual Installation
-Use this step if 'mac_dependency_loader.sh' is not working
1. On your current opened Terminal, type this commands below (NOTE: one at time, make sure to finish the download and installation before moving on the the next command)
>> pip install selenium
>> pip install google-api-python-client
>> pip install webdriver-manager
>> pip install gspread
>> pip install python-dotenv

## Run Service Deployment
1. Once you've completed installing all the prerequisites. 
2. Double click 'matcode_run_service.bat' file to start the Matcode Definition.
3. Double click 'promocode_run_service.bat' file to start the Promo Code Definition.

==========================================================

## For MAC

### Python Installation
### NOTE: If the PC/Laptop has already a python installed, ignore this step
1. Go to this link https://www.python.org/downloads/ to download .pkg installer for mac
2. Once download is done, proceed to install python then follow the instruction

### Configure .sh Files to Run with Terminal and Execute Permission
1. On your current directory, control-click 'mac_dependency_loader.sh' then select get info
2. Under Open with: section, select Terminal. If Terminal does not exist in selection, select Others then search for Terminal
3. Repeat from step 1 and 2 for 'run_service_mac.sh' file

4. Once your done, open Terminal and navigate to the directory where your script is located.
5. Then, run the commands below for your OS mac to give permission to run the files: 
>> chmod +x mac_dependency_loader.sh 
>> chmod +x run_service_mac.sh 

### Dependencies Installation
1. Run the dependency loader script in your machine. Select mac_dependency_loader.sh and run

### Dependencies Manual Installation
-Use this step if 'mac_dependency_loader.sh' is not working
1. On your current opened Terminal, type this commands below (NOTE: one at time, make sure to finish the download and installation before moving on the the next command)
>> pip3 install selenium
>> pip3 install google-api-python-client
>> pip3 install webdriver-manager
>> pip3 install gspread
>> pip3 install python-dotenv

## Run Service Deployment
1. Once you've completed installing all the prerequisites. 
2. Double click 'matcode_run_service.bat' file to start the Matcode Definition.
3. Double click 'promocode_run_service.bat' file to start the Promo Code Definition.
