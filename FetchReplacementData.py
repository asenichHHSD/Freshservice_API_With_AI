import requests
from datetime import datetime, timedelta
import json
import csv
import random
import time
import os
from tqdm import tqdm
import subprocess
import atexit
import signal
import sys

try:
    import ollama
    ollamaEnabled = True
except ImportError:
    print("Ollama library not found, AI features will be disabled...")
    ollamaEnabled = False


# Define the file where the data will be stored
userDataFile = 'user_data.json'
# Check if the file exists and read the stored data if it does
if os.path.exists(userDataFile):
    with open(userDataFile, 'r') as f:
        user_data = json.load(f)
else:
    user_data = {}

# Questions and default answers from previous runs
questions = {
    'freshserviceBaseURL': "Freshservice base URL Ex. https://xxx.freshservice.com/? ",
    'apiKey': "What is your API key?  ",
    'lenovoKey': "Please enter your Lenovo API key: ",
    'days': "How many days of service requests would you like to export?  ",
    'ai': "Would you like to rephrase descriptions with AI (Requires Ollama), yes/no? ",
    'model': "Please select enter the name of the model you want to use (must already be installed, select n/a if not using AI): ",
    'debug' : "Enable debugging, yes/no: "
}


# Ask questions and allow the user to keep previous answers or update them
for key, question in questions.items():
    if key in user_data:
        new_value = input(f"{question} (press Enter to keep '{user_data[key]}'): ")
        if new_value:
            user_data[key] = new_value  # Update value if new input is provided
    else:
        user_data[key] = input(question)  # Ask for the value if it's not stored

# Save the updated data back to the JSON file
with open(userDataFile, 'w') as f:
    json.dump(user_data, f, indent=4)

# Get user's API key from text file
file_name = "api_key.txt"

# Initialize a variable to store the API key

# Define your API endpoint and authentication details
base_url = str(user_data['freshserviceBaseURL']) + "api/v2/tickets"
password = ""

# Get the current time and format it
current_time = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')

# Define the filename with the current time appended
csvFilename = f'output_data_{current_time}.csv'

# Define Ollama process
global ollamaProcess

def checkOllama():
    # Check if the OS is Windows using os.name
    is_windows = os.name == "nt"
    
    # Check if the ollama.exe file is in the current working directory
    ollama_exists = os.path.isfile(os.path.join(os.getcwd(), "ollama.exe"))
    
    # Check all conditions
    if (user_data["ai"] == "yes" and
        ollamaEnabled == True and
        is_windows and
        ollama_exists):
        return True
    else:
        return False

def terminate_exe():
    """Ensure the .exe process is killed when the script exits."""
    if ollamaProcess.poll() is None:  # Check if the process is still running
        print("Terminating the exe file...")
        ollamaProcess.terminate()  # Send termination signal
        ollamaProcess.wait()  # Wait for it to fully terminate
        print("Exe terminated.")


def startOllama():
    # Get the current working directory
    current_dir = os.getcwd()

    # Set the environment variable OLLAMA_MODELS
    os.environ['OLLAMA_MODELS'] = os.path.join(current_dir, 'models')
    print(os.environ['OLLAMA_MODELS'])

    # Now proceed with the rest of your code
    ollamaProcess = subprocess.Popen(['ollama.exe', 'serve'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    # Register the termination function to be called on normal exit
    atexit.register(terminate_exe)
    # Register signal handlers to ensure termination on script interruption
    signal.signal(signal.SIGINT, signal_handler)  # Handles Ctrl+C
    signal.signal(signal.SIGTERM, signal_handler)  # Handles termination signals
    for i in tqdm(range(30), desc='Starting the Llama', unit='s', bar_format='{l_bar}{bar}| {remaining}', colour="green", leave=False):
        time.sleep(1)  # Sleep for 1 second (adjust this for finer progress)


def signal_handler(sig, frame):
    """Handle signals like SIGINT (Ctrl+C) and ensure cleanup."""
    print("Script interrupted, exiting...")
    terminate_exe()  # Clean up and terminate the exe
    sys.exit(0)

def logging(*args):
    if user_data['debug'] == 'yes':
        message = ' '.join(str(arg) for arg in args)
        print(message)

def coolDown(cooldownTime):

    # Create a progress bar that spans `total_duration` seconds
    for i in tqdm(range(cooldownTime), desc='Waiting for API Cooldown', unit='s', bar_format='{l_bar}{bar}| {remaining}', colour="red", leave=False):
        time.sleep(1)  # Sleep for 1 second (adjust this for finer progress)


def readAPIKey():
    # Try to open the file and read the API key
    try:
        with open(file_name, "r") as file:
            fileKey = str(file.read().strip())
        # Close the file automatically when the block is exited

        # Check if the API key was successfully read
        if fileKey:
            print("API key read successfully:", fileKey)
            return fileKey
        else:
            print("API key is empty. Make sure the file contains the API key.")
            input("Press Enter to exit....")
            exit()
    except FileNotFoundError:
        print(f"File '{file_name}' not found. Make sure the file is in the same directory as your Python script.")
        input("Press Enter to exit....")
        exit()
    except Exception as e:
        print("An error occurred:", str(e))
        input("Press Enter to exit....")
        exit()

def rephraseText(inputString):
    systemRole = "You are a bot that rephrases the explanation of device damage that is provided into breif clear and concise statements. Only provide the revised statment and nothing else. Omit any parts of the statement that indicate intentional damage. If the wording is already consise do not change it much."
    
    response = ollama.chat(model=user_data["model"], messages=[
    {"role": "system", "content": systemRole},
    {"role": "user", "content": str(inputString)},
    ])
    logging("-----------------------------")
    logging("Original Description: ")
    logging(inputString)
    logging("")
    logging("AI Description: ")
    logging(response['message']['content'])
    logging("-----------------------------")
    return response['message']['content']

def createCSVFile():

    # Open the file in write mode ('w'), which creates the file if it doesn't exist
    with open(csvFilename, mode='w', newline='') as file:
        writer = csv.writer(file)
        
        # Writing header (optional)
        if user_data["ai"] == "yes":
            writer.writerow(["Replacement Date", "Item", "Category" , "Sub_Category", "Item_Category", "Model", "Building", "Service Request", "Technician", "Asset Number", "Serial", "Warranty Status","Username", "Full Description", "AI Description"])
        else:
            writer.writerow(["Replacement Date", "Item", "Category" , "Sub_Category", "Item_Category", "Model", "Building", "Service Request", "Technician", "Asset Number", "Serial", "Warranty Status","Username", "Full Description"])

def randomSleep():
    time.sleep(random.uniform(0.1,0.6))

def get_serial_from_asset_tag(asset_tag):
    # Find the CSV file
    current_dir = os.getcwd()
    files = os.listdir(current_dir)
    
    csv_files = [f for f in files if 'assets-report' in f.lower() and f.lower().endswith('.csv')]

    if not csv_files:
        print("Error: No CSV file containing 'asset-report' found in the current directory.")
        sys.exit(1)  # Exit the script

    csv_file = csv_files[0]  # Use the first matching file

    # Try to open the CSV file
    try:
        with open(csv_file, 'r', newline='') as csvfile:
            reader = csv.DictReader(csvfile)
            # Ensure required fields are present
            if 'Asset Tag' not in reader.fieldnames or 'Serial' not in reader.fieldnames:
                logging("Error: CSV file does not contain 'Asset Tag' and 'Serial' fields.")

            # Search for the asset tag
            for row in reader:
                if row.get('Asset Tag') == asset_tag:
                    return row.get('Serial')
            # Asset tag not found
            logging("Warning: Asset tag not found in the CSV.")
            return None
    except IOError:
        logging("Error: Cannot read file csv file for serial number lookup. It may be open or inaccessible.")

def get_warranty_status(serial_number):
    url = f"https://supportapi.lenovo.com/v2.5/warranty?Serial={serial_number}"
    headers = {
        "ClientID": user_data["lenovoKey"]  # Replace with your actual Client ID
    }

    rerunLookup = True
    errorCount = 0

    while rerunLookup:


        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            data = response.json()

            if data.get('InWarranty') is not None:
                in_warranty = data['InWarranty']
                if in_warranty:
                    # Find the warranty with the latest 'End' date
                    warranties = data.get('Warranty', [])
                    if warranties:
                        latest_end_date = max(
                            datetime.strptime(w['End'], "%Y-%m-%dT%H:%M:%SZ")
                            for w in warranties
                        )
                        today = datetime.now()
                        delta = latest_end_date - today
                        months_left = delta.days // 30  # Approximate months left
                        rerunLookup = False
                        return f"In warranty: {months_left} months left"
                    else:
                        rerunLookup = False
                        return "In warranty: Warranty details not available"
                else:
                    rerunLookup = False
                    return "Out of Warranty"
            else:
                errorCount += 1
                if errorCount == 3:
                    rerunLookup = False
                    return "Error: Invalid response data"

        except requests.exceptions.RequestException as err:
            errorCount += 1

            if errorCount == 3:
                rerunLookup = False
                logging("Lenovo API Error")

                return "Error: Unable to retrieve warranty status"

def fetchReplacementData(days: int):
    api_Key = user_data["apiKey"]
    # Calculate the date 30 days ago
    end_date = datetime.now()
    start_date = end_date - timedelta(days=days)

    # Format the date strings in ISO 8601 format
    end_date_str = end_date.strftime("%Y-%m-%dT%H:%M:%S")
    start_date_str = start_date.strftime("%Y-%m-%dT%H:%M:%S")

    logging("Exporting data from past " + str(days) + " days")

    pageList = []

    for x in range(1001):
        pageList.append(x)

    pageList.remove(0)
    # print(pageList)

    blankPageCount = 0

    os.system('cls' if os.name == 'nt' else 'clear')
    for page in tqdm(pageList, desc="Processing Ticket Pages", unit="Page"):
        randomSleep()
        service_request_ids = []
        # Define the parameters for the API request
        params = {
        "updated_since": start_date_str,
        "per_page": 100,
        "page": page
        }

        # Send the GET request with basic authentication

        response = requests.get(base_url, params=params, auth=(api_Key, password))
        reRunPageFetch = True


        while reRunPageFetch:
            if response.status_code == 200:
                reRunPageFetch = False
                try:
                    data = response.json()['tickets']  # Try to parse the JSON response
                    # print(data)
                except json.decoder.JSONDecodeError:
                    print("Unable to parse the JSON response")
                    data = []  # Set data as an empty list if parsing fails
                    input("Press Enter to exit....")
                    exit()
            else:
                # response.close()
                logging("Waiting for API cooldown")
                # time.sleep(30)
                coolDown(60)

        logging("Fetching Data from Freshservice page: "+  str(page))
        for ticket in data:
                if isinstance(ticket, dict) and ticket.get("type") == "Service Request":
                    service_request_ids.append(ticket["id"])

        if service_request_ids == []:
            logging("No valid SRs on current page...")
            blankPageCount += 1
            if blankPageCount == 75:
                print("Encountered 50 consecutive blank ticket pages, exiting script...")
                break
            else:
                continue
        blankPageCount = 0
        response.close()
        
        logging("Service Request IDs:", service_request_ids)

        assetNumbers = []
        problemDescriptions = []
        building = []
        userName = []
        technician = []
        srNumber = []
        replacementDate = []
        item = []
        itemModel = []
        serialNumber = []
        warrantyStatus = []

        category = []
        subCategory = []
        itemCategory = []


        # print("Fetching specific replacement data...")

        for ticketNumber in tqdm(service_request_ids, desc="Fetching SR Data", unit="Service Request", leave=False, colour="blue"):
            randomSleep()
            logging(str(service_request_ids.index(ticketNumber)+1) + "/" + str(len(service_request_ids)))
            itemsURL = base_url+"/"+str(ticketNumber)+"/requested_items"
            # print(itemsURL)
            reRunSrFetch = True

            while reRunSrFetch:
                response = requests.get(itemsURL, auth=(api_Key, password))
                if response.status_code == 200:
                    reRunSrFetch = False
                    try:
                        data = response.json()['requested_items'][0]  # Try to parse the JSON response
                        # print(data)
                        if data.get('service_item_id') == 47:
                            custom_fields = data.get('custom_fields', {})
                            assetNumbers.append(custom_fields.get('asset'))
                            problemDescriptions.append(custom_fields.get('full_issue_description'))
                            building.append(custom_fields.get('building'))
                            userName.append(custom_fields.get('student_username'))
                            technician.append(custom_fields.get('technician_initials'))
                            srNumber.append(str(ticketNumber))
                            replacementDate.append(data.get('created_at'))
                            # print(replacementDate)
                            item.append(custom_fields.get('item'))
                            itemModel.append(custom_fields.get('model'))
                            machineSerial = get_serial_from_asset_tag(custom_fields.get('asset'))
                            serialNumber.append(machineSerial)
                            if machineSerial != None and len(machineSerial) ==8:
                                warrantyStatus.append(get_warranty_status(machineSerial))
                            else:
                                warrantyStatus.append("No Data for Lookup")

                            ticketDataURL = base_url+"/"+str(ticketNumber)
                            ticketDataResponse = requests.get(ticketDataURL, auth=(api_Key, password))
                            if ticketDataResponse.status_code == 200:
                                category.append(ticketDataResponse.json()['ticket']['category'])
                                subCategory.append(ticketDataResponse.json()['ticket']['sub_category'])
                                itemCategory.append(ticketDataResponse.json()['ticket']['item_category'])
                            else:
                                category.append("")
                                subCategory.append("")
                                itemCategory.append("")

                



                    except json.decoder.JSONDecodeError:
                        print("Unable to parse the JSON response")
                        data = []  # Set data as an empty list if parsing fails
                        input("Press Enter to exit....")
                        exit()
                    except IndexError:
                        logging("No Item Data found for " + str(ticketNumber))
                        continue

                else:
                    # response.close()
                    logging("Error: " + str(response.status_code))
                    logging("Waiting for API cooldown")
                    # time.sleep(30)
                    coolDown(60)
                response.close()

            

        logging("Writing data to CSV...")

        # Create a CSV file and write the data
        logging(csvFilename)
        try:
            with open(csvFilename, mode='a', newline='') as csv_file:
                writer = csv.writer(csv_file)
            
                # Write the header row (optional)
                # writer.writerow(["Replacement Date", "Item", "Model", "Building", "Service Request", "Technician", "Asset Number", "Username", "Full Description"])
            
                # Write the data rows
                for replaceDate, itemType, ticketCat, subCat, itemCat, itemTypeModel, replacementBuilding, replacementSRNum, replacementTech, replacementAsset, replacementSerial, replacementWarranty, replacementUsername, replacementDesc in tqdm(zip(replacementDate, item, category, subCategory, itemCategory, itemModel, building, srNumber, technician, assetNumbers, serialNumber, warrantyStatus, userName, problemDescriptions),desc="Writing Data to CSV", unit="Row", colour="yellow", leave=False, total=len(srNumber)):
                    formattedString = ' '.join(replacementDesc.splitlines())
                    # writer.writerow([asset, formattedString])
                    try:
                        if user_data["ai"] == "yes":
                            logging("Revising Description With AI...")
                            ai_description = rephraseText(formattedString)
                            writer.writerow([replaceDate, itemType, ticketCat, subCat, itemCat, itemTypeModel, replacementBuilding, replacementSRNum, replacementTech, replacementAsset, replacementSerial, replacementWarranty, replacementUsername, formattedString, ai_description])
                        else:
                            writer.writerow([replaceDate, itemType, ticketCat, subCat, itemCat, itemTypeModel, replacementBuilding, replacementSRNum, replacementTech, replacementAsset, replacementSerial, replacementWarranty, replacementUsername, formattedString])
                    except UnicodeEncodeError as e:
                        # print("Encoding issue with " + serviceRequestNumber + " resolving issue...")
                        problematic_character_index = e.start
                        problematic_string = e.object
                        cleaned_string = problematic_string[:problematic_character_index] + ' ' + problematic_string[problematic_character_index + 1:]
                        writer.writerow([replaceDate, itemType, ticketCat, subCat, itemCat, itemTypeModel, replacementBuilding, replacementSRNum, replacementTech, replacementAsset, replacementSerial, replacementWarranty, replacementUsername, cleaned_string])
                # for asset, description, build, tech, fullN in zip(assetNumbers, problemDescriptions, building, technician, fullName):
                #     formattedString = ' '.join(description.splitlines())
                #     writer.writerow([asset, formattedString, build, tech, fullN])
        except PermissionError:
            print("Please close or delete the output_data.csv file...")
            input("Press Enter to exit....")
            exit()
        except Exception as e:
            print(f"An error occurred: {e}")
            print("Other error, please read details...")
            input("Press Enter to exit....")
            exit()


# api_key = readAPIKeyFromFile()

# Try to start ollama
if checkOllama():
    logging("Local Llama is available")
    startOllama()
    

createCSVFile()
fetchReplacementData(int(user_data["days"]))
if checkOllama():
    terminate_exe()
print("Output file generated: " + str(csvFilename))
input("Press Enter to exit....")