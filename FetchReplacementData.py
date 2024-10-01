import requests
from datetime import datetime, timedelta
import json
import csv
import random
import time
import os

try:
    import ollama
except ImportError:
    print("Ollama library not found, AI features will be disabled...")


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
    'days': "How many days of service requests would you like to export?  ",
    'ai': "Would you like to rephrase descriptions with AI (Requires Ollama), yes/no? ",
    'model': "Please select enter the name of the model you want to use (must already be installed, select n/a if not using AI): "
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
    print("-----------------------------")
    print("Original Description: ")
    print(inputString)
    print("")
    print("AI Description: ")
    print(response['message']['content'])
    print("-----------------------------")
    return response['message']['content']

def createCSVFile():

    # Open the file in write mode ('w'), which creates the file if it doesn't exist
    with open(csvFilename, mode='w', newline='') as file:
        writer = csv.writer(file)
        
        # Writing header (optional)
        if user_data["ai"] == "yes":
            writer.writerow(["Replacement Date", "Item", "Category" , "Sub_Category", "Item_Category", "Model", "Building", "Service Request", "Technician", "Asset Number", "Username", "Full Description", "AI Description"])
        else:
            writer.writerow(["Replacement Date", "Item", "Category" , "Sub_Category", "Item_Category", "Model", "Building", "Service Request", "Technician", "Asset Number", "Username", "Full Description"])

def randomSleep(maxSleep):
    time.sleep(random.uniform(0.1,maxSleep))

def fetchReplacementData(days: int):
    api_Key = user_data["apiKey"]
    # Calculate the date 30 days ago
    end_date = datetime.now()
    start_date = end_date - timedelta(days=days)

    # Format the date strings in ISO 8601 format
    end_date_str = end_date.strftime("%Y-%m-%dT%H:%M:%S")
    start_date_str = start_date.strftime("%Y-%m-%dT%H:%M:%S")

    print("Exporting data from past " + str(days) + " days")

    pageList = []

    for x in range(301):
        pageList.append(x)

    pageList.remove(0)
    # print(pageList)

    for page in pageList:
        service_request_ids = []
        # Define the parameters for the API request
        params = {
        "updated_since": start_date_str,
        "per_page": 100,
        "page": page
        }

        # Send the GET request with basic authentication

        response = requests.get(base_url, params=params, auth=(api_Key, password))
        # print
        

        if response.status_code == 200:
            try:
                data = response.json()['tickets']  # Try to parse the JSON response
                # print(data)
            except json.decoder.JSONDecodeError:
                print("Unable to parse the JSON response")
                data = []  # Set data as an empty list if parsing fails
                input("Press Enter to exit....")
                exit()

        print("Fetching Data from Freshservice page: "+  str(page))
        for ticket in data:
                if isinstance(ticket, dict) and ticket.get("type") == "Service Request":
                    service_request_ids.append(ticket["id"])

        response.close()
        
        print("Service Request IDs:", service_request_ids)

        assetNumbers = []
        problemDescriptions = []
        building = []
        userName = []
        technician = []
        srNumber = []
        replacementDate = []
        item = []
        itemModel = []

        category = []
        subCategory = []
        itemCategory = []


        print("Fetching specific replacement data...")

        for ticketNumber in service_request_ids:
            randomSleep(1.1)
            print(str(service_request_ids.index(ticketNumber)+1) + "/" + str(len(service_request_ids)))
            itemsURL = base_url+"/"+str(ticketNumber)+"/requested_items"
            # print(itemsURL)
            response = requests.get(itemsURL, auth=(api_Key, password))
            if response.status_code == 200:
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
                    print("No Item Data found for " + str(ticketNumber))
                    continue

            else:
                print("Error: " + str(response.status_code))
            response.close()

            

        print("Writing data to CSV...")

        # Create a CSV file and write the data
        print(csvFilename)
        try:
            with open(csvFilename, mode='a', newline='') as csv_file:
                writer = csv.writer(csv_file)
            
                # Write the header row (optional)
                # writer.writerow(["Replacement Date", "Item", "Model", "Building", "Service Request", "Technician", "Asset Number", "Username", "Full Description"])
            
                # Write the data rows
                for replaceDate, itemType, ticketCat, subCat, itemCat, itemTypeModel, replacementBuilding, replacementSRNum, replacementTech, replacementAsset, replacementUsername, replacementDesc in zip(replacementDate, item, category, subCategory, itemCategory, itemModel, building, srNumber, technician, assetNumbers, userName, problemDescriptions):
                    formattedString = ' '.join(replacementDesc.splitlines())
                    # writer.writerow([asset, formattedString])
                    try:
                        if user_data["ai"] == "yes":
                            print("Revising Description With AI...")
                            ai_description = rephraseText(formattedString)
                            writer.writerow([replaceDate, itemType, ticketCat, subCat, itemCat, itemTypeModel, replacementBuilding, replacementSRNum, replacementTech, replacementAsset, replacementUsername, formattedString, ai_description])
                        else:
                            writer.writerow([replaceDate, itemType, ticketCat, subCat, itemCat, itemTypeModel, replacementBuilding, replacementSRNum, replacementTech, replacementAsset, replacementUsername, formattedString])
                    except UnicodeEncodeError as e:
                        # print("Encoding issue with " + serviceRequestNumber + " resolving issue...")
                        problematic_character_index = e.start
                        problematic_string = e.object
                        cleaned_string = problematic_string[:problematic_character_index] + ' ' + problematic_string[problematic_character_index + 1:]
                        writer.writerow([replaceDate, itemType, ticketCat, subCat, itemCat, itemTypeModel, replacementBuilding, replacementSRNum, replacementTech, replacementAsset, replacementUsername, cleaned_string])
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
createCSVFile()
fetchReplacementData(int(user_data["days"]))