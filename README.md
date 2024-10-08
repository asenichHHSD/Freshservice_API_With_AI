# HHSD Freshservice Python Script
---------------------------------
## Note

This script was developed specifically for use by HHSD. Please feel free to modify and utilize this code however you see fit.
---------------------------------

This Python script gathers replacement request data from Freshservice for the last specified number of days and exports it to a CSV file. Additionally, it can use AI to rephrase the descriptions of the replacement requests, making them more concise and clear. The AI rephrasing feature uses the Ollama API and supports multiple Llama models.

## Features
-----------

-   **Exports replacement data**: Fetches data from Freshservice API and exports it to a CSV file in the format `output_data_<timestamp>.csv`.
-   **AI-assisted description rephrasing**: Rephrases descriptions using AI (via Ollama), helping to clarify and simplify the content.
-   **Customizable export**: Allows you to specify the number of days for which you want to export data.
-   **CSV format**: Data is stored in a CSV file for easy review and analysis.

## Requirements
------------

- Python 3.8 or above
- Freshservice API key
- Ollama API for AI rephrasing (if enabled)

## Installation
------------

1.  Install Python 3.8 or above.
2.  Install required packages:

`pip install requests ollama tqdm`


## Running the Script
------------------

1.  Run FetchReplacementData.py:

    `python FetchReplacementData.py`

2.  Follow the prompts to enter the Freshservice base URL, API key, and the number of days of data you would like to export. You will also be asked if you'd like to enable AI-powered description rephrasing.

    If AI rephrasing is enabled, the script will revise the original descriptions using the chosen Llama model.

3.  Open the generated CSV file (`output_data_<timestamp>.csv`) to view the data.

### AI Rephrasing Option

-   When prompted, you can choose whether or not to use AI to rephrase the replacement descriptions.
-   You can select between any Ollama supported model for text rephrasing, which will generate clearer and more concise descriptions.

## Packaging the Script into an EXE
--------------------------------

If you want to distribute this script as an executable, follow these steps:

1.  Install `pyinstaller`:

    `pip install pyinstaller`

2.  Package the script:

    `pyinstaller.exe -F FetchReplacementData.py`

## Output
---------

The output file, `output_data_<timestamp>.csv`, will contain the following fields:

-   **Replacement Date**
-   **Item**
-   **Model**
-   **Building**
-   **Service Request**
-   **Technician**
-   **Asset Number**
-   **Username**
-   **Full Description**
-   **AI Description** (Optional, generated by the AI if enabled)

## Setting up Ollama
-------------------
### 
![alt text](https://ollama.com/public/ollama.png)

- Install Ollama on your system by following the system specific directions found here:
[Ollama Install](https://ollama.com/download)

- Download and run your desired model (script is written for llama3.1):
[Ollama Models](https://ollama.com/library)

- Enter model name when asked when running the script



## Notes
--------

-   AI rephrasing requires an active Ollama setup, and you will be prompted to select a Llama model during runtime.
-   CSV file output is created in the script directory and is time-stamped for reference.

------------------
![alt text](llama.png)
