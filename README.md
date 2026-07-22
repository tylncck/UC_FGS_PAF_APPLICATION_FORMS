# Qualtrics PAF Application Processor
An automated, survey-agnostic Python pipeline designed to extract Qualtrics survey submissions, dynamically map metadata, validate student records against an internal Data Warehouse, and render standardized PDF application forms.

## Project Metadata
- **Author:** Taylan Cicek
- **Role:** Lead, Data Analytics & Reporting - Faculty of Graduate Studies
- **Date:** July 2026 

## Project Overview
The Qualtrics PAF Application Processor automates the end-to-end reporting flow for program application forms (PAFs). By dynamically resolving human-readable question labels (such as UCID prompts and field exclusion lists) into exact Qualtrics Question IDs (QID), the system eliminates hardcoded keys and operates seamlessly across annual survey updates. The pipeline securely connects to the Qualtrics API, queries student data from the Data Warehouse, and uses HTML/CSS templating to generate clean, print-ready PDF reports.

## Cloning the Repository
To clone the repository using HTTPS, open your terminal, navigate to your project folder and run the following commands in your terminal:
``` bash
git clone https://git.ucalgary.ca/fgs-analytics/fgs_paf_application_forms.git
```

## Setup & Prerequisites
Before running the application, ensure all dependencies, environment configurations, and network permissions are properly set up.

### 1. Environment
Create and activate a Python virtual environment:
``` bash
python -m venv .venv
```

Activate the virtual environment: 

**macOS / Linux:**
``` bash
source .venv/bin/activate
```
**Windows:**
``` bash
.\.venv\Scripts\activate
```

**Upgrade pip**
``` bash
pip install --upgrade pip
```

### 2. WeasyPrint Installation
This project utilizes WeasyPrint for HTML-to-PDF rendering. WeasyPrint requires system-level C libraries (Pango, GObject, and Cairo) to be installed on your machine.

**macOS (via Homebrew):**
``` bash
brew install python pango libffi
```

**Ubuntu / Debian:**
``` bash
sudo apt-get update
sudo apt-get install build-essential python3-dev python3-pip python3-setuptools python3-wheel python3-cffi libcairo2 libpango-1.0-0 libpangocairo-1.0-0 libgdk-pixbuf2.0-0 libffi-dev shared-mime-info
```

**Windows:**

- **Step 1: Install MSYS2**

    Visit [weasyprint web page](https://doc.courtbouillon.org/weasyprint/stable/first_steps.html#windows) to see the installation steps for Windows. 
    - Install [MSYS2](https://www.msys2.org/#installation) keeping the default options.
    - After installation, in MSYS2’s shell, execute `pacman -S mingw-w64-x86_64-pango`. When prompted to confirm the installation, type `Y` and press Enter.
    - Close MSYS2’s shell. 

- **Step 2: Add MSYS2 to your Windows Environment Path**
    
    To let your normal Python/Jupyter environment find these libraries:
    - Press the Windows Key and search for "Environment Variables" (or choose "Edit the system environment variables").
    - Click the Environment Variables... button at the bottom.
    - In the "User variables" (or "System variables") box, find the variable named Path (or PATH), select it, and click Edit....
    - Click New and add this exact path where MSYS2 installed the library binaries: `C:\msys64\ucrt64\bin`
    - Click OK on all windows to save the changes.

- **Step 3: Verify the Setup**
    
    In your terminal run `python -m pip install weasyprint` to install the package and `python -m weasyprint --info` to check if installation is successful.

### 3. Dependecies
Run the following command in your terminal to install package dependencies:
``` bash
pip install -r requirements.txt
```

### 4. Qualtrics API Credentials & Survey ID
You will need an active Qualtrics OAuth client or API key with permissions to read survey metadata and responses.

**How to Obtain Qualtrics API Credentials:**
- Log in to your Qualtrics account.
- Navigate to Account Settings (top right icon) $\rightarrow$ Qualtrics IDs.
- Under User Settings, you will see ther Datacenter ID. 
- Under OAuth Client Manager, generate a new Client ID and Client Secret.
    - While generating new Client ID and Client Secret, make sure you select all read:.... items (especially read:survey_responses) in the Scopes area.   
- Locate your Survey ID (surveyId=SV_...) from the survey URL or the Surveys section inside Qualtrics IDs.

You will need to paste these data into the appropriate places in the [Jupyter Notebook](/notebooks/PAF_PDF_Generation.ipynb). 

### 5. Administrative VPN & Data Warehouse Server Access
To fetch student records (check_dw=True), your environment must have network connectivity to the internal Data Warehouse server.

Network Requirement: You must be connected to the Administrative UC Networks VPN prior to executing the script, if you have check_dw=True in your class initialization. 

**Default Server & Database Configuration:** Server hostnames and database names are defined directly as class constants inside PAFApplicationForms (e.g., SERVER and DATABASE).

**Custom Configurations:** If you need to point to a staging/test database server or update target endpoints, update these class constants directly inside the source code or override them prior to execution.

## Class Overview & Use Cases
### PAFApplicationForms Orchestrator

The main class, `PAFApplicationForms`, acts as an orchestrator managing API setup, metadata resolution, query execution, and PDF rendering.

#### Core Capabilities:
- **Dynamic QID Mapping:** Resolves readable question labels (e.g., "UCID#:") into internal QIDs automatically.
- **Question Exclusions:** Hides specific survey sections or sensitive questions by passing exact question texts in exclude_questions.
- **Database Integration:** Executes custom .sql queries against the Data Warehouse to append verified student attributes to survey outputs.
- **Strict Validation:** Incorporates layered type-checking, input guardrails, and file path assertions to ensure robust error handling.

#### Use Case 
Please refer to the [PAF_PDF_Generation.ipynb](/notebooks/PAF_PDF_Generation.ipynb) notebook for the use case.

