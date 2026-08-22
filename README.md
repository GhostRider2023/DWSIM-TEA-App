# DWSIM TEA App — Equipment Cost Estimator

## 📌 Overview

**DWSIM TEA App** is a Techno-Economic Analysis (TEA) application developed as part of the **FOSSEE DWSIM** project.

The application uses **Streamlit** to provide a user-friendly interface for estimating equipment purchase and capital costs from DWSIM process simulation data.

The application uses established chemical engineering cost correlations, including methods based on **Turton et al.** and **Seider et al.**

The application calculates:

- **Purchased Equipment Cost (Cp)**
- **Bare Module Cost (CBM)**
- **Total Capital Cost**
- Equipment-wise cost breakdowns
- Interactive cost visualizations

---

# 🚀 Features

- Upload DWSIM simulation data
- Read DWSIM Excel exports
- Automatic equipment identification
- Equipment type mapping
- Equipment cost estimation
- Cost correlations for supported equipment
- Purchased equipment cost calculation
- Bare module cost calculation
- Total capital cost estimation
- Interactive Plotly visualizations
- Export cost estimation reports
- Local DWSIM API integration

---

# ⚠️ Important Requirement

This application requires **DWSIM to be installed on the user's computer**.

The application uses the **DWSIM API** from the local DWSIM installation.

Therefore, every user must configure the **DWSIM API path according to their own computer**.

The DWSIM API path used by one user may not work on another user's computer.

**Do not copy the developer's DWSIM API path.**

---

# 💻 System Requirements

Before running the application, install the following:

- Python 3.9 or later
- Git
- DWSIM
- A modern web browser

The application is primarily intended to be run locally using Streamlit.

---

# 📥 Installation

## Step 1 — Clone the Repository

Open **Command Prompt**, **PowerShell**, or a terminal.

Run:

    git clone https://github.com/GhostRider2023/DWSIM-TEA-App.git

After cloning, move into the repository:

    cd DWSIM-TEA-App

---

# Step 2 — Enter the Application Folder

The final application is located inside:

    FOSSEE DWSIM FINAL SUBMISION

Enter the directory using:

    cd "FOSSEE DWSIM FINAL SUBMISION"

The complete path should now be:

    DWSIM-TEA-App/FOSSEE DWSIM FINAL SUBMISION

All commands from this point should be executed from this directory unless otherwise specified.

---

# Step 3 — Verify Python Installation

Check that Python is installed:

    python --version

You should see something similar to:

    Python 3.9.x

or:

    Python 3.10.x

or:

    Python 3.11.x

or a newer supported Python version.

If the `python` command does not work, try:

    python3 --version

If Python is not installed, install Python before continuing.

---

# Step 4 — Create a Virtual Environment

Creating a virtual environment is recommended.

On Windows:

    python -m venv venv

Activate the environment:

    venv\Scripts\activate

On Linux/macOS:

    python3 -m venv venv

Activate the environment:

    source venv/bin/activate

After activation, your terminal should normally show:

    (venv)

at the beginning of the command prompt.

---

# Step 5 — Install Python Dependencies

Make sure you are inside:

    DWSIM-TEA-App/FOSSEE DWSIM FINAL SUBMISION

Install the required packages:

    pip install -r requirements.txt

If `pip` does not work, use:

    python -m pip install -r requirements.txt

The required dependencies are specified in the project's `requirements.txt` file.

---

# ⚙️ DWSIM API Configuration

## Step 6 — Install DWSIM

DWSIM must be installed on the same computer where the TEA application is being executed.

The application requires access to the DWSIM API.

If DWSIM is already installed, continue to the next step.

---

# Step 7 — Find Your DWSIM Installation Directory

You must find the location where DWSIM is installed.

For example, DWSIM may be installed in a directory similar to:

    C:\Program Files\DWSIM

However, this is only an example.

Your installation may be located somewhere else.

For example:

    C:\DWSIM

or:

    D:\Software\DWSIM

or another directory selected during DWSIM installation.

### Windows

To find the DWSIM installation:

1. Open the Windows Start Menu.
2. Search for `DWSIM`.
3. Right-click the DWSIM application.
4. Select `Open file location`.
5. Right-click the DWSIM shortcut.
6. Select `Properties`.
7. Check the target/installation location.
8. Locate the DWSIM API files required by the application.

---

# Step 8 — Locate the DWSIM API

The application requires the appropriate DWSIM API files from your DWSIM installation.

The exact API location depends on:

- Your DWSIM version
- Your operating system
- Your DWSIM installation directory
- How DWSIM was installed

Therefore, do not assume that the API path shown in an example is correct for your computer.

Use the API location associated with your local DWSIM installation.

---

# Step 9 — Configure the DWSIM API Path

Open the Python source file in the application where the DWSIM API path is defined.

Locate the DWSIM API path configuration.

It should be configured using your own local path.

For example:

    DWSIM_API_PATH = r"YOUR_DWSIM_API_PATH"

Replace:

    YOUR_DWSIM_API_PATH

with the actual API path on your computer.

For example, if your API is located at:

    C:\MySoftware\DWSIM\API

you would configure:

    DWSIM_API_PATH = r"C:\MySoftware\DWSIM\API"

The above path is only an example.

**Use the actual DWSIM API path on your computer.**

---

# ⚠️ Very Important: DWSIM API Path

The DWSIM API path is **machine-specific**.

For example, the developer may have DWSIM installed at:

    C:\Program Files\DWSIM

Another user may have DWSIM installed at:

    D:\Applications\DWSIM

Another user may have:

    C:\DWSIM

All of these are valid depending on the individual installation.

Therefore:

- Do not copy the developer's API path.
- Do not copy another user's API path.
- Do not assume DWSIM is installed in `C:\Program Files`.
- Find the DWSIM installation on your own computer.
- Configure the application using your own API path.

---

# ▶️ Running the Application

## Step 10 — Start the Streamlit Application

After:

1. Cloning the repository
2. Entering the `FOSSEE DWSIM FINAL SUBMISION` directory
3. Creating the virtual environment
4. Installing the dependencies
5. Installing DWSIM
6. Locating your DWSIM API
7. Configuring your DWSIM API path

run:

    streamlit run app.py

If the `streamlit` command is not recognized, run:

    python -m streamlit run app.py

---

# 🌐 Step 11 — Open the Application

After starting Streamlit, the terminal will display a local address.

Typically:

    http://localhost:8501

Open this address in your browser.

The DWSIM TEA application should now be available locally.

---

# 📊 Using the Application

## Step 12 — Prepare Your DWSIM Simulation

Open DWSIM and create or open your process simulation.

Make sure the simulation contains the required equipment and process information.

---

## Step 13 — Export the DWSIM Simulation

Export the required simulation information using the format supported by the application.

The application is designed to process DWSIM simulation data exported in Excel format.

---

## Step 14 — Upload the DWSIM Data

Open the Streamlit application in your browser.

Upload the required DWSIM Excel file through the application's upload interface.

---

## Step 15 — Equipment Detection

The application reads the uploaded simulation data and identifies the supported equipment.

Equipment is mapped to the appropriate equipment categories used by the cost estimation engine.

---

## Step 16 — Equipment Cost Calculation

The application applies the appropriate cost correlation for each supported equipment type.

The calculated results include:

- Purchased Equipment Cost
- Bare Module Cost
- Equipment-wise costs
- Total Capital Cost

---

## Step 17 — Review the Results

The application provides interactive visualizations for reviewing the equipment cost distribution.

The results can be used to compare the estimated costs of different pieces of equipment.

---

## Step 18 — Export the Results

If the application provides an export option, use it to download the generated equipment cost report.

---

# 🗂️ Project Structure

The application is located in:

    DWSIM-TEA-App/
    └── FOSSEE DWSIM FINAL SUBMISION/

The main project structure is:

    DWSIM-TEA-App/
    │
    └── FOSSEE DWSIM FINAL SUBMISION/
        │
        ├── app.py
        ├── cost_engine.py
        ├── equipment_models.py
        ├── excel_parser.py
        ├── requirements.txt
        │
        └── data/
            ├── dwsim_equipment_cost_dataset.json
            ├── dwsim_equipment_mapping.csv
            ├── equipment_cost_correlations.csv
            └── material_and_bare_module_factors.csv

---

# 📄 Main Files

## app.py

The main Streamlit application.

It provides the user interface and connects the different components of the application.

---

## cost_engine.py

Contains the core cost calculation logic used by the application.

---

## equipment_models.py

Contains equipment-specific cost models and calculations.

---

## excel_parser.py

Handles parsing of DWSIM Excel data.

---

## requirements.txt

Contains the Python dependencies required to run the application.

---

## data/

Contains the datasets and cost correlation information used by the application.

---

# 💰 Cost Estimation

The application implements established chemical engineering cost estimation correlations.

The major calculated quantities include:

## Purchased Equipment Cost — Cp

The estimated cost of purchasing the equipment before applying the appropriate installation and bare-module factors.

---

## Bare Module Cost — CBM

The estimated equipment cost after accounting for relevant factors such as:

- Material of construction
- Installation
- Pressure
- Other applicable bare-module factors

---

## Total Capital Cost

The application combines the estimated equipment costs to provide an overall capital cost estimate.

---

# 📦 Dependencies

The project dependencies are specified in:

    requirements.txt

Install them with:

    pip install -r requirements.txt

The application uses packages including:

- Streamlit
- Pandas
- OpenPyXL
- Plotly

---

# 🐛 Troubleshooting

## Problem: DWSIM API Not Found

If the application cannot find the DWSIM API:

1. Verify that DWSIM is installed.
2. Find the DWSIM installation directory.
3. Locate the required API files.
4. Check the API path configured in the application.
5. Make sure the path is the path from your own computer.
6. Check for spelling mistakes in the path.
7. Restart the Streamlit application.

---

## Problem: Streamlit Is Not Recognized

If you receive an error similar to:

    'streamlit' is not recognized as an internal or external command

install Streamlit:

    pip install streamlit

Then run:

    streamlit run app.py

Alternatively:

    python -m streamlit run app.py

---

## Problem: Python Is Not Recognized

Check Python:

    python --version

If that does not work:

    python3 --version

Make sure Python is installed and added to the system PATH.

---

## Problem: Requirements Installation Fails

First upgrade pip:

    python -m pip install --upgrade pip

Then run:

    pip install -r requirements.txt

---

## Problem: Port 8501 Is Already in Use

If Streamlit reports that port `8501` is already being used, run the application using another port:

    streamlit run app.py --server.port 8502

Then open:

    http://localhost:8502

---

# ⚡ Complete Quick Start

For a user who already has **Python, Git, and DWSIM installed**, the complete process is:

## 1. Clone the repository

    git clone https://github.com/GhostRider2023/DWSIM-TEA-App.git

## 2. Enter the repository

    cd DWSIM-TEA-App

## 3. Enter the application directory

    cd "FOSSEE DWSIM FINAL SUBMISION"

## 4. Create a virtual environment

    python -m venv venv

## 5. Activate the virtual environment

### Windows

    venv\Scripts\activate

### Linux/macOS

    source venv/bin/activate

## 6. Install dependencies

    pip install -r requirements.txt

## 7. Configure your DWSIM API path

Update the DWSIM API configuration in the application using the path to the DWSIM API on your own computer.

Example:

    DWSIM_API_PATH = r"YOUR_DWSIM_API_PATH"

## 8. Run the application

    streamlit run app.py

## 9. Open the application

    http://localhost:8501

---

# 🔗 Repository

The application is available at:

https://github.com/GhostRider2023/DWSIM-TEA-App/tree/master/FOSSEE%20DWSIM%20FINAL%20SUBMISION

---

# ⚠️ Important User Checklist

Before running the application, make sure:

- [ ] Python is installed.
- [ ] Git is installed.
- [ ] DWSIM is installed.
- [ ] The repository has been cloned.
- [ ] The `FOSSEE DWSIM FINAL SUBMISION` directory has been opened.
- [ ] The Python virtual environment has been created.
- [ ] The virtual environment has been activated.
- [ ] `requirements.txt` has been installed.
- [ ] The DWSIM API location has been identified.
- [ ] The DWSIM API path has been configured.
- [ ] The DWSIM API path belongs to your own computer.
- [ ] The Streamlit application has been started.
- [ ] `http://localhost:8501` has been opened in a browser.

---

# 📝 Notes

1. The application requires a local DWSIM installation.
2. The application requires access to the DWSIM API.
3. The DWSIM API path must be configured by each user.
4. The API path is machine-specific.
5. The example API paths in this README are placeholders only.
6. Cost estimates are based on the cost correlations implemented in the application.
7. The calculated costs should be considered engineering estimates and not vendor quotations.

---

# 👩‍💻 Author

Developed as part of the **FOSSEE (Free/Libre and Open Source Software for Education)** initiative.

---

# 📄 License

This project is open-source and available under the **MIT License**.
