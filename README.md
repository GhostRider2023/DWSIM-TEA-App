# DWSIM TEA App — Equipment Cost Estimator

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://dwsim-tea-app-w4iyzufsdbvsftseruasto.streamlit.app/)

## 📌 Overview

**DWSIM TEA App** is a Techno-Economic Analysis (TEA) tool built with [Streamlit](https://streamlit.io) that estimates equipment purchase costs for chemical process simulations exported from [DWSIM](https://dwsim.org/).

It uses established cost-correlation methods (e.g., Turton et al., Seider et al.) to calculate:
- **Purchased Equipment Cost (Cp)**
- **Bare Module Cost (CBM)**
- **Total Capital Cost Estimation**

---

## 🚀 Features

- 📤 Upload DWSIM simulation exports (Excel format)
- 🔧 Automatic equipment type detection & mapping
- 💰 Cost correlations for pumps, compressors, heat exchangers, vessels, columns, and more
- 📊 Interactive cost breakdown visualizations (Plotly)
- 📥 Export cost reports

---

## 🗂️ Project Structure

```
DWSIM-TEA-App/
├── app.py                  # Main Streamlit application
├── cost_engine.py          # Core cost calculation engine
├── equipment_models.py     # Equipment-specific cost models
├── excel_parser.py         # DWSIM Excel file parser
├── requirements.txt        # Python dependencies
└── data/
    ├── dwsim_equipment_cost_dataset.json
    ├── dwsim_equipment_mapping.csv
    ├── equipment_cost_correlations.csv
    └── material_and_bare_module_factors.csv
```

---

## ⚙️ Local Setup

```bash
# 1. Clone the repository
git clone https://github.com/GhostRider2023/DWSIM-TEA-App.git
cd DWSIM-TEA-App

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run the app
streamlit run app.py
```

---

## ☁️ Deploy on Streamlit Cloud

1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Connect your GitHub account
3. Select `GhostRider2023/DWSIM-TEA-App`
4. Set **Main file path** to `app.py`
5. Click **Deploy!**

---

## 📦 Dependencies

| Package | Version |
|---------|---------|
| streamlit | ≥ 1.35.0 |
| pandas | ≥ 2.0.0 |
| openpyxl | ≥ 3.1.0 |
| plotly | ≥ 5.20.0 |

---

## 👩‍💻 Author

Developed as part of the **FOSSEE** (Free/Libre and Open Source Software for Education) initiative.

---

## 📄 License

This project is open-source and available under the [MIT License](LICENSE).
