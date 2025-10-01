# Automated Excel Template
A smart Excel template specially built for financial teams to conduct consistent benchmarking and help their clients save costs.
It automates raw bank data transformation, applies data validation with dropdowns for smooth data entry, and allows instant highlighting of search results.
Designed for accuracy, speed, and ease of use, it eliminates repetitive manual work and keeps your analysis reliable every time clean data is imported into the Master Sheet.

---

## Table of Contents

- [About the Project](#about-the-project)  
- [Features](#features)  
- [Project Structure](#project-structure)  
- [Getting Started](#getting-started)  
  - [Prerequisites](#prerequisites)  
- [How to Use](#how-to-use)  
  - [1️⃣ Prepare & Paste Raw Data](#prepare--paste-raw-data)  
  - [2️⃣ Clean Data](#clean-data)  
  - [3️⃣ Import Clean Data](#import-clean-data)  
  - [4️⃣ Data Validation](#data-validation)  
  - [5️⃣ Search & Highlight](#search--highlight)  
- [Challenges & Fixes](#challenges--fixes)  
- [Download Resources](#download-resources)  
- [Snapshots](#snapshots)  
- [License](#license)  
- [Acknowledgements](#acknowledgements)  
- [Support](#support)

---

## About the Project
This project replicates a real-world fee benchmarking process for banks and financial consulting firms.  
It allows users to:
- Import and clean raw transaction data.  
- Map bank fee descriptions to standardized services.  
- Benchmark service charges against reference prices.  
- Automate repetitive steps through Excel VBA macros.

---

## Features
- Automated **Power Query refresh** for data cleaning.  
- **Data import** from Clean Data into the Master Sheet with one click.  
- **Data validation dropdowns** for standardized service names, descriptions, and prices.  
- **Benchmarking engine** that calculates variance and % deviation.  
- Quick **search & highlight** feature to find items instantly.  
- User-friendly buttons to trigger macros.

---

## Project Structure

```bash
├── Demo_Datasets
│   ├── Raw_Bank_Data.csv
│   └── Master_Sheet.csv
├── Snapshots
│   └── Master_Sheet_Snapshot.png   <-- 📸 Screenshot of the Master Sheet
├── Macros
│   └── VBA_Code.pdf
├── README.md
└── User_Guide.pdf
```
---

## Getting Started

### Prerequisites
- Microsoft Excel (2016 or later recommended).
- Macros must be enabled
- Developer tab activated in Excel
- Power Query installed and configured
   
---

## How to Use

### 1️⃣ Prepare & Paste Raw Data
- Copy and paste your transactions into the **Raw_Bank_Data** sheet.  
- Replace any old entries with new ones to keep the sheet clean and up to date.

### 2️⃣ Clean Data
- Click the **Clean Data** button to automatically refresh Power Query.  
- This step processes and prepares the raw data for import.

### 3️⃣ Import Clean Data
- Click the **Import Data** button to load the cleaned dataset into the **Master Sheet**.  
- Existing rows shift down so the latest records always stay at the top.

### 4️⃣ Data Validation
- Use the **Data Validation** button to apply dropdown lists on key columns.  
- This ensures consistent entry of service names, descriptions, and other fields.

### 5️⃣ Search & Highlight
- Type any keyword in the **Search Bar (E3)** and press the **Search Button**.  
- All matching rows in the Master Sheet will be highlighted in yellow for quick review.

---

  ## Challenges & Fixes

| **Challenge** | **Fix Implemented** |
|---------------|----------------------|
| New data imported into Master Sheet was replacing old rows | Inserted blank rows first so existing data (incl. formulas in J–N) shifted down before pasting new data |
| Dropdown lists (Data Validation) disappeared after import | Automated re-application of data validation with a trigger button |
| Power Query refresh macro failed with `Selection.ListObject.QueryTable.Refresh` | Fixed by explicitly referencing sheet (`Clean_Data_Sheet`) and table name (`Table3_2`) |
| `If Not IsError(benchPrice)` line flagged error in VBA | Corrected loop logic and added error handling for missing benchmark values |
| Search button was too slow when filtering | Switched from filtering to conditional highlighting for instant feedback |
| Needed search box with placeholder text (“Search here”) | Used VBA to show input message that disappears once typing begins |

---

## Download Resources

<a href= "https://github.com/brightboy373/Automated-Excel-Template-for-a-Financial-Consulting-Firm/blob/main/Raw%20American%20Bank%20Statement.xlsx">Download Sample Dataset</a>

<a href= "https://github.com/brightboy373/Automated-Excel-Template-for-a-Financial-Consulting-Firm/blob/main/Manual%20for%20Excel%20Template.pdf">Download Template Manual</a>

<a href= "https://github.com/brightboy373/Automated-Excel-Template-for-a-Financial-Consulting-Firm/blob/main/Macro%20documentations.pdf">Download Macro Documentations</a>

<a href= "https://github.com/brightboy373/Automated-Excel-Template-for-a-Financial-Consulting-Firm/blob/main/Virtual%20Admin%20Project.xlsm">Download Excel Automated Template</a>

---

## Snapshots
<img width="944" height="488" alt="Virtual Admin Excel Project" src="https://github.com/user-attachments/assets/8a8eb394-d0c9-4b2e-b7b8-4e99ce9c100c" />


---

## License

---

## Acknowledgements
Built and designed by Bright Ihechukwu Enwere in Lagos, Nigeria 🇳🇬
Inspired by real-world financial reconciliation needs and designed for speed, accuracy, and simplicity.

---

## Support

For questions or troubleshooting, contact: Bright via

<a href= "enwerebright@gmail.com">enwerebright@gmail.com</a>
 
<a href= "https://github.com/brightboy373">github</a>

<a href= "https://www.linkedin.com/in/brightenwere">linkedin</a>








