# Excel Bank Data Automation 
A smart Excel template specially built for financial teams to conduct consistent benchmarking and help their clients save costs.
It automates raw bank data transformation, applies data validation with dropdowns for smooth data entry, and allows instant highlighting of search results.
Designed for accuracy, speed, and ease of use, it eliminates repetitive manual work and keeps your analysis reliable every time clean data is imported into the Master Sheet.

---

## Table of Contents
1. [About the Project](#about-the-project)  
2. [Features](#features)  
3. [Project Structure](#project-structure)  
4. [Getting Started](#getting-started)  
   - [Prerequisites](#prerequisites)  
   - [Installation](#installation)  
5. [How to Use](#how-to-use)  
   - [Data Import](#data-import)  
   - [Data Validation](#data-validation)  
   - [Benchmark Processing](#benchmark-processing)  
   - [Search & Highlight](#search--highlight)  
6. [Macros Reference](#macros-reference)  
7. [Demo Dataset](#demo-dataset)  
8. [Contributing](#contributing)  
9. [License](#license)  
10. [Acknowledgments](#acknowledgments)

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
- Basic knowledge of enabling Macros (VBA).

### Installation
1. Clone this repository:
   ```bash
   git clone https://github.com/your-username/automated-excel-fee-benchmarking.git
2. Open the Excel file.
3. Enable Macros when prompted.
   
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

## Download Resources
<a href= "https://github.com/brightboy373/Automated-Excel-Template-for-a-Financial-Consulting-Firm/blob/main/Raw%20American%20Bank%20Statement.xlsx">Download Sample Dataset</a>
<a href= "https://github.com/brightboy373/Automated-Excel-Template-for-a-Financial-Consulting-Firm/blob/main/Manual%20for%20Excel%20Template.pdf">Download Template Manual</a>
<a href= "https://github.com/brightboy373/Automated-Excel-Template-for-a-Financial-Consulting-Firm/blob/main/Macro%20documentations.pdf">Download Macro Documentations</a>
<a href= "[https://github.com/brightboy373/Automated-Excel-Template-for-a-Financial-Consulting-Firm/blob/main/Macro%20documentations.pdf](https://github.com/brightboy373/Automated-Excel-Template-for-a-Financial-Consulting-Firm/blob/main/Virtual%20Admin%20Project.xlsm)">Download Excel Automated Template</a>




