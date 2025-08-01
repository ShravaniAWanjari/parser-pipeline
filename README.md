# Parser Pipeline

This repository contains a robust and efficient data parsing pipeline designed to process Microsoft Excel workbooks, extract key performance indicators (KPIs), and generate insightful summaries. The pipeline is built using **FastAPI** for a fast and reliable API, leveraging the power of **Azure OpenAI** for natural language processing and data structuring.

---

## Features

- **FastAPI Endpoint:** A simple and intuitive API endpoint (`/upload_file`) to upload your workbooks.

- **Data Preprocessing:** Automatically converts workbook sheets into clean CSV files, handling common discrepancies like null values and extra spaces.

- **Azure OpenAI Integration:** Utilizes the Azure OpenAI API to intelligently parse and structure data, transforming raw numbers into meaningful JSON objects.

- **KPI-Wise JSON Output:** Generates a `final_supplier_kpis.json` file with a clear, hierarchical structure:

  ```json
  {
    "KPI_Name_1": {
      "Company_A": {
        "Month_1": "Value",
        "Month_2": "Value"
      }
    }
  }
  ```

- **Company-Specific Insights:** The `insights.py` module analyzes the structured KPI data to provide a summary of each company's performance, highlighting **5 key points** per company.

- **General Comparative Summary:** The `general_summary.py` module performs a comparative study of all companies, generating **5-10 key points** of generalized insights in a `general-info.json` file.

- **Comprehensive API Response:** The final API response consolidates all generated information—structured KPIs, company-specific insights, and the general summary—for a complete and actionable output.

---

## How It Works

The pipeline follows a sequential process to turn raw workbook data into valuable insights:

1. **File Upload:** You upload a workbook (e.g., `.xlsx`, `.xls`) to the `/upload_file` endpoint.
2. **Data Cleaning:** The workbook is processed, and each sheet is converted to a CSV file. The data is cleaned to fix discrepancies and ensure a consistent format.
3. **KPI Extraction:** The cleaned CSV data for each sheet is sent to the Azure OpenAI API. The model extracts KPIs and structures them into the `final_supplier_kpis.json` file.
4. **Company Insights:** The `final_supplier_kpis.json` is passed to the `insights.py` script, which analyzes the data and summarizes each company's performance with 5 key points.
5. **General Summary:** The company-specific insights are then processed by `general_summary.py`. This module performs a comparative analysis and provides a high-level summary of all companies, capturing key trends in a `general-info.json` file.
6. **Final Response:** The FastAPI endpoint gathers all the generated information and returns it as a single, comprehensive JSON response.

---

## Getting Started

To get started, you'll need:

- A running instance of the FastAPI application.
- Access to the Azure OpenAI API with the necessary credentials.
- The required Python libraries listed in `requirements.txt`.

### API Endpoint

- **Endpoint:** `/upload_file`
- **Method:** `POST`
- **Request Body:** A multipart form data with the file to be uploaded.

```python
import requests

# Replace with your API URL
url = "http://localhost:8000/upload_file"
files = {'file': open('your_workbook.xlsx', 'rb')}

response = requests.post(url, files=files)
print(response.json())
```

### Example Output

The JSON response will contain the structured KPI data, company-wise insights, and a general summary, giving you a complete overview of the workbook's content.

```json
{
  "kpis": {
    "KPI_Name_1": { 
      "company_name":{"month": value, "month": value ,
     },
    "KPI_Name_2": { ... }
  },
{company :
        "Keypoint 1 for Company A.",
        "Keypoint 2 for Company A."
      
},
  "general_summary": [
    "Overall keypoint 1.",
    "Overall keypoint 2."
  ]
}
```

