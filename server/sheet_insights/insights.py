import json
from sheet_insights.config import client
from pathlib import Path
import os
import time
import random


INSIGHT_PROMPT = """
You are a senior data analyst. Based on the json file content given in the input, generate exactly five concise insights per company as a raw JSON array of strings.
You must be fully accurate with dates, figures, and trends — no assumptions or extrapolations.

Instructions:
- The table contains monthly data in left-to-right order from January to June. Use exact month names from the table when describing trends or numbers.
- Do NOT infer trends beyond June, and do NOT confuse column positions or assume patterns.
- Each insight must highlight patterns, trends, gaps, or performance observations strictly from the actual data.
- Use clear, factual, and precise language — avoid vague phrases like "the data shows" or "it can be seen".
- Do NOT include object wrappers or labels like "insight"; only return a plain JSON array of 5 strings.

here's what the input looks like:
{
  "generatedOn": "2025-07-30",
  "kpiMetadata": {
    "unitDescriptions": {
      "accidents": "Number of safety incidents reported",
      "productionLossHrs": "Production hours lost due to supplier-caused material shortage",
      "okDeliveryPercent": "Percentage of OK deliveries based on ACMA standards",
      "trips": "Number of shipment trips completed per month",
      "quantityShipped": "Number of parts shipped by the supplier",
      "partsPerTrip": "Efficiency metric showing avg. parts shipped per trip",
      "vehicleTAT": "Average vehicle turnaround time at the plant (in hours)",
      "machineDowntimeHrs": "Machine breakdown time (in hours)",
      "machineBreakdowns": "Number of machine breakdowns"
    }
  },
  "accidents": {
    "Acute_Wiring": {
      "Jan": 0,
      "Feb": 0,
      "Mar": 0,
      "Apr": 0,
      "May": 0,
      "Jun": 0,
      "Jul": null,
      "Aug": null,
      "Sep": null,
      "Oct": null,
      "Nov": null,
      "Dec": null
    },
    "Ankita_Auto": {
      "Jan": null,
      "Feb": null,
      "Mar": 0,
      "Apr": 0,
      "May": 0,
      "Jun": 0,
      "Jul": null,
      "Aug": null,
      "Sep": null,
      "Oct": null,
      "Nov": null,
      "Dec": null
    },
    "CAM": {
      "Jan": 0,
      "Feb": 2,
      "Mar": 1,
      "Apr": 0,
      "May": 0,
      "Jun": null,
      "Jul": null,
      "Aug": null,
      "Sep": null,
      "Oct": null,
      "Nov": null,
      "Dec": null
    },
    "Daxter": {
      "Jan": 0,
      "Feb": 0,
      "Mar": 0,
      "Apr": 0,
      "May": 0,
      "Jun": null,
      "Jul": null,
      "Aug": null,
      "Sep": null,
      "Oct": null,
      "Nov": null,
      "Dec": 0
    },
    "JJ_Tecnoplast": {
      "Jan": null,
      "Feb": null,
      "Mar": 0,
      "Apr": 0,
      "May": 0,
      "Jun": 0,
      "Jul": null,
      "Aug": null,
      "Sep": null,
      "Oct": null,
      "Nov": null,
      "Dec": null
    },
    "Kamal": {
      "Jan": 0,
      "Feb": 0,
      "Mar": 0,
      "Apr": 1,
      "May": 0,
      "Jun": null,
      "Jul": null,
      "Aug": null,
      "Sep": null,
      "Oct": null,
      "Nov": null,
      "Dec": null
    },
    "Laxmi_SPRINGS": {
      "Jan": null,
      "Feb": null,
      "Mar": null,
      "Apr": 0,
      "May": 0,
      "Jun": 0,
      "Jul": null,
      "Aug": null,
      "Sep": null,
      "Oct": null,
      "Nov": null,
      "Dec": null
    },
    "Makarjyothi": {
      "Jan": 0,
      "Feb": 0,
      "Mar": 0,
      "Apr": 0,
      "May": 0,
      "Jun": 0,
      "Jul": null,
      "Aug": null,
      "Sep": null,
      "Oct": null,
      "Nov": null,
      "Dec": null
    },
    "Shree_Stamping": {
      "Jan": null,
      "Feb": null,
      "Mar": null,
      "Apr": null,
      "May": null,
      "Jun": null,
      "Jul": null,
      "Aug": null,
      "Sep": null,
      "Oct": null,
      "Nov": null,
      "Dec": null
    },
    "S_B_Precision_Springs": {
      "Jan": null,
      "Feb": 0,
      "Mar": 0,
      "Apr": 0,
      "May": 0,
      "Jun": 0,
      "Jul": null,
      "Aug": null,
      "Sep": null,
      "Oct": null,
      "Nov": null,
      "Dec": null
    },
    "Unique_Systems": {
      "Jan": null,
      "Feb": 0,
      "Mar": 0,
      "Apr": 0,
      "May": 0,
      "Jun": 0,
      "Jul": null,
      "Aug": null,
      "Sep": null,
      "Oct": null,
      "Nov": null,
      "Dec": null
    },
    "Victor_Engineers_ASAL": {
      "Jan": null,
      "Feb": 0,
      "Mar": 0,
      "Apr": 0,
      "May": 0,
      "Jun": 0,
      "Jul": null,
      "Aug": null,
      "Sep": null,
      "Oct": null,
      "Nov": null,
      "Dec": null
    }
  },
  "productionLossHrs": {
    "Acute_Wiring": {
      "Jan": 0,
      "Feb": 0,
      "Mar": 0,
      "Apr": 0,
      "May": 0,
      "Jun": 0,
      "Jul": null,
      "Aug": null,
      "Sep": null,
      "Oct": null,
      "Nov": null,
      "Dec": null
    },
    "Ankita_Auto": {
      "Jan": null,
      "Feb": null,
      "Mar": 0,
      "Apr": 0,
      "May": 0,
      "Jun": 0,
      "Jul": null,
      "Aug": null,
      "Sep": null,
      "Oct": null,
      "Nov": null,
      "Dec": null
    },
    "CAM": {
      "Jan": 0,
      "Feb": 0,
      "Mar": 0,
      "Apr": 0,
      "May": 0,
      "Jun": null,
      "Jul": null,
      "Aug": null,
      "Sep": null,
      "Oct": null,
      "Nov": null,
      "Dec": null
    },

As you can see the object is KPI wise and I need the output company wise, every kpi has company as values and every company has months as values and the months have the actual values.
But i need the output company wise with kpis as its value with each kpi with months and values

Only return the JSON array — no markdown, no formatting, no explanations.

Example output Dummy content:
  "CAM": [
    "Safety accidents peaked in February with 2 incidents, while January, May, and June reported zero accidents.",
    "OK delivery cycles percentage was lowest in April at 54% and highest in May at 80%.",
    "Production loss due to material shortage was consistently zero from January to June.",
    "The number of trips per month decreased from 26 in January to 19 in May.",
    "Machine breakdown hours were highest in April at 3.5 hours, with no breakdowns reported in January, February, or March."
  ],
  "Ankita Auto": [
    "No safety accidents were reported from March to June.",
    "Production loss due to material shortage was zero for every month from March to June.",
    "OK delivery cycles maintained a consistent 100% rate from March to June.",
    "The number of trips peaked at 42 in May, compared to 25 in both March and June.",
    "Machine breakdown hours dropped from 4 in March to 0 in June, with the number of machine breakdowns also decreasing from 0.5 to 0 over the same period."
  ],
"""




def get_insights():
    """Generate insights for a single sheet with retry logic"""

    kpi_path = os.path.abspath("results/final_supplier_kpis.json")
    output_path = 'results/insights.json'

    try:
        with open(kpi_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except json.JSONDecodeError as e:
        print(f"❌ JSON decode error in {kpi_path}: {e}")
        return None
    
    json_data_str = json.dumps(data, ensure_ascii=False, indent=2)


            
    deployment_name = os.getenv("AZURE_OPENAI_DEPLOYMENT")
    if not deployment_name:
        print(f"❌ AZURE_OPENAI_DEPLOYMENT environment variable not set")
        return None

    response = client.chat.completions.create(
    model=deployment_name,
    messages=[
        {"role": "system", "content": "You are a helpful data analyst. Respond quickly and concisely."},
        {"role": "user", "content": INSIGHT_PROMPT + f"\n```\n{json_data_str}\n```"}
    ],
    temperature=0.0,
    max_tokens=1500,
    timeout=30
    )
    reply = response.choices[0].message.content.strip()
    return json.loads(reply)