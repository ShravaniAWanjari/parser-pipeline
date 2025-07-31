import json
from sheet_insights.config import client
import os
from pathlib import Path

SUMMARY_PROMPT = """
You are an expert data analyst.

Given the following JSON object of sheet-wise insights, generate exactly 10 deep and comparative insights across all sheets. Make sure to :
- Keep the word limit 10-15 words per point.
- Be specific and grounded in the input data.
- Compare and contrast patterns across sheets where applicable.
- Include exact numbers, percentages, or statistical findings.
- Reveal underlying trends, anomalies, or correlations.
- Avoid generic or vague summaries.
- Skip sheets that only say "No data available".

Return your answer as a **valid JSON list with exactly 10 strings**.
Return only JSON. No markdown, no prose, no explanations, no bullet points.

If there is not enough data, return: ["Not enough data available"].
"""


def generate_general_insights(all_insights_path: str, output_path: str = "General-info.json"):
    # Check if general insights file already exists
    output_file = Path(output_path)
    if output_file.exists():
        print(f"⏭️ Skipping general insights generation - {output_file.name} already exists")
        with open(output_file, "r", encoding="utf-8") as f:
            return json.load(f)
    
    # Load insights data
    with open(all_insights_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    input_text = json.dumps(data, indent=2)

    response = client.chat.completions.create(
        model=os.getenv("AZURE_OPENAI_DEPLOYMENT"),
        messages=[
            {"role": "system", "content": "You are a helpful business analyst."},
            {"role": "user", "content": SUMMARY_PROMPT + f"\n```\n{input_text}\n```"}
        ],
        temperature=0.3,
        max_tokens=1200
    )
    reply = response.choices[0].message.content.strip()
    try:
        general = json.loads(reply)
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(general, f, indent=2)
        print(f"💾 Generated and saved general insights to: {output_path}")
        return general
    except json.JSONDecodeError:
        with open("general_summary_raw.txt", "w", encoding="utf-8") as f:
            f.write(reply)
        return []