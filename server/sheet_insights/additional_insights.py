import json
from sheet_insights.config import client
import os

ADDITIONAL_INSIGHTS_PROMPT = """
You are an expert data analyst with deep analytical skills.

You have been provided with:
1. Individual sheet-wise insights (5 insights per supplier/sheet)
2. General comparative insights (10 insights across all sheets)

Your task is to generate exactly 5 NEW and DEEPER insights that were NOT covered in the previous insights. Focus on:

- **Hidden patterns and correlations** between different metrics and suppliers
- **Seasonal trends and temporal patterns** that weren't highlighted before
- **Performance gaps and improvement opportunities** not mentioned previously
- **Risk factors and warning signals** that need attention
- **Operational efficiency insights** beyond what was already covered
- **Benchmarking and best practices** from top performers
- **Data quality issues and missing information** that impact analysis
- **Cross-supplier dependencies and relationships**
- **Predictive insights** based on current trends
- **Strategic recommendations** for supply chain optimization

Requirements:
- Each insight must be 15-25 words long
- Be specific with numbers, percentages, and supplier names
- Avoid repeating any information from the previous insights
- Focus on actionable intelligence and strategic value
- Include statistical observations and data-driven conclusions
- Highlight anomalies, outliers, and unexpected patterns

Return your answer as a **valid JSON list with exactly 5 strings**.
Return only JSON. No markdown, no prose, no explanations.

If there is not enough new data to generate meaningful additional insights, return: ["Insufficient new data patterns available for additional insights"].
"""


def generate_additional_insights(insights_path: str, general_insights_path: str, output_path: str = "additional-insights.json"):
    """
    Generate additional insights based on existing insights and general insights
    
    Args:
        insights_path: Path to the individual sheet insights JSON file
        general_insights_path: Path to the general insights JSON file
        output_path: Path to save the additional insights
    
    Returns:
        List of additional insights
    """
    try:
        # Load existing insights
        with open(insights_path, "r", encoding="utf-8") as f:
            sheet_insights = json.load(f)
        
        with open(general_insights_path, "r", encoding="utf-8") as f:
            general_insights = json.load(f)
        
        # Prepare the input for the AI model
        input_data = {
            "sheet_insights": sheet_insights,
            "general_insights": general_insights
        }
        
        input_text = json.dumps(input_data, indent=2)
        
        print(f"🤖 Calling AI model for additional insights generation...")
        
        response = client.chat.completions.create(
            model=os.getenv("AZURE_OPENAI_DEPLOYMENT"),
            messages=[
                {"role": "system", "content": "You are an expert business analyst specializing in supply chain and operational analytics."},
                {"role": "user", "content": ADDITIONAL_INSIGHTS_PROMPT + f"\n\nExisting Data:\n```json\n{input_text}\n```"}
            ],
            temperature=0.4,  # Slightly higher temperature for more creative insights
            max_tokens=2000
        )
        
        reply = response.choices[0].message.content.strip()
        print(f"📝 Raw AI response received")
        
        # Parse the JSON response
        try:
            additional_insights = json.loads(reply)
            
            # Validate that we got exactly 5 insights
            if not isinstance(additional_insights, list):
                raise ValueError("Response is not a list")

            if len(additional_insights) != 5:
                print(f"⚠️ Warning: Expected 5 insights, got {len(additional_insights)}")
            
            # Save to file
            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(additional_insights, f, indent=2, ensure_ascii=False)
            
            print(f"💾 Saved additional insights to: {output_path}")
            return additional_insights
            
        except json.JSONDecodeError as e:
            print(f"❌ Failed to parse AI response as JSON: {e}")
            # Save raw response for debugging
            with open("additional_insights_raw_output.txt", "w", encoding="utf-8") as f:
                f.write(reply)
            print(f"🔍 Raw response saved to additional_insights_raw_output.txt for debugging")
            return ["Error: Failed to generate valid additional insights"]
            
    except Exception as e:
        print(f"❌ Error in generate_additional_insights: {e}")
        return ["Error: Failed to generate additional insights"]


def get_additional_insights_summary(additional_insights_path: str):
    """
    Get a summary of the additional insights
    
    Args:
        additional_insights_path: Path to the additional insights JSON file
    
    Returns:
        Dictionary with summary information
    """
    try:
        if not os.path.exists(additional_insights_path):
            return {"error": "Additional insights file not found"}
        
        with open(additional_insights_path, "r", encoding="utf-8") as f:
            insights = json.load(f)
        
        return {
            "total_insights": len(insights),
            "insights": insights,
            "status": "available"
        }
        
    except Exception as e:
        return {"error": f"Failed to load additional insights: {str(e)}"}
