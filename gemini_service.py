# gemini_service.py
import json
import traceback
from config import GEMINI_API_KEY
from google import genai
from google.genai import types

# Initialize client
client = genai.Client(api_key=GEMINI_API_KEY)


def analyze_notice_from_pdf(pdf_bytes: bytes):
    """
    Analyze a notice by sending the PDF bytes directly to Gemini.
    Returns the same structured JSON as before:
    {
      "short_summary": ...,
      "detailed_summary": ...,
      "category": ...,
      "key_entities": {...},
      "criticality": ...,
      "criticality_reason": ...,
      "required_actions": ...
    }
    """
    print("🤖 Starting Gemini analysis (PDF input)...")

    prompt = """
You are a utility notice analyzer. Analyze the PDF provided (the file contains the full notice).
Provide the analysis in the following JSON format:
{
  "detailed_summary": "Detailed summary explaining all key points",
  "category": "One of: Late Notice, Maintenance, Address Change, Cheque Received, Disconnect Notice, Rate Change, Revert to Owner, 3rd Party Audit, Others",
  "key_entities": {
    "Entity Name 1": "Value 1",
    "Entity Name 2": "Value 2"
  },
  "criticality": "High, Medium, or Low",
  "criticality_reason": "Brief explanation for the criticality level",
  "required_actions": "Single string with all actions separated by newlines or semicolons"
}

Important:
- Category MUST be exactly one of the specified types. Categorize the notice based on the content of the whole notice. Be accurate in categorising it.
- Extract 5-10 most important key_entities; include mandatory key entities like vendor name, service address, account number, notice date and if these entities are not present put "N/A" for that specific entity.
- Use proper case for entity names.
- Always include required_actions.
- If text mentions 'inspection' or 'testing' treat as Maintenance.
Return ONLY valid JSON (no markdown).
"""


#   "short_summary": "2-3 sentence summary here",
    try:
        # Attach PDF as bytes + prompt
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=[
                types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
                prompt
            ],
        )

        raw = getattr(response, "text", None) or str(response)
        raw = raw.strip()

        # strip code fences if present
        if raw.startswith("```json"):
            raw = raw.replace("```json", "").replace("```", "").strip()
        elif raw.startswith("```"):
            raw = raw.replace("```", "").strip()

        result = json.loads(raw)

        # normalize/ensure required fields
        defaults = {
            # 'short_summary': 'Summary not available',
            'detailed_summary': 'Detailed analysis not available',
            'category': 'Others',
            'key_entities': {},
            'criticality': 'Medium',
            'required_actions': 'No action specified'
        }
        for k, v in defaults.items():
            if k not in result or result[k] is None:
                result[k] = v

        if not isinstance(result.get('key_entities'), dict):
            result['key_entities'] = {}

        print("✓ Gemini PDF analysis completed")
        return result

    except Exception as e:
        print(f"❌ Error during Gemini PDF analysis: {e}")
        traceback.print_exc()
        return {
            # 'short_summary': 'Error analyzing notice',
            'detailed_summary': '',
            'category': 'Others',
            'key_entities': {},
            'criticality': 'Medium',
            'required_actions': 'Manual review required'
        }

def extract_testing_fields_from_pdf(pdf_bytes: bytes) -> dict:
        """
        Testing-only extractor: returns just the fields needed for Excel comparison.
        Does NOT affect the main analyze_notice_from_pdf flow.
        """
        print("🧪 Starting Gemini TESTING extraction (PDF input)...")

        prompt = """
        You are a utility notice analyzer. Analyze the attached PDF (the file contains the full notice).

        Extract ONLY the following fields and return VALID JSON:

        {
        "vendor_account_number": "...",
        "vendor_name": "...",
        "service_address": "...",
        "notice_category": "One of: Maintenance, Address Change, Cheque Received, Disconnect Notice, Late Notice, Rate Change, Revert to Owner, 3rd Party Audit, Others",
        "notice_date": "...",
        "impact_date": "...",
        "impact_amount": "..."
        }

        Rules:
        - Extract the notice_date and impact_date in the specified format only (YYYY-MM-DD).
        - Categorize the notice with full accuracy keeping in mind the context of the whole notice. Don't categorize the notice based on just a single word or few words.
        - If text mentions 'inspection' or 'testing' treat the notice_category as "Maintenance".
        - If any field is missing in the notice, set its value to "NA".
        - Only if notice_category is "Late Notice" OR "Disconnect Notice":
        - Extract "impact_date" and "impact_amount" from the text if present. Otherwise set both "impact_date" and "impact_amount" to "NA".
        - Use plain strings only (no arrays, no nested objects).
        - Return ONLY valid JSON, with no explanations or markdown.
        """

        try:
            # Call Gemini (same pattern as analyze_notice_from_pdf)
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=[
                    types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
                    prompt
                ],
            )

            raw = getattr(response, "text", None) or str(response)
            raw = raw.strip()

            # Strip code fences if present
            if raw.startswith("```json"):
                raw = raw.replace("```json", "").replace("```", "").strip()
            elif raw.startswith("```"):
                raw = raw.replace("```", "").strip()

            data = json.loads(raw)

            # Ensure all keys exist and default to "NA"
            defaults = {
                "vendor_account_number": "NA",
                "vendor_name": "NA",
                "service_address": "NA",
                "notice_category": "Others",
                "notice_date": "NA",
                "impact_date": "NA",
                "impact_amount": "NA",
            }

            for k, v in defaults.items():
                if k not in data or data[k] is None or str(data[k]).strip() == "":
                    data[k] = v
                else:
                    # normalize to string
                    data[k] = str(data[k]).strip()

            # If category is not Late/Disconnect, force impact fields to "NA"
            category = data.get("notice_category", "Others")
            if category not in ["Late Notice", "Disconnect Notice"]:
                data["impact_date"] = "NA"
                data["impact_amount"] = "NA"

            print("✓ Gemini TESTING extraction completed")
            return data

        except Exception as e:
            print(f"❌ Error during Gemini TESTING extraction: {e}")
            traceback.print_exc()
            # In testing mode, still return a consistent dict
            return {
                "vendor_account_number": "NA",
                "vendor_name": "NA",
                "service_address": "NA",
                "notice_category": "Others",
                "notice_date": "NA",
                "impact_date": "NA",
                "impact_amount": "NA",
            }


def extract_maintenance_details_from_pdf(pdf_bytes: bytes):
    """
    Extract maintenance-specific details (vendor_list_url, service_type, location, etc.)
    by sending the PDF to Gemini directly.
    Returns the same dict structure you used before.
    """
    print("🔍 Extracting maintenance details (PDF input)...")

    prompt = """
Analyze this maintenance notice (PDF attached) and extract key details for finding qualified vendors.

CRITICAL TASK - FIND THE VENDOR LIST URL:
Look VERY CAREFULLY for any web address, URL, or link that points to a list of qualified/certified vendors, testers, or service providers.

Provide the response in JSON format:
{
    "vendor_list_url": "FULL URL to vendor/tester list - THIS IS THE MOST IMPORTANT FIELD - null if not found",
    "service_type": "Specific type of maintenance service (e.g., Backflow Testing, HVAC Repair, Electrical)",
    "location": "Full service address where work is needed",
    "city": "City name",
    "state": "State/Province (use 2-letter code like CA, TX)",
    "due_date": "MM/DD/YYYY or null",
    "urgency": "High/Medium/Low",
    "device_details": "Equipment info (manufacturer, model, serial number) or null",
    "certification_required": true or false,
    "special_requirements": "Any special compliance needs or null",
    "issuing_authority": "Organization issuing the notice or null"
}

Be exhaustive when looking for URLs (http, https, www, markdown links, partial broken links). If multiple URLs appear, choose the most likely vendor list URL. Return ONLY valid JSON.
"""

    try:
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=[
                types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
                prompt
            ],
            # generation_config=genai.types.GenerationConfig(
            #     temperature=0.0
            # )
        )

        raw = getattr(response, "text", None) or str(response)
        raw = raw.strip()
        if raw.startswith("```json"):
            raw = raw.replace("```json", "").replace("```", "").strip()
        elif raw.startswith("```"):
            raw = raw.replace("```", "").strip()

        maintenance_data = json.loads(raw)

        # Basic fallback if vendor_list_url missing: try to return None rather than "null"
        if maintenance_data.get("vendor_list_url") == "null":
            maintenance_data["vendor_list_url"] = None

        print("✓ Maintenance details extracted from PDF")
        return maintenance_data

    except Exception as e:
        print(f"❌ Error extracting maintenance details from PDF: {e}")
        traceback.print_exc()
        return None