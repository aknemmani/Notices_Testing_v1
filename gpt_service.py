# testing_app/gpt_service.py

import base64
import json
import traceback

from openai import OpenAI

from config import (
    OPENAI_API_KEY,
    OPENAI_GPT_5_1_MODEL,
    OPENAI_GPT_5_MINI_MODEL,
)

client = OpenAI(api_key=OPENAI_API_KEY)

# Shared prompt so all models follow the same rules as Gemini testing
COMMON_TESTING_PROMPT = """
You are a utility notice analyzer. Analyze the attached PDF (the file contains the full notice).

Extract ONLY the following fields and return VALID JSON:

"vendor_account_number": "...",
"vendor_name": "...",
"service_address": "...",
"notice_category": "One of: Late Notice, Maintenance, Address Change, Cheque Received, Disconnect Notice, Rate Change, Revert to Owner, 3rd Party Audit, Others",
"notice_date": "...",
"impact_date": "...",
"impact_amount": "..."

Rules:
- If text mentions 'inspection' or 'testing' treat the notice_category as "Maintenance".
- If any field is missing in the notice, set its value to "NA".
- Only if notice_category is "Late Notice" OR "Disconnect Notice":
  - Extract "impact_date" and "impact_amount" from the text if present.
  Otherwise set both "impact_date" and "impact_amount" to "NA".
- Use plain strings only (no arrays, no nested objects).
- Return ONLY valid JSON, with no explanations or markdown.
"""


def _normalize_testing_output(raw: str) -> dict:
    """
    Take raw text from GPT and coerce it into the testing dict
    with sane defaults and Gemini-compatible behavior.
    """
    raw = raw.strip()
    if raw.startswith("```json"):
        raw = raw.replace("```json", "").replace("```", "")
    elif raw.startswith("```"):
        raw = raw.replace("```", "")

    data = json.loads(raw)

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
            data[k] = str(data[k]).strip()

    # If category is not Late/Disconnect, force impact fields to NA
    category = data.get("notice_category", "Others")
    if category not in ["Late Notice", "Disconnect Notice"]:
        data["impact_date"] = "NA"
        data["impact_amount"] = "NA"

    return data


def _extract_testing_fields_from_pdf_gpt(pdf_bytes: bytes, model_name: str) -> dict:
    """
    Core GPT testing extractor used by both GPT‑5.1 and GPT‑5‑mini.
    Sends the PDF as base64 data via Responses API and returns normalized dict.
    """
    print(f"🧪 Starting GPT TESTING extraction (model={model_name})...")

    try:
        # Base64 encode PDF as data URL, as per OpenAI file-input docs[1]
        b64_data = base64.b64encode(pdf_bytes).decode("utf-8")
        data_url = f"data:application/pdf;base64,{b64_data}"

        response = client.responses.create(
            model=model_name,
            input=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "input_text",
                            "text": COMMON_TESTING_PROMPT,
                        },
                        {
                            "type": "input_file",
                            "filename": "notice.pdf",
                            "file_data": data_url,
                        },
                    ],
                }
            ],
        )  #[2][1]

        # Responses API: output_text is the plain text answer[3][4]
        raw = response.output_text  # type: ignore[attr-defined]
        data = _normalize_testing_output(raw)
        print(f"✓ GPT TESTING extraction completed (model={model_name})")
        return data

    except Exception as e:
        print(f"❌ Error during GPT TESTING extraction (model={model_name}): {e}")
        traceback.print_exc()
        return {
            "vendor_account_number": "NA",
            "vendor_name": "NA",
            "service_address": "NA",
            "notice_category": "Others",
            "notice_date": "NA",
            "impact_date": "NA",
            "impact_amount": "NA",
        }


def extract_testing_fields_from_pdf_gpt_5_1(pdf_bytes: bytes) -> dict:
    """
    Public helper for GPT 5.1 testing.
    """
    return _extract_testing_fields_from_pdf_gpt(pdf_bytes, OPENAI_GPT_5_1_MODEL)


def extract_testing_fields_from_pdf_gpt_5_mini(pdf_bytes: bytes) -> dict:
    """
    Public helper for GPT 5‑mini testing.
    """
    return _extract_testing_fields_from_pdf_gpt(pdf_bytes, OPENAI_GPT_5_MINI_MODEL)
