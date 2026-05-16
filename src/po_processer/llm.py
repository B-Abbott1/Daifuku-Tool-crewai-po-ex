import json
import os

from dotenv import load_dotenv
from openai import OpenAI

from po_processer.prompts import EXTRACTION_USER_TEMPLATE, SYSTEM_PROMPT

load_dotenv()

MODEL = "gpt-4o-mini"
MAX_ATTEMPTS = 3


def _validate_po_json(text: str) -> dict:
    data = json.loads(text)
    if not isinstance(data, dict):
        raise ValueError("Response must be a JSON object")
    if "header" not in data or "line_items" not in data:
        raise ValueError("Response must contain 'header' and 'line_items' keys")
    if not isinstance(data["line_items"], list):
        raise ValueError("'line_items' must be an array")
    return data


def extract_po_json(raw_po_text: str) -> str:
    """Extract PO header and line items as a JSON string via OpenAI."""
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY is not set")

    client = OpenAI(api_key=api_key)
    user_message = EXTRACTION_USER_TEMPLATE.format(raw_po_text=raw_po_text)
    last_error: Exception | None = None

    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            response = client.chat.completions.create(
                model=MODEL,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": user_message},
                ],
                response_format={"type": "json_object"},
                temperature=0,
            )
            content = response.choices[0].message.content
            if not content:
                raise ValueError("Empty response from model")
            _validate_po_json(content)
            return content
        except Exception as e:
            last_error = e
            if attempt < MAX_ATTEMPTS:
                print(f"[llm] Extraction attempt {attempt} failed, retrying...")

    raise RuntimeError(f"PO extraction failed after {MAX_ATTEMPTS} attempts: {last_error}")
