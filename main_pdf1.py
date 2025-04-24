import os
import uvicorn
import tempfile
import logging

from fastapi import FastAPI, HTTPException, Query, File, UploadFile
from fastapi.responses import JSONResponse

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
logger = logging.getLogger(__name__)

# Import custom modules
from parser import parse_email, parse_analysis_field, safe_json_load  # parser handles .msg & PDFs
from analyzer import get_system_prompt, invoke_TKD_api               # custom LLM/TKD API invocation

# Base directory and archive folder
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ARCHIVE_FOLDER = os.path.join(BASE_DIR, "emails_archive")
if not os.path.isdir(ARCHIVE_FOLDER):
    raise FileNotFoundError(f"Archive folder not found: {ARCHIVE_FOLDER}")

# Toolkit & model names
TKD_NAME = os.getenv("TKD_NAME", "EmailMonitor1")
LLM_NAME = os.getenv("LLM_NAME", "gpt-4")

app = FastAPI(title="Email Compliance Analyzer API", version="1.0")


# ---------------------------
# Endpoint 1: Analyze via Filename (GET)
# ---------------------------
@app.get("/analyze-email")
async def analyze_email_endpoint(
    filename: str = Query(..., description=".msg filename to analyze")
):
    file_path = os.path.join(ARCHIVE_FOLDER, filename)
    logger.debug("Analyzing file: %s", file_path)
    if not os.path.exists(file_path):
        raise HTTPException(404, "File not found")
    if not filename.lower().endswith('.msg'):
        raise HTTPException(400, "Only .msg files supported")

    try:
        # Parse the .msg (including PDFs & nested .msg)
        email_data = parse_email(file_path)
        logger.debug("Parsed email_data: %s", email_data)

        # Combine body + PDF texts
        combined_body = email_data.get('body', '')
        for pdf in email_data.get('pdf_attachments', []):
            combined_body += f"\n\nAttachment: {pdf['filename']}\n{pdf['content']}"
        logger.debug("Combined analysis input length: %d", len(combined_body))

        # Invoke TKD API
        system_prompt = get_system_prompt()
        raw = invoke_TKD_api(TKD_NAME, combined_body, system_prompt, LLM_NAME)
        analysis = safe_json_load(raw)
        if not isinstance(analysis, dict):
            raise ValueError("Invalid JSON from TKD API")
        parsed = parse_analysis_field(analysis)

        # Assemble result
        result = {
            'metadata':       email_data.get('metadata', {}),
            'classification': parsed.get('classification'),
            'category':       parsed.get('category'),
            'explanation':    parsed.get('explanation'),
            'context':        analysis.get('context')
        }
        if 'pdf_attachments' in email_data:
            result['pdf_attachments'] = email_data['pdf_attachments']
        if 'nested_emails' in email_data:
            result['nested_emails'] = email_data['nested_emails']

        logger.debug("Assembled /analyze-email result: %s", result)

    except Exception as e:
        logger.error("Error in /analyze-email: %s", e)
        raise HTTPException(500, str(e))

    return JSONResponse(content=result)


# ---------------------------
# Endpoint 2: Analyze via File Upload (POST)
# ---------------------------
@app.post("/analyze-file")
async def analyze_file_endpoint(
    file: UploadFile = File(..., description="Upload a .msg file for analysis")
):
    if not file.filename.lower().endswith(".msg"):
        raise HTTPException(400, "Only .msg files are supported")

    temp_path = None
    try:
        # 1) Save uploaded .msg to temp
        with tempfile.NamedTemporaryFile(delete=False, suffix='.msg') as tmp:
            temp_path = tmp.name
            tmp.write(await file.read())
        logger.debug("Saved upload to: %s", temp_path)

        # 2) Parse (extract body, PDFs, nested .msg)
        email_data = parse_email(temp_path)
        logger.debug("Parsed email_data: %s", email_data)

        # 3) Combine body + PDF texts
        combined_body = email_data.get('body', '')
        for pdf in email_data.get('pdf_attachments', []):
            combined_body += f"\n\nAttachment: {pdf['filename']}\n{pdf['content']}"
        logger.debug("Combined analysis input length: %d", len(combined_body))

        # 4) Invoke TKD API for main body
        system_prompt = get_system_prompt()
        raw_main = invoke_TKD_api(TKD_NAME, combined_body, system_prompt, LLM_NAME)
        analysis_main = safe_json_load(raw_main)
        if not isinstance(analysis_main, dict):
            raise ValueError("Invalid JSON from TKD API (main)")
        parsed_main = parse_analysis_field(analysis_main)

        # 5) Build main result
        result = {
            'metadata':       email_data.get('metadata', {}),
            'classification': parsed_main.get('classification'),
            'category':       parsed_main.get('category'),
            'explanation':    parsed_main.get('explanation'),
            'context':        analysis_main.get('context')
        }

        # 6) Include PDFs if present
        if 'pdf_attachments' in email_data:
            result['pdf_attachments'] = email_data['pdf_attachments']

        # 7) Analyze each nested .msg if present
        nested_inputs = email_data.get('nested_emails', [])
        if nested_inputs:
            nested_results = []
            for idx, nested in enumerate(nested_inputs, start=1):
                nested_body = nested.get('body', '')
                raw_n = invoke_TKD_api(TKD_NAME, nested_body, system_prompt, LLM_NAME)
                analysis_n = safe_json_load(raw_n)
                if not isinstance(analysis_n, dict):
                    logger.error("Nested #%d TKD API returned invalid JSON", idx)
                    continue
                parsed_n = parse_analysis_field(analysis_n)
                nested_results.append({
                    'metadata':       nested.get('metadata', {}),
                    'classification': parsed_n.get('classification'),
                    'category':       parsed_n.get('category'),
                    'explanation':    parsed_n.get('explanation'),
                    'context':        analysis_n.get('context')
                })
                logger.debug("Nested email #%d analyzed", idx)
            result['nested_emails'] = nested_results

        logger.debug("Assembled /analyze-file result: %s", result)

    except Exception as exc:
        logger.error("Error in /analyze-file: %s", exc)
        raise HTTPException(500, str(exc))
    finally:
        # cleanup temp file
        if temp_path and os.path.exists(temp
