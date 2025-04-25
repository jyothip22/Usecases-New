import os
import uvicorn
import tempfile
import logging

from fastapi import FastAPI, HTTPException, Query, File, UploadFile
from fastapi.responses import JSONResponse
from pydantic import BaseModel

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
logger = logging.getLogger(__name__)

# Import custom modules
from parser import (
    parse_email,
    parse_analysis_field,
    safe_json_load
)  # parser handles .msg, PDFs, Word, PPT, Excel attachments
from analyzer import get_system_prompt, invoke_TKD_api  # custom LLM/TKD API invocation

# Base path to stored emails
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ARCHIVE_FOLDER = os.path.join(BASE_DIR, "emails_archive")
if not os.path.isdir(ARCHIVE_FOLDER):
    raise FileNotFoundError(f"Archive folder not found: {ARCHIVE_FOLDER}")

# Toolkit & model names
TKD_NAME = os.getenv("TKD_NAME", "EmailMonitor1")
LLM_NAME = os.getenv("LLM_NAME", "gpt-4")

app = FastAPI(
    title="Email Compliance Analyzer API",
    version="1.0"
)

# ---------------------------
# Endpoint: Analyze via Filename (GET)
# ---------------------------
@app.get("/analyze-email")
async def analyze_email_endpoint(
    filename: str = Query(..., description=".msg filename to analyze")
):
    file_path = os.path.join(ARCHIVE_FOLDER, filename)
    logger.debug("Analyzing file: %s", file_path)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File not found")
    if not filename.lower().endswith('.msg'):
        raise HTTPException(status_code=400, detail="Only .msg files supported")

    try:
        # Parse .msg (includes all attachments)
        email_data = parse_email(file_path)
        logger.debug("Parsed email_data: %s", email_data)

        # Combine main body + all attachment texts
        combined_body = email_data.get('body', '')
        # PDFs
        for pdf in email_data.get('pdf_attachments', []):
            combined_body += f"\n\nAttachment (PDF): {pdf['filename']}\n{pdf['content']}"
        # Word docs
        for doc in email_data.get('docx_attachments', []):
            combined_body += f"\n\nAttachment (DOCX): {doc['filename']}\n{doc['content']}"
        # PowerPoints
        for ppt in email_data.get('pptx_attachments', []):
            combined_body += f"\n\nAttachment (PPTX): {ppt['filename']}\n{ppt['content']}"
        # Excel
        for xls in email_data.get('excel_attachments', []):
            combined_body += f"\n\nAttachment (Excel): {xls['filename']}\n{xls['content']}"
        logger.debug("Combined analysis input length: %d", len(combined_body))

        # Invoke TKD API
        system_prompt = get_system_prompt()
        raw = invoke_TKD_api(TKD_NAME, combined_body, system_prompt, LLM_NAME)
        analysis = safe_json_load(raw)
        if not isinstance(analysis, dict):
            raise ValueError("Invalid JSON from TKD API")
        parsed = parse_analysis_field(analysis)

        # Build base response
        result = {
            'metadata':       email_data.get('metadata', {}),
            'classification': parsed.get('classification'),
            'category':       parsed.get('category'),
            'explanation':    parsed.get('explanation'),
            'context':        analysis.get('context')
        }
        # Include attachments in response
        for key in ('pdf_attachments', 'docx_attachments', 'pptx_attachments', 'excel_attachments', 'nested_emails'):
            if key in email_data:
                result[key] = email_data[key]

        logger.debug("Assembled result: %s", result)
    except Exception as e:
        logger.error("Error in /analyze-email: %s", e)
        raise HTTPException(status_code=500, detail=str(e))

    return JSONResponse(content=result)

# ---------------------------
# Endpoint: Analyze via File Upload (POST)
# ---------------------------
class TextAnalysisRequest(BaseModel):
    text_input: str

@app.post("/analyze-file")
async def analyze_file_endpoint(
    file: UploadFile = File(..., description="Upload a .msg file for analysis")
):
    if not file.filename.lower().endswith('.msg'):
        raise HTTPException(status_code=400, detail="Only .msg files supported")

    temp_path = None
    try:
        # Save uploaded file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.msg') as tmp:
            temp_path = tmp.name
            tmp.write(await file.read())
        logger.debug("Saved upload to: %s", temp_path)

        # Parse and combine
        email_data = parse_email(temp_path)
        logger.debug("Parsed email_data: %s", email_data)
        combined_body = email_data.get('body', '')
        for key, label in [
            ('pdf_attachments', 'PDF'),
            ('docx_attachments', 'DOCX'),
            ('pptx_attachments', 'PPTX'),
            ('excel_attachments', 'Excel')
        ]:
            for att in email_data.get(key, []):
                combined_body += f"\n\nAttachment ({label}): {att['filename']}\n{att['content']}"
        logger.debug("Combined analysis input length: %d", len(combined_body))

        # Invoke TKD API
        system_prompt = get_system_prompt()
        raw = invoke_TKD_api(TKD_NAME, combined_body, system_prompt, LLM_NAME)
        analysis = safe_json_load(raw)
        if not isinstance(analysis, dict):
            raise ValueError("Invalid JSON from TKD API")
        parsed = parse_analysis_field(analysis)

        result = {
            'metadata':       email_data.get('metadata', {}),
            'classification': parsed.get('classification'),
            'category':       parsed.get('category'),
            'explanation':    parsed.get('explanation'),
            'context':        analysis.get('context')
        }
        for key in ('pdf_attachments', 'docx_attachments', 'pptx_attachments', 'excel_attachments', 'nested_emails'):
            if key in email_data:
                result[key] = email_data[key]

        logger.debug("Assembled upload result: %s", result)
    except Exception as e:
        logger.error("Error in /analyze-file: %s", e)
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
                logger.debug("Removed temp file: %s", temp_path)
            except Exception as rm:
                logger.error("Error removing temp file: %s", rm)

    return JSONResponse(content=result)

# ---------------------------
# Endpoint: Analyze via Text Input (POST)
# ---------------------------
@app.post("/analyze-text")
async def analyze_text_endpoint(request: TextAnalysisRequest):
    text_input = request.text_input
    logger.debug("Received text input (first 100 chars): %s...", text_input[:100])
    if not text_input.strip():
        raise HTTPException(status_code=400, detail="Text input is empty.")

    try:
        system_prompt = get_system_prompt()
        raw = invoke_TKD_api(TKD_NAME, text_input, system_prompt, LLM_NAME)
        analysis = safe_json_load(raw)
        if not isinstance(analysis, dict):
            raise ValueError("Invalid JSON from TKD API")
        parsed = parse_analysis_field(analysis)

        result = {
            'classification': parsed.get('classification'),
            'category':       parsed.get('category'),
            'explanation':    parsed.get('explanation'),
            'context':        analysis.get('context')
        }
        logger.debug("Assembled text result: %s", result)
    except Exception as e:
        logger.error("Error in /analyze-text: %s", e)
        raise HTTPException(status_code=500, detail=str(e))

    return JSONResponse(content=result)

if __name__ == '__main__':
    uvicorn.run(app, host='0.0.0.0', port=8000, log_level='info')
