import os
import uvicorn
import tempfile
import logging
rom fastapi import FastAPI, HTTPException, Query, File, UploadFile
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
from parser import parse_email, parse_analysis_field, safe_json_load  # parser handles .msg & PDF attachments
from analyzer import get_system_prompt, invoke_TKD_api                    # custom LLM/TKD API invocation

# Base path to stored emails
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ARCHIVE_FOLDER = os.path.join(BASE_DIR, "emails_archive")
if not os.path.isdir(ARCHIVE_FOLDER):
    raise FileNotFoundError(f"Archive folder not found: {ARCHIVE_FOLDER}")

# Toolkit & model names
tkd_name = os.getenv("TKD_NAME", "EmailMonitor1")
llm_name = os.getenv("LLM_NAME", "gpt-4")

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
        raise HTTPException(404, "File not found")
    if not filename.lower().endswith('.msg'):
        raise HTTPException(400, "Only .msg files supported")

    try:
        email_data = parse_email(file_path)
        logger.debug("Parsed email_data: %s", email_data)

        # Combine body + PDF content for analysis
        combined_body = email_data.get('body', '')
        for pdf in email_data.get('pdf_attachments', []):
            combined_body += f"\n\nAttachment: {pdf['filename']}\n{pdf['content']}"
        logger.debug("Combined analysis input length: %d", len(combined_body))

        # Invoke TKD API
        prompt = get_system_prompt()
        raw = invoke_TKD_api(tkd_name, combined_body, prompt, llm_name)
        analysis = safe_json_load(raw)
        if not isinstance(analysis, dict):
            raise ValueError("Invalid JSON from TKD API")
        parsed = parse_analysis_field(analysis)

        # Build response
        result = {
            'metadata':      email_data.get('metadata', {}),
            'classification': parsed.get('classification'),
            'category':       parsed.get('category'),
            'explanation':    parsed.get('explanation'),
            'context':        analysis.get('context')
        }
        if 'pdf_attachments' in email_data:
            result['pdf_attachments'] = email_data['pdf_attachments']
        if 'nested_emails' in email_data:
            result['nested_emails'] = email_data['nested_emails']

        logger.debug("Assembled result: %s", result)
    except Exception as e:
        logger.error("Error in analyze-email: %s", e)
        raise HTTPException(500, str(e))

    return JSONResponse(content=result)

# ---------------------------
# Other endpoints (analyze-text, analyze-file) would follow similar pattern,
# ensuring PDF content is appended and returned appropriately.

if __name__ == '__main__':
    uvicorn.run(app, host='0.0.0.0', port=8000, log_level='info')
