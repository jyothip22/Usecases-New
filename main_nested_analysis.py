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
from parser import parse_email  # handles nested .msg attachments and closes files
from analyzer import get_system_prompt, invoke_custom_api

# Base directory and archive folder
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ARCHIVE_FOLDER = os.path.join(BASE_DIR, "emails_archive")
if not os.path.isdir(ARCHIVE_FOLDER):
    raise FileNotFoundError(f"Archive folder not found: {ARCHIVE_FOLDER}")

# Toolkit/model name
TKD_NAME = os.getenv("TKD_NAME", "EmailMonitor1")

app = FastAPI(
    title="Email Compliance Analyzer API",
    version="1.0"
)

# ---------------------------
# Endpoint 1: Analyze via Filename
# ---------------------------
@app.get("/analyze-email")
async def analyze_email_endpoint(
    filename: str = Query(..., description=".msg filename to analyze")
):
    file_path = os.path.join(ARCHIVE_FOLDER, filename)
    logger.debug(f"Full file path: {file_path}")
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File not found")
    if not filename.lower().endswith(".msg"):
        raise HTTPException(status_code=400, detail="Only .msg files are supported")

    try:
        # Parse email (with nested attachments)
        email_data = parse_email(file_path)
        logger.debug(f"Parsed email data: {email_data}")
        
        # Analyze main body
        body = email_data.get("body", "")
        prompt = get_system_prompt()
        main_analysis = invoke_custom_api(TKD_NAME, body, prompt)
        logger.debug(f"Main analysis: {main_analysis[:100]}...")

        # Build response
        result = {"metadata": email_data.get("metadata", {}), "analysis": main_analysis}
        nested = email_data.get("nested_emails")
        if nested:
            # Analyze each nested email body
            nested_results = []
            for item in nested:
                nested_body = item.get("body", "")
                nested_meta = item.get("metadata", {})
                nested_analysis = invoke_custom_api(TKD_NAME, nested_body, prompt)
                nested_results.append({"metadata": nested_meta, "analysis": nested_analysis})
            result["nested_emails"] = nested_results
        
    except Exception as exc:
        logger.error(f"Error in /analyze-email: {exc}")
        raise HTTPException(status_code=500, detail=str(exc))

    return JSONResponse(content=result)

# ---------------------------
# Endpoint 2: Analyze via Text Input
# ---------------------------
class TextAnalysisRequest(BaseModel):
    text_input: str

@app.post("/analyze-text")
async def analyze_text(request: TextAnalysisRequest):
    text = request.text_input
    logger.debug(f"Received text: {text[:100]}...")
    if not text.strip():
        raise HTTPException(status_code=400, detail="Text input is empty.")
    try:
        prompt = get_system_prompt()
        analysis = invoke_custom_api(TKD_NAME, text, prompt)
        result = {"analysis": analysis}
    except Exception as exc:
        logger.error(f"Error in /analyze-text: {exc}")
        raise HTTPException(status_code=500, detail=str(exc))
    return JSONResponse(content=result)

# ---------------------------
# Endpoint 3: Analyze via File Upload
# ---------------------------
@app.post("/analyze-file")
async def analyze_file(file: UploadFile = File(..., description="Upload .msg file")):
    if not file.filename.lower().endswith(".msg"):
        raise HTTPException(status_code=400, detail="Only .msg files are supported")

    temp_path = None
    try:
        # save to temp
        with tempfile.NamedTemporaryFile(delete=False, suffix='.msg') as tmp:
            temp_path = tmp.name
            tmp.write(await file.read())
        logger.debug(f"Saved upload to {temp_path}")

        # parse and analyze
        email_data = parse_email(temp_path)
        body = email_data.get("body", "")
        prompt = get_system_prompt()
        main_analysis = invoke_custom_api(TKD_NAME, body, prompt)
        
        result = {"metadata": email_data.get("metadata", {}), "analysis": main_analysis}
        nested = email_data.get("nested_emails")
        if nested:
            nested_results = []
            for item in nested:
                nested_body = item.get("body", "")
                nested_meta = item.get("metadata", {})
                nested_analysis = invoke_custom_api(TKD_NAME, nested_body, prompt)
                nested_results.append({"metadata": nested_meta, "analysis": nested_analysis})
            result["nested_emails"] = nested_results
    except Exception as exc:
        logger.error(f"Error in /analyze-file: {exc}")
        raise HTTPException(status_code=500, detail=str(exc))
    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
                logger.debug(f"Removed temp file {temp_path}")
            except Exception as remove_err:
                logger.error(f"Error removing temp file: {remove_err}")

    return JSONResponse(content=result)

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
