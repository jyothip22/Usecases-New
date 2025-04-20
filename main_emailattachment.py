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

# Import custom modules.
from parser import parse_email        # parse_email includes nested .msg analysis
from analyzer import get_system_prompt, invoke_custom_api  # custom API invocation

# Determine base directory and archive folder
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ARCHIVE_FOLDER = os.path.join(BASE_DIR, "emails_archive")

if not os.path.isdir(ARCHIVE_FOLDER):
    raise FileNotFoundError(f"Archive folder not found: {ARCHIVE_FOLDER}")

# Toolkit/model name
tkd_name = os.getenv("TKD_NAME", "EmailMonitor1")

app = FastAPI(
    title="Email Compliance Analyzer API",
    description=(
        "Analyzes Outlook .msg email content for compliance and fraud detection "
        "by invoking a custom API with separate instructions and input text. "
        "Input can be provided via a file upload, a filename referencing an email stored locally, or plain text."
    ),
    version="1.0"
)

# ---------------------------
# Endpoint: Analyze via Filename
# ---------------------------
@app.get("/analyze-email")
async def analyze_email_endpoint(
    filename: str = Query(..., description=".msg filename to analyze")
):
    file_path = os.path.join(ARCHIVE_FOLDER, filename)
    logger.debug(f"Full file path: {file_path}")

    if not os.path.exists(file_path):
        logger.debug("File not found, raising 404.")
        raise HTTPException(status_code=404, detail="File not found")
    if not filename.lower().endswith(".msg"):
        logger.debug("File is not a .msg file, raising 400.")
        raise HTTPException(status_code=400, detail="Only .msg files are supported")

    try:
        # Parse top-level email (includes nested .msg attachments analysis)
        email_data = parse_email(file_path)
        logger.debug(f"Parsed email data (including nested): {email_data}")
        body = email_data.get("body", "")
        logger.debug(f"Extracted body (first 100 chars): {body[:100]}...")

        # Analyze main body
        system_prompt = get_system_prompt()
        analysis = invoke_custom_api(tkd_name, body, system_prompt)
        logger.debug(f"Custom API main analysis (first 100 chars): {analysis[:100]}...")

        # Build result
        result = {
            "metadata": email_data.get("metadata", {}),
            "analysis": analysis
        }

        # Include nested email analyses if present
        nested = email_data.get("nested_emails")
        if nested:
            result["nested_emails"] = nested
            logger.debug(f"Included nested_emails in result: {nested}")

        logger.debug(f"Final assembled result: {result}")
    except Exception as exc:
        logger.error(f"Exception in /analyze-email endpoint: {exc}")
        raise HTTPException(status_code=500, detail=str(exc))

    return JSONResponse(content=result)

# ---------------------------
# Endpoint: Analyze via Text Input
# ---------------------------
class TextAnalysisRequest(BaseModel):
    text_input: str

@app.post("/analyze-text")
async def analyze_text_endpoint(request: TextAnalysisRequest):
    text_input = request.text_input
    logger.debug(f"Received text input (first 100 chars): {text_input[:100]}...")

    if not text_input or text_input.strip() == "":
        raise HTTPException(status_code=400, detail="Text input is empty.")

    try:
        system_prompt = get_system_prompt()
        analysis = invoke_custom_api(tkd_name, text_input, system_prompt)
        logger.debug(f"Custom API analysis (first 100 chars): {analysis[:100]}...")
        result = {"analysis": analysis}
        logger.debug(f"Assembled result for text analysis: {result}")
    except Exception as exc:
        logger.error(f"Exception in /analyze-text endpoint: {exc}")
        raise HTTPException(status_code=500, detail=str(exc))

    return JSONResponse(content=result)

# ---------------------------
# Endpoint: Analyze via File Upload
# ---------------------------
@app.post("/analyze-file")
async def analyze_file_endpoint(
    file: UploadFile = File(..., description="Upload a .msg file for analysis")
):
    if not file.filename.lower().endswith(".msg"):
        raise HTTPException(status_code=400, detail="Only .msg files are supported")

    try:
        # Save upload to temp file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".msg") as tmp:
            tmp_path = tmp.name
            content = await file.read()
            tmp.write(content)
        logger.debug(f"Saved uploaded file to: {tmp_path}")

        # Parse (includes nested attachments)
        email_data = parse_email(tmp_path)
        logger.debug(f"Parsed email data (upload): {email_data}")
        body = email_data.get("body", "")
        logger.debug(f"Extracted body (first 100 chars): {body[:100]}...")

        # Main analysis
        system_prompt = get_system_prompt()
        analysis = invoke_custom_api(tkd_name, body, system_prompt)
        logger.debug(f"Custom API main analysis (first 100 chars): {analysis[:100]}...")

        result = {
            "metadata": email_data.get("metadata", {}),
            "analysis": analysis
        }
        nested = email_data.get("nested_emails")
        if nested:
            result["nested_emails"] = nested
            logger.debug(f"Included nested_emails in file-upload result: {nested}")

        logger.debug(f"Final assembled upload result: {result}")
    except Exception as exc:
        logger.error(f"Exception in /analyze-file endpoint: {exc}")
        raise HTTPException(status_code=500, detail=str(exc))
    finally:
        try:
            os.remove(tmp_path)
            logger.debug(f"Removed temporary file: {tmp_path}")
        except Exception as e:
            logger.error(f"Error removing temp file: {e}")

    return JSONResponse(content=result)

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000, log_level="info")
