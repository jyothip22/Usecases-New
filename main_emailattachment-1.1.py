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
from parser import parse_email  # Parses .msg, handles nested .msg attachments
from analyzer import get_system_prompt, invoke_custom_api  # Custom API invocation

# Determine the absolute base directory (project folder) and set the archive folder.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ARCHIVE_FOLDER = os.path.join(BASE_DIR, "emails_archive")

if not os.path.isdir(ARCHIVE_FOLDER):
    raise FileNotFoundError(f"Archive folder not found: {ARCHIVE_FOLDER}")

# Define a toolkit/model name (optionally set via environment variable)
TKD_NAME = os.getenv("TKD_NAME", "EmailMonitor1")

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
# Endpoint 1: Analyze via Filename (GET)
# ---------------------------
@app.get("/analyze-email")
async def analyze_email_endpoint(
    filename: str = Query(..., description="Filename of the .msg email to analyze")
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
        # Parse the email (includes nested .msg attachments)
        email_data = parse_email(file_path)
        logger.debug(f"Parsed email data: {email_data}")

        # Extract the main email body
        email_body = email_data.get("body", "")
        logger.debug(f"Extracted email body (first 100 chars): {email_body[:100]}...")

        # Perform the main body analysis via custom API
        system_prompt = get_system_prompt()
        analysis = invoke_custom_api(TKD_NAME, email_body, system_prompt)
        logger.debug(f"Custom API analysis received (first 100 chars): {analysis[:100]}...")

        # Assemble result with metadata and main analysis
        result = {
            "metadata": email_data.get("metadata", {}),
            "analysis": analysis
        }

        # Include nested email analyses if any
        nested = email_data.get("nested_emails")
        if nested:
            result["nested_emails"] = nested
            logger.debug(f"Added nested_emails to result: {nested}")

        logger.debug(f"Assembled result: {result}")
    except Exception as e:
        logger.error(f"Exception in /analyze-email endpoint: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

    return JSONResponse(content=result)

# ---------------------------
# Endpoint 2: Analyze via Text Input (POST)
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
        analysis = invoke_custom_api(TKD_NAME, text_input, system_prompt)
        logger.debug(f"Custom API analysis (first 100 chars): {analysis[:100]}...")
        result = {"analysis": analysis}
        logger.debug(f"Assembled result for text analysis: {result}")
    except Exception as e:
        logger.error(f"Exception in /analyze-text endpoint: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

    return JSONResponse(content=result)

# ---------------------------
# Endpoint 3: Analyze via File Upload (POST)
# ---------------------------
@app.post("/analyze-file")
async def analyze_file_endpoint(
    file: UploadFile = File(..., description="Upload a .msg file for analysis")
):
    if not file.filename.lower().endswith(".msg"):
        raise HTTPException(status_code=400, detail="Only .msg files are supported")

    try:
        # Save the uploaded file to a temporary location
        with tempfile.NamedTemporaryFile(delete=False, suffix=".msg") as tmp:
            temp_file_path = tmp.name
            tmp.write(await file.read())
        logger.debug(f"Saved uploaded file to temporary path: {temp_file_path}")

        # Parse the email (handles nested .msg attachments)
        email_data = parse_email(temp_file_path)
        logger.debug(f"Parsed email data: {email_data}")

        # Extract and analyze the main body
        email_body = email_data.get("body", "")
        logger.debug(f"Extracted email body (first 100 chars): {email_body[:100]}...")

        system_prompt = get_system_prompt()
        analysis = invoke_custom_api(TKD_NAME, email_body, system_prompt)
        logger.debug(f"Custom API analysis received (first 100 chars): {analysis[:100]}...")

        # Assemble the response
        result = {
            "metadata": email_data.get("metadata", {}),
            "analysis": analysis
        }

        # Attach nested email analyses if present
        nested = email_data.get("nested_emails")
        if nested:
            result["nested_emails"] = nested
            logger.debug(f"Added nested_emails to file upload result: {nested}")

        logger.debug(f"Assembled result for file upload: {result}")
    except Exception as e:
        logger.error(f"Exception in /analyze-file endpoint: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        # Clean up temp file
        try:
            os.remove(temp_file_path)
            logger.debug(f"Removed temporary file: {temp_file_path}")
        except Exception as e:
            logger.error(f"Error removing temporary file: {str(e)}")

    return JSONResponse(content=result)

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000, log_level="info")
