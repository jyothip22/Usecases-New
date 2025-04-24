import re
import os
import shutil
import tempfile
import extract_msg
import logging
import datetime
import json
from PyPDF2 import PdfReader  # pip install PyPDF2

# Module-level logger
logger = logging.getLogger(__name__)

def format_email_body(body: str) -> str:
    logger.debug("Formatting email body (raw length=%d)", len(body))
    body = body.replace('\r\n', '\n').replace('\r', '\n')
    body = re.sub(r'[ \t]+', ' ', body)
    body = re.sub(r'\n\s*\n', '\n\n', body)
    formatted = body.strip()
    logger.debug("Formatted email body (length=%d)", len(formatted))
    return formatted


def safe_json_load(raw):
    try:
        return json.loads(raw)
    except (json.JSONDecodeError, TypeError) as e:
        logger.debug("safe_json_load failed: %s", e)
        return raw


def parse_analysis_field(data: dict) -> dict:
    answer = data.get('answer')
    if not isinstance(answer, str):
        return {}
    matches = re.findall(r"([\w\s]+):\s*(.*?)(?=\n[\w\s]+:|$)", answer, re.DOTALL)
    parsed = {}
    for key, value in matches:
        key_clean = key.strip().lower().replace(' ', '_')
        parsed[key_clean] = value.strip()
        logger.debug("Parsed field '%s': '%s'", key_clean, parsed[key_clean])
    return parsed


def extract_pdf_text(pdf_path: str) -> str:
    logger.debug("Extracting text from PDF: %s", pdf_path)
    reader = PdfReader(pdf_path)
    texts = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            texts.append(text)
    combined = "\n".join(texts).strip()
    logger.debug("Extracted PDF text length=%d", len(combined))
    return combined


def parse_msg(file_path: str) -> dict:
    logger.debug("Starting parse_msg for file: %s", file_path)
    msg = extract_msg.Message(file_path)
    try:
        # 1) Metadata
        metadata = {
            "From": msg.sender,
            "To": msg.to,
            "Cc": msg.cc,
            "Bcc": "",
            "Date": msg.date
        }
        logger.debug("Raw metadata: %s", metadata)
        date_val = metadata["Date"]
        if isinstance(date_val, datetime.datetime):
            metadata["Date"] = date_val.isoformat()
        else:
            metadata["Date"] = str(date_val)
        logger.debug("Formatted metadata: %s", metadata)

        # 2) Body
        raw_body = msg.body or getattr(msg, "htmlBody", "") or ""
        logger.debug("Raw body length=%d", len(raw_body))
        body = format_email_body(raw_body) if raw_body else ""
        logger.debug("Formatted body length=%d", len(body))

        result = {"metadata": metadata, "body": body}
        pdf_texts = []
        nested_results = []

        # 3) Attachments
        attachments = msg.attachments or []
        logger.debug("Total attachments found: %d", len(attachments))
        for att in attachments:
            filename = att.getFilename() or ""
            logger.debug("Inspecting attachment: %s", filename)

            # PDF attachments
            if filename.lower().endswith('.pdf'):
                # save to temp file
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmpf:
                    tmpf.write(att.data)
                    tmp_pdf = tmpf.name
                try:
                    text = extract_pdf_text(tmp_pdf)
                    pdf_texts.append({"filename": filename, "content": text})
                except Exception as e:
                    logger.error("Error extracting PDF %s: %s", filename, e)
                finally:
                    os.remove(tmp_pdf)
                    logger.debug("Removed temp PDF file: %s", tmp_pdf)

            # Embedded .msg attachments
            if filename.lower().endswith('.msg'):
                logger.debug("Found nested .msg attachment: %s", filename)
                tmpdir = tempfile.mkdtemp()
                try:
                    save_type, paths = att.save(customPath=tmpdir, extractEmbedded=True)
                    if isinstance(paths, str):
                        paths = [paths]
                    for nested_path in paths:
                        logger.debug("Parsing nested .msg from: %s", nested_path)
                        nested_data = parse_msg(nested_path)
                        nested_results.append(nested_data)
                        logger.debug("Nested parsed successfully: %s", nested_path)
                except Exception as e:
                    logger.error("Error handling nested .msg '%s': %s", filename, e)
                finally:
                    shutil.rmtree(tmpdir, ignore_errors=True)
                    logger.debug("Removed temp dir: %s", tmpdir)

        if pdf_texts:
            result['pdf_attachments'] = pdf_texts
            logger.debug("Added pdf_attachments count=%d", len(pdf_texts))
        if nested_results:
            result['nested_emails'] = nested_results
            logger.debug("Added nested_emails count=%d", len(nested_results))

        logger.debug("Completed parse_msg for file: %s", file_path)
        return result
    finally:
        try:
            msg.close()
            logger.debug("Closed Message object: %s", file_path)
        except:
            pass


def parse_email(file_path: str) -> dict:
    logger.debug("In parse_email with file: %s", file_path)
    if file_path.lower().endswith('.msg'):
        return parse_msg(file_path)
    error_msg = "Unsupported file format. Only .msg files are supported."
    logger.error(error_msg)
    raise ValueError(error_msg)
