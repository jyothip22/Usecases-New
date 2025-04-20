import re
import os
import tempfile
import extract_msg
import logging
from PyPDF2 import PdfReader  # pip install PyPDF2

logger = logging.getLogger(__name__)

def format_email_body(body: str) -> str:
    """
    Normalize newlines, collapse spaces/tabs, collapse multiple blank lines,
    and trim leading/trailing whitespace.
    """
    body = body.replace('\r\n', '\n').replace('\r', '\n')
    body = re.sub(r'[ \t]+', ' ', body)
    body = re.sub(r'\n\s*\n', '\n\n', body)
    return body.strip()

def extract_pdf_text(pdf_path: str) -> str:
    """
    Extract all textual content from a PDF file.
    """
    logger.debug(f"Extracting text from PDF attachment: {pdf_path}")
    reader = PdfReader(pdf_path)
    text_chunks = []
    for page in reader.pages:
        page_text = page.extract_text()
        if page_text:
            text_chunks.append(page_text)
    return "\n".join(text_chunks).strip()

def parse_email(file_path: str) -> dict:
    """
    Parse an Outlook .msg file, extract:
      - metadata (From, To, Cc, Bcc, Date, Subject)
      - formatted body text
      - for each PDF attachment: extract its text and append to the body
    
    Returns:
        {
            "metadata": { ... },
            "body": "<email body>\\n\\n--- Attachment (foo.pdf) ---\\n<pdf text>..."
        }
    """
    logger.debug(f"Starting parse_email for: {file_path}")
    msg = extract_msg.Message(file_path)

    # 1) Metadata
    metadata = {
        "From":    msg.sender,
        "To":      msg.to,
        "Cc":      msg.cc,
        "Bcc":     "",            # .msg rarely contains Bcc
        "Date":    msg.date,
        "Subject": msg.subject
    }
    logger.debug(f"Extracted metadata: {metadata}")

    # 2) Body
    raw_body = msg.body or getattr(msg, "htmlBody", "") or ""
    body = format_email_body(raw_body)
    logger.debug(f"Formatted body length: {len(body)}")

    # 3) Attachments check
    attachments = msg.attachments or []
    if not attachments:
        logger.debug("No attachments found in this email.")
    else:
        for att in attachments:
            filename = att.longFilename or att.shortFilename or "attachment"
            ext = os.path.splitext(filename)[1].lower()
            logger.debug(f"Found attachment: {filename} (extension {ext})")
            
            if ext == ".pdf":
                # write attachment to temp file
                with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                    tmp.write(att.data)
                    tmp_path = tmp.name

                try:
                    pdf_text = extract_pdf_text(tmp_path)
                    if pdf_text:
                        separator = f"\n\n--- Attachment: {filename} (PDF) ---\n"
                        body += separator + pdf_text
                        logger.debug(f"Appended PDF text for {filename}, length {len(pdf_text)}")
                    else:
                        logger.debug(f"No text extracted from PDF {filename}")
                except Exception as e:
                    logger.error(f"Failed to extract PDF {filename}: {e}")
                finally:
                    try:
                        os.remove(tmp_path)
                    except OSError:
                        pass
            else:
                logger.debug(f"Skipping non-PDF attachment: {filename}")

    return {
        "metadata": metadata,
        "body": body
    }
