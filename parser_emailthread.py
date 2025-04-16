import re
import extract_msg
import logging

logger = logging.getLogger(__name__)

def format_email_body(body: str) -> str:
    """
    Formats an email body by:
      - Normalizing newline characters (converting Windows-style "\r\n" and "\r" to Unix "\n"),
      - Collapsing multiple spaces/tabs into a single space,
      - Collapsing multiple newlines into a single blank line,
      - Trimming leading and trailing whitespace.
    """
    # Normalize newlines.
    body = body.replace('\r\n', '\n').replace('\r', '\n')
    # Collapse multiple spaces or tabs into one space.
    body = re.sub(r'[ \t]+', ' ', body)
    # Collapse multiple newlines into a single blank line.
    body = re.sub(r'\n\s*\n', '\n\n', body)
    formatted = body.strip()
    logger.debug(f"Formatted email body (length {len(formatted)})")
    return formatted

def parse_email(file_path: str) -> dict:
    """
    Parses an Outlook .msg file to extract metadata and the email body.
    
    If the email body contains a thread delimiter (e.g., "-----Original Message-----"),
    the body is split into multiple parts.
    
    Returns:
        dict: Contains:
            - "metadata": a dictionary containing email header details (From, To, Cc, Bcc, Date, Subject).
            - "body": either a single string or a list of strings if a thread is detected.
    """
    logger.debug(f"Starting parse_email for file: {file_path}")
    
    try:
        msg = extract_msg.Message(file_path)
    except Exception as e:
        logger.error(f"Could not parse file {file_path}: {e}")
        raise e

    metadata = {
        "From": msg.sender,
        "To": msg.to,
        "Cc": msg.cc,
        "Bcc": "",  # .msg files typically do not include Bcc information.
        "Date": msg.date,
        "Subject": msg.subject
    }
    logger.debug(f"Extracted metadata: {metadata}")
    
    # Extract the body. Use plain text if available; fall back to HTML if not.
    body = msg.body
    if not body and hasattr(msg, "htmlBody"):
        body = msg.htmlBody
        logger.debug("Using HTML body fallback.")
    else:
        if body:
            logger.debug(f"Raw MSG body length: {len(body)}")
        else:
            logger.debug("No body found in the .msg file.")
    
    formatted_body = format_email_body(body) if body else ""
    logger.debug(f"Formatted MSG body length: {len(formatted_body)}")
    
    # Check for an email thread delimiter.
    thread_delimiter = "-----Original Message-----"
    if thread_delimiter in formatted_body:
        # Split by the delimiter and remove any empty parts.
        thread_parts = [part.strip() for part in formatted_body.split(thread_delimiter) if part.strip()]
        logger.debug(f"Email thread detected. Found {len(thread_parts)} parts.")
        result_body = thread_parts
    else:
        result_body = formatted_body
    
    return {"metadata": metadata, "body": result_body}
