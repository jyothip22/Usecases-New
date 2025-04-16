/*Summary
Normalization:
The function format_email_body() converts Windows (\r\n) and Mac (\r) newline characters to Unix (\n) for consistency, collapses multiple spaces/tabs, removes excess newline gaps, and trims the text.

Parsing:
The parse_email() function extracts metadata (sender, recipient, date, subject, etc.) using extract_msg. It then extracts and formats the email body.

Thread Handling:
It checks for a specific delimiter ("-----Original Message-----"). If found, it splits the formatted body into multiple parts, each assumed to represent an individual message within the thread. Each part is packaged as a dictionary (along with metadata) inside a list, which is returned under the key "thread". If no thread is detected, the body is returned as a single string under the key "body".

Logging:
Debug-level logging is used throughout the module to track progress and help with troubleshooting.

*/



import re
import extract_msg
import logging

# Create a logger for this module.
logger = logging.getLogger(__name__)

def format_email_body(body: str) -> str:
    """
    Formats an email body by:
      - Normalizing newline characters: converts Windows-style "\r\n" and Mac-style "\r" to Unix "\n".
      - Collapsing multiple spaces or tabs into a single space.
      - Collapsing multiple newlines into a single blank line.
      - Trimming leading and trailing whitespace.
    
    Args:
        body (str): The raw email body text.
        
    Returns:
        str: The formatted email body.
    """
    # Normalize newlines.
    body = body.replace('\r\n', '\n').replace('\r', '\n')
    # Collapse multiple spaces or tabs into a single space.
    body = re.sub(r'[ \t]+', ' ', body)
    # Collapse multiple newlines into a single blank line.
    body = re.sub(r'\n\s*\n', '\n\n', body)
    formatted = body.strip()
    logger.debug(f"Formatted email body (length {len(formatted)})")
    return formatted

def parse_email(file_path: str) -> dict:
    """
    Parses an Outlook .msg file to extract metadata and the email body.
    
    - It uses the extract_msg package to open the file and extract details.
    - It formats the email body by normalizing the newlines, spaces, and trimming extra whitespace.
    - If a thread delimiter is found (e.g., "-----Original Message-----"), it splits the email body
      into multiple parts. In that case, the result contains a key "thread" with a list of
      dictionaries (each with "metadata" and "body" for each thread message).
    - Otherwise, the result contains a "body" key with a single formatted string.
    
    Args:
        file_path (str): The full path to the .msg file.
        
    Returns:
        dict: A dictionary with keys:
              - "metadata": a dictionary with keys like From, To, Cc, Bcc, Date, Subject.
              - "body": a single string (if no thread is detected) or
              - "thread": a list of dictionaries, each with "metadata" and "body".
    """
    logger.debug(f"Starting parse_email for file: {file_path}")
    
    try:
        msg = extract_msg.Message(file_path)
    except Exception as e:
        logger.error(f"Error parsing file {file_path}: {e}")
        raise e

    metadata = {
        "From": msg.sender,
        "To": msg.to,
        "Cc": msg.cc,
        "Bcc": "",            # .msg files typically do not provide Bcc.
        "Date": msg.date,
        "Subject": msg.subject
    }
    logger.debug(f"Extracted metadata: {metadata}")
    
    # Get the raw body: prefer plain text; if empty, attempt HTML.
    body = msg.body
    if not body and hasattr(msg, "htmlBody"):
        body = msg.htmlBody
        logger.debug("Using HTML body fallback.")
    elif body:
        logger.debug(f"Raw MSG body length: {len(body)}")
    else:
        logger.debug("No body found in the .msg file.")
    
    formatted_body = format_email_body(body) if body else ""
    logger.debug(f"Formatted MSG body length: {len(formatted_body)}")
    
    # Check if this email contains a thread by looking for a delimiter.
    # Adjust the delimiter as needed. Commonly, Outlook threads include "-----Original Message-----".
    thread_delimiter = "-----Original Message-----"
    
    if thread_delimiter in formatted_body:
        # Split the text on the delimiter and remove empty parts.
        parts = [part.strip() for part in formatted_body.split(thread_delimiter) if part.strip()]
        logger.debug(f"Email thread detected. Found {len(parts)} parts.")
        # Create a list of thread messages.
        # In this example, we simply use the same metadata for each thread part.
        thread_messages = []
        for idx, part in enumerate(parts):
            logger.debug(f"Thread part {idx} length: {len(part)}")
            thread_messages.append({
                "metadata": metadata,  # Alternatively, extract per-part metadata if available.
                "body": part
            })
        return {"metadata": metadata, "thread": thread_messages}
    else:
        return {"metadata": metadata, "body": formatted_body}
