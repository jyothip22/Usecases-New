import re
import os
import tempfile
import extract_msg
import logging
import datetime

# Module-level logger
logger = logging.getLogger(__name__)


def format_email_body(body: str) -> str:
    """
    Formats an email body by normalizing newline characters,
    collapsing extra spaces and newlines, and ensuring clean, human-readable text.
    """
    logger.debug("Formatting email body (raw length=%d)", len(body))
    # Normalize newlines
    body = body.replace('\r\n', '\n').replace('\r', '\n')
    # Collapse spaces/tabs
    body = re.sub(r'[ \t]+', ' ', body)
    # Collapse blank lines
    body = re.sub(r'\n\s*\n', '\n\n', body)
    formatted = body.strip()
    logger.debug("Formatted email body (length=%d)", len(formatted))
    return formatted


def parse_msg(file_path: str) -> dict:
    """
    Parses an Outlook .msg file to extract metadata, formatted body,
    and recursively parse any attached .msg files.
    Ensures file handles are closed so temp files can be removed on Windows.
    """
    logger.debug("Starting parse_msg for file: %s", file_path)
    msg = extract_msg.Message(file_path)
    try:
        # 1) Extract metadata
        metadata = {
            "From": msg.sender,
            "To": msg.to,
            "Cc": msg.cc,
            "Bcc": "",
            "Date": msg.date
        }
        logger.debug("Extracted metadata before conversion: %s", metadata)

        # Convert Date to ISO string if datetime
        date_val = metadata.get("Date")
        if isinstance(date_val, datetime.datetime):
            logger.debug("Converting datetime object to ISO string: %s", date_val)
            metadata["Date"] = date_val.isoformat()
        else:
            metadata["Date"] = str(date_val)
        logger.debug("Final metadata: %s", metadata)

        # 2) Extract and format body
        raw_body = msg.body or getattr(msg, 'htmlBody', '') or ''
        logger.debug("Raw MSG body length: %d", len(raw_body))
        body = format_email_body(raw_body) if raw_body else ''
        logger.debug("Formatted MSG body length: %d", len(body))

        result = {"metadata": metadata, "body": body}

        # 3) Nested .msg attachments
        attachments = msg.attachments or []
        logger.debug("Total attachments found: %d", len(attachments))
        nested_results = []
        for att in attachments:
            # Use getFilename() to get the true filename, including extension
            filename = att.getFilename() or ''
            logger.debug("Inspecting attachment via getFilename(): %s", filename)

            # Only recurse into .msg attachments
            if filename.lower().endswith('.msg'):
                logger.debug("Found nested .msg attachment: %s", filename)
                # Write attachment to a temp file
                with tempfile.NamedTemporaryFile(delete=False, suffix='.msg') as tmp:
                    tmp.write(att.data)
                    tmp_path = tmp.name
                try:
                    # Recursively parse nested .msg
                    nested_data = parse_msg(tmp_path)
                    nested_results.append(nested_data)
                    logger.debug("Nested email parsed: %s", filename)
                except Exception as e:
                    logger.error("Error processing nested .msg '%s': %s", filename, e)
                finally:
                    # Clean up the temp file
                    try:
                        os.remove(tmp_path)
                        logger.debug("Removed nested temp file: %s", tmp_path)
                    except OSError as oe:
                        logger.error("Error removing nested temp file '%s': %s", tmp_path, oe)

        if nested_results:
            result['nested_emails'] = nested_results
            logger.debug("Added nested_emails: %d items", len(nested_results))

        logger.debug("Completed parse_msg for file: %s", file_path)
        return result

    finally:
        # Close to release any file handles
        try:
            msg.close()
            logger.debug("Closed Message object for file: %s", file_path)
        except Exception:
            pass


def parse_email(file_path: str) -> dict:
    """
    Dispatch: currently supports only .msg files via parse_msg().
    """
    logger.debug("In parse_email with file: %s", file_path)
    if file_path.lower().endswith('.msg'):
        return parse_msg(file_path)
    else:
        error_msg = "Unsupported file format. Only .msg files are supported."
        logger.error(error_msg)
        raise ValueError(error_msg)
