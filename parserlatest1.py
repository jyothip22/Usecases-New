import re
import os
import tempfile
import extract_msg
import logging
import datetime

logger = logging.getLogger(__name__)

def format_email_body(body: str) -> str:
    logger.debug("Formatting email body (raw length=%d)", len(body))
    body = body.replace('\r\n', '\n').replace('\r', '\n')
    body = re.sub(r'[ \t]+', ' ', body)
    body = re.sub(r'\n\s*\n', '\n\n', body)
    formatted = body.strip()
    logger.debug("Formatted email body (length=%d)", len(formatted))
    return formatted

def parse_msg(file_path: str) -> dict:
    logger.debug("Starting parse_msg for file: %s", file_path)
    msg = extract_msg.Message(file_path)
    try:
        # ... metadata and body extraction as before ...

        result = {"metadata": metadata, "body": body}

        # 3) Nested .msg attachments
        attachments = msg.attachments or []
        logger.debug("Total attachments found: %d", len(attachments))
        nested_results = []

        for att in attachments:
            filename = att.getFilename() or ""
            logger.debug("Inspecting attachment via getFilename(): %s", filename)

            # Only process embedded .msg
            if filename.lower().endswith(".msg"):
                logger.debug("Found nested .msg attachment: %s", filename)

                # Instead of tmp.write(att.data), use export()
                # att.data is a MSGFile instance :contentReference[oaicite:0]{index=0}
                embedded_msg: extract_msg.Message = att.data
                # Create a temp path and export the embedded message
                with tempfile.NamedTemporaryFile(delete=False, suffix=".msg") as tmp:
                    temp_path = tmp.name
                embedded_msg.export(temp_path)  # writes a proper .msg to disk :contentReference[oaicite:1]{index=1}
                logger.debug("Exported embedded .msg to: %s", temp_path)

                try:
                    # Recursively parse the newly written .msg
                    nested_data = parse_msg(temp_path)
                    nested_results.append(nested_data)
                    logger.debug("Nested email parsed: %s", filename)
                except Exception as e:
                    logger.error("Error processing nested .msg '%s': %s", filename, e)
                finally:
                    try:
                        os.remove(temp_path)
                        logger.debug("Removed nested temp file: %s", temp_path)
                    except OSError as oe:
                        logger.error("Error removing nested temp file '%s': %s", temp_path, oe)

        if nested_results:
            result["nested_emails"] = nested_results
            logger.debug("Added nested_emails: %d items", len(nested_results))

        logger.debug("Completed parse_msg for file: %s", file_path)
        return result

    finally:
        # Always close to release file handles
        try:
            msg.close()
            logger.debug("Closed Message object for file: %s", file_path)
        except Exception:
            pass

def parse_email(file_path: str) -> dict:
    logger.debug("In parse_email with file: %s", file_path)
    if file_path.lower().endswith(".msg"):
        return parse_msg(file_path)
    error = "Unsupported file format. Only .msg files are supported."
    logger.error(error)
    raise ValueError(error)
