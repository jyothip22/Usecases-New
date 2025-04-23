import re
import os
import shutil
import tempfile
import extract_msg
import logging
import datetime

logger = logging.getLogger(__name__)

def format_email_body(body: str) -> str:
    """
    Normalize and clean the email body:
      - Convert CRLF/CR to LF
      - Collapse multiple spaces/tabs/newlines
      - Trim leading/trailing whitespace
    """
    logger.debug("Formatting email body (raw length=%d)", len(body))
    body = body.replace('\r\n', '\n').replace('\r', '\n')
    body = re.sub(r'[ \t]+', ' ', body)
    body = re.sub(r'\n\s*\n', '\n\n', body)
    formatted = body.strip()
    logger.debug("Formatted body length=%d", len(formatted))
    return formatted

def parse_msg(file_path: str) -> dict:
    """
    Parse a .msg file:
      1) Extract metadata and format the main body.
      2) Detect any embedded .msg attachments, save them out,
         recursively parse each, then clean up.
    """
    logger.debug("Starting parse_msg: %s", file_path)
    msg = extract_msg.Message(file_path)
    try:
        # 1) Metadata
        metadata = {
            "From":    msg.sender,
            "To":      msg.to,
            "Cc":      msg.cc,
            "Bcc":     "",
            "Date":    msg.date
        }
        logger.debug("Metadata raw: %s", metadata)
        # Ensure Date is ISO string
        date_val = metadata["Date"]
        if isinstance(date_val, datetime.datetime):
            metadata["Date"] = date_val.isoformat()
        else:
            metadata["Date"] = str(date_val)
        logger.debug("Metadata final: %s", metadata)

        # 2) Body
        raw_body = msg.body or getattr(msg, "htmlBody", "") or ""
        logger.debug("Raw body length=%d", len(raw_body))
        body = format_email_body(raw_body) if raw_body else ""
        logger.debug("Formatted body length=%d", len(body))

        result = {"metadata": metadata, "body": body}

        # 3) Embedded .msg attachments
        attachments = msg.attachments or []
        logger.debug("Total attachments found: %d", len(attachments))
        nested_results = []

        for att in attachments:
            # this gives the true filename (with extension) used when saving
            filename = att.getFilename() or ""
            logger.debug("Inspecting attachment: %s", filename)

            # if it’s an embedded MSG, save it out
            if filename.lower().endswith(".msg"):
                logger.debug("Found nested .msg: %s", filename)

                # create a temp dir to hold the extracted attachment
                tmpdir = tempfile.mkdtemp()
                try:
                    # save the attachment into tmpdir (extractEmbedded=True to get a .msg)
                    save_type, paths = att.save(customPath=tmpdir, extractEmbedded=True)
                    # paths can be a str or list of str
                    if isinstance(paths, str):
                        paths = [paths]

                    for nested_path in paths:
                        logger.debug("Parsing nested file: %s", nested_path)
                        nested_data = parse_msg(nested_path)
                        nested_results.append(nested_data)
                        logger.debug("Nested parsed successfully: %s", nested_path)
                except Exception as e:
                    logger.error("Error saving/parsing nested .msg '%s': %s", filename, e)
                finally:
                    # clean up the entire temp dir
                    shutil.rmtree(tmpdir, ignore_errors=True)
                    logger.debug("Removed temp dir for nested attachments: %s", tmpdir)

        if nested_results:
            result["nested_emails"] = nested_results
            logger.debug("Added nested_emails count=%d", len(nested_results))

        logger.debug("Completed parse_msg: %s", file_path)
        return result

    finally:
        # ensure file handles are closed on Windows
        try:
            msg.close()
            logger.debug("Closed Message object: %s", file_path)
        except Exception:
            pass

def parse_email(file_path: str) -> dict:
    """
    Dispatch: only .msg supported.
    """
    logger.debug("In parse_email: %s", file_path)
    if file_path.lower().endswith(".msg"):
        return parse_msg(file_path)
    msg = "Unsupported file format; only .msg is supported."
    logger.error(msg)
    raise ValueError(msg)
