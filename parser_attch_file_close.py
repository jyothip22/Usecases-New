import re
import os
import tempfile
import extract_msg
import logging
import datetime

logger = logging.getLogger(__name__)


def format_email_body(body: str) -> str:
    """
    Formats an email body by normalizing newline characters,
    collapsing extra spaces and newlines, and ensuring clean, human-readable text.
    """
    body = body.replace('\r\n', '\n').replace('\r', '\n')
    body = re.sub(r'[ \t]+', ' ', body)
    body = re.sub(r'\n\s*\n', '\n\n', body)
    return body.strip()


def parse_msg(file_path: str) -> dict:
    """
    Parses an Outlook .msg file to extract metadata, formatted body,
    and recursively parse any attached .msg files.
    Ensures file handles are closed so temp files can be removed on Windows.
    """
    print(f"DEBUG: Starting parse_msg for file: {file_path}")
    msg = extract_msg.Message(file_path)
    try:
        # 1) Extract metadata
        metadata = {
            "From": msg.sender,
            "To":   msg.to,
            "Cc":   msg.cc,
            "Bcc":  "",
            "Date": msg.date
        }
        print(f"DEBUG: Extracted metadata before conversion: {metadata}")

        # Convert Date to ISO string if it's a datetime
        date_val = metadata.get("Date")
        if isinstance(date_val, datetime.datetime):
            print(f"DEBUG: Converting datetime object {date_val} to string.")
            metadata["Date"] = date_val.isoformat()
        else:
            metadata["Date"] = str(date_val)
        print(f"DEBUG: Final metadata: {metadata}")

        # 2) Extract body
        raw_body = msg.body or getattr(msg, 'htmlBody', '') or ''
        if raw_body:
            print(f"DEBUG: Raw MSG body length: {len(raw_body)}")
        else:
            print("DEBUG: No body found; defaulting to empty string.")
        body = format_email_body(raw_body) if raw_body else ''
        print(f"DEBUG: Formatted MSG body length: {len(body)}")

        result = {"metadata": metadata, "body": body}

        # 3) Handle nested .msg attachments
        nested = []
        for att in msg.attachments:
            name = att.longFilename or att.shortFilename or ''
            if name.lower().endswith('.msg'):
                print(f"DEBUG: Found nested .msg attachment: {name}")
                # Write to temp file
                with tempfile.NamedTemporaryFile(delete=False, suffix='.msg') as tmp:
                    tmp.write(att.data)
                    tmp_path = tmp.name
                try:
                    # Recursively parse nested email
                    nested_data = parse_msg(tmp_path)
                    nested.append(nested_data)
                    print(f"DEBUG: Nested email parsed: {name}")
                except Exception as e:
                    print(f"ERROR processing nested .msg '{name}': {e}")
                finally:
                    # Ensure nested temp file is closed and removed
                    try:
                        os.remove(tmp_path)
                        print(f"DEBUG: Removed nested temp file: {tmp_path}")
                    except OSError as oe:
                        print(f"ERROR removing nested temp file '{tmp_path}': {oe}")
        if nested:
            result["nested_emails"] = nested
            print(f"DEBUG: Added nested_emails: {len(nested)} items")

        print(f"DEBUG: Completed parse_msg for file: {file_path}")
        return result

    finally:
        # Always close the message file handle to allow deletion
        try:
            msg.close()
            print(f"DEBUG: Closed Message object for file: {file_path}")
        except Exception:
            pass


def parse_email(file_path: str) -> dict:
    """
    Dispatch function: handles only .msg files via parse_msg().
    """
    print(f"DEBUG: In parse_email with file: {file_path}")
    if file_path.lower().endswith('.msg'):
        return parse_msg(file_path)
    else:
        error_msg = "Unsupported file format. Only .msg files are supported."
        print(f"ERROR: {error_msg}")
        raise ValueError(error_msg)
