import re
import os
import tempfile
import datetime

try:
    import extract_msg
except ImportError:
    raise ImportError("The extract_msg package is required for parsing .msg files. Install it via 'pip install extract_msg'.")

logger = __import__('logging').getLogger(__name__)


def format_email_body(body: str) -> str:
    """
    Formats an email body by normalizing newline characters,
    collapsing extra spaces and newlines, and ensuring clean, human-readable text.
    """
    # Normalize newlines to Unix-style.
    body = body.replace('\r\n', '\n').replace('\r', '\n')
    # Collapse multiple spaces or tabs into a single space.
    body = re.sub(r'[ \t]+', ' ', body)
    # Collapse multiple newlines into a single blank line.
    body = re.sub(r'\n\s*\n', '\n\n', body)
    return body.strip()


def parse_msg(file_path: str) -> dict:
    """
    Parses an Outlook .msg file using extract_msg, returns metadata, body,
    and recursively parses any attached .msg files (nested_emails).
    """
    print(f"DEBUG: Starting parse_msg for file: {file_path}")
    msg = extract_msg.Message(file_path)

    # 1) Metadata
    metadata = {
        "From": msg.sender,
        "To": msg.to,
        "Cc": msg.cc,
        "Bcc": "",  # .msg typically lacks Bcc
        "Date": msg.date
    }
    print(f"DEBUG: Extracted metadata before conversion: {metadata}")

    # Convert Date to ISO string if datetime
    date_val = metadata.get("Date")
    if isinstance(date_val, datetime.datetime):
        print(f"DEBUG: Converting datetime object {date_val} to string.")
        metadata["Date"] = date_val.isoformat()
    else:
        metadata["Date"] = str(date_val)
    print(f"DEBUG: Final metadata: {metadata}")

    # 2) Body
    body = msg.body or getattr(msg, "htmlBody", "") or ""
    if not body:
        print("DEBUG: No plain text body; body is empty.")
    else:
        print(f"DEBUG: Raw MSG body length: {len(body)}")
    formatted_body = format_email_body(body) if body else ""
    print(f"DEBUG: Formatted MSG body length: {len(formatted_body)}")

    result = {
        "metadata": metadata,
        "body": formatted_body
    }

    # 3) Nested .msg attachments
    nested = []
    for att in msg.attachments:
        name = att.longFilename or att.shortFilename or ''
        if name.lower().endswith('.msg'):
            print(f"DEBUG: Found nested .msg attachment: {name}")
            # write to temp file
            with tempfile.NamedTemporaryFile(delete=False, suffix='.msg') as tmp:
                tmp.write(att.data)
                tmp_path = tmp.name
            try:
                # recursively parse
                nested_data = parse_msg(tmp_path)
                nested.append(nested_data)
                print(f"DEBUG: Nested email parsed: {name}")
            except Exception as e:
                print(f"ERROR processing nested .msg '{name}': {e}")
            finally:
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass

    if nested:
        result["nested_emails"] = nested
        print(f"DEBUG: Added nested_emails: {len(nested)} items")

    print(f"DEBUG: Completed parse_msg for file: {file_path}")
    return result


def parse_email(file_path: str) -> dict:
    """
    Dispatch function: currently handles only .msg files via parse_msg().
    """
    print(f"DEBUG: In parse_email with file: {file_path}")
    if file_path.lower().endswith('.msg'):
        return parse_msg(file_path)
    else:
        error_msg = "Unsupported file format. Only .msg files are supported."
        print(f"ERROR: {error_msg}")
        raise ValueError(error_msg)
