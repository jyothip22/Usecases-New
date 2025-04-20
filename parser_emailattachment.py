import re
import os
import tempfile
import logging
import extract_msg
from analyzer import get_system_prompt, invoke_custom_api

logger = logging.getLogger(__name__)
TKD_NAME = os.getenv("TKD_NAME", "EmailMonitor1")


def format_email_body(body: str) -> str:
    """
    Normalize newline characters, collapse extra whitespace, and trim.
    """
    body = body.replace('\r\n', '\n').replace('\r', '\n')
    body = re.sub(r'[ \t]+', ' ', body)
    body = re.sub(r'\n\s*\n', '\n\n', body)
    return body.strip()


def parse_email(file_path: str) -> dict:
    """
    Parses an Outlook .msg file to extract metadata, body, and analyze nested .msg attachments.
    """
    logger.debug(f"Parsing .msg file: {file_path}")
    msg = extract_msg.Message(file_path)

    # 1) Extract metadata
    metadata = {
        "From": msg.sender,
        "To": msg.to,
        "Cc": msg.cc,
        "Bcc": "",
        "Date": msg.date,
        "Subject": msg.subject
    }

    # 2) Extract and format body
    raw_body = msg.body or getattr(msg, 'htmlBody', '') or ''
    body = format_email_body(raw_body)
    logger.debug(f"Formatted body length: {len(body)}")

    # 3) Analyze nested .msg attachments
    nested_results = []
    for att in msg.attachments:
        # only consider .msg attachments
        filename = att.longFilename or att.shortFilename or ''
        if filename.lower().endswith('.msg'):
            # write attachment to temp file
            with tempfile.NamedTemporaryFile(delete=False, suffix='.msg') as tmp:
                tmp.write(att.data)
                tmp_path = tmp.name
            try:
                # parse nested email
                nested = parse_email(tmp_path)
                nested_body = nested.get('body', '')
                # analyze via LLM
                system_prompt = get_system_prompt()
                analysis = invoke_custom_api(TKD_NAME, nested_body, system_prompt)
                nested_results.append({
                    'metadata': nested.get('metadata', {}),
                    'analysis': analysis
                })
                logger.debug(f"Nested email '{filename}' analyzed.")
            except Exception as e:
                logger.error(f"Error processing nested .msg '{filename}': {e}")
            finally:
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass

    result = {
        'metadata': metadata,
        'body': body
    }
    if nested_results:
        result['nested_emails'] = nested_results

    return result
