import re
import logging

logger = logging.getLogger(__name__)

def parse_analysis_field(data: dict) -> dict:
    """
    Extracts structured fields from LLM response in either of these formats:

    - Numbered format:
        '1. Classification: ... 2. Category: ... 3. Explanation: ... 4. Citation: ...'

    - Plain format:
        'Classification: ...\nCategory: ...\nExplanation: ...\nCitation: ...'
    """
    answer = data.get('answer')
    if not isinstance(answer, str):
        return {}

    parsed = {}

    # First try numbered format: "1. Key: Value 2. Key: Value ..."
    numbered_matches = re.findall(
        r"\d+\.\s*([\w\s]+):\s*(.*?)(?=\n?\d+\.\s*[\w\s]+:|$)",
        answer,
        re.DOTALL
    )
    if numbered_matches:
        for key, value in numbered_matches:
            key_clean = key.strip().lower().replace(' ', '_')
            parsed[key_clean] = value.strip()
            logger.debug("Parsed field '%s': '%s'", key_clean, parsed[key_clean])
        return parsed

    # Fallback: Key: Value (newline-separated)
    fallback_matches = re.findall(
        r"([\w\s]+):\s*(.*?)(?=\n[\w\s]+:|$)",
        answer,
        re.DOTALL
    )
    for key, value in fallback_matches:
        key_clean = key.strip().lower().replace(' ', '_')
        parsed[key_clean] = value.strip()
        logger.debug("Parsed field '%s': '%s'", key_clean, parsed[key_clean])

    return parsed
