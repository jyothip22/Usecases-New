def parse_analysis_field(data: dict) -> dict:
    """
    Extracts structured fields from LLM response in either of these formats:
    
    - Inline-numbered:
      '1. Result: ... 2. Category: ... 3. Explanation: ... 4. Citation: ...'
    
    - Newline-separated:
      'Result: ...\nCategory: ...\nExplanation: ...\nCitation: ...'
    """
    answer = data.get('answer')
    if not isinstance(answer, str):
        return {}

    parsed = {}

    # Try numbered format first: 1. Result: ... 2. Category: ...
    numbered_matches = re.findall(r"\d+\.\s*([\w\s]+):\s*(.*?)(?=\d+\.|$)", answer, re.DOTALL)

    if numbered_matches:
        for key, value in numbered_matches:
            key_clean = key.strip().lower().replace(' ', '_')
            parsed[key_clean] = value.strip()
            logger.debug("Parsed field '%s': '%s'", key_clean, parsed[key_clean])
        return parsed

    # Fallback: plain "Key: value" format with newlines
    line_matches = re.findall(r"([\w\s]+):\s*(.*?)(?=\n[\w\s]+:|$)", answer, re.DOTALL)
    for key, value in line_matches:
        key_clean = key.strip().lower().replace(' ', '_')
        parsed[key_clean] = value.strip()
        logger.debug("Parsed field '%s': '%s'", key_clean, parsed[key_clean])

    return parsed
