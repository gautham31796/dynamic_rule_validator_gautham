import logging
import os
import re
import json
from jsonpath_ng import parse as jp_parse

# Logging setup
log_file = os.path.join(os.path.dirname(__file__), 'rule_engine.log')
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

if logger.hasHandlers():
    logger.handlers.clear()

file_handler = logging.FileHandler(log_file, mode='w')
file_handler.setLevel(logging.DEBUG)

console_handler = logging.StreamHandler()
console_handler.setLevel(logging.DEBUG)

formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

logger.addHandler(file_handler)
logger.addHandler(console_handler)

# Normalize text for comparison (not used in strict match but kept for condition logic)
def normalize(text):
    return re.sub(r'[^a-z0-9]', '', str(text).lower())

# Resolve placeholder using jsonpath
def resolve_placeholder(json_data, key, field_map):
    json_path = field_map.get(key, key)
    try:
        expr = jp_parse(json_path if json_path.startswith('$') else f'$.{json_path}')
        logger.debug(f"[TRACE] Scanning JSON for <{key}> using path: {json_path}")
        logger.debug(f"[TRACE] JSON sample: {json.dumps(json_data, indent=2)[:500]}")
        matches = expr.find(json_data)
        logger.debug(f"[TRACE] Matches found: {matches}")
        if not matches:
            logger.warning(f"No matches for <{key}> with path '{json_path}'")
            return ""
        raw_value = matches[0].value
        return ", ".join(map(str, raw_value)) if isinstance(raw_value, list) else str(raw_value)
    except Exception as e:
        logger.error(f"Failed to resolve <{key}>: {e}")
        return ""

# Rule evaluation function with strict matching
def evaluate_rules(rules_df, document_data, input_data, field_map):
    results = []

    for _, row in rules_df.iterrows():
        rule_id = row.get('Rule No', row.get('Output Identifier', 'Unknown'))
        expected = row.get('Output Language', '')
        match_type = row.get('Match Type', 'paragraph')
        formula = row.get('Formula', '')
        input_conditions = row.get('Input Value', '')

        # Evaluate input conditions
        if input_conditions:
            all_conditions_pass = True
            for cond in str(input_conditions).splitlines():
                if '=' in cond:
                    cond_key, cond_val = cond.split('=', 1)
                    cond_key = cond_key.strip()
                    cond_val = cond_val.strip().strip('"')
                    actual_val = resolve_placeholder(input_data, cond_key, field_map)
                    logger.debug(f"[COND] Evaluating: {cond_key} == {cond_val} -> Actual: {actual_val}")

                    expected_vals = [normalize(v.strip()) for v in cond_val.split(",")]
                    actual_vals = [normalize(v.strip()) for v in actual_val.split(",")]

                    if not all(val in actual_vals for val in expected_vals):
                        logger.info(f"[SKIP] Rule {rule_id} skipped due to condition mismatch: {cond_key}")
                        all_conditions_pass = False
                        break
            if not all_conditions_pass:
                results.append({"Rule No": rule_id, "Status": "SKIPPED", "Reason": "Condition mismatch"})
                continue

        # Resolve placeholders in expected output
        placeholders = re.findall(r"<(.*?)>", expected)
        for key in placeholders:
            val = resolve_placeholder(input_data, key, field_map)
            if isinstance(val, str):
                expected = expected.replace(f"<{key}>", val.strip())

        logger.debug(f"[CHECK] Final expected string after placeholder resolution: {expected}")
        found = False

        if match_type.lower() == "paragraph":
            for para in document_data["paragraphs"]:
                if expected.strip() in para["text"]:
                    logger.debug(f"[MATCH] Rule {rule_id} found in paragraph: {para['text']}")
                    found = True
                    break

        elif match_type.lower() == "table":
            for table in document_data["tables"]:
                for row in table:
                    for cell in row:
                        if expected.strip() in cell:
                            logger.debug(f"[MATCH] Rule {rule_id} found in table cell: {cell}")
                            found = True
                            break
                    if found:
                        break

        if found:
            results.append({"Rule No": rule_id, "Status": "PASS", "Reason": ""})
        else:
            reason = f"Expected text not found exactly. Expected: '{expected}'"
            results.append({"Rule No": rule_id, "Status": "FAIL", "Reason": reason})
            logger.debug(f"DEBUG FAIL - Rule {rule_id} =>")
            logger.debug(f"Expected (raw): {expected}")
            logger.debug("-------------")

    return results
