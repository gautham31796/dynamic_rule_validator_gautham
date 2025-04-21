import pandas as pd
import json
import fitz  # PyMuPDF
import re

def extract_text_from_pdf(pdf_path):
    with fitz.open(pdf_path) as doc:
        text = ""
        for page in doc:
            text += page.get_text()
    return text

def normalize_text(text):
    text = text.lower()
    text = re.sub(r'[^a-z0-9\s]', '', text)
    return ' '.join(text.split())

def evaluate_rule(rule_row, pdf_text, input_data):
    input_val = rule_row.get('Input Value', '')
    expected = rule_row['Output Language']

    # ✅ Handle conditional input logic like: Key=Value
    if '=' in input_val:
        condition_key, condition_val = input_val.split('=', 1)
        condition_key = condition_key.strip()
        condition_val = condition_val.strip()

        # Case-insensitive key lookup
        input_data_lower = {k.lower(): v for k, v in input_data.items()}
        input_actual = input_data_lower.get(condition_key.lower(), "")

        # Handle nested dicts (fallback)
        if isinstance(input_actual, dict):
            input_actual = list(input_actual.keys())[0] if input_actual else ""

        actual_norm = normalize_text(str(input_actual))
        expected_norm = normalize_text(str(condition_val))

        print(f"Checking conditional match for Rule {rule_row['Rule No']}:")
        print(f"Raw Input Value: {input_val}")
        print(f"From JSON: {condition_key} = {input_actual}")
        print(f"Normalized: actual='{actual_norm}' vs expected='{expected_norm}'")

        if actual_norm != expected_norm:
            print(f"Skipping Rule {rule_row['Rule No']} — Condition Mismatch.")
            return 'SKIPPED', expected

    # 🔄 Replace placeholders like <Key>
    placeholders = re.findall(r"<(.*?)>", expected)
    print(f"Placeholders found in rule: {placeholders}")

    for key in placeholders:
        val = input_data.get(key, "")
        if isinstance(val, dict):
            val = list(val.values())[0] if val else ""
        expected = expected.replace(f"<{key}>", str(val))

    print("Expected before normalize:", expected)

    expected_clean = normalize_text(expected)
    pdf_text_clean = normalize_text(pdf_text)

    result = 'PASS' if expected_clean in pdf_text_clean else 'FAIL'

    if result == 'FAIL':
        print("\nDEBUG FAIL >>>")
        print("Expected Clean:", expected_clean)
        print("PDF Text Contains Expected:", expected_clean in pdf_text_clean)
        print("First 500 chars of PDF Text:\n", pdf_text_clean[:500])
        print("-------------")

    return result, expected

def load_rules(excel_path):
    df = pd.read_excel(excel_path, engine='openpyxl')
    df.columns = df.columns.str.strip()
    print("Loaded columns:", df.columns.tolist())
    return df

def main():
    excel_path = "Rules.xlsx"
    pdf_path = "New_York_Life_Insurance.pdf"
    json_path = "testdata.json"

    rules_df = load_rules(excel_path)
    pdf_text = extract_text_from_pdf(pdf_path)

    with open(json_path, 'r') as f:
        raw_data = json.load(f)
        input_data = raw_data.get("testData", {})  
    print("Loaded input data keys:", list(input_data.keys()))

    results = []
    for _, row in rules_df.iterrows():
        rule_id = row.get('Rule No', 'N/A')
        result, expected = evaluate_rule(row, pdf_text, input_data)
        results.append(result)
        print(f"Rule {rule_id}: {result}")

    rules_df['Result'] = results
    rules_df.to_excel("rule_results.xlsx", index=False)

if __name__ == "__main__":
    main()