import pandas as pd
import json
import fitz  # PyMuPDF
import docx
import re
import os

# Import both rule evaluators
import text_validation as text_validator
import table_validation as table_validator

def extract_text(document_path):
    if document_path.lower().endswith('.pdf'):
        return text_validator.extract_text_from_pdf(document_path)
    elif document_path.lower().endswith('.docx'):
        return text_validator.extract_text_from_word(document_path)
    else:
        raise ValueError("Unsupported document type. Use PDF or Word (.docx)")

def load_input_data(json_path):
    with open(json_path, 'r') as f:
        raw_data = json.load(f)
        return raw_data.get("testData", raw_data)

def main():
    excel_path = "testdata_1.xlsx"
    document_path = "New_York_Life_Insurance.docx"
    json_path = "test.json"
    
    # Load everything
    rules_df = pd.read_excel(excel_path, engine='openpyxl')
    rules_df.columns = rules_df.columns.str.strip()
    document_text = extract_text(document_path)
    input_data = load_input_data(json_path)

    output_data = []

    for _, row in rules_df.iterrows():
        rule_type = row.get('Type', '').strip().lower()
        rule_id = row.get('Output Identifier', 'N/A')

        if rule_type == 'text':
            result, reason = text_validator.evaluate_rule(row, document_text, input_data, document_path)
        elif rule_type == 'table':
            result, reason = table_validator.evaluate_rule(row, document_text, input_data, document_path)
        else:
            result, reason = 'SKIPPED', f"Unknown rule type '{rule_type}'"

        output_data.append({
            "Output Identifier": rule_id,
            "Status": result,
            "Reason": reason
        })
        print(f"Rule {rule_id}: {result} — {reason}")

    result_df = pd.DataFrame(output_data)
    result_df.to_excel("rule_results.xlsx", index=False)

if __name__ == "__main__":
    main()