from engine.document_parser import parse_document
from engine.rule_loader import load_rules
from engine.json_resolver import resolve_input
from engine.rule_evaluator import evaluate_rules
from engine.report_writer import write_report
import json

def main():
    doc_path = "New_York_Life_Insurance.docx"  # or "sample.pdf"
    json_path = "test.json"
    rules_path = "testdata_1.xlsx"
    mapping_path = "field_mappings.json"
    output_path = "validation_report.xlsx"

    # 🔍 Step 1: Parse the document
    document_data = parse_document(doc_path)

    # ✅ Step 2: Load full input data (no slicing at "txn")
    with open(json_path, 'r') as f:
        raw_data = json.load(f)
        input_data = raw_data # <-- FIXED: keep full structure

    # ✅ Step 3: Load rulebook and field mappings
    rules_df = load_rules(rules_path)
    with open(mapping_path) as f:
        field_map = json.load(f)

    # 🔍 Step 4: Evaluate rules and write report
    results = evaluate_rules(rules_df, document_data, input_data, field_map)
    write_report(results, output_path)

if __name__ == "__main__":
    main()