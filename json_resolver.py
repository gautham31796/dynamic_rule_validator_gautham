import json

def resolve_input(json_path):
    with open(json_path, 'r') as f:
        return json.load(f)
