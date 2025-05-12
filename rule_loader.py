import pandas as pd

def load_rules(path):
    df = pd.read_excel(path, engine='openpyxl')
    df.columns = df.columns.str.strip()
    df.rename(columns={"Output Identifier": "Rule No"}, inplace=True)
    return df
