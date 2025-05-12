import pandas as pd

def write_report(results, output_path):
    df = pd.DataFrame(results)
    df.to_excel(output_path, index=False)
