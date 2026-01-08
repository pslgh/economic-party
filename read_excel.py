import pandas as pd
import json
import sys
import os

def excel_to_json(file_path, output_path=None):
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return

    try:
        # Read all sheets
        xls = pd.ExcelFile(file_path)
        data = {}
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            # Replace NaN with null and non-finite values for valid JSON
            df = df.where(pd.notnull(df), None)
            
            # Convert timestamp/datetime objects to strings
            for col in df.columns:
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    df[col] = df[col].astype(str)
                    
            data[sheet_name] = df.to_dict(orient='records')
        
        json_output = json.dumps(data, indent=2, ensure_ascii=False)
        
        if output_path:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(json_output)
            print(f"Successfully wrote JSON to {output_path}")
        else:
            print(json_output)
        
    except Exception as e:
        print(f"Error reading Excel: {str(e)}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python read_excel.py <path> [output_path]")
    else:
        output = sys.argv[2] if len(sys.argv) > 2 else None
        excel_to_json(sys.argv[1], output)
