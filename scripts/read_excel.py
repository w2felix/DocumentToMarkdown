import pandas as pd
import sys
from pathlib import Path

# Get project root (parent of scripts folder)
project_root = Path(__file__).parent.parent

try:
    df = pd.read_excel(project_root / 'test_poster' / 'AACR26_Selected_Apr13.xlsx')
    print("=== Excel File Structure ===")
    print(f"Total rows: {len(df)}")
    print(f"\nColumns: {df.columns.tolist()}")
    print("\n=== First 5 rows ===")
    print(df.head(5).to_string())
    print("\n=== Sample row for poster 160 (if exists) ===")
    poster_160 = df[df.apply(lambda row: '160' in str(row.values), axis=1)]
    if not poster_160.empty:
        print(poster_160.to_string())
except Exception as e:
    print(f"Error: {e}")
    sys.exit(1)
