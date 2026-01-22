import pandas as pd

def read_loads_excel(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)
    
    # Rename columns to standardized names if they have different labels
    column_mapping = {
        'x (m)': 'x',
        'value (N)': 'value',
        'x': 'x',
        'value': 'value'
    }
    
    # Rename columns that match the mapping
    df = df.rename(columns={col: column_mapping[col] for col in df.columns if col in column_mapping})
    
    required = {"x", "type", "value"}
    if not required.issubset(df.columns):
        raise ValueError(f"Excel must contain columns: {required}")
    df["type"] = df["type"].str.strip().str.lower()
    df = df.sort_values(by="x").reset_index(drop=True)
    return df

if __name__ == "__main__":
    df = read_loads_excel("loads.xlsx")
    print(df)
