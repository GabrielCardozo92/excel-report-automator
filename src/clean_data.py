import pandas as pd
from pathlib import Path

DATA_DIR = Path(__file__).resolve().parent.parent / "data"

def load_data():
    """Load the raw sales data from the Excel file."""
    file_path = DATA_DIR / "raw_sales.xlsx"
    if not file_path.exists():
        raise FileNotFoundError(f"Data file not found: {file_path}")
    
    df = pd.read_excel(file_path)
    print(f"Loaded {len(df)} rows from {file_path}")
    return df
    

def clean_dates(df):
    """Parse dates from 3 known formats."""
    df = df.copy()
    raw = df["Date"]
    parsed = pd.to_datetime(raw, format="%d/%m/%Y", errors="coerce")
    parsed = parsed.fillna(pd.to_datetime(raw, format="%Y-%m-%d", errors="coerce"))
    parsed = parsed.fillna(pd.to_datetime(raw, format="%b %d, %Y", errors="coerce"))
    if parsed.isna().any():
        raise ValueError(f"{int(parsed.isna().sum())} dates did not match any known format")
    df["Date"] = parsed
    print(f"[dates] Parsed {len(df)} dates from 3 known formats")
    return df

def clean_prices(df):
    """Standardize the price format to float."""
    df = df.copy()
    df["Price"] = df["Price"].str.replace(r"[^\d.]", "", regex=True).astype(float)
    print(f"[prices] Standardized price format for {len(df)} rows")
    return df

def clean_products(df):
    """Standardize the product names."""
    df = df.copy()
    df["Product"] = df["Product"].str.strip().str.title()
    print(f"[products] Standardized product names for {len(df)} rows")
    return df

def remove_duplicates(df):
    """Remove duplicate rows from the DataFrame."""
    before = len(df)
    df = df.drop_duplicates()
    print(f"[duplicates] Removed {before - len(df)} duplicate rows")
    return df

def handle_missing(df):
    """Handle missing values by filling or dropping them."""
    df = df.copy()
    df["Seller"] = df["Seller"].fillna("Unknown")
    df["Region"] = df["Region"].fillna("Unknown")
    missing = int(df["Quantity"].isna().sum())
    recovered = (df["Total"] / df["Price"]).round()
    df["Quantity"] = df["Quantity"].fillna(recovered)
    print(f"[missing] Filled missing quantities for {missing} rows")
    return df

def fix_totals(df):
    """Recalculate the Total column based on Price and Quantity."""
    df = df.copy()
    mismatch = (df["Total"] - df["Price"] * df["Quantity"]).abs() > 0.01
    df["Total"] = (df["Price"] * df["Quantity"]).round(2)
    print(f"[totals] Found and fixed {int(mismatch.sum())} incorrect totals")
    return df

def main():
    df = load_data()
    df = clean_dates(df)
    df = clean_prices(df)
    df = clean_products(df)
    df = remove_duplicates(df)
    df = handle_missing(df)
    df = fix_totals(df)
    
    cleaned_file_path = DATA_DIR / "cleaned_sales.xlsx"
    out = df.copy()
    out["Date"] = out["Date"].dt.date
    out.to_excel(cleaned_file_path, index=False)
    print(f"Cleaned data saved to {cleaned_file_path}")

if __name__ == "__main__":
    main()