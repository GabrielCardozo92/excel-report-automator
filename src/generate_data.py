import random
import pandas as pd
from faker import Faker
from pathlib import Path

fake = Faker()
fake.seed_instance(42)
random.seed(42)

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
PRODUCTS = ["Laptop", "Smartphone", "Tablet", "Headphones", "Smartwatch", "Camera", "Printer", "Monitor", "Keyboard"]
REGIONS = ["North America", "Europe", "Asia", "South America", "Africa", "Australia"]
SELLERS = ["Parker", "Anna", "Kevin", "Sarah", "Ronald", "Emily", "David", "Sophia", "Michael", "Olivia"]
NUM_ROWS = 1000


def generate_row():
    product = random.choice(PRODUCTS)
    region = random.choice(REGIONS)
    seller = random.choice(SELLERS)
    price = round(random.uniform(50, 2000), 2)
    quantity = random.randint(1, 20)
    date = fake.date_between(start_date='-1y', end_date='today')
    
    return {
        "Date": date,
        "Region": region,
        "Seller": seller,
        "Product": product,
        "Price": price,
        "Quantity": quantity,
        "Total": round(price * quantity, 2),
    }

def make_messy(df):
    """Deliberately corrupt clean data to simulate real-world files."""
    df = df.copy()
    df["Date"] = df["Date"].apply(lambda x: x.strftime(random.choice(["%Y-%m-%d", "%d/%m/%Y", "%b %d, %Y"])))
    df["Price"] = df["Price"].apply(lambda x: random.choice([f"${x:,.2f}", f"{x:.2f}", f"USD {x:.0f}"]))
    df["Product"] = df["Product"].apply(lambda x: random.choice([x.lower(), x.upper(), x.title(), f" {x} ", f"{x}  "]))
    bad_rows = df.sample(n=25).index
    df.loc[bad_rows, "Total"] = df.loc[bad_rows, "Total"].apply(lambda x: round(x * random.uniform(0.5, 1.5), 2))
    duplicates = df.sample(n=30)
    df = pd.concat([df, duplicates], ignore_index=True)
    for _ in range(40):
        row = random.randint(0, len(df) - 1)
        col = random.choice(["Seller", "Region", "Quantity"])
        df.at[row, col] = None
    df = df.sample(frac=1).reset_index(drop=True)
    return df

def main():
    df = pd.DataFrame([generate_row() for _ in range(NUM_ROWS)])
    df = make_messy(df)
    df.to_excel(DATA_DIR / "raw_sales.xlsx", index=False)
    print(f"Generated {len(df)} rows -> {DATA_DIR / 'raw_sales.xlsx'}")

if __name__ == "__main__":
    main()