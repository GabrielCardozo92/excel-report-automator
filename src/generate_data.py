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

def main():
    df = pd.DataFrame([generate_row() for _ in range(NUM_ROWS)])
    df.to_excel(DATA_DIR / "raw_sales.xlsx", index=False)
    print(f"Generated {len(df)} rows -> {DATA_DIR / 'raw_sales.xlsx'}")

if __name__ == "__main__":
    main()