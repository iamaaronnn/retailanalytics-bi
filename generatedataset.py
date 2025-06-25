# Aaron Mascarenhas
import pandas as pd
import random
from datetime import datetime, timedelta

# Constants
companies = ['Amazon India', 'Flipkart', 'Myntra', 'Snapdeal']
categories = {
    'Electronics': ['Watches', 'Cameras', 'Mobiles', 'Laptops'],
    'Home Appliances': ['Furniture', 'Kitchen Appliances', 'Lighting'],
    'Fashion': ["Men's Clothing", "Women's Clothing", 'Footwear', 'Accessories'],
    'Books': ['Fiction', 'Non-Fiction', 'Academic', 'Comics'],
    'Groceries': ['Vegetables', 'Fruits', 'Dairy', 'Snacks']
}

cities = ['Mumbai', 'Delhi', 'Bhopal', 'Hyderabad', 'Chennai', 'Ahmedabad']
states = ['Maharashtra', 'Delhi', 'Rajasthan', 'Tamil Nadu', 'Gujarat', 'West Bengal', 'Uttar Pradesh']
regions = ['North', 'South', 'East', 'West', 'Central']

def generate_order_id():
    return f"ORD{random.randint(100000, 999999)}"

def random_order_date():
    start = datetime(2020, 1, 1)
    end = datetime(2024, 12, 31)
    return start + timedelta(days=random.randint(0, (end - start).days))

def random_shipping_date(order_date):
    return order_date - timedelta(days=random.randint(2, 20))

def generate_row():
    company = random.choice(companies)
    category = random.choice(list(categories.keys()))
    sub_category = random.choice(categories[category])
    order_date = random_order_date()
    shipping_date = random_shipping_date(order_date)
    return [
        generate_order_id(),
        company,
        category,
        sub_category,
        order_date.strftime('%Y-%m-%d'),
        shipping_date.strftime('%Y-%m-%d'),
        f"Customer {random.randint(1, 5000)}",
        city := random.choice(cities),
        state := random.choice(states),
        region := random.choice(regions),
        random.randint(5000, 100000),
        random.randint(1, 10),
        random.randint(0, 30),
        random.randint(1000, 25000)
    ]

def generate_data(n=500):
    return [generate_row() for _ in range(n)]

columns = ["Order ID", "Company", "Category", "Sub Category", "Order Date", "Shipping Date",
           "Customer Name", "City", "State", "Region", "Sales (₹)", "Quantity", "Discount (%)", "Profit (₹)"]

df = pd.DataFrame(generate_data(500), columns=columns)

# Save as Excel (.xlsx)
df.to_excel("random_orders.xlsx", index=False, engine='openpyxl')
print("✅ 'random_orders.xlsx' has been created successfully.")
