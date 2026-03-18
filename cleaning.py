import pandas as pd


file = "raw_data/Raw Data.xlsx"

df_orders = pd.read_excel(file, sheet_name="orders")
df_customers = pd.read_excel(file, sheet_name="customers")
df_products = pd.read_excel(file, sheet_name="products")


# ---------- STANDARD CHECKS ----------

def sheet_report(df, name):
    """Compact dataframe report"""
    
    print(f"\n--- {name.upper()} REPORT ---")

    # drop fully empty columns
    empty_cols = df.columns[(df.isna() | (df == '')).all()]
    if len(empty_cols) > 0:
        print("Dropped empty columns:", list(empty_cols))
        df = df.drop(columns=empty_cols)

    # missing values
    missing = (df.isna() | (df == '')).sum()
    missing = missing[missing > 0]

    if len(missing) > 0:
        print("Missing values:", missing.to_dict())
    else:
        print("Missing values: none")

    # datatype inconsistencies
    multi_types = [
        col for col in df.columns
        if df[col].apply(type).nunique() > 1
    ]

    if multi_types:
        print("Mixed datatypes:", multi_types)
    else:
        print("Mixed datatypes: none")

    return df


def check_unique(df, col):
    dup = df[col].duplicated().sum()
    print(f"{col} duplicates:", dup)


# ---------- RUN CHECKS ----------

df_orders = sheet_report(df_orders, "orders")


df_customers = sheet_report(df_customers, "customers")
check_unique(df_customers, "Customer ID")

df_products = sheet_report(df_products, "products")
check_unique(df_products, "Product ID")


# ---------- SPECIFIC VALIDATIONS ----------
print("")
print("--- CONSISTENCY REPORT ---")
print("--- ONLY CHECKS NON NULL NON EMPTY VALUES ---")
# OrderID → CustomerID consistency
check = df_orders.groupby("Order ID")["Customer ID"].nunique()
mismatch = check[check > 1]
print("Order-Customer mismatch in orders:", len(mismatch))


# email format
mail_pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'

df_mail = df_customers[df_customers["Email"].notna() & (df_customers["Email"] != "")]
invalid_mail = df_mail[~df_mail["Email"].str.match(mail_pattern)]

print("Invalid emails in customers:", len(invalid_mail))


# phone format
phone_pattern = r'^\+\d{1,3} \(\d{3}\) \d{3}-\d{4}$'

df_phone = df_customers[df_customers["Phone Number"].notna() & (df_customers["Phone Number"] != "")]
invalid_phone = df_phone[~df_phone["Phone Number"].str.match(phone_pattern)]

print("Invalid phone numbers in customers:", len(invalid_phone))


# loyalty card values
valid_loyal = {"yes", "no"}
values = set(df_customers["Loyalty Card"].str.lower().dropna())

if values.issubset(valid_loyal):
    print("Loyalty Card values in customers OK")
else:
    print("Unexpected Loyalty Card values in customers:", values)

# ---------- ORDERS TABLE ----------
df_orders['Order Date'] = pd.to_datetime(df_orders['Order Date'])
df_orders = df_orders.assign(
    Year=df_orders['Order Date'].dt.year,
    Month=df_orders['Order Date'].dt.month,
    Day_of_Month=df_orders['Order Date'].dt.day
)

# ---------- CUSTOMERS TABLE ----------
    # make 1 / 0 markers
    
df_customers = df_customers.assign(
    Customer_has_Email=df_customers['Email'].str.contains('@', na=False).astype(int),
    Customer_has_Phone=df_customers['Phone Number'].notna().astype(int),
    Customer_has_loyal_card=(df_customers['Loyalty Card'].str.lower() == 'yes').astype(int)
)

    # drop unnecessary columns
df_customers = df_customers.drop(columns=[
    'Customer Name','Email','Phone Number','Address Line 1','Postcode','Loyalty Card'
])

# ---------- PRODUCTS TABLE ----------
    # drop unnecessary columns
df_products = df_products.drop(columns=['Price per 100g'])

# ---------- MERGE TABLES ----------
df_total = (
    df_orders
    .merge(df_customers, on='Customer ID')
    .merge(df_products, on='Product ID')
)

# ---------- CALCULATE MEASURES ----------
df_total = df_total.assign(
    Total_Sale_Value=(df_total['Unit Price'] * df_total['Quantity']).round(2),
    Total_Profit_Value=(df_total['Profit'] * df_total['Quantity']).round(2),
    Total_Sale_Weight=(df_total['Size'] * df_total['Quantity']).round(2)

)

# drop columns no longer needed
df_total = df_total.drop(columns=['Unit Price', 'Profit', 'Size'])

#standard check on fianl_table
sheet_report(df_total, 'final table')

#safe final table as csv
file = "final_data/final_table.csv"
df_total.to_csv(file, index=False, encoding='utf-8')
print("")
print("final table saved to", file)
print("")

