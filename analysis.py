import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from math import ceil
import os

# -------------------------------
# 1. Load Input Files
# -------------------------------
DATA_DIR = "Data"  # folder containing all Excel files

order_df = pd.read_excel(os.path.join(DATA_DIR, "Company X - Order Report.xlsx"), sheet_name="Sheet1")
pincode_df = pd.read_excel(os.path.join(DATA_DIR, "Company X - Pincode Zones.xlsx"), sheet_name="Sheet1")
sku_df = pd.read_excel(os.path.join(DATA_DIR, "Company X - SKU Master.xlsx"), sheet_name="Sheet1")
invoice_df = pd.read_excel(os.path.join(DATA_DIR, "Courier Company - Invoice.xlsx"), sheet_name="Sheet1")
rates_df = pd.read_excel(os.path.join(DATA_DIR, "Courier Company - Rates.xlsx"), sheet_name="Sheet2")

def clean_cols(df):
    df.columns = [str(c).strip() for c in df.columns]
    return df

order_df = clean_cols(order_df)
pincode_df = clean_cols(pincode_df)
sku_df = clean_cols(sku_df)
invoice_df = clean_cols(invoice_df)
rates_df = clean_cols(rates_df)

# -------------------------------
# 2. Compute Total Weight per Order (as per X)
# -------------------------------
sku_df['Weight (kg)'] = sku_df['Weight (g)'] / 1000.0
order_with_weights = order_df.merge(sku_df[['SKU', 'Weight (kg)']], on='SKU', how='left')
order_with_weights['Line Weight (kg)'] = order_with_weights['Order Qty'] * order_with_weights['Weight (kg)']

total_weight_x = (
    order_with_weights.groupby('ExternOrderNo', as_index=False)['Line Weight (kg)']
    .sum()
    .rename(columns={'ExternOrderNo': 'Order ID', 'Line Weight (kg)': 'Total weight as per X (KG)'})
)

# -------------------------------
# 3. Delivery Zone (as per X)
# -------------------------------
zone_as_per_x = invoice_df.merge(
    pincode_df.rename(columns={'Zone': 'Delivery Zone as per X'}),
    on=['Warehouse Pincode', 'Customer Pincode'],
    how='left'
)[['Order ID', 'Delivery Zone as per X']]

zone_as_per_x['Delivery Zone as per X'] = zone_as_per_x['Delivery Zone as per X'].astype(str).str.upper().str.strip()

# -------------------------------
# 4. Courier Invoice Cleanup
# -------------------------------
inv = invoice_df.rename(columns={
    'AWB Code': 'AWB Number',
    'Charged Weight': 'Total weight as per Courier Company (KG)',
    'Zone': 'Delivery Zone charged by Courier Company'
})
inv['Delivery Zone charged by Courier Company'] = inv['Delivery Zone charged by Courier Company'].astype(str).str.upper().str.strip()

# -------------------------------
# 5. Merge Base Data
# -------------------------------
base = inv.merge(total_weight_x, on='Order ID', how='left').merge(zone_as_per_x, on='Order ID', how='left')

rates = rates_df.rename(columns={'Weight Slabs': 'Slab Length (KG)'})
rates['Zone'] = rates['Zone'].astype(str).str.upper().str.strip()

base = base.merge(
    rates[['Zone', 'Slab Length (KG)']].rename(columns={'Zone': 'Delivery Zone as per X', 'Slab Length (KG)': 'Slab Length as per X (KG)'}),
    on='Delivery Zone as per X', how='left'
).merge(
    rates[['Zone', 'Slab Length (KG)']].rename(columns={'Zone': 'Delivery Zone charged by Courier Company', 'Slab Length (KG)': 'Slab Length charged by Courier (KG)'}),
    on='Delivery Zone charged by Courier Company', how='left'
)

# -------------------------------
# 6. Compute Weight Slabs
# -------------------------------
def round_up_to_slab(weight, slab):
    if pd.isna(weight) or pd.isna(slab) or slab <= 0:
        return np.nan
    return ceil(weight / slab) * slab

base['Weight slab as per X (KG)'] = base.apply(
    lambda r: round_up_to_slab(r['Total weight as per X (KG)'], r['Slab Length as per X (KG)']), axis=1
)
base['Weight slab charged by Courier Company (KG)'] = base.apply(
    lambda r: round_up_to_slab(r['Total weight as per Courier Company (KG)'], r['Slab Length charged by Courier (KG)']), axis=1
)

# -------------------------------
# 7. Expected Charges (Fixed Logic from Notebook)
# -------------------------------
charges = rates[['Zone', 'Slab Length (KG)', 'Forward Fixed Charge', 'Forward Additional Weight Slab Charge',
                 'RTO Fixed Charge', 'RTO Additional Weight Slab Charge']].rename(columns={'Zone': 'Delivery Zone as per X'})
calc = base.merge(charges, on='Delivery Zone as per X', how='left')

# Number of slabs
calc['n_slabs_expected'] = (calc['Weight slab as per X (KG)'] / calc['Slab Length (KG)']).fillna(0).astype(float)

# Expected Forward Charges
calc['Expected FWD'] = calc['Forward Fixed Charge'] + (calc['n_slabs_expected'] - 1).clip(lower=0) * calc['Forward Additional Weight Slab Charge']

# Expected RTO Charges
calc['has_rto'] = calc['Type of Shipment'].str.contains('RTO', case=False, na=False)
calc['Expected RTO'] = np.where(
    calc['has_rto'],
    calc['RTO Fixed Charge'] + (calc['n_slabs_expected'] - 1).clip(lower=0) * calc['RTO Additional Weight Slab Charge'],
    0.0
)

calc['Expected Charge as per X (Rs.)'] = (calc['Expected FWD'] + calc['Expected RTO']).round(2)

# -------------------------------
# 8. Final Dataset
# -------------------------------
result = calc[[
    'Order ID',
    'AWB Number',
    'Total weight as per X (KG)',
    'Weight slab as per X (KG)',
    'Total weight as per Courier Company (KG)',
    'Weight slab charged by Courier Company (KG)',
    'Delivery Zone as per X',
    'Delivery Zone charged by Courier Company',
    'Expected Charge as per X (Rs.)',
    'Billing Amount (Rs.)'
]].rename(columns={'Billing Amount (Rs.)': 'Charges Billed by Courier Company (Rs.)'})

result['Difference Between Expected Charges and Billed Charges (Rs.)'] = (
    result['Expected Charge as per X (Rs.)'] - result['Charges Billed by Courier Company (Rs.)']
).round(2)

# -------------------------------
# 9. Summary Table
# -------------------------------
def summarize(df):
    correct = df[np.isclose(df['Difference Between Expected Charges and Billed Charges (Rs.)'], 0.0)]
    over = df[df['Difference Between Expected Charges and Billed Charges (Rs.)'] < 0]
    under = df[df['Difference Between Expected Charges and Billed Charges (Rs.)'] > 0]

    summary_rows = [
        {"Metric": "Total Orders - Correctly Charged", "Count": len(correct), "Amount (Rs.)": round(correct['Charges Billed by Courier Company (Rs.)'].sum(), 2)},
        {"Metric": "Total Orders - Over Charged", "Count": len(over), "Amount (Rs.)": abs(round(over['Difference Between Expected Charges and Billed Charges (Rs.)'].sum(), 2))},
        {"Metric": "Total Orders - Under Charged", "Count": len(under), "Amount (Rs.)": abs(round(under['Difference Between Expected Charges and Billed Charges (Rs.)'].sum(), 2))},
    ]
    return pd.DataFrame(summary_rows)

summary_df = summarize(result)

# -------------------------------
# 10. Save to Excel
# -------------------------------
with pd.ExcelWriter("Reconciliation_Result.xlsx", engine="xlsxwriter") as writer:
    summary_df.to_excel(writer, sheet_name="Summary", index=False)
    result.to_excel(writer, sheet_name="Calculations", index=False)

print("✅ Reconciliation_Result.xlsx generated successfully!")

# -------------------------------
# 11. Analysis & Visualizations
# -------------------------------
# Bar Chart (Counts)
plt.figure(figsize=(6,4))
plt.bar(summary_df['Metric'], summary_df['Count'], color=['green','red','orange'])
plt.title("Order Count by Category")
plt.ylabel("Number of Orders")
plt.xticks(rotation=20)
plt.savefig("summary_bar.png")
plt.close()

# Pie Chart (Amounts - use absolute values)
plt.figure(figsize=(6,6))
amounts = summary_df['Amount (Rs.)'].abs()
plt.pie(amounts, labels=summary_df['Metric'], autopct='%1.1f%%', startangle=140)
plt.title("Distribution of Amounts by Category")
plt.savefig("summary_pie.png")
plt.close()

# Histogram of Differences
plt.figure(figsize=(6,4))
plt.hist(result['Difference Between Expected Charges and Billed Charges (Rs.)'], bins=20, color="skyblue", edgecolor="black")
plt.title("Distribution of Charge Differences")
plt.xlabel("Difference (Rs.)")
plt.ylabel("Number of Orders")
plt.savefig("diff_hist.png")
plt.close()

print("📊 Charts saved: summary_bar.png, summary_pie.png, diff_hist.png")

# -------------------------------
# 12. Business Summary
# -------------------------------
print("\n📌 Business Insights:")
print("- Correctly charged orders:", int(summary_df.iloc[0]['Count']))
print("- Overcharged orders:", int(summary_df.iloc[1]['Count']), "with total overcharge:", summary_df.iloc[1]['Amount (Rs.)'])
print("- Undercharged orders:", int(summary_df.iloc[2]['Count']), "with total undercharge:", summary_df.iloc[2]['Amount (Rs.)'])
print("\n✅ Recommendation: Focus on overcharged orders to recover money from courier partners. Undercharged orders may benefit the company but should be investigated for consistency.")
