import pandas as pd

# Load Excel
file_path = "logic.xlsx"
xls = pd.ExcelFile(file_path)

sheet1 = pd.read_excel(xls, "Sheet1")
sheet2 = pd.read_excel(xls, "Sheet2")

# Normalize text function
def normalize(text):
    return str(text).strip().casefold()

# Build rules dictionary from Sheet2
rules = {}
for _, row in sheet2.iterrows():
    statuses = [normalize(s) for s in row.drop("Final Answer").dropna().tolist()]
    statuses = tuple(sorted(set(statuses)))
    rules[statuses] = row["Final Answer"]

# Group statuses by request ID
grouped = sheet1.groupby("Assigned Request Ids")["Request Status"].apply(list).reset_index()

# Deduplicate statuses for both display and matching
def dedup_statuses(status_list):
    return sorted(set(status_list), key=str)

def get_final_answer(status_list):
    key = tuple(sorted(set(normalize(s) for s in status_list)))
    return rules.get(key, "❌ No matching rule")

# Apply deduplication
grouped["Request Status"] = grouped["Request Status"].apply(dedup_statuses)
grouped["Final Answer"] = grouped["Request Status"].apply(get_final_answer)

# Save output
output_file = "final_output.xlsx"
grouped.to_excel(output_file, index=False)

print(f"✅ Final results saved to {output_file}")
