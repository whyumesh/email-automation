import pandas as pd
import win32com.client as win32

# ---------- Step 1: Load recipients ----------
emails = pd.read_excel("email.xlsx")

to_list = emails["To"].dropna().tolist()
cc_list = emails["cc"].dropna().tolist()

# ---------- Step 2: Load the Excel table ----------
table_file = "zbm_summary_updated_20250930_111329.xlsx"

# Assuming table is in the first sheet; adjust sheet_name if needed
df = pd.read_excel(table_file, sheet_name=0)

# Drop completely empty rows and columns (just in case)
df = df.dropna(how="all").dropna(axis=1, how="all")

# ---------- Step 3: Convert to HTML ----------
html_table = df.to_html(index=False, border=1, justify="center")

# ---------- Step 4: Send email via Outlook ----------
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)

mail.Subject = "Automated ZBM Summary Report"
mail.HTMLBody = f"""
<p>Hello,</p>
<p>Please find below the updated ZBM Summary Report:</p>
{html_table}
<p>Regards,<br>Automation</p>
"""

mail.To = "; ".join(to_list)
mail.CC = "; ".join(cc_list)

mail.Send()
print("✅ Email sent successfully!")
