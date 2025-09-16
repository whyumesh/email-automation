import pandas as pd
import os
from jinja2 import Environment, FileSystemLoader
import win32com.client as win32
from datetime import datetime as dt

# Load Jinja2 Template
env = Environment(loader=FileSystemLoader('.'))
template = env.get_template("email_template_V4.html")

#pathc = "C:\\Users\\PAGARAX1\\OneDrive - Abbott\\Documents\\Project 3D email automation\\"
#df = pd.read_excel(os.path.join(pathc, 'Sample Request - Master File.xlsx'))

#df = pd.read_excel('Sap Master Dummy File.xlsx')

df = pd.read_excel('TBM & ABM Automation Email 1209252.xlsx')

emp_codes = [729919,29841]
#df = df[df['Input Sample Request: Created Alias'].isin(emp_codes)]

#Flag filter

"""
Docket Number	Transporter Name

#df['Date'] = df['Date'].dt.date
df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
df['Date'] = df['Date'].dt.strftime('%d/%m/%Y')
z = dt.today()
current_date = z.date()

df['Delivery Date'] = pd.to_datetime(df['Delivery Date'], errors='coerce')
df['Delivery Date'] = df['Delivery Date'].dt.strftime('%d/%m/%Y')
df['Delivery Date'].fillna('-', inplace=True)
"""

df.rename(columns={'Request Date':'Date'}, inplace=True)
df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
df['Date'] = df['Date'].dt.strftime('%d-%b-%Y')
z = dt.today()
current_date = z.date()

df['Delivery Date'] = pd.to_datetime(df['Delivery Date'], errors='coerce')
df['Delivery Date'] = df['Delivery Date'].dt.strftime('%d-%b-%Y')
df['Delivery Date'].fillna('-', inplace=True)

df['Dispatch Date'] = pd.to_datetime(df['Dispatch Date'], errors='coerce')
df['Dispatch Date'] = df['Dispatch Date'].dt.strftime('%d-%b-%Y')
df['Dispatch Date'].fillna('-', inplace=True)


df['Rto Reason'].fillna('-', inplace=True)
df['Invoice #'].fillna('-', inplace=True)
df['Docket Number'].fillna('-', inplace=True)
df['Transporter Name'].fillna('-', inplace=True)

# Initialize Outlook
outlook = win32.Dispatch("Outlook.Application")

# Group by Created Alias
grouped = df.groupby("Input Sample Request: Created Alias")
 
for alias, group in grouped:
    email_id = group["TBM EMAIL_ID"].iloc[0]
    requested_date = group["Date"].iloc[0]
 
    # Prepare rows data for the template
    rows = []
    for _, row in group.iterrows():
        rows.append({
            "Affiliate": row["AFFILIATE"],
            "Requested_Date": row["Date"],
            "Request_ID": row["Assigned Request Ids"],
            "Doctor_Code": row["Doctor: Customer Code"],
            "SAP_Customer_Code": row["Doctor: SAP Customer Code(New)"],
            "Doctor_Name": row["Doctor: Account Name"],
            "Item_Code": row["Item Code"],
            "SKU_Name": row["SKU"],
            "Requested_Quantity": row["Requested Quantity"],
            "Dispatch_Date":row["Dispatch Date"],
            "Delivery_Date":row["Delivery Date"],
            "Request_Status": row["Request Status"],
            "Rto_Reason":row["Rto Reason"],
            "Invoice_number":row["Invoice #"],
            "Docket_number":row["Docket Number"],
            "Transporter_name":row["Transporter Name"]
            
        })
 
    # Render HTML email content
    email_html = template.render(rows=rows)
 
    # Create and send email
    mail = outlook.CreateItem(0)
    mail.To = email_id
    #mail.BCC = 'vaibhav.nalawade@abbott.com'
    mail.Subject = f"Sample Direct Dispatch to Doctors - Request Status as of {current_date}"
    mail.HTMLBody = email_html
    mail.SentOnBehalfOfName = 'EPD_SFA@abbott.com'
    mail.Send()  # use mail.Send() to send automatically
       
print("All emails sent successfully!")    
    
    
    










