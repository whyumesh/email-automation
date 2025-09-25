import pandas as pd
import numpy as np
from datetime import datetime

def create_zbm_summary():
    """
    Automate the creation of ZBM summary sheet from master_tracker.csv
    """
    
    print("🔄 Starting ZBM Summary Automation...")
    
    # Read master tracker data
    print("📖 Reading master_tracker.csv...")
    try:
        # Try different encodings
        encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
        df = None
        
        for encoding in encodings:
            try:
                df = pd.read_csv('master_tracker.csv', encoding=encoding)
                print(f"✅ Successfully loaded {len(df)} records from master_tracker.csv using {encoding} encoding")
                break
            except UnicodeDecodeError:
                continue
        
        if df is None:
            print("❌ Could not read master_tracker.csv with any of the tried encodings")
            return
            
    except Exception as e:
        print(f"❌ Error reading master_tracker.csv: {e}")
        return
    
    # Clean and prepare data
    print("🧹 Cleaning and preparing data...")
    
    # Ensure required columns exist
    required_columns = ['ABM Terr Code', 'TBM HQ', 'Input Sample Request: Created By',
                        'Doctor: Customer Code', 'Assigned Request Ids', 'Request Status']
    missing = [c for c in required_columns if c not in df.columns]
    if missing:
        print(f"❌ Missing required columns in master_tracker.csv: {missing}")
        return

    # Remove rows where key fields are null or empty
    df = df.dropna(subset=['ABM Terr Code', 'TBM HQ', 'Input Sample Request: Created By'])
    df = df[df['ABM Terr Code'].astype(str).str.strip() != '']
    df = df[df['TBM HQ'].astype(str).str.strip() != '']
    df = df[df['Input Sample Request: Created By'].astype(str).str.strip() != '']

    # Restrict to specified TBM HQ cities
    allowed_hq = {"MUMBAI", "AHMEDABAD", "PUNE", "NAGPUR"}
    df['TBM HQ'] = df['TBM HQ'].astype(str)
    df = df[df['TBM HQ'].str.upper().isin(allowed_hq)]

    # Build Area Name: "TBM HQ - ABM Terr Code"
    df['Area Name'] = (
        df['TBM HQ'].astype(str).str.strip() + ' - ' + df['ABM Terr Code'].astype(str).str.strip()
    )

    # Compute Final Answer per unique request id using rules from logic.xlsx (same as test.py)
    print("🧠 Computing final status per unique Request Id using rules (test.py logic)...")
    try:
        xls_rules = pd.ExcelFile('logic.xlsx')
        sheet2 = pd.read_excel(xls_rules, 'Sheet2')

        def normalize(text):
            return str(text).strip().casefold()

        rules = {}
        for _, row in sheet2.iterrows():
            statuses = [normalize(s) for s in row.drop('Final Answer').dropna().tolist()]
            statuses = tuple(sorted(set(statuses)))
            rules[statuses] = row['Final Answer']

        # Group statuses by request id from master data
        grouped = df.groupby('Assigned Request Ids')['Request Status'].apply(list).reset_index()

        def get_final_answer(status_list):
            key = tuple(sorted(set(normalize(s) for s in status_list)))
            return rules.get(key, '❌ No matching rule')

        # Dedup display statuses as in test.py (not strictly needed, but helpful)
        grouped['Request Status'] = grouped['Request Status'].apply(lambda lst: sorted(set(lst), key=str))
        grouped['Final Answer'] = grouped['Request Status'].apply(get_final_answer)

        # Flag for Pending for Invoicing (D) based on presence of specific status in the list
        def has_action_pending(status_list):
            target = 'action pending / in process'
            return any(normalize(s) == target for s in status_list)
        grouped['Has D Pending'] = grouped['Request Status'].apply(has_action_pending)

        # Save final answers to Excel
        grouped.to_excel('final_output.xlsx', index=False)
        print("✅ Final Answer per request saved to final_output.xlsx")

        # Merge Final Answer back to main dataframe
        df = df.merge(grouped[['Assigned Request Ids', 'Final Answer']], on='Assigned Request Ids', how='left')
    except Exception as e:
        print(f"❌ Error computing final status from logic.xlsx: {e}")
        return
    
    print(f"📊 After cleaning: {len(df)} records remaining")
    
    # Group by ABM Terr Code and ABM Name
    print("📈 Aggregating data by ABM Terr Code...")
    
    # Create aggregation functions
    def count_unique_doctors(group):
        return group['Doctor: Customer Code'].nunique()
    
    def count_unique_requests(group):
        return group['Assigned Request Ids'].nunique()
    
    def count_status_category(group, status_list):
        """Count requests with specific statuses"""
        return group[group['Request Status'].isin(status_list)]['Assigned Request Ids'].nunique()
    
    # Define status categories based on the ZBM sheet structure
    status_categories = {
        'out_of_stock_on_hold': ['Out of stock', 'On hold', 'Not permitted'],
        'request_raised': ['Request Raised'],
        'delivered_return_action_pending': ['Delivered', 'Return', 'Action pending / In Process', 'Dispatched & In Transit', 'Dispatch Pending'],
        'action_pending': ['Action pending / In Process'],
        'dispatch_pending': ['Dispatch Pending'],
        'delivered': ['Delivered'],
        'dispatched_in_transit': ['Dispatched & In Transit'],
        'rto': ['RTO']
    }
    
    # Aggregate data
    # Determine which column to use for unique TBMs
    tbm_code_col = 'TBM Terr Code' if 'TBM Terr Code' in df.columns else 'ABM Terr Code'

    aggregated = df.groupby(['Area Name', 'Input Sample Request: Created By']).agg({
        tbm_code_col: 'nunique',             # Unique TBMs
        'Doctor: Customer Code': 'nunique',  # Unique HCPs
        'Assigned Request Ids': 'nunique',   # Unique Requests
        'Final Answer': lambda x: x.tolist()  # All statuses for further processing
    }).reset_index()
    
    # Rename columns
    # rename using index to handle dynamic tbm_code_col
    aggregated = aggregated.rename(columns={
        'Input Sample Request: Created By': 'ABM Name',
        tbm_code_col: 'Unique TBMs',
        'Doctor: Customer Code': 'Unique HCPs',
        'Assigned Request Ids': 'Unique Requests',
        'Final Answer': 'All Statuses'
    })
    
    # Calculate status-specific counts
    print("🔢 Calculating status-specific metrics...")
    
    for category_name, status_list in status_categories.items():
        print(f"   Processing {category_name}...")
        
        # Create a function to count statuses for each group
        def count_statuses_for_group(group):
            statuses = group['Final Answer'].tolist()
            return sum(1 for status in statuses if status in status_list)
        
        # Apply the counting function
        status_counts = df.groupby(['Area Name', 'Input Sample Request: Created By'], group_keys=False).apply(
            lambda x: count_statuses_for_group(x)
        ).reset_index(name=f'count_{category_name}')
        
        # Merge with aggregated data
        status_counts = status_counts.rename(columns={'Input Sample Request: Created By': 'ABM Name'})
        aggregated = aggregated.merge(status_counts, on=['Area Name', 'ABM Name'], how='left')
    
    # Calculate derived metrics
    print("📊 Calculating derived metrics...")
    
    # Requests Raised (A+B+C) - total unique requests
    aggregated['Requests Raised'] = aggregated['Unique Requests']
    
    # HO metrics
    aggregated['Request Cancelled Out of Stock'] = aggregated['count_out_of_stock_on_hold']
    aggregated['Action Pending at HO'] = aggregated['count_action_pending']
    
    # HUB metrics
    aggregated['Sent to HUB'] = aggregated['count_delivered_return_action_pending']
    # Pending for Invoicing (D): unique count of requests where Request Status contains Action Pending / in Process
    d_counts = df.merge(grouped[['Assigned Request Ids', 'Has D Pending']], on='Assigned Request Ids', how='left') \
                .groupby(['Area Name', 'Input Sample Request: Created By']) \
                .apply(lambda x: x.loc[x['Has D Pending'] == True, 'Assigned Request Ids'].nunique()) \
                .reset_index(name='Pending for Invoicing')
    d_counts = d_counts.rename(columns={'Input Sample Request: Created By': 'ABM Name'})
    aggregated = aggregated.merge(d_counts, on=['Area Name', 'ABM Name'], how='left')
    aggregated['Pending for Invoicing'] = aggregated['Pending for Invoicing'].fillna(0).astype(int)
    aggregated['Pending for Dispatch'] = aggregated['count_dispatch_pending']
    aggregated['Requests Dispatched'] = aggregated['count_delivered'] + aggregated['count_dispatched_in_transit'] + aggregated['count_rto']
    
    # Delivery Status
    aggregated['Delivered'] = aggregated['count_delivered']
    aggregated['Dispatched In Transit'] = aggregated['count_dispatched_in_transit']
    aggregated['RTO'] = aggregated['count_rto']
    
    # RTO Reasons (placeholder - might need more specific logic)
    aggregated['Incomplete Address'] = 0
    aggregated['Doctor Non Contactable'] = 0
    aggregated['Doctor Refused to Accept'] = 0
    aggregated['Hold Delivery'] = 0
    
    # Create the final summary dataframe
    print("📋 Creating final summary...")
    
    summary_columns = [
        'Area Name', 'ABM Name', 'Unique TBMs', 'Unique HCPs', 'Unique Requests', 'Requests Raised',
        'Request Cancelled Out of Stock', 'Action Pending at HO', 'Sent to HUB',
        'Pending for Invoicing', 'Pending for Dispatch', 'Requests Dispatched',
        'Delivered', 'Dispatched In Transit', 'RTO',
        'Incomplete Address', 'Doctor Non Contactable', 'Doctor Refused to Accept', 'Hold Delivery'
    ]
    
    final_summary = aggregated[summary_columns].copy()
    
    # Sort by Area Name
    final_summary = final_summary.sort_values('Area Name').reset_index(drop=True)
    
    # Save to CSV first for verification
    csv_output = f"zbm_summary_output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    final_summary.to_csv(csv_output, index=False)
    print(f"💾 Saved summary data to {csv_output} for verification")
    
    # Write values into template ZBM sheet preserving formatting
    print("💾 Writing values into template ZBM sheet (preserving formatting)...")
    
    try:
        from openpyxl import load_workbook
        from copy import copy as copy_style

        wb = load_workbook('zbm_summary.xlsx')
        ws = wb['ZBM']

        # Target starts at row 4 (1-based), column B (2): matches template layout
        start_row = 4
        start_col = 2

        # Columns order in sheet for data section
        col_order = [
            'Area Name', 'ABM Name', 'Unique TBMs', 'Unique HCPs', 'Unique Requests', 'Requests Raised',
            'Request Cancelled Out of Stock', 'Action Pending at HO', 'Sent to HUB',
            'Pending for Invoicing', 'Pending for Dispatch', 'Requests Dispatched',
            'Delivered', 'Dispatched In Transit', 'RTO',
            'Incomplete Address', 'Doctor Non Contactable', 'Doctor Refused to Accept', 'Hold Delivery'
        ]

        # Clear previous data rows (optional: clear 3000 rows)
        max_clear_rows = max(len(final_summary) + 50, 200)
        for r in range(start_row, start_row + max_clear_rows):
            for c in range(start_col, start_col + len(col_order)):
                ws.cell(row=r, column=c).value = None

        # Ensure enough styled rows: copy style from the first data row (row 4)
        def copy_row_style(src_row_idx, dst_row_idx):
            for c in range(start_col, start_col + len(col_order)):
                src = ws.cell(row=src_row_idx, column=c)
                dst = ws.cell(row=dst_row_idx, column=c)
                dst.number_format = src.number_format
                # Preserve fill/border/alignment, but enforce Calibri 10 Bold
                from openpyxl.styles import Font
                dst.font = Font(name='Calibri', size=10, bold=True, color=src.font.color)
                dst.alignment = copy_style(src.alignment)
                dst.border = copy_style(src.border)
                dst.fill = copy_style(src.fill)

        for i in range(len(final_summary)):
            target_row = start_row + i
            # If we're beyond existing styled rows, copy style from first data row
            if target_row > ws.max_row:
                ws.insert_rows(target_row)
            copy_row_style(start_row, target_row)
            for j, col_name in enumerate(col_order):
                ws.cell(row=target_row, column=start_col + j).value = (
                    None if col_name not in final_summary.columns else final_summary.at[i, col_name]
                )

        # Save as new timestamped file to avoid overwriting if Excel is open
        new_excel_file = f"zbm_summary_updated_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        wb.save(new_excel_file)
        print(f"✅ Successfully created updated Excel file with formatting preserved: {new_excel_file}")
        
        # Display summary statistics
        print("\n📊 Summary Statistics:")
        print(f"   Total ABM Territories: {len(final_summary)}")
        print(f"   Total Unique HCPs: {final_summary['Unique HCPs'].sum()}")
        print(f"   Total Unique Requests: {final_summary['Unique Requests'].sum()}")
        print(f"   Total Delivered: {final_summary['Delivered'].sum()}")
        print(f"   Total RTO: {final_summary['RTO'].sum()}")
        
        # Show first few rows
        print("\n📋 Sample of generated data:")
        print(final_summary.head(10).to_string(index=False))
        
    except Exception as e:
        print(f"❌ Error updating Excel file: {e}")
        return
    
    print("\n🎉 ZBM Summary automation completed successfully!")

def watch_and_build():
    import time
    import os
    path = 'master_tracker.csv'
    try:
        last_mtime = os.path.getmtime(path)
    except OSError:
        print("❌ master_tracker.csv not found for watch mode")
        return
    print("👀 Watching master_tracker.csv for changes. Press Ctrl+C to stop.")
    while True:
        try:
            time.sleep(2)
            mtime = os.path.getmtime(path)
            if mtime != last_mtime:
                print("🔁 Change detected. Rebuilding report...")
                last_mtime = mtime
                create_zbm_summary()
        except KeyboardInterrupt:
            print("👋 Watcher stopped.")
            break

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description='Generate ZBM summary or watch for changes')
    parser.add_argument('--watch', action='store_true', help='Watch master_tracker.csv and regenerate on change')
    args = parser.parse_args()
    if args.watch:
        watch_and_build()
    else:
        create_zbm_summary()
