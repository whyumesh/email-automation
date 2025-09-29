# ZBM Summary Automation - Data Sources & Business Logic Explanation

## 📊 **DATA SOURCES OVERVIEW**

### **1. Primary Data Source: `master_tracker.csv`**
- **Purpose**: Main repository of all sample request transactions
- **Size**: 74,576 records (as of current data)
- **Encoding**: Latin-1 (handles special characters in names/addresses)
- **Update Frequency**: Real-time updates from Abbworld system

#### **Key Data Fields Used:**

| Field Name | Purpose | Data Type | Example |
|------------|---------|-----------|---------|
| `Assigned Request Ids` | Unique identifier for each request | String | REQ-5101072025 |
| `ABM Terr Code` | Territory code for Area Business Manager | String | IA000009 |
| `TBM HQ` | Territory Business Manager Headquarters | String | MUMBAI, AHMEDABAD, PUNE, NAGPUR |
| `ABM Name` | Area Business Manager name | String | Ruchit Rajendrakumar Andharia |
| `Doctor: Customer Code` | Healthcare Professional identifier | String | 2015-50078275018-TC |
| `Request Status` | Current status of the request | String | Delivered, Out of stock, Action pending |
| `Final Status` | Computed final status | String | Calculated using business rules |

---

## 🧠 **BUSINESS LOGIC EXPLANATION**

### **2. Logic Rules: `logic.xlsx`**

#### **Purpose:**
The logic file contains **37 business rules** that determine the final status of requests based on their status combinations. This is critical because:

- **Single requests can have multiple statuses** (e.g., "Delivered" + "Action pending")
- **Business needs to know the "true" final status** for reporting
- **Different status combinations require different business actions**

#### **Structure:**
- **Sheet2**: Contains the rule matrix
- **Columns**: Various status combinations
- **Final Answer Column**: The computed final status for each combination

#### **Example Rule Logic:**
```
If Request Status = ["Delivered", "Action pending"] 
Then Final Answer = "Delivered"

If Request Status = ["Out of stock", "On hold"] 
Then Final Answer = "Out of stock"

If Request Status = ["Request Raised", "Dispatch Pending"] 
Then Final Answer = "Dispatch Pending"
```

#### **Why This Logic is Critical:**
1. **Eliminates Ambiguity**: Multiple statuses per request are resolved to single final status
2. **Business Priority**: Rules prioritize the most important status (e.g., "Delivered" over "Action pending")
3. **Consistent Reporting**: All stakeholders see the same final status
4. **Actionable Intelligence**: Management can make decisions based on clear status

---

## 🔄 **DATA PROCESSING WORKFLOW**

### **Step 1: Data Cleaning & Filtering**
```python
# Filter for specific TBM HQ cities only
allowed_hq = {"MUMBAI", "AHMEDABAD", "PUNE", "NAGPUR"}
df = df[df['TBM HQ'].str.upper().isin(allowed_hq)]

# Remove incomplete records
df = df.dropna(subset=['ABM Terr Code', 'TBM HQ', 'ABM Name'])
```

**Why This Filtering:**
- **Geographic Focus**: Only processes data for specified regions
- **Data Quality**: Ensures all required fields are present
- **Business Scope**: Aligns with ZBM territory coverage

### **Step 2: Territory Aggregation**
```python
# Create Area Name: "TBM HQ - ABM Terr Code"
df['Area Name'] = df['TBM HQ'] + ' - ' + df['ABM Terr Code']

# Group by territory and ABM
aggregated = df.groupby(['Area Name', 'ABM Name']).agg({
    'ABM Terr Code': 'nunique',           # Unique TBMs
    'Doctor: Customer Code': 'nunique',   # Unique HCPs
    'Assigned Request Ids': 'nunique',     # Unique Requests
})
```

**Why This Aggregation:**
- **Territory Management**: Groups data by sales territories
- **Performance Metrics**: Calculates key performance indicators
- **Management Reporting**: Provides territory-wise insights

### **Step 3: Status Categorization**
```python
status_categories = {
    'out_of_stock_on_hold': ['Out of stock', 'On hold', 'Not permitted'],
    'request_raised': ['Request Raised'],
    'delivered_return_action_pending': ['Delivered', 'Return', 'Action pending / In Process'],
    'action_pending': ['Action pending / In Process'],
    'dispatch_pending': ['Dispatch Pending'],
    'delivered': ['Delivered'],
    'dispatched_in_transit': ['Dispatched & In Transit'],
    'rto': ['RTO']
}
```

**Why These Categories:**
- **Business Process Mapping**: Each category represents a business process stage
- **Performance Tracking**: Enables measurement of process efficiency
- **Issue Identification**: Helps identify bottlenecks and problems

---

## 📈 **CALCULATION FORMULAS EXPLAINED**

### **Template Formula Implementation:**

#### **1. Requests Raised (A + B + C)**
```
Requests Raised = Request Cancelled Out of Stock + Action Pending at HO + Sent to HUB
```
**Business Logic:**
- **A (Cancelled)**: Requests that couldn't be fulfilled due to stock issues
- **B (Action Pending)**: Requests waiting for internal processing
- **C (Sent to HUB)**: Requests that have been processed and sent for delivery

#### **2. Sent to HUB (D + E + F)**
```
Sent to HUB = Pending for Invoicing + Pending for Dispatch + Requests Dispatched
```
**Business Logic:**
- **D (Pending Invoicing)**: Requests processed but invoicing pending
- **E (Pending Dispatch)**: Requests invoiced but dispatch pending
- **F (Dispatched)**: Requests that have been dispatched

#### **3. Requests Dispatched (G + H + I)**
```
Requests Dispatched = Delivered + Dispatched In Transit + RTO
```
**Business Logic:**
- **G (Delivered)**: Successfully delivered to HCPs
- **H (In Transit)**: Currently being delivered
- **I (RTO)**: Returned to origin due to delivery issues

---

## 🎯 **WHY THIS DATA STRUCTURE MATTERS**

### **1. Business Intelligence**
- **Territory Performance**: Compare performance across different ABM territories
- **Process Efficiency**: Identify bottlenecks in the sample request process
- **Resource Allocation**: Understand where to focus efforts

### **2. Management Reporting**
- **Executive Dashboards**: High-level summary for senior management
- **Operational Reports**: Detailed metrics for operations teams
- **Compliance Tracking**: Ensure all requests are properly tracked

### **3. Process Optimization**
- **Bottleneck Identification**: Find where requests get stuck
- **Success Rate Analysis**: Measure delivery success rates
- **Resource Planning**: Plan inventory and logistics based on demand

### **4. Stakeholder Communication**
- **HCP Updates**: Inform healthcare professionals about request status
- **Internal Updates**: Keep sales teams informed about territory performance
- **Management Briefings**: Provide accurate data for decision-making

---

## 🔍 **DATA QUALITY & VALIDATION**

### **Data Quality Checks:**
1. **Completeness**: All required fields must be present
2. **Consistency**: Territory codes must match TBM HQ assignments
3. **Accuracy**: Request IDs must be unique and valid
4. **Timeliness**: Data should be current and up-to-date

### **Business Rule Validation:**
1. **Status Logic**: All status combinations must have defined rules
2. **Formula Balance**: All calculations must balance correctly
3. **Territory Mapping**: All territories must be properly mapped
4. **Performance Metrics**: All KPIs must be calculable

---

## 📋 **TEMPLATE STRUCTURE EXPLANATION**

### **Why Template Format is Critical:**
1. **Consistency**: Ensures all reports look identical
2. **Professional Appearance**: Maintains corporate branding
3. **Ease of Use**: Familiar format for stakeholders
4. **Automation**: Enables automated report generation

### **Template Sections:**
- **Headers (Rows 1-2)**: Column definitions and sub-headers
- **Section Labels (Row 3)**: HO, HUB, Delivery Status, RTO Reasons
- **Data Rows (Row 4+)**: Territory-wise metrics
- **Total Row**: Summary calculations

---

## 🚀 **PRODUCTION READINESS**

### **System Reliability:**
- ✅ **Error Handling**: Graceful failure management
- ✅ **Data Validation**: Comprehensive data quality checks
- ✅ **Format Preservation**: Exact template formatting maintained
- ✅ **Calculation Accuracy**: All formulas verified and balanced

### **Business Value:**
- ✅ **Time Savings**: Automated report generation
- ✅ **Accuracy**: Eliminates manual calculation errors
- ✅ **Consistency**: Standardized reporting format
- ✅ **Scalability**: Handles large data volumes efficiently

This automation system transforms raw transaction data into actionable business intelligence, enabling data-driven decision making and process optimization in the pharmaceutical sample request management process.
