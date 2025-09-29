# ZBM Automation - Data Flow & Business Logic Diagram

## 🔄 **DATA FLOW PROCESS**

```
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│   master_tracker│    │   logic.xlsx    │    │ zbm_summary.xlsx│
│      .csv       │    │   (37 rules)    │    │   (template)    │
│  (74,576 recs)  │    │                 │    │                 │
└─────────┬───────┘    └─────────┬───────┘    └─────────┬───────┘
          │                      │                      │
          ▼                      ▼                      ▼
┌─────────────────────────────────────────────────────────────────┐
│                    DATA PROCESSING PIPELINE                     │
├─────────────────────────────────────────────────────────────────┤
│ 1. DATA LOADING & VALIDATION                                    │
│    • Load CSV with multiple encoding support                    │
│    • Validate required columns exist                           │
│    • Check data quality (nulls, duplicates)                    │
│                                                                 │
│ 2. DATA CLEANING & FILTERING                                   │
│    • Filter by TBM HQ: MUMBAI, AHMEDABAD, PUNE, NAGPUR        │
│    • Remove incomplete records                                  │
│    • Create Area Name: "TBM HQ - ABM Terr Code"               │
│                                                                 │
│ 3. BUSINESS LOGIC APPLICATION                                  │
│    • Group requests by Request ID                              │
│    • Apply 37 business rules from logic.xlsx                   │
│    • Calculate Final Answer for each request                   │
│                                                                 │
│ 4. TERRITORY AGGREGATION                                       │
│    • Group by Area Name and ABM Name                           │
│    • Calculate unique counts (TBMs, HCPs, Requests)           │
│    • Categorize by status types                                │
│                                                                 │
│ 5. FORMULA CALCULATIONS                                        │
│    • Requests Raised = A + B + C                               │
│    • Sent to HUB = D + E + F                                   │
│    • Requests Dispatched = G + H + I                           │
│                                                                 │
│ 6. TEMPLATE POPULATION                                         │
│    • Preserve exact formatting from template                   │
│    • Map data to correct columns                               │
│    • Add total row with calculations                           │
└─────────────────────────────────────────────────────────────────┘
          │
          ▼
┌─────────────────────────────────────────────────────────────────┐
│                    OUTPUT FILES                                │
├─────────────────────────────────────────────────────────────────┤
│ • zbm_summary_updated_TIMESTAMP.xlsx (Formatted Report)        │
│ • zbm_summary_output_TIMESTAMP.csv (Data Verification)          │
│ • final_output_TIMESTAMP.xlsx (Request Status Logic)           │
└─────────────────────────────────────────────────────────────────┘
```

## 🧠 **BUSINESS LOGIC DETAILED BREAKDOWN**

### **Status Categories & Their Business Meaning:**

```
┌─────────────────────────────────────────────────────────────────┐
│                    STATUS CATEGORIES                            │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│ 📦 OUT OF STOCK / ON HOLD                                       │
│    • Out of stock: Product not available                        │
│    • On hold: Request temporarily suspended                     │
│    • Not permitted: Request not allowed                         │
│                                                                 │
│ 📋 REQUEST RAISED                                               │
│    • Initial request submitted by HCP                          │
│    • Waiting for processing                                     │
│                                                                 │
│ ⚙️ ACTION PENDING / IN PROCESS                                 │
│    • Internal processing required                               │
│    • Documentation or approval pending                          │
│                                                                 │
│ 🚚 DISPATCH PENDING                                             │
│    • Ready for dispatch but waiting for logistics              │
│    • Packaging or labeling in progress                         │
│                                                                 │
│ ✅ DELIVERED                                                    │
│    • Successfully delivered to HCP                             │
│    • Request completed successfully                             │
│                                                                 │
│ 🚛 DISPATCHED & IN TRANSIT                                     │
│    • Currently being delivered                                 │
│    • In transit to destination                                  │
│                                                                 │
│ ↩️ RTO (RETURN TO ORIGIN)                                      │
│    • Delivery failed                                            │
│    • Returned due to various reasons                           │
└─────────────────────────────────────────────────────────────────┘
```

### **Formula Logic Explanation:**

```
┌─────────────────────────────────────────────────────────────────┐
│                    CALCULATION FORMULAS                         │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│ 📊 REQUESTS RAISED = A + B + C                                  │
│    A = Request Cancelled Out of Stock                          │
│    B = Action Pending at HO                                    │
│    C = Sent to HUB                                             │
│                                                                 │
│    Business Logic: Total requests that entered the system      │
│                                                                 │
│ 📦 SENT TO HUB = D + E + F                                     │
│    D = Pending for Invoicing                                   │
│    E = Pending for Dispatch                                    │
│    F = Requests Dispatched                                     │
│                                                                 │
│    Business Logic: Requests that left HO and went to HUB       │
│                                                                 │
│ 🚚 REQUESTS DISPATCHED = G + H + I                            │
│    G = Delivered                                               │
│    H = Dispatched In Transit                                   │
│    I = RTO                                                     │
│                                                                 │
│    Business Logic: Requests that left HUB for delivery         │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

## 🎯 **WHY THIS LOGIC IS CRITICAL FOR BUSINESS**

### **1. Process Visibility**
- **End-to-End Tracking**: From request to delivery
- **Bottleneck Identification**: Where requests get stuck
- **Performance Measurement**: Success rates and efficiency

### **2. Resource Management**
- **Inventory Planning**: Based on request patterns
- **Logistics Optimization**: Dispatch and delivery planning
- **Staff Allocation**: Focus efforts where needed

### **3. Stakeholder Communication**
- **HCP Updates**: Accurate status for healthcare professionals
- **Management Reporting**: Clear metrics for decision-making
- **Compliance**: Audit trail for regulatory requirements

### **4. Business Intelligence**
- **Territory Performance**: Compare ABM territories
- **Trend Analysis**: Identify patterns and opportunities
- **Process Improvement**: Data-driven optimization

## 🔍 **DATA QUALITY ASSURANCE**

### **Validation Points:**
1. **Input Validation**: File existence, encoding, structure
2. **Data Completeness**: Required fields present
3. **Business Rule Validation**: All status combinations covered
4. **Calculation Verification**: Formulas balance correctly
5. **Output Validation**: Template format preserved

### **Error Handling:**
- **Graceful Degradation**: System continues with partial data
- **Clear Error Messages**: Specific guidance for issues
- **Fallback Options**: Alternative processing paths
- **Audit Trail**: Complete logging of all operations

This comprehensive system ensures reliable, accurate, and professional reporting for pharmaceutical sample request management.
