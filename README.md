# Kimbal Demand Intelligence Dashboard

This version hosts the Excel file inside the Vercel project and uses one backend data layer for both the dashboard and AI Q&A.

## Architecture

```text
/data/Smart_Meter_B2B_Demand_Plan_FY2526.xlsx
   ↓
/api/data on Vercel parses Excel using xlsx
   ↓
Dashboard charts and tables

User question
   ↓
/api/ask on Vercel
   ↓
Same parsed Excel data becomes AI context
   ↓
Groq Llama model
   ↓
AI answer in dashboard
```

## Where to put the Excel file

Place your Excel file here:

```text
data/Smart_Meter_B2B_Demand_Plan_FY2526.xlsx
```

Then commit and deploy to Vercel.

## Required Vercel environment variables

```text
GROQ_API_KEY=your_groq_api_key
```

Optional:

```text
GROQ_MODEL=llama-3.3-70b-versatile
EXCEL_FILE_NAME=Smart_Meter_B2B_Demand_Plan_FY2526.xlsx
EXCEL_FILE_PATH=data/Smart_Meter_B2B_Demand_Plan_FY2526.xlsx
DATA_CACHE_MS=300000
```

## Required Excel structure

The parser supports these sheet names and columns.

### 1. CUSTOMERS

Columns supported:

- `id`
- `name` or `customer`
- `state`
- `plan` or `annual_plan` or `units`
- `rev` or `revenue_cr` or `revenue`

### 2. SKUS

Columns supported:

- `id` or `sku`
- `desc` or `description`
- `pct` or `portfolio_pct`
- `annual` or `annual_plan` or `plan`
- `actual`
- `color` optional

### 3. MONTHLY

Columns supported:

- `month`
- `units` or `volume` or `plan`

### 4. SKU_MONTHLY

Columns supported:

- `sku`
- One column for each month, for example `Apr-25`, `May-25`, etc.

## Important note

This is simple and good for an MVP. Every time you change the Excel file, you need to redeploy Vercel. For live updates without redeploy, move the Excel to Google Drive, Supabase, or a database later.
