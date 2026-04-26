import fs from "fs/promises";
import path from "path";
import * as XLSX from "xlsx";

const MONTHS = ["Apr-25","May-25","Jun-25","Jul-25","Aug-25","Sep-25","Oct-25","Nov-25","Dec-25","Jan-26","Feb-26","Mar-26"];

const FALLBACK_DATA = {
  meta: { source: "Fallback static sample", fileName: "Smart_Meter_B2B_Demand_Plan_FY2526.xlsx", loadedAt: new Date().toISOString(), warning: "Excel file not found in /data, using bundled sample data." },
  months: MONTHS,
  monthly: [777728,812057,870405,1018408,1093307,1074978,1011011,1020237,985691,883081,803990,832408],
  customers: [
    {id:"C01",name:"MSEDCL",state:"Maharashtra",plan:2640000,rev:770.72},
    {id:"C02",name:"UPPCL",state:"Uttar Pradesh",plan:2280000,rev:665.62},
    {id:"C03",name:"BSES Rajdhani",state:"Delhi",plan:1440000,rev:420.39},
    {id:"C04",name:"BESCOM",state:"Karnataka",plan:1200000,rev:350.33},
    {id:"C05",name:"TPDDL",state:"Delhi",plan:960000,rev:280.26},
    {id:"C06",name:"APEPDCL",state:"Andhra Pradesh",plan:1080000,rev:315.30},
    {id:"C07",name:"CESC Limited",state:"West Bengal",plan:840000,rev:245.23},
    {id:"C08",name:"BSES Yamuna",state:"Delhi",plan:720000,rev:210.20},
    {id:"C09",name:"TANGEDCO",state:"Tamil Nadu",plan:480000,rev:140.13},
    {id:"C10",name:"Adani Electricity",state:"Maharashtra",plan:360000,rev:105.10}
  ],
  skus: [
    {id:"SM-SP-NB",desc:"Single Phase NB-IoT",pct:38,annual:4560000,actual:4218707,color:"#00e5ff"},
    {id:"SM-SP-RF",desc:"Single Phase RF Mesh",pct:22,annual:2640000,actual:2471361,color:"#39d353"},
    {id:"SM-TP-NB",desc:"Three Phase NB-IoT",pct:14,annual:1680000,actual:1581790,color:"#ff6b35"},
    {id:"SM-TP-RF",desc:"Three Phase RF Mesh",pct:9,annual:1080000,actual:1002148,color:"#ffc107"},
    {id:"SM-CT-HT",desc:"CT Operated HT Meter",pct:5,annual:600000,actual:561809,color:"#a78bfa"},
    {id:"SM-PP-4G",desc:"Prepaid Smart 4G",pct:7,annual:840000,actual:787008,color:"#f472b6"},
    {id:"SM-PP-RF",desc:"Prepaid RF Mesh",pct:5,annual:600000,actual:560478,color:"#fdcb6e"}
  ],
  skuMonthly: {
    "SM-SP-NB":[290955,301738,332421,386336,412651,406758,378043,382676,376877,331301,302040,316911],
    "SM-SP-RF":[170492,180682,195392,225960,237641,237666,228134,221992,217763,198671,176812,180156],
    "SM-TP-NB":[110311,116509,119623,141224,157532,148419,146476,147807,135723,127302,116128,114736],
    "SM-TP-RF":[72263,74232,77721,90459,97819,95757,88127,93786,87707,77111,71899,75267],
    "SM-CT-HT":[39597,41186,42228,50238,54447,554418,51051,51673,48062,44944,40590,42375],
    "SM-PP-4G":[55107,56905,59824,72792,78143,76541,69903,71752,69378,61668,55423,59572],
    "SM-PP-RF":[39003,40805,43196,51399,55074,54419,49277,50551,50181,42084,41098,43391]
  }
};

let cache = null;
let cacheAt = 0;
const CACHE_MS = Number(process.env.DATA_CACHE_MS || 5 * 60 * 1000);

function cleanKey(k) {
  return String(k || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "");
}

function toNumber(v) {
  if (typeof v === "number") return v;
  if (v === null || v === undefined || v === "") return 0;
  return Number(String(v).replace(/[₹,\s]/g, "")) || 0;
}

function normaliseRows(rows) {
  return rows.map(row => {
    const out = {};
    Object.entries(row).forEach(([k, v]) => out[cleanKey(k)] = v);
    return out;
  });
}

function sheetRows(workbook, candidates) {
  const names = workbook.SheetNames;
  const found = names.find(n => candidates.map(cleanKey).includes(cleanKey(n)));
  if (!found) return [];
  return normaliseRows(XLSX.utils.sheet_to_json(workbook.Sheets[found], { defval: "" }));
}

function parseWorkbook(buffer, fileName) {
  const workbook = XLSX.read(buffer, { type: "buffer" });
  const customerRows = sheetRows(workbook, ["CUSTOMERS", "Customer", "Customer Plan", "Customer Ranking"]);
  const skuRows = sheetRows(workbook, ["SKUS", "SKU", "SKU Breakdown"]);
  const monthlyRows = sheetRows(workbook, ["MONTHLY", "Monthly", "Monthly Volume"]);
  const skuMonthlyRows = sheetRows(workbook, ["SKU_MONTHLY", "SKU Monthly", "Month SKU", "Month x SKU"]);

  const customers = customerRows.length ? customerRows.map((r, i) => ({
    id: String(r.id || r.customer_id || `C${String(i + 1).padStart(2, "0")}`),
    name: String(r.name || r.customer || r.customer_name || "").trim(),
    state: String(r.state || r.region || "").trim(),
    plan: toNumber(r.plan || r.annual_plan || r.units || r.volume),
    rev: toNumber(r.rev || r.revenue || r.revenue_cr || r.revenue_in_cr)
  })).filter(r => r.name) : FALLBACK_DATA.customers;

  const skus = skuRows.length ? skuRows.map((r, i) => ({
    id: String(r.id || r.sku || r.sku_id || "").trim(),
    desc: String(r.desc || r.description || r.sku_description || "").trim(),
    pct: toNumber(r.pct || r.portfolio_pct || r.portfolio_share),
    annual: toNumber(r.annual || r.annual_plan || r.plan),
    actual: toNumber(r.actual || r.actuals || r.achieved),
    color: String(r.color || FALLBACK_DATA.skus[i % FALLBACK_DATA.skus.length]?.color || "#00e5ff")
  })).filter(r => r.id) : FALLBACK_DATA.skus;

  const months = monthlyRows.length ? monthlyRows.map(r => String(r.month || r.months || r.period || "").trim()).filter(Boolean) : MONTHS;
  const monthly = monthlyRows.length ? monthlyRows.map(r => toNumber(r.units || r.volume || r.plan || r.monthly)).filter(n => n > 0) : FALLBACK_DATA.monthly;

  const skuMonthly = {};
  if (skuMonthlyRows.length) {
    skuMonthlyRows.forEach(r => {
      const sku = String(r.sku || r.id || r.sku_id || "").trim();
      if (!sku) return;
      skuMonthly[sku] = months.map(m => toNumber(r[cleanKey(m)] || r[m]));
    });
  }

  return {
    meta: { source: "Vercel-hosted Excel", fileName, loadedAt: new Date().toISOString() },
    months,
    monthly,
    customers,
    skus,
    skuMonthly: Object.keys(skuMonthly).length ? skuMonthly : FALLBACK_DATA.skuMonthly
  };
}

async function readExcelFromProject() {
  const fileName = process.env.EXCEL_FILE_NAME || "Smart_Meter_B2B_Demand_Plan_FY2526.xlsx";
  const relativePath = process.env.EXCEL_FILE_PATH || path.join("data", fileName);
  const absolutePath = path.isAbsolute(relativePath) ? relativePath : path.join(process.cwd(), relativePath);
  const buffer = await fs.readFile(absolutePath);
  return { buffer, fileName: path.basename(absolutePath) };
}

export async function getDashboardData() {
  if (cache && Date.now() - cacheAt < CACHE_MS) return cache;
  try {
    const { buffer, fileName } = await readExcelFromProject();
    cache = parseWorkbook(buffer, fileName);
  } catch (err) {
    cache = { ...FALLBACK_DATA, meta: { ...FALLBACK_DATA.meta, error: err.message, loadedAt: new Date().toISOString() } };
  }
  cacheAt = Date.now();
  return cache;
}

export function buildAIContext(data) {
  const totalUnits = data.customers.reduce((sum, c) => sum + Number(c.plan || 0), 0);
  const totalRevenue = data.customers.reduce((sum, c) => sum + Number(c.rev || 0), 0);
  const topCustomers = [...data.customers].sort((a, b) => b.plan - a.plan).slice(0, 20)
    .map(c => `${c.name} (${c.state}): ${(c.plan/100000).toFixed(2)}L units, Rs.${Number(c.rev || 0).toFixed(2)} Cr`).join("\n");
  const skuText = data.skus.map(s => `${s.id} - ${s.desc}: planned ${(s.annual/100000).toFixed(2)}L, actual ${(s.actual/100000).toFixed(2)}L, portfolio ${s.pct}%`).join("\n");
  const monthlyText = data.months.map((m, i) => `${m}: ${((data.monthly[i] || 0)/100000).toFixed(2)}L`).join(" | ");
  return `You are an expert demand planning analyst for Kimbal Technologies. Answer only using the supplied demand data. Be crisp, analytical, and use Indian number format.\n\nDATA SOURCE: ${data.meta?.source || "unknown"}\nFILE: ${data.meta?.fileName || "unknown"}\nLOADED AT: ${data.meta?.loadedAt || "unknown"}\nTOTAL ANNUAL PLAN: ${(totalUnits/10000000).toFixed(2)} Cr units\nTOTAL REVENUE: Rs.${totalRevenue.toFixed(2)} Cr\n\nCUSTOMERS:\n${topCustomers}\n\nSKUS:\n${skuText}\n\nMONTHLY DEMAND:\n${monthlyText}`;
}
