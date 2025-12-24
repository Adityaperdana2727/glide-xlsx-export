const express = require("express");
const XLSX = require("xlsx");
const path = require("path");

const app = express();
app.use(express.json({ limit: "20mb" }));
app.use(express.static("public"));

const NETLIFY_PREVIEW_BASE = "https://peppy-mandazi-47f02f.netlify.app/";
const MIN_W = 8;
const MAX_W = 40;

function excelDateSerial(dateStr) {
  if (!dateStr) return "";
  const m = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return dateStr;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return Math.floor((d - new Date(1899, 11, 30)) / 86400000);
}

function makeInvoiceLink(raw, rowId) {
  return `${NETLIFY_PREVIEW_BASE}?u=${encodeURIComponent(raw)}&row_id=${encodeURIComponent(rowId)}`;
}

app.post("/export-xlsx", (req, res) => {
  const rows = req.body.rows;
  if (!Array.isArray(rows) || rows.length === 0) {
    return res.status(400).json({ error: "rows tidak valid / kosong" });
  }

  const headers = [
    "Input Date","Sellout Date","Delivery Date","Purchase Time",
    "SO Number","Employee Name","Employee ID","Branch","Location",
    "Dealer","Category","Sub Category","Model",
    "Price","Qty","Amount","Inc Target",
    "Customer","Contact","Address","Link Invoice","Status"
  ];

  const aoa = [headers];
  const colMax = headers.map(h => h.length);

  rows.forEach(r => {
    const row = [
      excelDateSerial(r.input_date),
      excelDateSerial(r.sellout_date),
      excelDateSerial(r.delivery_date),
      r.purchase_time,
      r.so_number,
      r.employee_name,
      r.employee_id,
      r.branch_area,
      r.location,
      r.dealer,
      r.category,
      r.sub_category,
      r.model,
      Number(r.price || 0),
      Number(r.qty || 0),
      Number(r.amount || 0),
      Number(r.incentive_target || 0),
      r.customer_name,
      r.contact,
      r.address,
      "Link Invoice",
      r.status
    ];

    row.forEach((v, i) => {
      colMax[i] = Math.max(colMax[i], String(v ?? "").length);
    });

    aoa.push(row);
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  [0,1,2].forEach(c => {
    for (let r = 1; r < aoa.length; r++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && typeof cell.v === "number") {
        cell.t = "n";
        cell.z = "dd/mm/yyyy";
      }
    }
  });

  rows.forEach((r, i) => {
    const ref = XLSX.utils.encode_cell({ r: i + 1, c: 20 });
    ws[ref] = {
      t: "s",
      v: "Link Invoice",
      l: { Target: makeInvoiceLink(r.url_invoice, r.row_id) }
    };
  });

  ws["!cols"] = colMax.map(len => ({
    wch: Math.min(MAX_W, Math.max(MIN_W, len + 2))
  }));

  ws["!freeze"] = { ySplit: 1 };

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sell-Out");

  const filename = `SellOut-${Date.now()}.xlsx`;
  const filepath = path.join(__dirname, "public", filename);
  XLSX.writeFile(wb, filepath);

  res.json({
    success: true,
    download_url: `/${filename}`
  });
});

app.listen(3000, () => {
  console.log("Server running on port 3000");
});
