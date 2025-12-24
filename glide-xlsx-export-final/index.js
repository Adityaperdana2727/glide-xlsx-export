const express = require("express");
const XLSX = require("xlsx");
const app = express();

app.use(express.json({ limit: "20mb" }));

/**
 * SIMPAN XLSX TERAKHIR DI MEMORY
 * (tanpa token, tanpa DB)
 */
let lastXlsxBuffer = null;

app.get("/", (req, res) => {
  res.send("Backend XLSX Export is running");
});

/**
 * 1) ENDPOINT GENERATE
 * Glide hanya KIRIM DATA
 */
app.post("/export-xlsx", (req, res) => {
  let rows = req.body?.rows ?? req.body;

  // Jika rows dikirim sebagai string (kasus Glide)
  if (typeof rows === "string") {
    try {
      rows = JSON.parse(rows);
    } catch (e) {
      return res.status(400).json({ error: "rows string tapi bukan JSON valid" });
    }
  }

  // Jika bentuk { rows: [...] }
  if (rows && typeof rows === "object" && Array.isArray(rows.rows)) {
    rows = rows.rows;
  }

  if (!Array.isArray(rows) || rows.length === 0) {
    return res.status(400).json({ error: "rows tidak valid / kosong" });
  }

  const headers = [
    "Input Date","Sellout Date","Delivery Date","Purchase Time",
    "SO Number","Employee Name","Employee ID","Branch",
    "Location","Dealer","Category","Sub Category","Model",
    "Price","Qty","Amount","Inc Target",
    "Customer","Contact","Address","Link Invoice","Status"
  ];

  const data = [headers];

  rows.forEach(r => {
    data.push([
      r.input_date,
      r.sellout_date,
      r.delivery_date,
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
    ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(data);

  // Hyperlink invoice
  rows.forEach((r, i) => {
    const ref = XLSX.utils.encode_cell({ r: i + 1, c: 20 });
    ws[ref] = {
      t: "s",
      v: "Link Invoice",
      l: { Target: r.url_invoice || "" }
    };
  });

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sell-Out");

  lastXlsxBuffer = XLSX.write(wb, {
    bookType: "xlsx",
    type: "buffer"
  });

  res.json({ success: true });
});

/**
 * 2) ENDPOINT DOWNLOAD
 * Dibuka via WebView Glide
 */
app.get("/download", (req, res) => {
  if (!lastXlsxBuffer) {
    return res.status(404).send("Belum ada file. Silakan generate dulu.");
  }

  res.setHeader(
    "Content-Disposition",
    "attachment; filename=SellOut.xlsx"
  );
  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );

  res.send(lastXlsxBuffer);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log("Server running on port", PORT);
});
