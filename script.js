let originalData = [];
let currentData = [];
let kotaList = [];
let produkList = [];
let kotaColumn = "";
let produkColumn = "";
let orderCountMap = {};

document.getElementById('fileInput').addEventListener('change', handleFile);

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, {type: 'array', cellDates: true});
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet, {raw: false, defval: ""});

      if (json.length === 0) {
        alert("File Excel kosong!");
        return;
      }

      const keys = Object.keys(json[0]);
      const findCol = (words) => keys.find(k => words.some(w => k.toLowerCase().includes(w))) || null;

      const colNo      = findCol(["no","id"]) || keys[0];
      const colNama    = findCol(["nama","name"]) || keys[1];
      const colTelp    = findCol(["tlp","hp","wa","phone","nomor","telp"]) || keys[2];
      kotaColumn       = findCol(["kota","city","kab","daerah","region"]) || keys[3];
      produkColumn     = findCol(["produk","barang","item","product"]) || keys[4];
      const colTanggal = findCol(["tanggal","date","tgl","order"]) || keys[5];

      originalData = json.map(row => ({
        No:       row[colNo] ?? "",
        Nama:     row[colNama] ?? "",
        NomorTLP: (row[colTelp] ?? "").toString().trim(),
        Kota:     row[kotaColumn] ?? "",
        Produk:   row[produkColumn] ?? "",
        Tanggal:  row[colTanggal] ?? ""
      })).filter(r => r.NomorTLP);

      kotaList   = [...new Set(originalData.map(r => r.Kota).filter(Boolean))].sort((a,b) => a.localeCompare(b));
      produkList = [...new Set(originalData.map(r => r.Produk).filter(Boolean))].sort((a,b) => a.localeCompare(b));

      currentData = [...originalData];
      currentData.sort((a,b) => (a.Nama || "").localeCompare(b.Nama || ""));

      showTable(currentData);
      populateKotaFilter();
      populateProdukFilter();
      document.getElementById('controls').style.display = 'block';
      document.getElementById('status').innerHTML = `Berhasil load ${originalData.length} baris data!`;
      calculateOrderStats();

    } catch (err) {
      alert("Error baca file: " + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

function showTable(data) {
  if (data.length === 0) {
    document.getElementById('tableContainer').innerHTML = '<p style="text-align:center;color:#ff6666;font-size:1.3rem;margin:40px 0;">Semua data sudah diexport!</p>';
    return;
  }
  let html = '<table><tr><th>No</th><th>Nama</th><th>Nomor TLP</th><th>Kota</th><th>Produk</th><th>Tanggal</th></tr>';
  data.forEach(r => {
    html += `<tr>
      <td>${r.No}</td>
      <td>${r.Nama}</td>
      <td>${r.NomorTLP}</td>
      <td>${r.Kota}</td>
      <td>${r.Produk}</td>
      <td>${r.Tanggal}</td>
    </tr>`;
  });
  html += '</table>';
  document.getElementById('tableContainer').innerHTML = html;
}

// MODE 1 & 3: Tetap rapih A-Z
function applyFilterKota() {
  const k = document.getElementById('filterKota').value;
  currentData = k ? originalData.filter(r => r.Kota === k) : [...originalData];
  currentData.sort((a,b) => (a.Nama || "").localeCompare(b.Nama || ""));
  showTable(currentData);
  document.getElementById('btnExportKota').disabled = false;
}

function applyFilterProduk() {
  const p = document.getElementById('filterProduk').value;
  currentData = p ? originalData.filter(r => r.Produk === p) : [...originalData];
  currentData.sort((a,b) => (a.Nama || "").localeCompare(b.Nama || ""));
  showTable(currentData);
  document.getElementById('btnExportProduk').disabled = false;
}

// MODE 4: SUPER RAPIH → Loyal dulu + Nama A-Z
function calculateOrderStats() {
  orderCountMap = {};
  originalData.forEach(row => {
    const key = row.NomorTLP.trim();
    if (key) orderCountMap[key] = (orderCountMap[key] || 0) + 1;
  });
  const once = Object.values(orderCountMap).filter(c => c === 1).length;
  const twice = Object.values(orderCountMap).filter(c => c === 2).length;
  const threePlus = Object.values(orderCountMap).filter(c => c >= 3).length;
  document.getElementById('orderStats').innerHTML = `
    <strong>Statistik Order:</strong><br>
    1× Order: <b>${once}</b>  |  2× Order: <b>${twice}</b>  |  3×+ Order: <b>${threePlus}</b> (Loyal!)
  `;
}

function filterByOrderCount(countType) {
  let filtered = [];
  if (countType === 1) filtered = originalData.filter(r => (orderCountMap[r.NomorTLP.trim()] || 0) === 1);
  else if (countType === 2) filtered = originalData.filter(r => (orderCountMap[r.NomorTLP.trim()] || 0) === 2);
  else if (countType === 3) filtered = originalData.filter(r => (orderCountMap[r.NomorTLP.trim()] || 0) >= 3);

  // SUPER SORT: Loyal dulu → kalau sama → Nama A-Z
  filtered.sort((a, b) => {
    const countA = orderCountMap[a.NomorTLP.trim()] || 0;
    const countB = orderCountMap[b.NomorTLP.trim()] || 0;
    if (countB !== countA) {
      return countB - countA; // yang lebih sering di atas
    }
    return (a.Nama || "").localeCompare(b.Nama || ""); // kalau sama → urut nama
  });

  currentData = filtered;
  showTable(currentData);
  document.getElementById('btnExportOrder').disabled = filtered.length === 0;
  document.getElementById('status').innerHTML = `Menampilkan ${countType === 3 ? '3+' : countType + '×'} order → ${filtered.length} pelanggan (sudah urut rapih!)`;
}

function exportOrderGroup() {
  if (currentData.length === 0) return;

  // Pastikan urutan tetap rapih sebelum export
  currentData.sort((a, b) => {
    const countA = orderCountMap[a.NomorTLP.trim()] || 0;
    const countB = orderCountMap[b.NomorTLP.trim()] || 0;
    if (countB !== countA) return countB - countA;
    return (a.Nama || "").localeCompare(b.Nama || "");
  });

  const counts = [...new Set(currentData.map(r => orderCountMap[r.NomorTLP.trim()]))];
  const label = counts.length === 1
    ? (counts[0] === 1 ? "1x_Order" : counts[0] === 2 ? "2x_Order" : "3plus_Order_Loyal")
    : "Mixed_Order";

  const ws = XLSX.utils.json_to_sheet(currentData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Data");
  XLSX.writeFile(wb, `${label}_RAPIH_${new Date().toISOString().slice(0,10)}.xlsx`);

  originalData = originalData.filter(o => !currentData.some(c => c.NomorTLP === o.NomorTLP));
  currentData = [...originalData];
  showTable(currentData);
  document.getElementById('status').innerHTML = `Export ${label} selesai! → ${originalData.length} tersisa`;
  document.getElementById('btnExportOrder').disabled = true;
  calculateOrderStats();
}

// Fungsi lain tetap sama (populate, exportByKota, dll)
function populateKotaFilter() {
  const sel = document.getElementById('filterKota');
  sel.innerHTML = '<option value="">-- Pilih Kota --</option>';
  kotaList.forEach(k => {
    const cnt = originalData.filter(r => r.Kota === k).length;
    sel.innerHTML += `<option value="${k}">${k} (${cnt})</option>`;
  });
}

function populateProdukFilter() {
  const sel = document.getElementById('filterProduk');
  sel.innerHTML = '<option value="">-- Pilih Produk --</option>';
  produkList.forEach(p => {
    const cnt = originalData.filter(r => r.Produk === p).length;
    sel.innerHTML += `<option value="${p}">${p} (${cnt})</option>`;
  });
}

function exportKota() {
  const k = document.getElementById('filterKota').value || 'SemuaKota';
  const ws = XLSX.utils.json_to_sheet(currentData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Data");
  XLSX.writeFile(wb, `${k}_Export_${new Date().toISOString().slice(0,10)}.xlsx`);
  originalData = originalData.filter(o => !currentData.some(c => c.NomorTLP === o.NomorTLP));
  currentData = [...originalData];
  showTable(currentData);
  document.getElementById('status').innerHTML = `Export kota "${k}" selesai → ${originalData.length} tersisa`;
  document.getElementById('btnExportKota').disabled = true;
}

function exportProdukManual() {
  const p = document.getElementById('filterProduk').value || 'SemuaProduk';
  const safeName = p.replace(/[^a-zA-Z0-9]/g, '_');
  const ws = XLSX.utils.json_to_sheet(currentData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Data");
  XLSX.writeFile(wb, `${safeName}_Export_${new Date().toISOString().slice(0,10)}.xlsx`);
  originalData = originalData.filter(o => !currentData.some(c => c.NomorTLP === o.NomorTLP));
  currentData = [...originalData];
  showTable(currentData);
  document.getElementById('status').innerHTML = `Export produk "${p}" selesai → ${originalData.length} tersisa`;
  document.getElementById('btnExportProduk').disabled = true;
}

// Mode 2 & export all (sudah ada sort nama)
function exportByKota() {
  let c = 0;
  kotaList.forEach(k => {
    const d = originalData.filter(r => r.Kota === k);
    if (d.length > 0) {
      const ws = XLSX.utils.json_to_sheet(d.sort((a,b) => a.Nama.localeCompare(b.Nama)));
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Data");
      XLSX.writeFile(wb, `${k}_Auto_${new Date().toISOString().slice(0,10)}.xlsx`);
      c++;
    }
  });
  document.getElementById('status').innerHTML = `${c} file kota selesai!`;
}

function exportByKotaOneFile() {
  const wb = XLSX.utils.book_new();
  kotaList.forEach(k => {
    const d = originalData.filter(r => r.Kota === k);
    if (d.length > 0) {
      const ws = XLSX.utils.json_to_sheet(d.sort((a,b) => a.Nama.localeCompare(b.Nama)));
      XLSX.utils.book_append_sheet(wb, ws, k.substring(0,30));
    }
  });
  XLSX.writeFile(wb, `SemuaKota_${new Date().toISOString().slice(0,10)}.xlsx`);
}

function exportAllProdukSeparate() {
  let count = 0;
  produkList.forEach(p => {
    const dataProd = originalData.filter(r => r.Produk === p);
    if (dataProd.length > 0) {
      const ws = XLSX.utils.json_to_sheet(dataProd.sort((a,b) => a.Nama.localeCompare(b.Nama)));
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Data");
      const safeName = p.replace(/[^a-zA-Z0-9]/g, '_').substring(0,30);
      XLSX.writeFile(wb, `${safeName}_Auto_${new Date().toISOString().slice(0,10)}.xlsx`);
      count++;
    }
  });
  document.getElementById('status').innerHTML = `SELESAI! ${count} file produk terpisah!`;
}

function exportAllProdukOneFile() {
  const wb = XLSX.utils.book_new();
  let sheetCount = 0;
  produkList.forEach(p => {
    const dataProd = originalData.filter(r => r.Produk === p);
    if (dataProd.length > 0) {
      const ws = XLSX.utils.json_to_sheet(dataProd.sort((a,b) => a.Nama.localeCompare(b.Nama)));
      const sheetName = p.replace(/[^a-zA-Z0-9]/g, '_').substring(0,31);
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
      sheetCount++;
    }
  });
  if (sheetCount > 0) {
    XLSX.writeFile(wb, `Semua_Produk_${new Date().toISOString().slice(0,10)}.xlsx`);
    document.getElementById('status').innerHTML = `SELESAI! 1 file dengan ${sheetCount} sheet!`;
  }
}

function switchMode(m) {
  ['mode1','mode2','mode3','mode4'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.style.display = 'none';
  });
  document.getElementById('mode' + m).style.display = 'block';
  if (m === 4) calculateOrderStats();
}

function resetData() {
  if (confirm("Yakin reset semua dan upload ulang?")) location.reload();
}