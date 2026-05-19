import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── SUPABASE ─────────────────────────────────────────────────
// NOTE: Ganti dengan project Supabase baru khusus cetakan,
// atau pakai yang sama dengan tabel berbeda (prefix "ct_")
const SUPABASE_URL = "https://irhsorgoodxwflxivtot.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImlyaHNvcmdvb2R4d2ZseGl2dG90Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg2MTM2ODMsImV4cCI6MjA5NDE4OTY4M30.V9C2lPkV1w1VV8YLO-1eZ0t7XJew7NJSfuZLzK4DsVs";
const H = { "Content-Type": "application/json", "apikey": SUPABASE_KEY, "Authorization": `Bearer ${SUPABASE_KEY}`, "Prefer": "return=representation" };

// Tabel pakai prefix "ct_" biar ga bentrok sama dashboard ATK
const api = {
  get: async (table, params = "") => {
    const r = await fetch(`${SUPABASE_URL}/rest/v1/ct_${table}?${params}`, { headers: H });
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  },
  upsert: async (table, data) => {
    const r = await fetch(`${SUPABASE_URL}/rest/v1/ct_${table}`, {
      method: "POST",
      headers: { ...H, "Prefer": "resolution=merge-duplicates,return=representation" },
      body: JSON.stringify(Array.isArray(data) ? data : [data]),
    });
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  },
  delete: async (table, filter) => {
    const r = await fetch(`${SUPABASE_URL}/rest/v1/ct_${table}?${filter}`, { method: "DELETE", headers: H });
    if (!r.ok) throw new Error(await r.text());
  },
};

// ─── DATA BARANG CETAKAN ──────────────────────────────────────
const INITIAL_ITEMS = [
  { kode: 1,  nama: "Amplop Dinas 1/2 Folio",                   satuan: "Pax"   },
  { kode: 2,  nama: "Amplop Dinas Besar",                        satuan: "Pax"   },
  { kode: 3,  nama: "Amplop Dinas Folio",                        satuan: "Pax"   },
  { kode: 4,  nama: "Amplop Dinas Kecil",                        satuan: "Pax"   },
  { kode: 5,  nama: "Amplop Polos 1/2 Folio",                    satuan: "Pax"   },
  { kode: 6,  nama: "Amplop Polos Folio",                        satuan: "Pax"   },
  { kode: 7,  nama: "Amplop Polos Kecil",                        satuan: "Pax"   },
  { kode: 8,  nama: "Ban Uang 10.000",                           satuan: "Bal"   },
  { kode: 9,  nama: "Ban Uang 100.000",                          satuan: "Bal"   },
  { kode: 10, nama: "Ban Uang 1.000",                            satuan: "Bal"   },
  { kode: 11, nama: "Ban Uang 20.000",                           satuan: "Bal"   },
  { kode: 12, nama: "Ban Uang 2.000",                            satuan: "Bal"   },
  { kode: 13, nama: "Ban Uang 50.000",                           satuan: "Bal"   },
  { kode: 14, nama: "Ban Uang 5.000",                            satuan: "Bal"   },
  { kode: 15, nama: "Bilyet Deposito",                           satuan: "Blok"  },
  { kode: 16, nama: "Bilyet Giro 10 Lembar",                     satuan: "Buku"  },
  { kode: 17, nama: "Bilyet Giro 25 Lembar",                     satuan: "Buku"  },
  { kode: 18, nama: "Blangko BKK",                               satuan: "Blok"  },
  { kode: 19, nama: "Blangko BPKK",                              satuan: "Blok"  },
  { kode: 20, nama: "Blangko KU",                                satuan: "Blok"  },
  { kode: 21, nama: "Blangko Tanda Penyetoran",                  satuan: "Blok"  },
  { kode: 22, nama: "Buku Cek 10 Lembar",                        satuan: "Blok"  },
  { kode: 23, nama: "Buku Cek 25 Lembar",                        satuan: "Blok"  },
  { kode: 24, nama: "Buku Tabungan Firdaus",                     satuan: "buah"  },
  { kode: 25, nama: "Buku Tabungan Pensiun",                     satuan: "buah"  },
  { kode: 26, nama: "Buku Tabungan Seulanga",                    satuan: "buah"  },
  { kode: 27, nama: "Buku Tabungan Simpeda",                     satuan: "buah"  },
  { kode: 28, nama: "Buku Tabungan Simpel",                      satuan: "buah"  },
  { kode: 29, nama: "Buku Tabungan TabunganKu",                  satuan: "buah"  },
  { kode: 30, nama: "Buku Tabungan TAG",                         satuan: "buah"  },
  { kode: 31, nama: "Buku Tabungan Sahara",                      satuan: "buah"  },
  { kode: 32, nama: "Form Kuasa AutoDebet",                      satuan: "Blok"  },
  { kode: 33, nama: "Form Pembukaan Rekening Nasabah",           satuan: "rim"   },
  { kode: 34, nama: "Form Penerimaan Pengaduan Klaim",           satuan: "rim"   },
  { kode: 35, nama: "Blangko Perintah Order Barang",             satuan: "Blok"  },
  { kode: 36, nama: "Form Permohonan Penggantian PIN",           satuan: "rim"   },
  { kode: 37, nama: "Form Rahn",                                 satuan: "Blok"  },
  { kode: 38, nama: "Form Registrasi CMS",                       satuan: "Blok"  },
  { kode: 39, nama: "Kop Surat",                                 satuan: "rim"   },
  { kode: 40, nama: "Tanda Penerimaan",                          satuan: "Blok"  },
  { kode: 41, nama: "Slip Penarikan",                            satuan: "Blok"  },
  { kode: 42, nama: "Slip Setoran",                              satuan: "Blok"  },
  { kode: 43, nama: "Spektroline Signature & Overlay",           satuan: "Blok"  },
  { kode: 44, nama: "Surat Rahn",                                satuan: "Blok"  },
  { kode: 45, nama: "Specimen Giro",                             satuan: "Lembar"},
  { kode: 46, nama: "Plastik Bilyet Deposito",                   satuan: "Buah"  },
  { kode: 47, nama: "Form Penutupan Rekening",                   satuan: "rim"   },
  { kode: 48, nama: "Blangko CN DN",                             satuan: "Blok"  },
  { kode: 49, nama: "Blangko PO Konsumsi",                       satuan: "Blok"  },
  { kode: 50, nama: "Blangko Tanda Terima PO Barang",            satuan: "Blok"  },
  { kode: 51, nama: "Blangko Surat Pengantar",                   satuan: "Blok"  },
  { kode: 52, nama: "Formulir Internet Banking Corporate",       satuan: "rim"   },
  { kode: 53, nama: "Kop Bank Garansi",                          satuan: "rim"   },
  { kode: 54, nama: "Voucher Tabungan Seulanga",                 satuan: "Blok"  },
  { kode: 55, nama: "Spectrolite Invisible Signature",           satuan: "kotak" },
  { kode: 56, nama: "Blangko Surat Pos",                         satuan: "Blok"  },
  { kode: 57, nama: "Nota Debet",                                satuan: "Blok"  },
  { kode: 58, nama: "BPWD Kliring Penyerahan Syariah",           satuan: "Blok"  },
  { kode: 59, nama: "BPWD Kliring Pengembalian Syariah",         satuan: "Blok"  },
  { kode: 60, nama: "KBWD Syariah",                              satuan: "Blok"  },
  { kode: 61, nama: "Formulir Permohonan ATM",                   satuan: "rim"   },
  { kode: 62, nama: "Formulir Standing Instruction",             satuan: "rim"   },
  { kode: 63, nama: "Map Hijau Bank Aceh",                       satuan: "buah"  },
  { kode: 64, nama: "Blanko Samsat",                             satuan: "Blok"  },
  { kode: 65, nama: "Nota Pembelian/Penjualan Uang Kertas Asing",satuan: "rim"   },
  { kode: 66, nama: "Aplikasi Transfer",                         satuan: "rim"   },
  { kode: 67, nama: "Formulir Pelaporan Lalu Lintas Devisa Incoming", satuan: "rim" },
];

const INITIAL_STOCK = {
  1:10, 2:10, 3:10, 4:10, 5:10, 6:5,  7:5,  8:5,  9:10, 10:15,
  11:15,12:20,13:20,14:20,15:4, 16:100,17:100,18:50,19:100,20:100,
  21:50,22:100,23:100,24:400,25:100,26:250,27:300,28:100,29:100,30:200,
  31:0, 32:0, 33:50, 34:2,  35:20, 36:2,  37:1,  38:0,  39:10, 40:14,
  41:50,42:400,43:100,44:1, 45:500,46:1000,47:5, 48:100,49:3, 50:20,
  51:1, 52:10,53:1, 54:3,  55:100,56:50, 57:0,  58:0,  59:0,  60:0,
  61:0, 62:0, 63:0, 64:0,  65:0,  66:0,  67:0,
};

const BIDANG_OPTIONS = ["Pemimpin","Wakil Pimpinan","Operational","Funding","Pembiayaan","Legal","Umum","OB","Driver","Security"];
const SATUAN_OPTIONS = ["Pax","Bal","Blok","Buku","buah","Buah","rim","Lembar","kotak","Lusin","Pak","Set","Roll","Botol"];

// ─── THEME: EMERALD/HIJAU untuk Cetakan ──────────────────────
const THEME = {
  primary:    "#0F6E56",
  primaryDark:"#0A5241",
  primaryBg:  "#E1F5EE",
  primaryText:"#0F6E56",
  primaryMid: "#1D9E75",
};

const today = () => new Date().toISOString().split("T")[0];
const formatDate = (iso) => !iso ? "-" : new Date(iso).toLocaleDateString("id-ID", { day:"2-digit", month:"short", year:"numeric" });
const formatDateTime = (iso) => !iso ? "-" : new Date(iso).toLocaleString("id-ID", { day:"2-digit", month:"short", year:"numeric", hour:"2-digit", minute:"2-digit" });

const Badge = ({ children, color = "gray" }) => {
  const C = { green:{bg:"#EAF3DE",text:"#3B6D11"}, red:{bg:"#FCEBEB",text:"#A32D2D"}, amber:{bg:"#FAEEDA",text:"#854F0B"}, blue:{bg:"#E6F1FB",text:"#185FA5"}, gray:{bg:"#F1EFE8",text:"#5F5E5A"}, teal:{bg:"#E1F5EE",text:"#0F6E56"}, purple:{bg:"#EEEDFE",text:"#3C3489"}, emerald:{bg:"#D1FAE5",text:"#065F46"} };
  const c = C[color] || C.gray;
  return <span style={{ background:c.bg, color:c.text, fontSize:11, fontWeight:500, padding:"2px 8px", borderRadius:6, display:"inline-block", whiteSpace:"nowrap" }}>{children}</span>;
};

const MetricCard = ({ label, value, icon, color = "teal", sub }) => {
  const P = { teal:{bg:"#E1F5EE",text:"#0F6E56",ic:"#1D9E75"}, green:{bg:"#EAF3DE",text:"#3B6D11",ic:"#639922"}, red:{bg:"#FCEBEB",text:"#A32D2D",ic:"#E24B4A"}, amber:{bg:"#FAEEDA",text:"#854F0B",ic:"#BA7517"} };
  const c = P[color] || P.teal;
  return (
    <div style={{ background:c.bg, borderRadius:12, padding:"14px 16px", display:"flex", flexDirection:"column", gap:4 }}>
      <div style={{ display:"flex", alignItems:"center", gap:6 }}>
        <span style={{ fontSize:18, color:c.ic }}>{icon}</span>
        <span style={{ fontSize:11, color:c.text, fontWeight:500, textTransform:"uppercase", letterSpacing:"0.04em" }}>{label}</span>
      </div>
      <div style={{ fontSize:26, fontWeight:700, color:c.text, lineHeight:1.2 }}>{value}</div>
      {sub && <div style={{ fontSize:11, color:c.text, opacity:0.7 }}>{sub}</div>}
    </div>
  );
};

const Modal = ({ modal, onClose }) => {
  if (!modal) return null;
  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.5)", zIndex:9999, display:"flex", alignItems:"center", justifyContent:"center", padding:16 }}>
      <div style={{ background:"#fff", borderRadius:14, padding:"24px", maxWidth:360, width:"100%", boxShadow:"0 8px 40px rgba(0,0,0,0.2)" }}>
        {modal.icon && <div style={{ fontSize:28, marginBottom:10, textAlign:"center" }}>{modal.icon}</div>}
        <div style={{ fontSize:15, fontWeight:600, color:"#111", marginBottom: modal.message ? 8 : 20 }}>{modal.title}</div>
        {modal.message && <div style={{ fontSize:13, color:"#555", marginBottom:20, lineHeight:1.5 }}>{modal.message}</div>}
        <div style={{ display:"flex", gap:8, justifyContent:"flex-end" }}>
          {modal.type === "confirm" && <button onClick={onClose} style={{ background:"none", border:"1px solid #ddd", borderRadius:8, padding:"8px 16px", cursor:"pointer", fontSize:13, color:"#555" }}>Batal</button>}
          <button onClick={() => { modal.onOk?.(); onClose(); }} style={{ background: modal.danger ? "#A32D2D" : THEME.primary, color:"#fff", border:"none", borderRadius:8, padding:"8px 18px", cursor:"pointer", fontSize:13, fontWeight:600 }}>{modal.okLabel || "OK"}</button>
        </div>
      </div>
    </div>
  );
};

export default function App() {
  const [tab, setTab] = useState("overview");
  const [transactions, setTransactions] = useState([]);
  const [customItems, setCustomItems] = useState([]);
  const [itemOverrides, setItemOverrides] = useState({});
  const [deletedBase, setDeletedBase] = useState([]);
  const [modal, setModal] = useState(null);
  const [loading, setLoading] = useState(true);
  const [saveStatus, setSaveStatus] = useState("");
  const [lastUpdated, setLastUpdated] = useState(null);
  const [searchOverview, setSearchOverview] = useState("");
  const [sortField, setSortField] = useState("kode");
  const [sortDir, setSortDir] = useState("asc");
  const [searchHistory, setSearchHistory] = useState("");
  const [expandedTx, setExpandedTx] = useState({});
  const [searchKelola, setSearchKelola] = useState("");
  const [editingKode, setEditingKode] = useState(null);
  const [editForm, setEditForm] = useState({});
  const [masukDate, setMasukDate] = useState(today());
  const [masukKeterangan, setMasukKeterangan] = useState("");
  const [masukItems, setMasukItems] = useState([{ kode:"", jumlah:"" }]);
  const [keluarDate, setKeluarDate] = useState(today());
  const [keluarKeterangan, setKeluarKeterangan] = useState("");
  const [keluarPenerima, setKeluarPenerima] = useState("");
  const [keluarItems, setKeluarItems] = useState([{ kode:"", jumlah:"" }]);
  const [newNama, setNewNama] = useState("");
  const [newSatuan, setNewSatuan] = useState("buah");
  const [newSatuanCustom, setNewSatuanCustom] = useState("");
  const [newStokAwal, setNewStokAwal] = useState("0");

  const showAlert = (title, message, icon="ℹ️") => setModal({ type:"alert", title, message, icon });
  const showConfirm = (title, message, onOk, danger=false, icon="⚠️") =>
    setModal({ type:"confirm", title, message, onOk, danger, icon, okLabel: danger ? "Ya, Hapus" : "Konfirmasi" });
  const showStatus = (s) => { setSaveStatus(s); setTimeout(() => setSaveStatus(""), 2500); };

  const loadData = useCallback(async () => {
    try {
      const [txs, customs, overrides, deleted] = await Promise.all([
        api.get("transactions", "select=*&order=id.desc"),
        api.get("custom_items", "select=*&order=kode.asc"),
        api.get("item_overrides", "select=*"),
        api.get("deleted_base", "select=*"),
      ]);
      setTransactions(txs);
      setCustomItems(customs.map(c => ({ ...c, stokAwal: c.stok_awal, isCustom: true })));
      const ovMap = {};
      for (const o of overrides) ovMap[o.kode] = { nama: o.nama, satuan: o.satuan, stokAwal: o.stok_awal };
      setItemOverrides(ovMap);
      setDeletedBase(deleted.map(d => d.kode));
      setLastUpdated(new Date().toISOString());
    } catch (e) { showAlert("Gagal memuat data", e.message, "❌"); }
    setLoading(false);
  }, []);

  useEffect(() => { loadData(); }, [loadData]);
  useEffect(() => { const i = setInterval(loadData, 10000); return () => clearInterval(i); }, [loadData]);

  const allItems = [
    ...INITIAL_ITEMS.filter(i => !deletedBase.includes(i.kode)).map(i => { const ov = itemOverrides[i.kode]; return ov ? { ...i, ...ov } : i; }),
    ...customItems,
  ];

  const getBaseStok = (kode) => { const ov = itemOverrides[kode]; return (ov && ov.stokAwal !== undefined) ? ov.stokAwal : (INITIAL_STOCK[kode] || 0); };

  const computeStock = () => {
    const stock = {};
    for (const item of allItems) stock[item.kode] = item.isCustom ? (item.stokAwal || 0) : getBaseStok(item.kode);
    for (const tx of transactions) {
      for (const item of tx.items) {
        const k = parseInt(item.kode);
        if (stock[k] === undefined) stock[k] = 0;
        if (tx.type === "masuk") stock[k] += parseInt(item.jumlah) || 0;
        else stock[k] -= parseInt(item.jumlah) || 0;
      }
    }
    return stock;
  };

  const currentStock = computeStock();
  const totalMasuk = transactions.filter(t => t.type === "masuk").reduce((s,t) => s + t.items.reduce((ss,i) => ss + (parseInt(i.jumlah)||0), 0), 0);
  const totalKeluar = transactions.filter(t => t.type === "keluar").reduce((s,t) => s + t.items.reduce((ss,i) => ss + (parseInt(i.jumlah)||0), 0), 0);
  const totalInitial = allItems.reduce((s,i) => s + (i.isCustom ? (i.stokAwal||0) : getBaseStok(i.kode)), 0);
  const totalAvailable = Object.values(currentStock).reduce((a,b) => a + Math.max(0,b), 0);

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    const stokData = [
      ["STOCK OPNAME BARANG CETAKAN - BANK ACEH SYARIAH CABANG JAKARTA"],
      [`Diekspor: ${formatDateTime(new Date().toISOString())}`],
      [],
      ["No","Kode","Nama Barang","Satuan","Stok Awal","Masuk","Keluar","Stok Tersedia","Status"],
    ];
    allItems.forEach((item, i) => {
      const stokAwal = item.isCustom ? (item.stokAwal||0) : getBaseStok(item.kode);
      const masukQty = transactions.filter(t => t.type==="masuk").reduce((s,t) => s + t.items.filter(it => parseInt(it.kode)===item.kode).reduce((ss,it) => ss+(parseInt(it.jumlah)||0),0),0);
      const keluarQty = transactions.filter(t => t.type==="keluar").reduce((s,t) => s + t.items.filter(it => parseInt(it.kode)===item.kode).reduce((ss,it) => ss+(parseInt(it.jumlah)||0),0),0);
      const sisa = stokAwal + masukQty - keluarQty;
      stokData.push([i+1, item.kode, item.nama, item.satuan, stokAwal, masukQty, keluarQty, sisa, sisa<=0?"Habis":sisa<=5?"Menipis":"Tersedia"]);
    });
    stokData.push([], ["","","","TOTAL", totalInitial, totalMasuk, totalKeluar, totalAvailable,""]);
    const ws1 = XLSX.utils.aoa_to_sheet(stokData);
    ws1["!cols"] = [{wch:4},{wch:6},{wch:42},{wch:8},{wch:10},{wch:8},{wch:8},{wch:14},{wch:10}];
    XLSX.utils.book_append_sheet(wb, ws1, "Ringkasan Stok");

    const txData = [
      ["RIWAYAT TRANSAKSI CETAKAN - BANK ACEH SYARIAH CABANG JAKARTA"],
      [`Diekspor: ${formatDateTime(new Date().toISOString())}`],
      [],
      ["Tanggal","Tipe","Nama Barang","Satuan","Jumlah","Pemohon Bidang","Keterangan","Dicatat"],
    ];
    [...transactions].sort((a,b) => new Date(a.tanggal)-new Date(b.tanggal)).forEach(tx => {
      tx.items.forEach(item => {
        txData.push([formatDate(tx.tanggal), tx.type==="masuk"?"MASUK":"KELUAR", getItemName(item.kode), getItemSatuan(item.kode), parseInt(item.jumlah)||0, tx.penerima||"-", tx.keterangan||"-", formatDateTime(tx.created_at)]);
      });
    });
    const ws2 = XLSX.utils.aoa_to_sheet(txData);
    ws2["!cols"] = [{wch:14},{wch:8},{wch:42},{wch:8},{wch:8},{wch:18},{wch:25},{wch:20}];
    XLSX.utils.book_append_sheet(wb, ws2, "Riwayat Transaksi");
    XLSX.writeFile(wb, `Cetakan_BankAcehSyariah_${new Date().toISOString().split("T")[0]}.xlsx`);
  };

  const getItemName = (kode) => allItems.find(i => i.kode===parseInt(kode))?.nama || `Kode ${kode}`;
  const getItemSatuan = (kode) => allItems.find(i => i.kode===parseInt(kode))?.satuan || "";

  const addMasuk = async () => {
    const valid = masukItems.filter(i => i.kode && i.jumlah && parseInt(i.jumlah)>0);
    if (!masukDate || valid.length===0) { showAlert("Tidak lengkap","Lengkapi tanggal dan minimal 1 item.","⚠️"); return; }
    showStatus("saving");
    try {
      await api.upsert("transactions", { id:Date.now(), type:"masuk", tanggal:masukDate, keterangan:masukKeterangan, penerima:null, items:valid, created_at:new Date().toISOString() });
      showStatus("saved"); setMasukItems([{kode:"",jumlah:""}]); setMasukKeterangan(""); setMasukDate(today()); setTab("history"); loadData();
    } catch(e) { showStatus("error"); showAlert("Gagal simpan", e.message, "❌"); }
  };

  const addKeluar = async () => {
    const valid = keluarItems.filter(i => i.kode && i.jumlah && parseInt(i.jumlah)>0);
    if (!keluarDate || valid.length===0) { showAlert("Tidak lengkap","Lengkapi tanggal dan minimal 1 item.","⚠️"); return; }
    showStatus("saving");
    try {
      await api.upsert("transactions", { id:Date.now(), type:"keluar", tanggal:keluarDate, keterangan:keluarKeterangan, penerima:keluarPenerima, items:valid, created_at:new Date().toISOString() });
      showStatus("saved"); setKeluarItems([{kode:"",jumlah:""}]); setKeluarKeterangan(""); setKeluarPenerima(""); setKeluarDate(today()); setTab("history"); loadData();
    } catch(e) { showStatus("error"); showAlert("Gagal simpan", e.message, "❌"); }
  };

  const deleteTransaction = (id) => {
    showConfirm("Hapus transaksi?","Transaksi ini akan dihapus permanen.", async () => {
      showStatus("saving");
      try { await api.delete("transactions", `id=eq.${id}`); showStatus("saved"); loadData(); }
      catch(e) { showStatus("error"); showAlert("Gagal hapus", e.message, "❌"); }
    }, true, "🗑️");
  };

  const addBarangBaru = async () => {
    const nama = newNama.trim().toUpperCase();
    const satuan = newSatuan==="__custom__" ? newSatuanCustom.trim() : newSatuan;
    const stokAwal = parseInt(newStokAwal) || 0;
    if (!nama) { showAlert("Nama wajib diisi","","⚠️"); return; }
    if (!satuan) { showAlert("Satuan wajib diisi","","⚠️"); return; }
    if (allItems.some(i => i.nama===nama)) { showAlert("Duplikat","Nama barang sudah ada.","⚠️"); return; }
    const maxKode = Math.max(...allItems.map(i => i.kode), 100);
    showStatus("saving");
    try {
      await api.upsert("custom_items", { kode:maxKode+1, nama, satuan, stok_awal:stokAwal, added_at:new Date().toISOString() });
      showStatus("saved"); setNewNama(""); setNewSatuan("buah"); setNewSatuanCustom(""); setNewStokAwal("0"); loadData();
    } catch(e) { showStatus("error"); showAlert("Gagal simpan", e.message, "❌"); }
  };

  const startEdit = (item) => { setEditingKode(item.kode); setEditForm({ nama:item.nama, satuan:item.satuan, stokAwal: item.isCustom?(item.stokAwal||0):getBaseStok(item.kode) }); };

  const saveEdit = async () => {
    const kode = editingKode;
    const item = allItems.find(i => i.kode===kode);
    if (!editForm.nama.trim()) { showAlert("Nama tidak boleh kosong","","⚠️"); return; }
    showStatus("saving");
    try {
      if (item.isCustom) await api.upsert("custom_items", { kode, nama:editForm.nama.trim().toUpperCase(), satuan:editForm.satuan, stok_awal:parseInt(editForm.stokAwal)||0 });
      else await api.upsert("item_overrides", { kode, nama:editForm.nama.trim().toUpperCase(), satuan:editForm.satuan, stok_awal:parseInt(editForm.stokAwal)||0 });
      showStatus("saved"); setEditingKode(null); loadData();
    } catch(e) { showStatus("error"); showAlert("Gagal simpan", e.message, "❌"); }
  };

  const deleteItem = (kode) => {
    const item = allItems.find(i => i.kode===kode);
    showConfirm(`Hapus "${item?.nama}"?`,"Barang ini akan dihapus dari daftar dan semua hitungan stoknya dihilangkan.", async () => {
      showStatus("saving");
      try {
        const affected = transactions.filter(tx => tx.items.some(i => parseInt(i.kode)===kode));
        for (const tx of affected) await api.upsert("transactions", { ...tx, items: tx.items.filter(i => parseInt(i.kode)!==kode) });
        if (item.isCustom) await api.delete("custom_items", `kode=eq.${kode}`);
        else { await api.upsert("deleted_base", { kode }); try { await api.delete("item_overrides", `kode=eq.${kode}`); } catch {} }
        showStatus("saved"); loadData();
      } catch(e) { showStatus("error"); showAlert("Gagal hapus", e.message, "❌"); }
    }, true, "🗑️");
  };

  const addItemRow = (type) => { if (type==="masuk") setMasukItems([...masukItems,{kode:"",jumlah:""}]); else setKeluarItems([...keluarItems,{kode:"",jumlah:""}]); };
  const removeItemRow = (type, idx) => {
    if (type==="masuk") { const r=masukItems.filter((_,i)=>i!==idx); setMasukItems(r.length?r:[{kode:"",jumlah:""}]); }
    else { const r=keluarItems.filter((_,i)=>i!==idx); setKeluarItems(r.length?r:[{kode:"",jumlah:""}]); }
  };
  const updateItemRow = (type, idx, field, val) => {
    if (type==="masuk") { const r=[...masukItems]; r[idx]={...r[idx],[field]:val}; setMasukItems(r); }
    else { const r=[...keluarItems]; r[idx]={...r[idx],[field]:val}; setKeluarItems(r); }
  };
  const toggleExpand = (id) => setExpandedTx(p => ({ ...p, [id]: !p[id] }));
  const handleSort = (field) => { if (sortField===field) setSortDir(d=>d==="asc"?"desc":"asc"); else { setSortField(field); setSortDir("asc"); } };
  const SortIcon = ({ field }) => sortField!==field ? <span style={{color:"#ccc",marginLeft:3}}>↕</span> : <span style={{marginLeft:3}}>{sortDir==="asc"?"↑":"↓"}</span>;

  const filteredStock = allItems
    .filter(i => i.nama.toLowerCase().includes(searchOverview.toLowerCase()) || String(i.kode).includes(searchOverview))
    .map(item => {
      const stokAwal = item.isCustom ? (item.stokAwal||0) : getBaseStok(item.kode);
      const masukQty = transactions.filter(t=>t.type==="masuk").reduce((s,t)=>s+t.items.filter(it=>parseInt(it.kode)===item.kode).reduce((ss,it)=>ss+(parseInt(it.jumlah)||0),0),0);
      const keluarQty = transactions.filter(t=>t.type==="keluar").reduce((s,t)=>s+t.items.filter(it=>parseInt(it.kode)===item.kode).reduce((ss,it)=>ss+(parseInt(it.jumlah)||0),0),0);
      return { ...item, stokAwal, masukQty, keluarQty, sisa: stokAwal+masukQty-keluarQty };
    })
    .sort((a,b) => {
      if (sortField==="nama") return sortDir==="asc" ? a.nama.localeCompare(b.nama) : b.nama.localeCompare(a.nama);
      const f = {sisa:"sisa",stokAwal:"stokAwal",masuk:"masukQty",keluar:"keluarQty",kode:"kode"}[sortField]||"kode";
      return sortDir==="asc" ? a[f]-b[f] : b[f]-a[f];
    });

  const filteredKelola = allItems.filter(i => i.nama.toLowerCase().includes(searchKelola.toLowerCase()) || String(i.kode).includes(searchKelola));
  const filteredTx = [...transactions].sort((a,b)=>b.id-a.id).filter(tx => {
    const q = searchHistory.toLowerCase();
    if (!q) return true;
    return formatDate(tx.tanggal).toLowerCase().includes(q) || tx.type.includes(q) ||
      tx.items.map(i=>getItemName(i.kode).toLowerCase()).join(" ").includes(q) ||
      (tx.keterangan||"").toLowerCase().includes(q) || (tx.penerima||"").toLowerCase().includes(q);
  });

  const inp = { border:"1px solid #e0e0e0", borderRadius:8, padding:"10px 12px", fontSize:14, background:"#fff", color:"#111", width:"100%", boxSizing:"border-box" };
  const thS = { padding:"10px 12px", fontWeight:600, color:"#666", fontSize:11, textTransform:"uppercase", letterSpacing:"0.04em", borderBottom:"1px solid #eee", cursor:"pointer", userSelect:"none", whiteSpace:"nowrap", background:"#f5faf8" };
  const btnPrimary = { background:THEME.primary, color:"#fff", border:"none", borderRadius:8, padding:"11px 22px", fontSize:14, fontWeight:600, cursor:"pointer", width:"100%" };
  const btnDanger = { background:"#A32D2D", color:"#fff", border:"none", borderRadius:8, padding:"11px 22px", fontSize:14, fontWeight:600, cursor:"pointer", width:"100%" };
  const tabs = [
    { id:"overview", label:"Ringkasan", icon:"🖨️" },
    { id:"masuk",    label:"Masuk",     icon:"📥" },
    { id:"keluar",   label:"Keluar",    icon:"📤" },
    { id:"history",  label:"Riwayat",   icon:"📋" },
    { id:"kelola",   label:"Kelola",    icon:"⚙️" },
  ];

  const ItemFormRow = ({ items, type }) => (
    <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
      {items.map((item, idx) => {
        const sel = allItems.find(i => i.kode===parseInt(item.kode));
        const stAvail = item.kode ? (currentStock[parseInt(item.kode)]||0) : null;
        return (
          <div key={idx} style={{ display:"flex", flexDirection:"column", gap:6, background:"#f5faf8", borderRadius:8, padding:"10px 12px", border:"1px solid #d1ece3" }}>
            <div style={{ display:"flex", gap:8, alignItems:"center" }}>
              <select value={item.kode} onChange={e=>updateItemRow(type,idx,"kode",e.target.value)} style={{ ...inp, flex:1 }}>
                <option value="">— Pilih barang cetakan —</option>
                {allItems.map(i => <option key={i.kode} value={i.kode}>{i.kode}. {i.nama}</option>)}
              </select>
              <button onClick={()=>removeItemRow(type,idx)} style={{ background:"none", border:"none", cursor:"pointer", color:"#E24B4A", fontSize:22, padding:0, flexShrink:0, lineHeight:1 }}>×</button>
            </div>
            <div style={{ display:"flex", gap:8, alignItems:"center" }}>
              <input type="number" min="1" placeholder="Jumlah" value={item.jumlah} onChange={e=>updateItemRow(type,idx,"jumlah",e.target.value)} style={{ ...inp, width:100, flexShrink:0 }}/>
              <span style={{ fontSize:13, color:"#888" }}>{sel?.satuan||""}</span>
              {type==="keluar" && stAvail!==null && <Badge color={stAvail>0?"teal":"red"}>stok: {stAvail}</Badge>}
            </div>
          </div>
        );
      })}
      <button onClick={()=>addItemRow(type)} style={{ display:"flex", alignItems:"center", justifyContent:"center", gap:6, background:"none", border:`1px dashed ${THEME.primary}`, borderRadius:8, padding:"10px 12px", cursor:"pointer", color:THEME.primary, fontSize:13 }}>
        ＋ Tambah item
      </button>
    </div>
  );

  if (loading) return (
    <div style={{ minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center", background:"#f0f7f4" }}>
      <div style={{ textAlign:"center", color:"#888" }}>
        <div style={{ fontSize:40, marginBottom:12 }}>⏳</div>
        <div style={{ fontSize:14 }}>Memuat data cetakan...</div>
      </div>
    </div>
  );

  return (
    <div style={{ minHeight:"100vh", background:"#edf5f1", fontFamily:"-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif" }}>
      <Modal modal={modal} onClose={()=>setModal(null)}/>

      {/* TOP NAV — EMERALD */}
      <div style={{ background:`linear-gradient(135deg, ${THEME.primaryDark} 0%, ${THEME.primary} 100%)`, padding:"0 16px", position:"sticky", top:0, zIndex:100 }}>
        <div style={{ maxWidth:960, margin:"0 auto", display:"flex", alignItems:"center", justifyContent:"space-between", height:52 }}>
          <div>
            <div style={{ fontSize:10, color:"rgba(255,255,255,0.6)", letterSpacing:"0.06em", textTransform:"uppercase" }}>Bank Aceh Syariah · Cabang Jakarta</div>
            <div style={{ fontSize:15, fontWeight:700, color:"#fff" }}>🖨️ Stock Opname Cetakan</div>
          </div>
          <div style={{ display:"flex", alignItems:"center", gap:8 }}>
            {saveStatus==="saving" && <Badge color="amber">🔄</Badge>}
            {saveStatus==="saved"  && <Badge color="green">✓</Badge>}
            {saveStatus==="error"  && <Badge color="red">✗</Badge>}
            <button onClick={loadData} style={{ background:"rgba(255,255,255,0.15)", border:"none", borderRadius:8, padding:"6px 10px", cursor:"pointer", fontSize:16, color:"#fff" }}>🔄</button>
            <button onClick={exportExcel} style={{ background:"rgba(255,255,255,0.2)", border:"1px solid rgba(255,255,255,0.4)", borderRadius:8, padding:"6px 12px", cursor:"pointer", fontSize:12, fontWeight:600, color:"#fff", display:"flex", alignItems:"center", gap:4 }}>
              ⬇ Excel
            </button>
          </div>
        </div>
      </div>

      {lastUpdated && (
        <div style={{ background:THEME.primaryDark }}>
          <div style={{ maxWidth:960, margin:"0 auto", padding:"4px 16px", fontSize:11, color:"rgba(255,255,255,0.5)", textAlign:"right" }}>
            🕐 Terakhir diperbarui: {formatDateTime(lastUpdated)}
          </div>
        </div>
      )}

      <div style={{ maxWidth:960, margin:"0 auto", padding:"16px 12px 80px" }}>

        {/* STOCK ALERT PANELS */}
        {(() => {
          const habis = filteredStock.filter(i => i.sisa <= 0);
          const menipis = filteredStock.filter(i => i.sisa > 0 && i.sisa <= 5);
          if (habis.length === 0 && menipis.length === 0) return null;
          return (
            <div style={{ display: "grid", gridTemplateColumns: habis.length > 0 && menipis.length > 0 ? "1fr 1fr" : "1fr", gap: 10, marginBottom: 12 }}>
              {habis.length > 0 && (
                <div style={{ background: "#fff", borderRadius: 12, border: "1.5px solid #E24B4A", overflow: "hidden" }}>
                  <div style={{ background: "linear-gradient(135deg, #A32D2D, #E24B4A)", padding: "10px 14px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <span style={{ fontSize: 18 }}>🚨</span>
                      <span style={{ fontSize: 12, fontWeight: 700, color: "#fff", textTransform: "uppercase", letterSpacing: "0.05em" }}>Stok Habis</span>
                    </div>
                    <div style={{ background: "rgba(255,255,255,0.25)", borderRadius: 20, padding: "2px 10px", fontSize: 18, fontWeight: 800, color: "#fff" }}>{habis.length}</div>
                  </div>
                  <div style={{ maxHeight: 220, overflowY: "auto" }}>
                    {habis.map((item, i) => (
                      <div key={item.kode} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 14px", borderBottom: "1px solid #FEF2F2", background: i % 2 === 0 ? "#fff" : "#FFF5F5" }}>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ fontSize: 12, fontWeight: 600, color: "#111", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{item.nama}</div>
                          <div style={{ fontSize: 10, color: "#aaa" }}>{item.satuan}</div>
                        </div>
                        <div style={{ background: "#FCEBEB", color: "#A32D2D", fontSize: 11, fontWeight: 700, padding: "2px 8px", borderRadius: 6, flexShrink: 0, marginLeft: 8 }}>0</div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
              {menipis.length > 0 && (
                <div style={{ background: "#fff", borderRadius: 12, border: "1.5px solid #F59E0B", overflow: "hidden" }}>
                  <div style={{ background: "linear-gradient(135deg, #92400E, #D97706)", padding: "10px 14px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <span style={{ fontSize: 18 }}>⚠️</span>
                      <span style={{ fontSize: 12, fontWeight: 700, color: "#fff", textTransform: "uppercase", letterSpacing: "0.05em" }}>Stok Menipis</span>
                    </div>
                    <div style={{ background: "rgba(255,255,255,0.25)", borderRadius: 20, padding: "2px 10px", fontSize: 18, fontWeight: 800, color: "#fff" }}>{menipis.length}</div>
                  </div>
                  <div style={{ maxHeight: 220, overflowY: "auto" }}>
                    {menipis.map((item, i) => (
                      <div key={item.kode} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 14px", borderBottom: "1px solid #FFFBEB", background: i % 2 === 0 ? "#fff" : "#FFFDF5" }}>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ fontSize: 12, fontWeight: 600, color: "#111", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{item.nama}</div>
                          <div style={{ fontSize: 10, color: "#aaa" }}>{item.satuan}</div>
                        </div>
                        <div style={{ background: "#FAEEDA", color: "#854F0B", fontSize: 11, fontWeight: 700, padding: "2px 8px", borderRadius: 6, flexShrink: 0, marginLeft: 8 }}>{item.sisa}</div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          );
        })()}

        {/* METRICS */}
        <div style={{ display:"grid", gridTemplateColumns:"repeat(2,1fr)", gap:8, marginBottom:12 }}>
          <MetricCard label="Stok Awal"    value={totalInitial.toLocaleString("id-ID")}   icon="🗂️" color="teal"  sub="opname terakhir"/>
          <MetricCard label="Tersedia"     value={totalAvailable.toLocaleString("id-ID")} icon="✅" color="green" sub="saat ini"/>
          <MetricCard label="Total Masuk"  value={totalMasuk.toLocaleString("id-ID")}     icon="📥" color="teal"  sub={`${transactions.filter(t=>t.type==="masuk").length} transaksi`}/>
          <MetricCard label="Total Keluar" value={totalKeluar.toLocaleString("id-ID")}    icon="📤" color="amber" sub={`${transactions.filter(t=>t.type==="keluar").length} transaksi`}/>
        </div>

        {/* MAIN CARD */}
        <div style={{ background:"#fff", borderRadius:14, border:`1px solid #c8e6d8`, overflow:"hidden" }}>

          {/* TAB BAR */}
          <div style={{ display:"flex", borderBottom:`1px solid #c8e6d8`, overflowX:"auto" }}>
            {tabs.map(t => (
              <button key={t.id} onClick={()=>setTab(t.id)}
                style={{ flex:1, background: tab===t.id ? "#f0faf5" : "none", border:"none", borderBottom: tab===t.id ? `2px solid ${THEME.primary}` : "2px solid transparent", padding:"12px 8px", cursor:"pointer", fontSize:12, fontWeight: tab===t.id ? 700 : 400, color: tab===t.id ? THEME.primary : "#999", whiteSpace:"nowrap", minWidth:60 }}>
                <div style={{ fontSize:18 }}>{t.icon}</div>
                <div>{t.label}</div>
              </button>
            ))}
          </div>

          <div style={{ padding:"16px" }}>

            {/* OVERVIEW */}
            {tab==="overview" && (
              <div>
                <div style={{ display:"flex", gap:8, marginBottom:12, alignItems:"center", flexWrap:"wrap" }}>
                  <input type="text" placeholder="🔍 Cari barang cetakan..." value={searchOverview} onChange={e=>setSearchOverview(e.target.value)} style={{ ...inp, flex:1, minWidth:160 }}/>
                  <div style={{ fontSize:11, color:"#aaa", whiteSpace:"nowrap" }}>{filteredStock.length} item</div>
                </div>
                <div style={{ overflowX:"auto" }}>
                  <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                    <thead>
                      <tr>
                        <th onClick={()=>handleSort("kode")}    style={thS}># <SortIcon field="kode"/></th>
                        <th onClick={()=>handleSort("nama")}    style={{...thS,textAlign:"left"}}>Nama <SortIcon field="nama"/></th>
                        <th onClick={()=>handleSort("stokAwal")} style={{...thS,textAlign:"center"}}>Awal <SortIcon field="stokAwal"/></th>
                        <th onClick={()=>handleSort("masuk")}   style={{...thS,textAlign:"center"}}>Masuk <SortIcon field="masuk"/></th>
                        <th onClick={()=>handleSort("keluar")}  style={{...thS,textAlign:"center"}}>Keluar <SortIcon field="keluar"/></th>
                        <th onClick={()=>handleSort("sisa")}    style={{...thS,textAlign:"center"}}>Sisa <SortIcon field="sisa"/></th>
                        <th style={{...thS,textAlign:"center"}}>Status</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredStock.map((item,i) => {
                        const status = item.sisa<=0?"Habis":item.sisa<=5?"Menipis":"Tersedia";
                        const sc = item.sisa<=0?"red":item.sisa<=5?"amber":"teal";
                        return (
                          <tr key={item.kode} style={{ borderBottom:"1px solid #f0f9f4", background: i%2===0?"#fff":"#f8fdf9" }}>
                            <td style={{ padding:"8px 10px", textAlign:"center", color:"#bbb", fontSize:11 }}>{item.kode}</td>
                            <td style={{ padding:"8px 10px", color:"#111", fontWeight:500 }}>
                              {item.nama}
                              {item.isCustom && <span style={{ marginLeft:5 }}><Badge color="purple">baru</Badge></span>}
                              <div style={{ fontSize:10, color:"#bbb", fontWeight:400 }}>{item.satuan}</div>
                            </td>
                            <td style={{ padding:"8px 10px", textAlign:"center", color:"#666" }}>{item.stokAwal}</td>
                            <td style={{ padding:"8px 10px", textAlign:"center", color:THEME.primary }}>{item.masukQty>0?`+${item.masukQty}`:"-"}</td>
                            <td style={{ padding:"8px 10px", textAlign:"center", color: item.keluarQty>0?"#A32D2D":"#ccc" }}>{item.keluarQty>0?`-${item.keluarQty}`:"-"}</td>
                            <td style={{ padding:"8px 10px", textAlign:"center", fontWeight:700, color: item.sisa<=0?"#A32D2D":item.sisa<=5?"#854F0B":"#111" }}>{item.sisa}</td>
                            <td style={{ padding:"8px 10px", textAlign:"center" }}><Badge color={sc}>{status}</Badge></td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* MASUK */}
            {tab==="masuk" && (
              <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
                <h2 style={{ margin:0, fontSize:16, fontWeight:700, color:THEME.primary }}>📥 Input Stok Masuk</h2>
                <div><label style={{ fontSize:13, color:"#666", display:"block", marginBottom:6 }}>Tanggal Penerimaan *</label><input type="date" value={masukDate} onChange={e=>setMasukDate(e.target.value)} style={inp}/></div>
                <div><label style={{ fontSize:13, color:"#666", display:"block", marginBottom:6 }}>Keterangan / No. Nota</label><input type="text" placeholder="contoh: Pengiriman dari kantor pusat" value={masukKeterangan} onChange={e=>setMasukKeterangan(e.target.value)} style={inp}/></div>
                <div><label style={{ fontSize:13, color:"#666", display:"block", marginBottom:8 }}>Daftar Barang *</label><ItemFormRow items={masukItems} type="masuk"/></div>
                <button onClick={addMasuk} style={btnPrimary}>💾 Simpan Stok Masuk</button>
              </div>
            )}

            {/* KELUAR */}
            {tab==="keluar" && (
              <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
                <h2 style={{ margin:0, fontSize:16, fontWeight:700, color:"#A32D2D" }}>📤 Input Stok Keluar</h2>
                <div><label style={{ fontSize:13, color:"#666", display:"block", marginBottom:6 }}>Tanggal Pengeluaran *</label><input type="date" value={keluarDate} onChange={e=>setKeluarDate(e.target.value)} style={inp}/></div>
                <div>
                  <label style={{ fontSize:13, color:"#666", display:"block", marginBottom:6 }}>Pemohon dari Bidang</label>
                  <select value={keluarPenerima} onChange={e=>setKeluarPenerima(e.target.value)} style={inp}>
                    <option value="">— Pilih bidang —</option>
                    {BIDANG_OPTIONS.map(b => <option key={b} value={b}>{b}</option>)}
                  </select>
                </div>
                <div><label style={{ fontSize:13, color:"#666", display:"block", marginBottom:6 }}>Keterangan</label><input type="text" placeholder="contoh: Permintaan rutin bulanan" value={keluarKeterangan} onChange={e=>setKeluarKeterangan(e.target.value)} style={inp}/></div>
                <div>
                  <label style={{ fontSize:13, color:"#666", display:"block", marginBottom:8 }}>Daftar Barang *</label>
                  <div style={{ background:"#f5faf8", borderRadius:8, padding:"10px 12px", marginBottom:8, fontSize:12, color:"#888" }}>ℹ️ Stok tersedia ditampilkan per item</div>
                  <ItemFormRow items={keluarItems} type="keluar"/>
                </div>
                <button onClick={addKeluar} style={btnDanger}>💾 Simpan Stok Keluar</button>
              </div>
            )}

            {/* HISTORY */}
            {tab==="history" && (
              <div>
                <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:12, flexWrap:"wrap" }}>
                  <input type="text" placeholder="🔍 Cari transaksi..." value={searchHistory} onChange={e=>setSearchHistory(e.target.value)} style={{ ...inp, flex:1, minWidth:160 }}/>
                  <span style={{ fontSize:12, color:"#aaa", whiteSpace:"nowrap" }}>{filteredTx.length} transaksi</span>
                </div>
                {filteredTx.length===0 ? (
                  <div style={{ textAlign:"center", padding:"40px 0", color:"#aaa" }}><div style={{ fontSize:32 }}>📋</div><div style={{ marginTop:8 }}>Belum ada transaksi</div></div>
                ) : (
                  <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
                    {filteredTx.map(tx => {
                      const isExp = expandedTx[tx.id];
                      const totalQty = tx.items.reduce((s,i) => s+(parseInt(i.jumlah)||0), 0);
                      return (
                        <div key={tx.id} style={{ border:"1px solid #eee", borderLeft:`4px solid ${tx.type==="masuk"?THEME.primary:"#E24B4A"}`, borderRadius:10, overflow:"hidden", background:"#fff" }}>
                          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", gap:8, padding:"12px", cursor:"pointer" }} onClick={()=>toggleExpand(tx.id)}>
                            <div style={{ flex:1, minWidth:0 }}>
                              <div style={{ display:"flex", alignItems:"center", gap:6, flexWrap:"wrap", marginBottom:4 }}>
                                <Badge color={tx.type==="masuk"?"teal":"red"}>{tx.type==="masuk"?"📥 MASUK":"📤 KELUAR"}</Badge>
                                <span style={{ fontSize:13, fontWeight:700, color:"#111" }}>{formatDate(tx.tanggal)}</span>
                              </div>
                              <div style={{ display:"flex", flexWrap:"wrap", gap:6 }}>
                                {tx.penerima && <Badge color="teal">🏢 {tx.penerima}</Badge>}
                                <span style={{ fontSize:11, color:"#aaa" }}>{tx.items.length} jenis · {totalQty} item</span>
                              </div>
                              {tx.keterangan && <div style={{ fontSize:11, color:"#aaa", marginTop:2 }}>— {tx.keterangan}</div>}
                            </div>
                            <div style={{ display:"flex", flexDirection:"column", alignItems:"flex-end", gap:4, flexShrink:0 }}>
                              <button onClick={e=>{ e.stopPropagation(); deleteTransaction(tx.id); }} style={{ background:"none", border:"none", cursor:"pointer", color:"#ddd", fontSize:16, padding:0 }}>🗑️</button>
                              <span style={{ fontSize:11, color:"#bbb" }}>{isExp?"▲":"▼"}</span>
                            </div>
                          </div>
                          {isExp && (
                            <div style={{ borderTop:"1px solid #f0f0f0", padding:"10px 12px", background:"#f8fdf9" }}>
                              <div style={{ display:"flex", flexWrap:"wrap", gap:12, marginBottom:10, fontSize:12 }}>
                                {tx.created_at && <span style={{ color:"#888" }}>📅 {formatDateTime(tx.created_at)}</span>}
                                {tx.penerima && <span style={{ color:THEME.primary, fontWeight:600 }}>🏢 {tx.penerima}</span>}
                                {tx.keterangan && <span style={{ color:"#666" }}>📝 {tx.keterangan}</span>}
                              </div>
                              {tx.items.map((item,i) => (
                                <div key={i} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"6px 0", borderBottom:"1px solid #f0f0f0" }}>
                                  <span style={{ fontSize:13, color:"#111" }}>{getItemName(item.kode)}</span>
                                  <span style={{ fontSize:13, fontWeight:700, color: tx.type==="masuk"?THEME.primary:"#A32D2D" }}>
                                    {tx.type==="masuk"?"+":"-"}{item.jumlah} {getItemSatuan(item.kode)}
                                  </span>
                                </div>
                              ))}
                              <div style={{ display:"flex", justifyContent:"space-between", padding:"6px 0", marginTop:2 }}>
                                <span style={{ fontSize:12, fontWeight:700, color:"#888" }}>TOTAL</span>
                                <span style={{ fontSize:13, fontWeight:700, color: tx.type==="masuk"?THEME.primary:"#A32D2D" }}>{tx.type==="masuk"?"+":"-"}{totalQty} item</span>
                              </div>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            )}

            {/* KELOLA */}
            {tab==="kelola" && (
              <div style={{ display:"flex", flexDirection:"column", gap:20 }}>
                <div style={{ background:"#f5faf8", borderRadius:10, padding:"16px", border:`1px solid #c8e6d8` }}>
                  <h2 style={{ margin:"0 0 4px", fontSize:15, fontWeight:700, color:THEME.primary }}>➕ Tambah Jenis Barang Baru</h2>
                  <p style={{ margin:"0 0 12px", fontSize:12, color:"#888" }}>Langsung tersedia untuk semua pengguna.</p>
                  <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
                    <div><label style={{ fontSize:13, color:"#666", display:"block", marginBottom:6 }}>Nama Barang *</label><input type="text" placeholder="contoh: BUKU TABUNGAN BARU" value={newNama} onChange={e=>setNewNama(e.target.value)} style={inp}/></div>
                    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
                      <div>
                        <label style={{ fontSize:13, color:"#666", display:"block", marginBottom:6 }}>Satuan *</label>
                        <select value={newSatuan} onChange={e=>setNewSatuan(e.target.value)} style={inp}>
                          {SATUAN_OPTIONS.map(s => <option key={s} value={s}>{s}</option>)}
                          <option value="__custom__">Lainnya...</option>
                        </select>
                      </div>
                      <div><label style={{ fontSize:13, color:"#666", display:"block", marginBottom:6 }}>Stok Awal</label><input type="number" min="0" value={newStokAwal} onChange={e=>setNewStokAwal(e.target.value)} style={inp}/></div>
                    </div>
                    {newSatuan==="__custom__" && <input type="text" placeholder="Ketik satuan baru..." value={newSatuanCustom} onChange={e=>setNewSatuanCustom(e.target.value)} style={inp}/>}
                    <button onClick={addBarangBaru} style={btnPrimary}>➕ Tambah Barang</button>
                  </div>
                </div>

                <div>
                  <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", gap:8, marginBottom:10, flexWrap:"wrap" }}>
                    <h2 style={{ margin:0, fontSize:15, fontWeight:700 }}>📋 Semua Barang ({allItems.length})</h2>
                    <input type="text" placeholder="🔍 Cari..." value={searchKelola} onChange={e=>setSearchKelola(e.target.value)} style={{ ...inp, width:180 }}/>
                  </div>
                  <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
                    {filteredKelola.map((item,i) => {
                      const stokAwal = item.isCustom ? (item.stokAwal||0) : getBaseStok(item.kode);
                      const isEditing = editingKode===item.kode;
                      return (
                        <div key={item.kode} style={{ border:"1px solid #eee", borderRadius:10, padding:"12px 14px", background: isEditing?"#f0fdf4": i%2===0?"#fff":"#f8fdf9" }}>
                          {isEditing ? (
                            <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
                              <input value={editForm.nama} onChange={e=>setEditForm(f=>({...f,nama:e.target.value}))} style={inp} placeholder="Nama barang"/>
                              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8 }}>
                                <select value={editForm.satuan} onChange={e=>setEditForm(f=>({...f,satuan:e.target.value}))} style={inp}>{SATUAN_OPTIONS.map(s=><option key={s} value={s}>{s}</option>)}</select>
                                <input type="number" min="0" value={editForm.stokAwal} onChange={e=>setEditForm(f=>({...f,stokAwal:e.target.value}))} style={inp} placeholder="Stok awal"/>
                              </div>
                              <div style={{ display:"flex", gap:8 }}>
                                <button onClick={saveEdit} style={{ ...btnPrimary, padding:"8px 16px", width:"auto", flex:1 }}>✓ Simpan</button>
                                <button onClick={()=>setEditingKode(null)} style={{ background:"none", border:"1px solid #ddd", borderRadius:8, padding:"8px 16px", cursor:"pointer", fontSize:14, color:"#666", flex:1 }}>Batal</button>
                              </div>
                            </div>
                          ) : (
                            <div style={{ display:"flex", alignItems:"flex-start", justifyContent:"space-between", gap:8 }}>
                              <div style={{ flex:1, minWidth:0 }}>
                                <div style={{ display:"flex", alignItems:"center", gap:6, flexWrap:"wrap", marginBottom:3 }}>
                                  <span style={{ fontWeight:600, fontSize:13, color:"#111" }}>{item.nama}</span>
                                  {item.isCustom ? <Badge color="purple">Tambahan</Badge> : <Badge color="teal">Bawaan</Badge>}
                                </div>
                                <div style={{ fontSize:12, color:"#888" }}>{item.satuan} · Stok awal: <strong>{stokAwal}</strong></div>
                              </div>
                              <div style={{ display:"flex", gap:6, flexShrink:0 }}>
                                <button onClick={()=>startEdit(item)} style={{ background:"none", border:`1px solid #c8e6d8`, borderRadius:8, padding:"6px 10px", cursor:"pointer", fontSize:13, color:THEME.primary }}>✏️</button>
                                <button onClick={()=>deleteItem(item.kode)} style={{ background:"none", border:"1px solid #fcc", borderRadius:8, padding:"6px 10px", cursor:"pointer", fontSize:13, color:"#A32D2D" }}>🗑️</button>
                              </div>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>
            )}

          </div>
        </div>

        <div style={{ marginTop:16, fontSize:11, color:"#9db", textAlign:"center" }}>
          ⚡ Powered by Supabase · auto-refresh tiap 10 detik
        </div>
      </div>
    </div>
  );
}
