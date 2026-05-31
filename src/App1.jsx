// ═══════════════════════════════════════════════════════════════
// МАГАЗИН ЕСАБЫ — ТОЛЫҚ ПРОГРАММА
// Xprinter XP-58IIH | Supabase | Role-Based | Excel | Telegram
// ═══════════════════════════════════════════════════════════════

import { useState, useEffect, useRef, useCallback } from "react";
import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer
} from "recharts";

// ─── SUPABASE CONFIG ────────────────────────────────────────────
// .env.local файлына қойың:
// VITE_SUPABASE_URL=https://xxxx.supabase.co
// VITE_SUPABASE_ANON_KEY=eyJhbGc...
// VITE_TELEGRAM_BOT_TOKEN=7xxxxxxx:AAF...
// VITE_TELEGRAM_CHAT_ID=-100xxxxxxxxx
// VITE_SHOP_NAME=Магазин атыңыз

const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
);

const SHOP_NAME = import.meta.env.VITE_SHOP_NAME || "Магазин";
const TG_TOKEN = import.meta.env.VITE_TELEGRAM_BOT_TOKEN;
const TG_CHAT  = import.meta.env.VITE_TELEGRAM_CHAT_ID;

// ─── HELPERS ────────────────────────────────────────────────────
const fmt  = (n) => Number(n || 0).toLocaleString("uz-UZ") + " сўм";
const today = () => new Date().toISOString().slice(0, 10);
const now   = () => new Date().toLocaleTimeString("uz-UZ", { hour: "2-digit", minute: "2-digit" });
const COLORS = ["#f59e0b","#10b981","#3b82f6","#ec4899","#8b5cf6","#ef4444","#14b8a6","#f97316"];

const ROLES = { director: "Директор", seller: "Сатыўшы", supply: "Снабженец" };
const PAYMENT = { cash: "💵 Нақ", card: "💳 Терминал", qr: "📱 QR", debt: "📒 Қарыз" };

// ─── PRINT (ESC/POS via browser) ────────────────────────────────
async function printReceipt(sale, items) {
  const w = window.open("", "_blank", "width=400,height=600");
  const lines = items.map(i =>
    `<tr><td>${i.product_name}</td><td style="text-align:right">${i.qty} × ${fmt(i.sell_price)}</td><td style="text-align:right">${fmt(i.qty * i.sell_price)}</td></tr>`
  ).join("");
  w.document.write(`
    <html><head><style>
      body{font-family:monospace;font-size:14px;width:95%;margin:1px;padding:1px}
      h2{text-align:center;font-size:14px;margin:4px 0}
      p{text-align:center;margin:2px;font-size:11px}
      table{width:100%;border-collapse:collapse}
      td{padding:2px 0;font-size:12px}
      hr{border:1px dashed #000}
      .total{font-size:14px;font-weight:bold;text-align:right}
    </style></head><body>
      <h2>${SHOP_NAME}</h2>
      <hr/>
      <p>📅 ${sale.date}  ⏰ ${now()}</p>
      <p>${PAYMENT[sale.payment_type] || sale.payment_type}</p>
      <hr/>
      <table>${lines}</table>
      <hr/>
      <div class="total">ЖӘМИ: ${fmt(sale.total)}</div>
      <hr/>
      <p>Рахмет! Қайта келиң!</p>
    </body></html>
  `);
  w.document.close();
  w.focus();
  setTimeout(() => { w.print(); w.close(); }, 500);
}

// ─── TELEGRAM ────────────────────────────────────────────────────
async function sendTelegram(text) {
  if (!TG_TOKEN || !TG_CHAT) return;
  try {
    await fetch(`https://api.telegram.org/bot${TG_TOKEN}/sendMessage`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ chat_id: TG_CHAT, text, parse_mode: "HTML" })
    });
  } catch {}
}

async function sendTelegramFile(blob, filename, caption) {
  if (!TG_TOKEN || !TG_CHAT) return;
  try {
    const fd = new FormData();
    fd.append("chat_id", TG_CHAT);
    fd.append("document", blob, filename);
    fd.append("caption", caption);
    await fetch(`https://api.telegram.org/bot${TG_TOKEN}/sendDocument`, { method: "POST", body: fd });
  } catch {}
}

// ─── EXCEL EXPORT ────────────────────────────────────────────────
function exportStock(products) {
  const data = products.map((p, i) => ({
    "№": i + 1,
    "Товар аты": p.name,
    "Штрих-код": p.barcode || "",
    "Саны": p.stock,
    "Өлшем": p.unit,
    "Кирис баҳасы": p.buy_price,
    "Сатыў баҳасы": p.sell_price,
    "Жәми қуны (кирис)": p.stock * p.buy_price,
    "Жәми қуны (сатыў)": p.stock * p.sell_price,
  }));
  const totalBuy  = products.reduce((s, p) => s + p.stock * p.buy_price, 0);
  const totalSell = products.reduce((s, p) => s + p.stock * p.sell_price, 0);
  data.push({ "Товар аты": "ЖӘМИ", "Жәми қуны (кирис)": totalBuy, "Жәми қуны (сатыў)": totalSell });

  const ws = XLSX.utils.json_to_sheet(data);
  ws["!cols"] = [{ wch: 4 },{ wch: 30 },{ wch: 15 },{ wch: 8 },{ wch: 8 },{ wch: 14 },{ wch: 14 },{ wch: 18 },{ wch: 18 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Қалдық товар");
  const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  return new Blob([buf], { type: "application/octet-stream" });
}

function downloadExcel(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

function exportPurchaseTemplate() {
  const data = [
    { "Товар аты": "Мысал: Шекер 1кг", "Штрих-код": "4600123456", "Саны": 100, "Кирис баҳасы": 12000, "Сатыў баҳасы": 15000, "Өлшем": "кг" },
  ];
  const ws = XLSX.utils.json_to_sheet(data);
  ws["!cols"] = [{ wch: 30 },{ wch: 15 },{ wch: 8 },{ wch: 14 },{ wch: 14 },{ wch: 8 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Приход");
  XLSX.writeFile(wb, "приход_шаблон.xlsx");
}

// ═══════════════════════════════════════════════════════════════
// MAIN APP
// ═══════════════════════════════════════════════════════════════
export default function App() {
  const [user, setUser]       = useState(null);
  const [profile, setProfile] = useState(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => {
      setUser(data?.session?.user || null);
      setLoading(false);
    });
    const { data: listener } = supabase.auth.onAuthStateChange((_e, session) => {
      setUser(session?.user || null);
    });
    return () => listener.subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (!user) { setProfile(null); return; }
    supabase.from("profiles").select("*").eq("id", user.id).single()
      .then(({ data }) => setProfile(data));
  }, [user]);

  if (loading) return <Splash text="Жүкленип атыр…" />;
  if (!user)   return <AuthPage />;
  if (!profile) return <Splash text="Профиль жүкленип атыр…" />;

  return <MainApp profile={profile} />;
}

// ─── SPLASH ─────────────────────────────────────────────────────
function Splash({ text }) {
  return (
    <div style={{ display:"flex", alignItems:"center", justifyContent:"center", height:"100vh", background:"#0f172a", color:"#f59e0b", fontSize:20 }}>
      {text}
    </div>
  );
}

// ─── AUTH PAGE ──────────────────────────────────────────────────
function AuthPage() {
  const [email, setEmail]   = useState("");
  const [pass,  setPass]    = useState("");
  const [err,   setErr]     = useState("");
  const [load,  setLoad]    = useState(false);

  const login = async () => {
    setLoad(true); setErr("");
    const { error } = await supabase.auth.signInWithPassword({ email, password: pass });
    if (error) setErr("Логин ямаса парол қате!");
    setLoad(false);
  };

  return (
    <div style={{ minHeight:"100vh", background:"#0f172a", display:"flex", alignItems:"center", justifyContent:"center", padding:16 }}>
      <div style={{ background:"#1e293b", borderRadius:16, padding:28, width:"100%", maxWidth:360 }}>
        <div style={{ textAlign:"center", marginBottom:24 }}>
          <div style={{ fontSize:48 }}>🏪</div>
          <div style={{ fontSize:22, fontWeight:700, color:"#f59e0b" }}>{SHOP_NAME}</div>
          <div style={{ fontSize:13, color:"#64748b", marginTop:4 }}>Магазин есабы</div>
        </div>
        {err && <div style={{ background:"#7f1d1d", color:"#fca5a5", borderRadius:8, padding:10, marginBottom:12, fontSize:13 }}>{err}</div>}
        <Inp placeholder="Email" value={email} onChange={setEmail} type="email" />
        <Inp placeholder="Парол" value={pass}  onChange={setPass}  type="password" />
        <Btn label={load ? "Кирип атыр…" : "Кириў"} onClick={login} disabled={load} />
      </div>
    </div>
  );
}

// ─── MAIN APP ────────────────────────────────────────────────────
const NAV_ALL     = [
  { id:"dashboard", icon:"📊", label:"Бас бет" },
  { id:"sell",      icon:"💰", label:"Сатыў" },
  { id:"purchase",  icon:"🛒", label:"Кирис" },
  { id:"products",  icon:"📦", label:"Товарлар" },
  { id:"clients",   icon:"👥", label:"Клиентлер" },
  { id:"requests",  icon:"📋", label:"Сораныс" },
  { id:"stats",     icon:"📈", label:"Статистика" },
  { id:"reports",   icon:"📤", label:"Есабат" },
  { id:"settings",  icon:"⚙️", label:"Параметр" },
];

const NAV_ROLES = {
  seller:   ["dashboard","sell","requests"],
  supply:   ["dashboard","purchase","products","requests","reports"],
  director: ["dashboard","sell","purchase","products","clients","requests","stats","reports","settings"],
};

function MainApp({ profile }) {
  const [tab, setTab] = useState("dashboard");
  const [products,  setProducts]  = useState([]);
  const [clients,   setClients]   = useState([]);
  const [categories, setCategories] = useState([]);

  const allowed = NAV_ROLES[profile.role] || NAV_ROLES.seller;
  const nav = NAV_ALL.filter(n => allowed.includes(n.id));

  useEffect(() => {
    supabase.from("products").select("*").order("name").then(({ data }) => setProducts(data || []));
    supabase.from("clients").select("*").order("name").then(({ data }) => setClients(data || []));
    supabase.from("categories").select("*").order("name").then(({ data }) => setCategories(data || []));
  }, []);

  const refreshProducts = () =>
    supabase.from("products").select("*").order("name").then(({ data }) => setProducts(data || []));
  const refreshClients = () =>
    supabase.from("clients").select("*").order("name").then(({ data }) => setClients(data || []));

  const logout = () => supabase.auth.signOut();

  const pages = { dashboard: Dashboard, sell: Sell, purchase: Purchase, products: Products,
    clients: Clients, requests: Requests, stats: Stats, reports: Reports, settings: Settings };
  const Page = pages[tab] || Dashboard;

  return (
    <div style={{ fontFamily:"'Segoe UI',sans-serif", background:"#0f172a", minHeight:"100vh", color:"#e2e8f0", paddingBottom:72 }}>
      {/* Header */}
      <div style={{ background:"linear-gradient(135deg,#1e3a5f,#0f172a)", padding:"12px 16px", display:"flex", justifyContent:"space-between", alignItems:"center", borderBottom:"1px solid #1e293b" }}>
        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
          <span style={{ fontSize:26 }}>🏪</span>
          <div>
            <div style={{ fontWeight:700, fontSize:16, color:"#f59e0b" }}>{SHOP_NAME}</div>
            <div style={{ fontSize:11, color:"#64748b" }}>{ROLES[profile.role]} — {profile.full_name}</div>
          </div>
        </div>
        <button onClick={logout} style={{ background:"none", border:"1px solid #334155", borderRadius:8, color:"#64748b", padding:"6px 12px", cursor:"pointer", fontSize:12 }}>Шығыў</button>
      </div>

      {/* Content */}
      <div style={{ padding:"14px 12px" }}>
        <Page profile={profile} products={products} clients={clients} categories={categories}
          setProducts={setProducts} refreshProducts={refreshProducts} refreshClients={refreshClients} />
      </div>

      {/* Bottom Nav */}
      <div style={{ position:"fixed", bottom:0, left:0, right:0, background:"#0f172a", borderTop:"1px solid #1e293b", display:"flex", zIndex:100, overflowX:"auto" }}>
        {nav.map(n => (
          <button key={n.id} onClick={() => setTab(n.id)}
            style={{ flex:1, minWidth:56, padding:"7px 2px 9px", background:"none", border:"none", cursor:"pointer",
              color: tab===n.id ? "#f59e0b" : "#475569", display:"flex", flexDirection:"column", alignItems:"center", gap:1 }}>
            <span style={{ fontSize:19 }}>{n.icon}</span>
            <span style={{ fontSize:8, fontWeight: tab===n.id ? 700 : 400, whiteSpace:"nowrap" }}>{n.label}</span>
          </button>
        ))}
      </div>
    </div>
  );
}

// ─── DASHBOARD ───────────────────────────────────────────────────
function Dashboard({ products }) {
  const [todaySales, setTodaySales] = useState([]);
  const [todayItems, setTodayItems] = useState([]);

  useEffect(() => {
    supabase.from("sales").select("*").eq("date", today()).then(({ data }) => setTodaySales(data || []));
    supabase.from("sale_items").select("*, sales!inner(date)").eq("sales.date", today())
      .then(({ data }) => setTodayItems(data || []));
  }, []);

  const revenue = todaySales.reduce((s, x) => s + Number(x.total), 0);
  const cost    = todayItems.reduce((s, x) => s + x.qty * x.buy_price, 0);
  const profit  = revenue - cost;
  const byPay   = { cash:0, card:0, qr:0, debt:0 };
  todaySales.forEach(s => { byPay[s.payment_type] = (byPay[s.payment_type]||0) + Number(s.total); });
  const lowStock = products.filter(p => p.stock <= p.min_stock);

  // top products
  const pMap = {};
  todayItems.forEach(i => { pMap[i.product_name] = (pMap[i.product_name]||0) + i.qty; });
  const top3 = Object.entries(pMap).sort((a,b)=>b[1]-a[1]).slice(0,3);

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
      <div style={{ fontSize:12, color:"#64748b" }}>📅 Бүгин: {today()}</div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
        <Card icon="💰" label="Бүгинги сатыў" value={fmt(revenue)} color="#10b981" />
        <Card icon="📈" label="Бүгинги пайда" value={fmt(profit)} color="#f59e0b" />
        <Card icon="🛒" label="Сатыўлар саны" value={todaySales.length+" рет"} color="#3b82f6" />
        <Card icon="📦" label="Товар түрлери" value={products.length+" түр"} color="#8b5cf6" />
      </div>

      {/* Төлем түрлери */}
      <div style={{ background:"#1e293b", borderRadius:12, padding:12 }}>
        <div style={{ fontWeight:700, color:"#f59e0b", marginBottom:8, fontSize:13 }}>💳 Бүгинги төлем түрлери</div>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8 }}>
          {Object.entries(byPay).map(([k,v]) => (
            <div key={k} style={{ background:"#0f172a", borderRadius:8, padding:"8px 10px" }}>
              <div style={{ fontSize:11, color:"#64748b" }}>{PAYMENT[k]}</div>
              <div style={{ fontWeight:700, fontSize:13, color:"#e2e8f0", marginTop:2 }}>{fmt(v)}</div>
            </div>
          ))}
        </div>
      </div>

      {lowStock.length > 0 && (
        <div style={{ background:"#7f1d1d", borderRadius:12, padding:12 }}>
          <div style={{ fontWeight:700, color:"#fca5a5", marginBottom:8, fontSize:13 }}>⚠️ Тауысылып атырған товарлар ({lowStock.length})</div>
          {lowStock.slice(0,5).map(p => (
            <div key={p.id} style={{ display:"flex", justifyContent:"space-between", fontSize:12, color:"#fecaca", padding:"3px 0" }}>
              <span>{p.name}</span><span style={{ fontWeight:700 }}>{p.stock} {p.unit}</span>
            </div>
          ))}
        </div>
      )}

      {top3.length > 0 && (
        <div style={{ background:"#1e293b", borderRadius:12, padding:12 }}>
          <div style={{ fontWeight:700, color:"#f59e0b", marginBottom:8, fontSize:13 }}>🏆 Бүгинги топ товарлар</div>
          {top3.map(([name, qty], i) => (
            <div key={name} style={{ display:"flex", justifyContent:"space-between", padding:"5px 0", borderBottom:"1px solid #0f172a", fontSize:13 }}>
              <span>{["🥇","🥈","🥉"][i]} {name}</span>
              <span style={{ color:"#f59e0b" }}>{qty} дана</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ─── SELL ────────────────────────────────────────────────────────
function Sell({ profile, products, clients, refreshProducts, refreshClients }) {
  const [cart, setCart]         = useState([]);
  const [date, setDate]         = useState(today());
  const [productId, setProductId] = useState("");
  const [qty, setQty]           = useState("");
  const [payType, setPayType]   = useState("cash");
  const [clientId, setClientId] = useState("");
  const [comment, setComment]   = useState("");
  const [msg, setMsg]           = useState("");
  const [saving, setSaving]     = useState(false);
  const [handoverAmt, setHandoverAmt] = useState("");
  const [handoverComment, setHandoverComment] = useState("");
  const [handoverMsg, setHandoverMsg] = useState("");
  const [recentSales, setRecentSales] = useState([]);

  const loadRecent = () =>
    supabase.from("sales").select("*, sale_items(*)").order("created_at", { ascending:false }).limit(10)
      .then(({ data }) => setRecentSales(data || []));

  useEffect(() => { loadRecent(); }, []);

  const addToCart = () => {
    const p = products.find(x => x.id === +productId);
    if (!p || !qty) return;
    const q = +qty;
    if (q > p.stock) { setMsg("❌ Қалдықта жеткиликсиз!"); setTimeout(()=>setMsg(""),3000); return; }
    setCart(c => {
      const ex = c.find(i => i.product_id === p.id);
      if (ex) return c.map(i => i.product_id===p.id ? {...i, qty:i.qty+q} : i);
      return [...c, { product_id:p.id, product_name:p.name, qty:q, sell_price:p.sell_price, buy_price:p.buy_price, unit:p.unit }];
    });
    setQty(""); setProductId("");
  };

  const total = cart.reduce((s,i) => s + i.qty*i.sell_price, 0);

  const submit = async () => {
    if (!cart.length) return;
    setSaving(true);
    const { data: sale, error } = await supabase.from("sales").insert({
      seller_id: profile.id, client_id: clientId||null,
      payment_type: payType, total, comment, date
    }).select().single();
    if (error) { setMsg("❌ Қате болды!"); setSaving(false); return; }
    await supabase.from("sale_items").insert(cart.map(i => ({ ...i, sale_id: sale.id })));
    // stock update
    for (const i of cart) {
      const p = products.find(x => x.id === i.product_id);
      if (p) await supabase.from("products").update({ stock: p.stock - i.qty }).eq("id", p.id);
    }
    // debt update
    if (payType === "debt" && clientId) {
      const cl = clients.find(c => c.id === +clientId);
      if (cl) await supabase.from("clients").update({ debt: cl.debt + total }).eq("id", cl.id);
      await supabase.from("client_history").insert({
        client_id: +clientId, type: "debt",
        amount: total, comment: comment||null, date
      });
      refreshClients();
    }
    await printReceipt(sale, cart);
    await sendTelegram(`🛒 <b>Жаңа сатыў</b>\n👤 ${profile.full_name}\n💰 ${fmt(total)}\n${PAYMENT[payType]}\n📅 ${date}`);
    refreshProducts();
    setCart([]); setComment(""); setClientId("");
    setMsg("✅ Сатыў сақланды!");
    setTimeout(()=>setMsg(""),3000);
    loadRecent();
    setSaving(false);
  };

  const submitHandover = async () => {
    if (!handoverAmt) return;
    await supabase.from("cash_handovers").insert({ seller_id:profile.id, amount:+handoverAmt, comment:handoverComment, date:today() });
    await sendTelegram(`💵 <b>Касса тапсырылды</b>\n👤 ${profile.full_name}\n💰 ${fmt(handoverAmt)}\n💬 ${handoverComment||"—"}`);
    setHandoverAmt(""); setHandoverComment("");
    setHandoverMsg("✅ Тапсырылды!"); setTimeout(()=>setHandoverMsg(""),3000);
  };

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
      {msg && <Alert msg={msg} />}
      <div style={{ background:"#1e293b", borderRadius:12, padding:14 }}>
        <div style={{ fontWeight:700, color:"#10b981", marginBottom:10 }}>💰 Жаңа сатыў</div>
        <label style={{ fontSize:11, color:"#94a3b8" }}>Күни</label>
        <input type="date" value={date} onChange={e=>setDate(e.target.value)}
          style={inputStyle} />
        <SearchPicker products={products} value={productId} onChange={setProductId} />
        <div style={{ display:"flex", gap:8, marginBottom:8 }}>
          <input type="number" min="1" value={qty} onChange={e=>setQty(e.target.value)} placeholder="Саны"
            style={{ flex:1, ...inputStyle, marginBottom:0 }} />
          <button onClick={addToCart} style={btnStyle("#10b981")}>Қос</button>
        </div>
        {/* Төлем түри */}
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:6, marginBottom:8 }}>
          {Object.entries(PAYMENT).map(([k,v]) => (
            <button key={k} onClick={()=>setPayType(k)}
              style={{ padding:"8px 4px", background:payType===k?"#10b981":"#0f172a", border:`1px solid ${payType===k?"#10b981":"#334155"}`,
                borderRadius:8, color:payType===k?"#0f172a":"#94a3b8", fontWeight:payType===k?700:400, cursor:"pointer", fontSize:12 }}>
              {v}
            </button>
          ))}
        </div>
        {/* Клиент (қарызда) */}
        {payType === "debt" && (
          <select value={clientId} onChange={e=>setClientId(e.target.value)} style={inputStyle}>
            <option value="">Клиент таңлаң...</option>
            {clients.map(c => <option key={c.id} value={c.id}>{c.name} (қарыз: {fmt(c.debt)})</option>)}
          </select>
        )}
        <Inp placeholder="Комментарий (ихтиярий)" value={comment} onChange={setComment} />
      </div>

      {cart.length > 0 && (
        <div style={{ background:"#1e293b", borderRadius:12, padding:12 }}>
          <div style={{ fontWeight:700, color:"#f59e0b", marginBottom:8 }}>🛒 Себет</div>
          {cart.map(i => (
            <div key={i.product_id} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"6px 0", borderBottom:"1px solid #0f172a", fontSize:13 }}>
              <span style={{ flex:1 }}>{i.product_name}</span>
              <span style={{ color:"#10b981", marginRight:8 }}>{i.qty}×{fmt(i.sell_price)}</span>
              <button onClick={()=>setCart(c=>c.filter(x=>x.product_id!==i.product_id))}
                style={{ background:"none",border:"none",color:"#ef4444",cursor:"pointer",fontSize:16 }}>✕</button>
            </div>
          ))}
          <div style={{ display:"flex", justifyContent:"space-between", marginTop:10, fontWeight:700, fontSize:16 }}>
            <span>Жәми:</span><span style={{ color:"#f59e0b" }}>{fmt(total)}</span>
          </div>
          <button onClick={submit} disabled={saving}
            style={{ width:"100%", padding:12, background:"#10b981", border:"none", borderRadius:8, color:"#0f172a", fontWeight:700, cursor:"pointer", marginTop:8, fontSize:14 }}>
            {saving ? "Сақланып атыр…" : "✅ Сатыўды сақлаў ҳәм чек"}
          </button>
        </div>
      )}

      {/* Касса тапсырыў */}
      <div style={{ background:"#1e293b", borderRadius:12, padding:12 }}>
        <div style={{ fontWeight:700, color:"#3b82f6", marginBottom:8, fontSize:13 }}>💵 Касса тапсырыў</div>
        {handoverMsg && <Alert msg={handoverMsg} />}
        <Inp placeholder="Сумма (сўм)" value={handoverAmt} onChange={setHandoverAmt} type="number" />
        <Inp placeholder="Комментарий" value={handoverComment} onChange={setHandoverComment} />
        <Btn label="Тапсырыў" onClick={submitHandover} color="#3b82f6" />
      </div>

      {/* Соңғы сатыўлар */}
      <div style={{ background:"#1e293b", borderRadius:12, padding:12 }}>
        <div style={{ fontWeight:700, color:"#f59e0b", marginBottom:8, fontSize:13 }}>📋 Соңғы сатыўлар</div>
        {recentSales.map(s => (
          <div key={s.id} style={{ padding:"8px 0", borderBottom:"1px solid #0f172a" }}>
            <div style={{ display:"flex", justifyContent:"space-between", fontSize:13 }}>
              <span style={{ color:"#94a3b8" }}>{s.date}</span>
              <span style={{ color:"#10b981", fontWeight:700 }}>{fmt(s.total)}</span>
              <span style={{ fontSize:11, color:"#64748b" }}>{PAYMENT[s.payment_type]}</span>
            </div>
            {s.comment && <div style={{ fontSize:11, color:"#64748b" }}>💬 {s.comment}</div>}
          </div>
        ))}
      </div>
    </div>
  );
}

// ─── PURCHASE ────────────────────────────────────────────────────
function Purchase({ profile, products, categories, refreshProducts }) {
  const [cart, setCart]         = useState([]);
  const [date, setDate]         = useState(today());
  const [productId, setProductId] = useState("");
  const [qty, setQty]           = useState("");
  const [buyPrice, setBuyPrice] = useState("");
  const [sellPrice, setSellPrice] = useState("");
  const [comment, setComment]   = useState("");
  const [msg, setMsg]           = useState("");
  const [saving, setSaving]     = useState(false);

  // Excel import
  const fileRef = useRef();
  const importExcel = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const wb = XLSX.read(ev.target.result, { type:"array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws);
      const items = rows.map(r => ({
        product_name: r["Товар аты"] || "",
        barcode:      String(r["Штрих-код"] || ""),
        qty:          +r["Саны"] || 0,
        buy_price:    +r["Кирис баҳасы"] || 0,
        sell_price:   +r["Сатыў баҳасы"] || 0,
        unit:         r["Өлшем"] || "дана",
      })).filter(r => r.product_name && r.qty > 0);
      setCart(items.map(i => ({ ...i, product_id: products.find(p=>p.barcode===i.barcode||p.name===i.product_name)?.id || null })));
      setMsg(`✅ ${items.length} товар жүкленди`);
      setTimeout(()=>setMsg(""),3000);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const addToCart = () => {
    const p = products.find(x => x.id === +productId);
    if (!p || !qty || !buyPrice) return;
    setCart(c => {
      const ex = c.find(i => i.product_id === p.id);
      if (ex) return c.map(i => i.product_id===p.id ? {...i, qty:i.qty+(+qty), buy_price:+buyPrice, sell_price:+sellPrice||p.sell_price} : i);
      return [...c, { product_id:p.id, product_name:p.name, qty:+qty, buy_price:+buyPrice, sell_price:+sellPrice||p.sell_price, unit:p.unit }];
    });
    setQty("1"); setBuyPrice(""); setSellPrice(""); setProductId("");
  };

  const submit = async () => {
    if (!cart.length) return;
    setSaving(true);
    const { data: purchase } = await supabase.from("purchases").insert({
      supply_id: profile.id, comment, date
    }).select().single();
    await supabase.from("purchase_items").insert(cart.map(i => ({ ...i, purchase_id: purchase.id })));
    for (const i of cart) {
      if (i.product_id) {
        const p = products.find(x => x.id === i.product_id);
        if (p) await supabase.from("products").update({ stock: p.stock + i.qty, buy_price: i.buy_price, sell_price: i.sell_price }).eq("id", p.id);
      } else if (i.product_name) {
        await supabase.from("products").insert({ name:i.product_name, barcode:i.barcode||null, buy_price:i.buy_price, sell_price:i.sell_price, stock:i.qty, unit:i.unit });
      }
    }
    await sendTelegram(`📦 <b>Жаңа кирис</b>\n👤 ${profile.full_name}\n📦 ${cart.length} түр товар\n📅 ${date}`);
    refreshProducts();
    setCart([]); setComment("");
    setMsg("✅ Кирис сақланды!"); setTimeout(()=>setMsg(""),3000);
    setSaving(false);
  };

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
      {msg && <Alert msg={msg} />}
      <div style={{ display:"flex", gap:8 }}>
        <button onClick={exportPurchaseTemplate} style={{ flex:1, padding:10, background:"#1e293b", border:"1px solid #334155", borderRadius:8, color:"#94a3b8", cursor:"pointer", fontSize:13 }}>
          📥 Шаблон жүклеў
        </button>
        <button onClick={()=>fileRef.current.click()} style={{ flex:1, padding:10, background:"#3b82f6", border:"none", borderRadius:8, color:"#fff", cursor:"pointer", fontWeight:700, fontSize:13 }}>
          📤 Excel-ден кирис
        </button>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={importExcel} style={{ display:"none" }} />
      </div>

      <div style={{ background:"#1e293b", borderRadius:12, padding:14 }}>
        <div style={{ fontWeight:700, color:"#3b82f6", marginBottom:10 }}>🛒 Қолдан кирис</div>
        <input type="date" value={date} onChange={e=>setDate(e.target.value)} style={inputStyle} />
        <SearchPicker products={products} value={productId} onChange={(id)=>{
          setProductId(id);
          const p = products.find(x=>x.id===+id);
          if(p){ setBuyPrice(String(p.buy_price)); setSellPrice(String(p.sell_price)); }
        }} />
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:6, marginBottom:8 }}>
          <input type="number" value={qty} onChange={e=>setQty(e.target.value)} placeholder="Саны" style={{ ...inputStyle, marginBottom:0 }} />
          <input type="number" value={buyPrice} onChange={e=>setBuyPrice(e.target.value)} placeholder="Кирис баҳа" style={{ ...inputStyle, marginBottom:0 }} />
          <input type="number" value={sellPrice} onChange={e=>setSellPrice(e.target.value)} placeholder="Сатыў баҳа" style={{ ...inputStyle, marginBottom:0 }} />
        </div>
        <Btn label="Қосыў" onClick={addToCart} color="#3b82f6" />
      </div>

      {cart.length > 0 && (
        <div style={{ background:"#1e293b", borderRadius:12, padding:12 }}>
          <div style={{ fontWeight:700, color:"#f59e0b", marginBottom:8 }}>📦 Кирис тизими ({cart.length} товар)</div>
          {cart.map((i,idx) => (
            <div key={idx} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"6px 0", borderBottom:"1px solid #0f172a", fontSize:12 }}>
              <div>
                <div style={{ color:"#e2e8f0" }}>{i.product_name}</div>
                <div style={{ color:"#64748b" }}>{i.qty} {i.unit} × {fmt(i.buy_price)}</div>
              </div>
              <button onClick={()=>setCart(c=>c.filter((_,j)=>j!==idx))} style={{ background:"none",border:"none",color:"#ef4444",cursor:"pointer",fontSize:16 }}>✕</button>
            </div>
          ))}
          <Inp placeholder="Комментарий" value={comment} onChange={setComment} />
          <button onClick={submit} disabled={saving} style={{ width:"100%", padding:12, background:"#3b82f6", border:"none", borderRadius:8, color:"#fff", fontWeight:700, cursor:"pointer", marginTop:6 }}>
            {saving ? "Сақланып атыр…" : "✅ Киристи сақлаў"}
          </button>
        </div>
      )}
    </div>
  );
}

// ─── PRODUCTS ────────────────────────────────────────────────────
function Products({ products, categories, refreshProducts }) {
  const [search, setSearch] = useState("");
  const [form, setForm]     = useState({ name:"", barcode:"", category_id:"", buy_price:"", sell_price:"", stock:"", min_stock:"5", unit:"дана" });
  const [editId, setEditId] = useState(null);
  const [msg, setMsg]       = useState("");

  const filtered = products.filter(p =>
    p.name.toLowerCase().includes(search.toLowerCase()) ||
    (p.barcode||"").includes(search)
  );

  const submit = async () => {
    if (!form.name || !form.sell_price) return;
    const obj = { name:form.name, barcode:form.barcode||null, category_id:form.category_id||null,
      buy_price:+form.buy_price, sell_price:+form.sell_price, stock:+form.stock, min_stock:+form.min_stock, unit:form.unit };
    if (editId) {
      await supabase.from("products").update(obj).eq("id", editId);
    } else {
      await supabase.from("products").insert(obj);
    }
    refreshProducts();
    setForm({ name:"", barcode:"", category_id:"", buy_price:"", sell_price:"", stock:"", min_stock:"5", unit:"дана" });
    setEditId(null);
    setMsg("✅ Сақланды!"); setTimeout(()=>setMsg(""),2000);
  };

  const del = async (id) => {
    if (!window.confirm("Өшире аласаңба?")) return;
    await supabase.from("products").delete().eq("id", id);
    refreshProducts();
  };

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
      {msg && <Alert msg={msg} />}
      <div style={{ background:"#1e293b", borderRadius:12, padding:14 }}>
        <div style={{ fontWeight:700, color:"#f59e0b", marginBottom:10 }}>{editId?"✏️ Өзгертиў":"➕ Жаңа товар"}</div>
        <Inp placeholder="Товар аты *" value={form.name} onChange={v=>setForm(f=>({...f,name:v}))} />
        <div style={{ display:"flex", gap:8 }}>
          <input placeholder="Штрих-код" value={form.barcode} onChange={e=>setForm(f=>({...f,barcode:e.target.value}))}
            style={{ flex:1, ...inputStyle, marginBottom:0 }} />
          <ScanBtn onScan={code=>setForm(f=>({...f,barcode:code}))} label="📷" />
        </div>
        <div style={{ height:8 }}/>
        <select value={form.category_id} onChange={e=>setForm(f=>({...f,category_id:e.target.value}))} style={inputStyle}>
          <option value="">Категория...</option>
          {categories.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}
        </select>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:6 }}>
          <input type="number" placeholder="Кирис баҳасы" value={form.buy_price} onChange={e=>setForm(f=>({...f,buy_price:e.target.value}))} style={{ ...inputStyle, marginBottom:0 }} />
          <input type="number" placeholder="Сатыў баҳасы *" value={form.sell_price} onChange={e=>setForm(f=>({...f,sell_price:e.target.value}))} style={{ ...inputStyle, marginBottom:0 }} />
          <input type="number" placeholder="Баслапқы саны" value={form.stock} onChange={e=>setForm(f=>({...f,stock:e.target.value}))} style={{ ...inputStyle, marginBottom:0 }} />
          <input type="number" placeholder="Мин. қалдық" value={form.min_stock} onChange={e=>setForm(f=>({...f,min_stock:e.target.value}))} style={{ ...inputStyle, marginBottom:0 }} />
        </div>
        <div style={{ height:8 }}/>
        <input placeholder="Өлшем (дана/кг...)" value={form.unit} onChange={e=>setForm(f=>({...f,unit:e.target.value}))} style={inputStyle} />
        <Btn label={editId?"Сақлаў":"Қосыў"} onClick={submit} />
        {editId && <Btn label="Бийкар" onClick={()=>{setEditId(null);setForm({name:"",barcode:"",category_id:"",buy_price:"",sell_price:"",stock:"",min_stock:"5",unit:"дана"})}} secondary />}
      </div>

      <input placeholder="🔍 Аты ямаса штрих-код..." value={search} onChange={e=>setSearch(e.target.value)} style={inputStyle} />

      {filtered.map(p => (
        <div key={p.id} style={{ background:"#1e293b", borderRadius:12, padding:12 }}>
          <div style={{ display:"flex", justifyContent:"space-between" }}>
            <div>
              <div style={{ fontWeight:600, fontSize:14 }}>{p.name}</div>
              {p.barcode && <div style={{ fontSize:10, color:"#64748b" }}>🔢 {p.barcode}</div>}
            </div>
            <div style={{ display:"flex", gap:8 }}>
              <button onClick={()=>{setEditId(p.id);setForm({name:p.name,barcode:p.barcode||"",category_id:p.category_id||"",buy_price:p.buy_price,sell_price:p.sell_price,stock:p.stock,min_stock:p.min_stock,unit:p.unit})}}
                style={{ background:"none",border:"none",cursor:"pointer",fontSize:18 }}>✏️</button>
              <button onClick={()=>del(p.id)} style={{ background:"none",border:"none",cursor:"pointer",fontSize:18 }}>🗑️</button>
            </div>
          </div>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:6, marginTop:8 }}>
            <Mini label="Алыў" value={fmt(p.buy_price)} />
            <Mini label="Сатыў" value={fmt(p.sell_price)} color="#10b981" />
            <Mini label="Қалдық" value={`${p.stock} ${p.unit}`} color={p.stock<=p.min_stock?"#ef4444":"#f59e0b"} />
          </div>
        </div>
      ))}
    </div>
  );
}

// ─── DEBT PAYMENT ────────────────────────────────────────────────
function DebtPayment({ client, onPaid }) {
  const [amt, setAmt]         = useState("");
  const [comment, setComment] = useState("");
  const [date, setDate]       = useState(today());
  const [show, setShow]       = useState(false);

  const pay = async () => {
    if (!amt) return;
    const newDebt = Math.max(0, client.debt - +amt);
    await supabase.from("clients").update({ debt: newDebt }).eq("id", client.id);
    await supabase.from("client_history").insert({
      client_id: client.id, type: "payment",
      amount: +amt, comment: comment||null, date
    });
    //await supabase.from("cash_handovers").insert({
      //seller_id: null, amount: +amt,
      //comment: `Қарыз төлеў: ${client.name}${comment ? " — " + comment : ""}`, date
      //});
    setAmt(""); setComment(""); setShow(false);
    onPaid();
  };

  return (
    <div style={{ marginTop:8 }}>
      {!show ? (
        <button onClick={()=>setShow(true)}
          style={{ padding:"6px 14px", background:"#10b981", border:"none", borderRadius:8, color:"#0f172a", fontWeight:700, cursor:"pointer", fontSize:12 }}>
          💵 Қарыз төлеў
        </button>
      ) : (
        <div style={{ marginTop:8, display:"flex", flexDirection:"column", gap:6 }}>
          <input type="date" value={date} onChange={e=>setDate(e.target.value)}
            style={{ padding:"8px 10px", background:"#0f172a", border:"1px solid #334155", borderRadius:8, color:"#e2e8f0", fontSize:13 }} />
          <input type="number" placeholder="Сумма (сўм)" value={amt} onChange={e=>setAmt(e.target.value)}
            style={{ padding:"8px 10px", background:"#0f172a", border:"1px solid #334155", borderRadius:8, color:"#e2e8f0", fontSize:13 }} />
          <input type="text" placeholder="Комментарий (ихтиярий)" value={comment} onChange={e=>setComment(e.target.value)}
            style={{ padding:"8px 10px", background:"#0f172a", border:"1px solid #334155", borderRadius:8, color:"#e2e8f0", fontSize:13 }} />
          <div style={{ display:"flex", gap:6 }}>
            <button onClick={pay} style={{ flex:1, padding:"8px 14px", background:"#10b981", border:"none", borderRadius:8, color:"#0f172a", fontWeight:700, cursor:"pointer" }}>✅ Сақлаў</button>
            <button onClick={()=>setShow(false)} style={{ padding:"8px 14px", background:"#ef4444", border:"none", borderRadius:8, color:"#fff", fontWeight:700, cursor:"pointer" }}>✕</button>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── CLIENT HISTORY ──────────────────────────────────────────────
function ClientHistory({ client }) {
  const [list, setList] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    supabase.from("client_history").select("*")
      .eq("client_id", client.id)
      .order("date", { ascending: false })
      .then(({ data }) => { setList(data || []); setLoading(false); });
  }, [client.id]);

  return (
    <div style={{ marginTop:10, borderTop:"1px solid #334155", paddingTop:10 }}>
      <div style={{ fontWeight:700, color:"#f59e0b", marginBottom:8, fontSize:12 }}>📋 Тарийх</div>
      {loading && <div style={{ color:"#64748b", fontSize:12 }}>Жүктелип атыр…</div>}
      {!loading && list.length === 0 && <div style={{ color:"#475569", fontSize:12 }}>Тарийх жоқ</div>}
      {list.map(h => (
        <div key={h.id} style={{ display:"flex", justifyContent:"space-between", padding:"5px 0", borderBottom:"1px solid #0f172a", fontSize:12 }}>
          <div>
            <span style={{ color: h.type==="payment" ? "#10b981" : "#ef4444" }}>
              {h.type==="payment" ? "✅ Төлеў" : "📒 Қарыз"}
            </span>
            {h.comment && <span style={{ color:"#64748b", marginLeft:6 }}>— {h.comment}</span>}
          </div>
          <div style={{ textAlign:"right" }}>
            <div style={{ fontWeight:700, color: h.type==="payment"?"#10b981":"#ef4444" }}>{fmt(h.amount)}</div>
            <div style={{ fontSize:10, color:"#475569" }}>{h.date}</div>
          </div>
        </div>
      ))}
    </div>
  );
}

// ─── CLIENTS ────────────────────────────────────────────────────
function Clients({ clients, refreshClients }) {
  const [form, setForm]       = useState({ name:"", phone:"" });
  const [msg, setMsg]         = useState("");
  const [openId, setOpenId]   = useState(null);

  const submit = async () => {
    if (!form.name) return;
    await supabase.from("clients").insert({ name:form.name, phone:form.phone||null });
    refreshClients();
    setForm({ name:"", phone:"" });
    setMsg("✅ Клиент қосылды!"); setTimeout(()=>setMsg(""),2000);
  };

  const totalDebt = clients.reduce((s,c)=>s+Number(c.debt),0);

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
      {msg && <Alert msg={msg} />}
      <Card icon="📒" label="Жәми қарыз" value={fmt(totalDebt)} color="#ef4444" />
      <div style={{ background:"#1e293b", borderRadius:12, padding:14 }}>
        <div style={{ fontWeight:700, color:"#f59e0b", marginBottom:10 }}>➕ Жаңа клиент</div>
        <Inp placeholder="Аты *" value={form.name} onChange={v=>setForm(f=>({...f,name:v}))} />
        <Inp placeholder="Телефон" value={form.phone} onChange={v=>setForm(f=>({...f,phone:v}))} />
        <Btn label="Қосыў" onClick={submit} />
      </div>
      {clients.map(c => (
        <div key={c.id} style={{ background:"#1e293b", borderRadius:12, padding:12 }}>
          <div style={{ display:"flex", justifyContent:"space-between", cursor:"pointer" }}
            onClick={()=>setOpenId(openId===c.id ? null : c.id)}>
            <div>
              <div style={{ fontWeight:600 }}>{c.name} <span style={{ fontSize:11, color:"#64748b" }}>{openId===c.id?"▲":"▼"}</span></div>
              {c.phone && <div style={{ fontSize:12, color:"#64748b" }}>📞 {c.phone}</div>}
            </div>
            <div style={{ textAlign:"right" }}>
              <div style={{ fontWeight:700, color:c.debt>0?"#ef4444":"#10b981", fontSize:14 }}>{fmt(c.debt)}</div>
              <div style={{ fontSize:10, color:"#64748b" }}>қарыз</div>
            </div>
          </div>
          {c.debt > 0 && <DebtPayment client={c} onPaid={refreshClients} />}
          {openId === c.id && <ClientHistory client={c} />}
        </div>
      ))}
    </div>
  );
}

// ─── REQUESTS ────────────────────────────────────────────────────
function Requests({ profile, products }) {
  const [list, setList]         = useState([]);
  const [name, setName]         = useState("");
  const [qty, setQty]           = useState("");
  const [comment, setComment]   = useState("");
  const [msg, setMsg]           = useState("");
  const isSeller  = profile.role === "seller";
  const isSupply  = profile.role === "supply" || profile.role === "director";

  const load = () =>
    supabase.from("requests").select("*, profiles(full_name)").order("created_at",{ascending:false})
      .then(({ data }) => setList(data || []));

  useEffect(() => { load(); }, []);

  const submit = async () => {
    if (!name) return;
    await supabase.from("requests").insert({ seller_id:profile.id, product_name:name, qty:+qty||null, comment });
    await sendTelegram(`📋 <b>Жаңа сораныс</b>\n👤 ${profile.full_name}\n📦 ${name}${qty?" ("+qty+" дана)":""}\n💬 ${comment||"—"}`);
    setName(""); setQty(""); setComment("");
    setMsg("✅ Сораныс жиберилди!"); setTimeout(()=>setMsg(""),2000);
    load();
  };

  const updateStatus = async (id, status) => {
    await supabase.from("requests").update({ status }).eq("id", id);
    load();
  };

  const STATUS_COLOR = { new:"#f59e0b", ordered:"#3b82f6", done:"#10b981" };
  const STATUS_LABEL = { new:"Жаңа", ordered:"Заказ берилди", done:"Орынланды" };

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
      {msg && <Alert msg={msg} />}
      {isSeller && (
        <div style={{ background:"#1e293b", borderRadius:12, padding:14 }}>
          <div style={{ fontWeight:700, color:"#f59e0b", marginBottom:10 }}>📋 Жаңа сораныс</div>
          <Inp placeholder="Товар аты *" value={name} onChange={setName} />
          <Inp placeholder="Саны (ихтиярий)" value={qty} onChange={setQty} type="number" />
          <Inp placeholder="Комментарий" value={comment} onChange={setComment} />
          <Btn label="Жибериў" onClick={submit} />
        </div>
      )}
      {list.map(r => (
        <div key={r.id} style={{ background:"#1e293b", borderRadius:12, padding:12 }}>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start" }}>
            <div>
              <div style={{ fontWeight:600 }}>{r.product_name} {r.qty?`(${r.qty} дана)`:""}</div>
              <div style={{ fontSize:11, color:"#64748b" }}>👤 {r.profiles?.full_name} | {r.created_at?.slice(0,10)}</div>
              {r.comment && <div style={{ fontSize:12, color:"#94a3b8" }}>💬 {r.comment}</div>}
            </div>
            <span style={{ background:STATUS_COLOR[r.status]+"33", color:STATUS_COLOR[r.status], borderRadius:6, padding:"3px 8px", fontSize:11, fontWeight:700 }}>
              {STATUS_LABEL[r.status]}
            </span>
          </div>
          {isSupply && r.status !== "done" && (
            <div style={{ display:"flex", gap:6, marginTop:8 }}>
              {r.status === "new" && <button onClick={()=>updateStatus(r.id,"ordered")} style={btnStyle("#3b82f6","small")}>Заказ берилди</button>}
              <button onClick={()=>updateStatus(r.id,"done")} style={btnStyle("#10b981","small")}>Орындалды ✅</button>
            </div>
          )}
        </div>
      ))}
    </div>
  );
}

// ─── STATS ────────────────────────────────────────────────────────
function Stats() {
  const [days7, setDays7]         = useState([]);
  const [pieQty, setPieQty]       = useState([]);
  const [pieSum, setPieSum]       = useState([]);
  const [pieProfit, setPieProfit] = useState([]);
  const [monthly, setMonthly]     = useState([]);

  useEffect(() => {
    const d7 = Array.from({length:7},(_,i)=>{
      const d=new Date(); d.setDate(d.getDate()-6+i); return d.toISOString().slice(0,10);
    });
    // 7 күнлик
    supabase.from("sales").select("date,total").gte("date",d7[0])
      .then(({ data }) => {
        const map={};
        (data||[]).forEach(s=>{ map[s.date]=(map[s.date]||0)+Number(s.total); });
        setDays7(d7.map(d=>({ date:d.slice(5), rev:map[d]||0 })));
      });
    // Пирог диаграммалар
    supabase.from("sale_items").select("product_name, qty, sell_price, buy_price")
      .then(({ data }) => {
        const qtyMap={}, sumMap={}, profitMap={};
        let totalProfit = 0;
        (data||[]).forEach(i=>{
          const name = i.product_name || "Белгисиз";
          const qty = Number(i.qty||0);
          const sell = Number(i.sell_price||0);
          const buy = Number(i.buy_price||0);
          const profit = (sell - buy) * qty;
          qtyMap[name]    = (qtyMap[name]||0) + qty;
          sumMap[name]    = (sumMap[name]||0) + sell * qty;
          profitMap[name] = (profitMap[name]||0) + profit;
          totalProfit += profit;
        });
        const mkPie = (map) => Object.entries(map).sort((a,b)=>b[1]-a[1]).slice(0,6).map(([name,value])=>({name,value:Math.round(value)}));
        setPieQty(mkPie(qtyMap));
        setPieSum(mkPie(sumMap));
        setPieProfit(mkPie(profitMap));
      });
    // Айлық
    supabase.from("sales").select("date,total")
      .then(({ data }) => {
        const map={};
        (data||[]).forEach(s=>{ const m=s.date.slice(0,7); map[m]=(map[m]||0)+Number(s.total); });
        setMonthly(Object.entries(map).slice(-6).map(([m,rev])=>({month:m.slice(5),rev})));
      });
  }, []);

  const tooltipStyle = { contentStyle:{background:"#1e293b",border:"1px solid #f59e0b",borderRadius:8,color:"#e2e8f0"}, itemStyle:{color:"#e2e8f0"}, labelStyle:{color:"#f59e0b"} };

  const PieBlock = ({ data, title, formatter }) => (
    data.length > 0 ? (
      <div style={{ background:"#1e293b", borderRadius:12, padding:14 }}>
        <div style={{ fontWeight:700, color:"#10b981", marginBottom:12, fontSize:13 }}>{title}</div>
        <ResponsiveContainer width="100%" height={220}>
          <PieChart>
            <Pie data={data} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={80}
              label={({name, percent}) => `${(percent*100).toFixed(0)}%`}
              labelLine={true}>
              {data.map((_,i)=><Cell key={i} fill={COLORS[i%COLORS.length]} />)}
            </Pie>
            <Tooltip formatter={formatter||((v)=>v)} {...tooltipStyle} />
          </PieChart>
        </ResponsiveContainer>
        <div style={{ display:"flex", flexWrap:"wrap", gap:6, marginTop:4 }}>
          {data.map((d,i)=>(
            <div key={d.name} style={{ fontSize:11, color:COLORS[i%COLORS.length], background:"#0f172a", padding:"2px 8px", borderRadius:6 }}>
              ● {d.name}: {formatter ? formatter(d.value) : d.value}
            </div>
          ))}
        </div>
      </div>
    ) : null
  );

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
      <div style={{ background:"#1e293b", borderRadius:12, padding:14 }}>
        <div style={{ fontWeight:700, color:"#f59e0b", marginBottom:12, fontSize:13 }}>📅 Соңғы 7 күн</div>
        <ResponsiveContainer width="100%" height={180}>
          <BarChart data={days7}>
            <CartesianGrid strokeDasharray="3 3" stroke="#334155" />
            <XAxis dataKey="date" tick={{fill:"#94a3b8",fontSize:11}} />
            <YAxis tick={{fill:"#94a3b8",fontSize:10}} tickFormatter={v=>(v/1000)+"к"} />
            <Tooltip formatter={v=>fmt(v)} {...tooltipStyle} />
            <Bar dataKey="rev" fill="#f59e0b" radius={[4,4,0,0]} name="Сатыў" />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <PieBlock data={pieQty} title="🏆 Товар үлеси (саны бойынша)" formatter={v=>v+" дана"} />
      <PieBlock data={pieSum} title="💰 Товар үлеси (сумма бойынша)" formatter={fmt} />
      <PieBlock data={pieProfit} title="📈 Товар үлеси (пайда бойынша)" formatter={fmt} />

      {monthly.length > 0 && (
        <div style={{ background:"#1e293b", borderRadius:12, padding:14 }}>
          <div style={{ fontWeight:700, color:"#3b82f6", marginBottom:12, fontSize:13 }}>📆 Айлық сатыў</div>
          <ResponsiveContainer width="100%" height={180}>
            <LineChart data={monthly}>
              <CartesianGrid strokeDasharray="3 3" stroke="#334155" />
              <XAxis dataKey="month" tick={{fill:"#94a3b8",fontSize:11}} />
              <YAxis tick={{fill:"#94a3b8",fontSize:10}} tickFormatter={v=>(v/1000)+"к"} />
              <Tooltip formatter={v=>fmt(v)} {...tooltipStyle} />
              <Line type="monotone" dataKey="rev" stroke="#3b82f6" strokeWidth={2} dot={{fill:"#3b82f6"}} name="Сатыў" />
            </LineChart>
          </ResponsiveContainer>
        </div>
      )}
    </div>
  );
}

// ─── REPORTS ─────────────────────────────────────────────────────
function Reports({ profile, products }) {
  const [period, setPeriod]   = useState("today");
  const [sales, setSales]     = useState([]);
  const [items, setItems]     = useState([]);
  const [handovers, setHandovers] = useState([]);
  const [closing, setClosing] = useState(false);
  const [closeMsg, setCloseMsg] = useState("");

  useEffect(() => {
    const filter = period === "today" ? today()
      : period === "week" ? new Date(Date.now()-7*86400000).toISOString().slice(0,10)
      : new Date(Date.now()-30*86400000).toISOString().slice(0,10);
    supabase.from("sales").select("*").gte("date", filter)
      .then(({ data }) => setSales(data || []));
    supabase.from("sales").select("*, sale_items(*)").gte("date", filter)
      .then(({ data }) => {
        const allItems = (data||[]).flatMap(s => s.sale_items || []);
        setItems(allItems);
      });
    supabase.from("cash_handovers").select("*").gte("date", filter)
      .then(({ data }) => setHandovers(data || []));
  }, [period]);

  const revenue    = sales.reduce((s,x)=>s+Number(x.total),0);
  const cost       = items.reduce((s,x)=>s+Number(x.qty||0)*Number(x.buy_price||0),0);
  const profit     = revenue - cost;
  const byPay      = { cash:0, card:0, qr:0, debt:0 };
  sales.forEach(s=>{ byPay[s.payment_type]=(byPay[s.payment_type]||0)+Number(s.total); });
  const totalHandover = handovers.reduce((s,h)=>s+Number(h.amount),0);
  //const cashInRegister = byPay.cash - totalHandover;
  // Қарыз төлеўлерди алайық
const debtPayments = handovers.filter(h => h.comment?.startsWith("Қарыз төлеў")).reduce((s,h)=>s+Number(h.amount),0);
const realHandovers = totalHandover - debtPayments;
const cashInRegister = byPay.cash + debtPayments - realHandovers;

  const downloadStock = () => {
    const blob = exportStock(products);
    downloadExcel(blob, `қалдық_${today()}.xlsx`);
  };

  const closeDay = async () => {
    setClosing(true);
    await supabase.from("day_closings").insert({
      closed_by: profile.id, total_sales: revenue, total_cash: byPay.cash,
      total_card: byPay.card, total_qr: byPay.qr, total_debt: byPay.debt,
      total_profit: profit, date: today()
    });
    const blob = exportStock(products);
    await sendTelegramFile(blob, `қалдық_${today()}.xlsx`,
      `📊 Күн жабылды — ${today()}\n💰 Сатыў: ${fmt(revenue)}\n📈 Пайда: ${fmt(profit)}\n💵 Нақ: ${fmt(byPay.cash)}\n💳 Терминал: ${fmt(byPay.card)}\n📱 QR: ${fmt(byPay.qr)}\n📒 Қарыз: ${fmt(byPay.debt)}`
    );
    setClosing(false);
    setCloseMsg("✅ Күн жабылды! Telegram-ға жиберилди."); setTimeout(()=>setCloseMsg(""),4000);
  };

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
      {closeMsg && <Alert msg={closeMsg} />}
      <div style={{ display:"flex", gap:6 }}>
        {[["today","Бүгин"],["week","7 күн"],["month","30 күн"]].map(([k,l])=>(
          <button key={k} onClick={()=>setPeriod(k)}
            style={{ flex:1, padding:"8px 4px", background:period===k?"#f59e0b":"#1e293b", border:"none", borderRadius:8, color:period===k?"#0f172a":"#94a3b8", fontWeight:period===k?700:400, cursor:"pointer", fontSize:12 }}>
            {l}
          </button>
        ))}
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
        <Card icon="💰" label="Сатыў" value={fmt(revenue)} color="#10b981" />
        <Card icon="📈" label="Пайда" value={fmt(profit)} color="#f59e0b" />
        <Card icon="📦" label="Шығын (өзиндей)" value={fmt(cost)} color="#ef4444" />
        <Card icon="🛒" label="Сатыўлар саны" value={sales.length+" рет"} color="#3b82f6" />
        <Card icon="🏦" label="Кассадагы нақт" value={fmt(cashInRegister)} color="#10b981" />
        <Card icon="💸" label="Тапсырылған" value={fmt(totalHandover)} color="#8b5cf6" />
        <Card icon="📊" label="Маржа %" value={revenue ? ((profit/revenue)*100).toFixed(1)+"%" : "0%"} color="#14b8a6" />
      </div>
      <div style={{ background:"#1e293b", borderRadius:12, padding:12 }}>
        <div style={{ fontWeight:700, color:"#f59e0b", marginBottom:8, fontSize:13 }}>💳 Төлем түрлери</div>
        {Object.entries(byPay).map(([k,v])=>(
          <div key={k} style={{ display:"flex", justifyContent:"space-between", padding:"5px 0", borderBottom:"1px solid #0f172a", fontSize:13 }}>
            <span>{PAYMENT[k]}</span><span style={{ fontWeight:700 }}>{fmt(v)}</span>
          </div>
        ))}
      </div>
      {handovers.length > 0 && (
        <div style={{ background:"#1e293b", borderRadius:12, padding:12 }}>
          <div style={{ fontWeight:700, color:"#8b5cf6", marginBottom:8, fontSize:13 }}>💸 Тапсырыў тарийхы</div>
          {handovers.map(h => (
            <div key={h.id} style={{ padding:"6px 0", borderBottom:"1px solid #0f172a" }}>
              <div style={{ display:"flex", justifyContent:"space-between", fontSize:13 }}>
                <span style={{ color:"#94a3b8" }}>{h.date}</span>
                <span style={{ fontWeight:700, color:"#8b5cf6" }}>{fmt(h.amount)}</span>
              </div>
              {h.comment && <div style={{ fontSize:11, color:"#64748b" }}>💬 {h.comment}</div>}
            </div>
          ))}
        </div>
      )}
      <button onClick={downloadStock} style={{ padding:12, background:"#1e293b", border:"1px solid #334155", borderRadius:10, color:"#e2e8f0", cursor:"pointer", fontWeight:700, fontSize:13 }}>
        📥 Қалдық товар Excel жүклеў
      </button>
      {(profile.role === "director" || profile.role === "seller") && (
        <button onClick={closeDay} disabled={closing}
          style={{ padding:14, background:"#7c3aed", border:"none", borderRadius:10, color:"#fff", cursor:"pointer", fontWeight:700, fontSize:14 }}>
          {closing ? "Жабылып атыр…" : "🔒 Күнди жабыў → Telegram-ға жибериў"}
        </button>
      )}
    </div>
  );
}

// ─── SETTINGS ────────────────────────────────────────────────────
function Settings({ profile }) {
  const [users, setUsers]   = useState([]);
  const [email, setEmail]   = useState("");
  const [pass,  setPass]    = useState("");
  const [name,  setName]    = useState("");
  const [role,  setRole]    = useState("seller");
  const [msg,   setMsg]     = useState("");

  useEffect(() => {
    supabase.from("profiles").select("*").order("full_name").then(({ data }) => setUsers(data || []));
  }, []);

  const createUser = async () => {
    if (!email || !pass || !name) return;
    const { error } = await supabase.auth.admin.createUser({
      email, password: pass,
      user_metadata: { full_name: name, role },
      email_confirm: true
    });
    if (error) { setMsg("❌ " + error.message); return; }
    setEmail(""); setPass(""); setName(""); setRole("seller");
    setMsg("✅ Пайдаланыўшы қосылды!"); setTimeout(()=>setMsg(""),3000);
    supabase.from("profiles").select("*").order("full_name").then(({ data }) => setUsers(data || []));
  };

  if (profile.role !== "director") return (
    <div style={{ textAlign:"center", padding:40, color:"#64748b" }}>⛔ Тек директор ушын</div>
  );

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
      {msg && <Alert msg={msg} />}
      <div style={{ background:"#1e293b", borderRadius:12, padding:14 }}>
        <div style={{ fontWeight:700, color:"#f59e0b", marginBottom:10 }}>👤 Жаңа пайдаланыўшы қосыў</div>
        <Inp placeholder="Аты-жөни *" value={name} onChange={setName} />
        <Inp placeholder="Email *" value={email} onChange={setEmail} type="email" />
        <Inp placeholder="Парол *" value={pass} onChange={setPass} type="password" />
        <select value={role} onChange={e=>setRole(e.target.value)} style={inputStyle}>
          <option value="seller">Сатыўшы</option>
          <option value="supply">Снабженец</option>
          <option value="director">Директор</option>
        </select>
        <Btn label="Пайдаланыўшы қосыў" onClick={createUser} />
      </div>
      <div style={{ background:"#1e293b", borderRadius:12, padding:12 }}>
        <div style={{ fontWeight:700, color:"#f59e0b", marginBottom:8, fontSize:13 }}>👥 Пайдаланыўшылар</div>
        {users.map(u=>(
          <div key={u.id} style={{ display:"flex", justifyContent:"space-between", padding:"8px 0", borderBottom:"1px solid #0f172a", fontSize:13 }}>
            <span>{u.full_name}</span>
            <span style={{ background:u.role==="director"?"#7c3aed33":u.role==="supply"?"#3b82f633":"#10b98133",
              color:u.role==="director"?"#c4b5fd":u.role==="supply"?"#93c5fd":"#6ee7b7",
              borderRadius:6, padding:"2px 8px", fontSize:11 }}>
              {ROLES[u.role]}
            </span>
          </div>
        ))}
      </div>
    </div>
  );
}

// ─── SEARCHABLE PICKER ───────────────────────────────────────────
function SearchPicker({ products, value, onChange }) {
  const [search, setSearch] = useState("");
  const [open, setOpen]     = useState(false);
  const ref = useRef();

  const selected = products.find(p => p.id === +value);
  const filtered = search.trim()
    ? products.filter(p => p.name.toLowerCase().includes(search.toLowerCase()) || (p.barcode||"").includes(search)).slice(0,30)
    : products.slice(0,30);

  useEffect(() => {
    const h = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", h);
    return () => document.removeEventListener("mousedown", h);
  }, []);

  const pick = (p) => { onChange(String(p.id)); setSearch(""); setOpen(false); };

  const handleScan = useCallback((code) => {
    const p = products.find(x => x.barcode === code);
    if (p) { onChange(String(p.id)); setSearch(""); setOpen(false); }
  }, [products, onChange]);

  return (
    <div ref={ref} style={{ position:"relative", marginBottom:8 }}>
      <div style={{ display:"flex", gap:8 }}>
        <div style={{ flex:1, position:"relative" }}>
          <input placeholder={selected ? selected.name : "🔍 Товар аты ямаса штрих-код..."}
            value={search} onChange={e=>{setSearch(e.target.value);setOpen(true);}} onFocus={()=>setOpen(true)}
            style={{ width:"100%", ...inputStyle, marginBottom:0, border:`1px solid ${open?"#f59e0b":"#334155"}`, boxSizing:"border-box" }} />
          {selected && !search && (
            <button onClick={()=>{onChange("");setSearch("");}} style={{ position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",color:"#64748b",cursor:"pointer",fontSize:16 }}>✕</button>
          )}
        </div>
        <ScanBtn onScan={handleScan} label="📷" />
      </div>
      {open && filtered.length > 0 && (
        <div style={{ position:"absolute", top:"calc(100% + 4px)", left:0, right:50, background:"#1e293b",
          border:"1px solid #334155", borderRadius:10, zIndex:300, maxHeight:220, overflowY:"auto", boxShadow:"0 8px 24px rgba(0,0,0,0.6)" }}>
          {filtered.map(p=>(
            <div key={p.id} onClick={()=>pick(p)}
              style={{ padding:"10px 14px", cursor:"pointer", borderBottom:"1px solid #0f172a",
                background:p.id===+value?"#0f172a":"transparent", display:"flex", justifyContent:"space-between" }}>
              <div>
                <div style={{ fontSize:13, color:"#e2e8f0" }}>{p.name}</div>
                {p.barcode&&<div style={{ fontSize:10, color:"#475569" }}>🔢 {p.barcode}</div>}
              </div>
              <div style={{ textAlign:"right" }}>
                <div style={{ fontSize:12, color:"#10b981" }}>{fmt(p.sell_price)}</div>
                <div style={{ fontSize:10, color:p.stock<=p.min_stock?"#ef4444":"#64748b" }}>қалдық: {p.stock}</div>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ─── SCAN BUTTON ─────────────────────────────────────────────────
function ScanBtn({ onScan, label="📷" }) {
  const videoRef = useRef();
  const [open, setOpen] = useState(false);

  const start = async () => {
    setOpen(true);
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode:"environment" } });
      videoRef.current.srcObject = stream;
      videoRef.current.play();
      scan(stream);
    } catch { setOpen(false); alert("Камераға рухсат жоқ"); }
  };

  const scan = async (stream) => {
    if (!("BarcodeDetector" in window)) { stream.getTracks().forEach(t=>t.stop()); setOpen(false); alert("Браузериңиз сканерди қоллап-қуўатламайды"); return; }
    const detector = new window.BarcodeDetector({ formats:["ean_13","ean_8","code_128","code_39","upc_a","upc_e","qr_code"] });
    const interval = setInterval(async () => {
      if (!videoRef.current) { clearInterval(interval); return; }
      try {
        const codes = await detector.detect(videoRef.current);
        if (codes.length > 0) {
          clearInterval(interval);
          stream.getTracks().forEach(t=>t.stop());
          setOpen(false);
          onScan(codes[0].rawValue);
        }
      } catch {}
    }, 300);
  };

  const close = () => {
    if (videoRef.current?.srcObject) videoRef.current.srcObject.getTracks().forEach(t=>t.stop());
    setOpen(false);
  };

  return (
    <>
      <button onClick={start} style={{ padding:"10px 14px", background:"#1e293b", border:"1px solid #334155", borderRadius:8, color:"#e2e8f0", cursor:"pointer", fontSize:18, whiteSpace:"nowrap" }}>{label}</button>
      {open && (
        <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.9)", zIndex:999, display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center" }}>
          <div style={{ color:"#f59e0b", marginBottom:12, fontWeight:700 }}>Штрих-кодды сканерлеңиз</div>
          <video ref={videoRef} style={{ width:"90%", maxWidth:400, borderRadius:12, border:"2px solid #f59e0b" }} />
          <button onClick={close} style={{ marginTop:16, padding:"10px 32px", background:"#ef4444", border:"none", borderRadius:8, color:"#fff", fontWeight:700, cursor:"pointer" }}>Жабыў</button>
        </div>
      )}
    </>
  );
}

// ─── UI ATOMS ────────────────────────────────────────────────────
const inputStyle = {
  width:"100%", padding:"10px 12px", background:"#0f172a", border:"1px solid #334155",
  borderRadius:8, color:"#e2e8f0", marginBottom:8, fontSize:14, boxSizing:"border-box"
};

function btnStyle(color="#f59e0b", size) {
  return { padding: size==="small"?"7px 12px":"10px 18px", background:color, border:"none", borderRadius:8, color:color==="#f59e0b"?"#0f172a":"#fff", fontWeight:700, cursor:"pointer", fontSize:size==="small"?12:14 };
}

function Inp({ placeholder, value, onChange, type="text" }) {
  return <input type={type} placeholder={placeholder} value={value} onChange={e=>onChange(e.target.value)} style={inputStyle} />;
}
function Btn({ label, onClick, secondary, color, disabled }) {
  return (
    <button onClick={onClick} disabled={disabled}
      style={{ width:"100%", padding:12, background:secondary?"#1e293b":(color||"#f59e0b"), border:secondary?"1px solid #334155":"none",
        borderRadius:8, color:secondary?"#94a3b8":(color&&color!=="#f59e0b"?"#fff":"#0f172a"), fontWeight:700, cursor:"pointer", marginTop:6, fontSize:14, opacity:disabled?0.6:1 }}>
      {label}
    </button>
  );
}
function Card({ icon, label, value, color }) {
  return (
    <div style={{ background:"#1e293b", borderRadius:12, padding:12, borderLeft:`3px solid ${color||"#f59e0b"}` }}>
      <div style={{ fontSize:22 }}>{icon}</div>
      <div style={{ fontSize:11, color:"#64748b", marginTop:4 }}>{label}</div>
      <div style={{ fontWeight:700, color:color||"#e2e8f0", fontSize:14, marginTop:2 }}>{value}</div>
    </div>
  );
}
function Mini({ label, value, color }) {
  return (
    <div style={{ background:"#0f172a", borderRadius:8, padding:"6px 8px" }}>
      <div style={{ fontSize:10, color:"#64748b" }}>{label}</div>
      <div style={{ fontSize:12, fontWeight:600, color:color||"#e2e8f0", marginTop:2 }}>{value}</div>
    </div>
  );
}
function Alert({ msg }) {
  const ok = msg.startsWith("✅");
  return (
    <div style={{ background:ok?"#064e3b":"#7f1d1d", color:ok?"#6ee7b7":"#fca5a5", borderRadius:10, padding:12, textAlign:"center", fontSize:13, fontWeight:600 }}>
      {msg}
    </div>
  );
}
