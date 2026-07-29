import { useState, useEffect, useMemo } from "react";
import { Store, User, LogOut, Plus, TrendingUp, Package, ShieldCheck, ArrowLeft, Loader2, Lock, Users, Download } from "lucide-react";
import * as XLSX from "xlsx";
import { databases, DATABASE_ID, TABLE_ID } from "./appwriteClient";

const NAVY = "#0B3D56";
const ORANGE = "#C0501E";

function naira(n) {
  return "₦" + Number(n || 0).toLocaleString("en-NG", { maximumFractionDigits: 0 });
}
function uid() {
  return Math.random().toString(36).slice(2, 10);
}
// Random 4-digit numeric code (for PINs / shop portal codes).
function genPin() {
  return String(Math.floor(1000 + Math.random() * 9000));
}
// Random recovery code, avoiding ambiguous characters (0/O, 1/I) — shown once
// at account creation, and the only way to reset a forgotten master code or
// PIN afterward, since we never store the original in a recoverable form.
function genRecoveryCode() {
  const chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  let s = "";
  for (let i = 0; i < 10; i++) s += chars[Math.floor(Math.random() * chars.length)];
  return s.slice(0, 5) + "-" + s.slice(5);
}
// Builds a real .xlsx workbook client-side (via SheetJS) and triggers a
// browser download — no server round-trip needed.
function exportToExcel(filename, sheets) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(({ name, rows }) => {
    const ws = XLSX.utils.json_to_sheet(rows.length ? rows : [{ "No data": "No sales recorded yet" }]);
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31)); // Excel sheet names cap at 31 chars
  });
  XLSX.writeFile(wb, filename);
}
function saleRow(s, opts = {}) {
  const row = { Date: s.date, Item: s.product, Category: s.category };
  if (opts.shop) row.Shop = opts.shopName;
  if (opts.worker) row.Worker = s.workerName;
  row["Qty"] = s.qty;
  row["Unit Price (₦)"] = s.unitPrice;
  row["Cost/Item (₦)"] = s.cost;
  row["Total (₦)"] = s.total;
  row["Profit (₦)"] = s.profit;
  row["Payment Method"] = s.paymentMethod;
  return row;
}
// SHA-256 hash via the browser's built-in Web Crypto API — no secret ever
// touches storage or the source code in plaintext.
async function hashSecret(value) {
  const enc = new TextEncoder().encode(String(value));
  const buf = await crypto.subtle.digest("SHA-256", enc);
  return Array.from(new Uint8Array(buf)).map((b) => b.toString(16).padStart(2, "0")).join("");
}
// Each key ("shops", "users", "sales", "solo_businesses", "boss_code_hash")
// is stored as its own document, using the key itself as the document ID,
// with the JSON payload stringified into a single "value" text field.
async function loadList(key, fallback) {
  try {
    const doc = await databases.getDocument(DATABASE_ID, TABLE_ID, key);
    return JSON.parse(doc.value);
  } catch (e) {
    // 404 just means this key hasn't been saved yet — that's expected on first run.
    if (e?.code !== 404) console.error("storage load failed", key, e);
    return fallback;
  }
}
async function saveList(key, value) {
  const payload = { value: JSON.stringify(value) };
  try {
    await databases.createDocument(DATABASE_ID, TABLE_ID, key, payload);
  } catch (createErr) {
    try {
      await databases.updateDocument(DATABASE_ID, TABLE_ID, key, payload);
    } catch (updateErr) {
      console.error("storage save failed", key, updateErr);
    }
  }
}

export default function RetailTrackerApp() {
  const [loading, setLoading] = useState(true);
  const [shops, setShops] = useState([]);
  const [users, setUsers] = useState([]);
  const [sales, setSales] = useState([]);
  const [soloBiz, setSoloBiz] = useState([]);
  const [bossCodeHash, setBossCodeHash] = useState(null);
  const [bossRecoveryHash, setBossRecoveryHash] = useState(null);
  const [screen, setScreen] = useState("mode");
  const [currentUser, setCurrentUser] = useState(null);
  const [activeShopId, setActiveShopId] = useState(null);

  useEffect(() => {
    (async () => {
      const [s, u, sl, sb, bch, brh] = await Promise.all([
        loadList("shops", []),
        loadList("users", []),
        loadList("sales", []),
        loadList("solo_businesses", []),
        loadList("boss_code_hash", null),
        loadList("boss_recovery_hash", null),
      ]);
      setShops(s);
      setUsers(u);
      setSales(sl);
      setSoloBiz(sb);
      setBossCodeHash(bch);
      setBossRecoveryHash(brh);
      setLoading(false);
    })();
  }, []);

  function goHome() {
    setCurrentUser(null);
    setActiveShopId(null);
    setScreen("mode");
  }

  if (loading) {
    return (
      <div className="min-h-[500px] flex items-center justify-center bg-slate-50 text-slate-500">
        <Loader2 className="w-5 h-5 animate-spin mr-2" /> Loading tracker…
      </div>
    );
  }

  return (
    <div className="min-h-[640px] bg-slate-50 font-sans relative">
      {screen === "mode" && (
        <ModeSelect onPick={(m) => setScreen(m === "solo" ? "solo-login" : "role")} />
      )}

      {/* -------- SOLO MODE -------- */}
      {screen === "solo-login" && (
        <SoloLogin
          soloBiz={soloBiz}
          onBack={goHome}
          onSuccess={(b) => { setCurrentUser(b); setScreen("solo-dash"); }}
          onCreate={async (b) => {
            const next = [...soloBiz, b];
            setSoloBiz(next);
            await saveList("solo_businesses", next);
          }}
          onResetPin={async (name, recoveryCode, newPin) => {
            const biz = soloBiz.find((x) => x.name.trim().toLowerCase() === name.trim().toLowerCase());
            if (!biz) return { ok: false, reason: "notfound" };
            const ok = (await hashSecret(recoveryCode)) === biz.recoveryHash;
            if (!ok) return { ok: false, reason: "code" };
            const pinHash = await hashSecret(newPin);
            const updated = { ...biz, pinHash };
            const next = soloBiz.map((x) => (x.id === biz.id ? updated : x));
            setSoloBiz(next);
            await saveList("solo_businesses", next);
            return { ok: true, business: updated };
          }}
        />
      )}
      {screen === "solo-dash" && currentUser && (
        <SoloDashboard
          business={currentUser}
          sales={sales}
          onLogout={goHome}
          onAddSale={async (sale) => {
            const next = [...sales, sale];
            setSales(next);
            await saveList("sales", next);
          }}
        />
      )}

      {/* -------- TEAM MODE -------- */}
      {screen === "role" && (
        <RoleSelect onBack={goHome} onPick={(r) => setScreen(r === "boss" ? "boss-login" : "shop-select")} />
      )}
      {screen === "shop-select" && (
        <ShopSelect
          shops={shops}
          onBack={() => setScreen("role")}
          onPick={(id) => { setActiveShopId(id); setScreen("shop-code"); }}
        />
      )}
      {screen === "shop-code" && (
        <ShopCodeGate
          shop={shops.find((s) => s.id === activeShopId)}
          onBack={() => setScreen("shop-select")}
          onSuccess={() => setScreen("worker-login")}
        />
      )}
      {screen === "worker-login" && (
        <WorkerLogin
          shop={shops.find((s) => s.id === activeShopId)}
          users={users.filter((u) => u.shopId === activeShopId)}
          onBack={() => setScreen("shop-code")}
          onSuccess={(u) => { setCurrentUser(u); setScreen("worker-dash"); }}
          onCreateUser={async (u) => {
            const next = [...users, u];
            setUsers(next);
            await saveList("users", next);
          }}
        />
      )}
      {screen === "worker-dash" && currentUser && (
        <WorkerDashboard
          user={currentUser}
          shop={shops.find((s) => s.id === activeShopId)}
          sales={sales}
          onLogout={goHome}
          onAddSale={async (sale) => {
            const next = [...sales, sale];
            setSales(next);
            await saveList("sales", next);
          }}
        />
      )}
      {screen === "boss-login" && (
        <BossLogin
          hasCode={!!bossCodeHash}
          onBack={() => setScreen("role")}
          onSuccess={() => { setCurrentUser({ role: "boss" }); setScreen("boss-dash"); }}
          onVerify={async (code) => (await hashSecret(code)) === bossCodeHash}
          onCreateCode={async (code, recoveryCode) => {
            const h = await hashSecret(code);
            const rh = await hashSecret(recoveryCode);
            setBossCodeHash(h);
            setBossRecoveryHash(rh);
            await saveList("boss_code_hash", h);
            await saveList("boss_recovery_hash", rh);
          }}
          onResetCode={async (recoveryCode, newCode) => {
            const ok = (await hashSecret(recoveryCode)) === bossRecoveryHash;
            if (!ok) return false;
            const h = await hashSecret(newCode);
            setBossCodeHash(h);
            await saveList("boss_code_hash", h);
            return true;
          }}
        />
      )}
      {screen === "boss-dash" && currentUser && (
        <BossDashboard
          shops={shops}
          users={users}
          sales={sales}
          onLogout={goHome}
          onUpdateShopCode={async (shopId, code) => {
            const hash = await hashSecret(code);
            const next = shops.map((s) => (s.id === shopId ? { ...s, portalCodeHash: hash } : s));
            setShops(next);
            await saveList("shops", next);
          }}
          onAddShop={async (name, code) => {
            const hash = await hashSecret(code);
            const next = [...shops, { id: uid(), name: name.trim(), portalCodeHash: hash }];
            setShops(next);
            await saveList("shops", next);
          }}
          onRenameShop={async (shopId, name) => {
            const next = shops.map((s) => (s.id === shopId ? { ...s, name: name.trim() } : s));
            setShops(next);
            await saveList("shops", next);
          }}
          onRemoveShop={async (shopId) => {
            const next = shops.filter((s) => s.id !== shopId);
            setShops(next);
            await saveList("shops", next);
          }}
          onResetWorkerPin={async (userId, newPin) => {
            const pinHash = await hashSecret(newPin);
            const next = users.map((u) => (u.id === userId ? { ...u, pinHash } : u));
            setUsers(next);
            await saveList("users", next);
          }}
        />
      )}
    </div>
  );
}

// ============================================================== MODE SELECT
function ModeSelect({ onPick }) {
  return (
    <div className="min-h-[640px] flex flex-col items-center justify-center px-6 py-16">
      <div className="w-14 h-14 rounded-2xl flex items-center justify-center mb-5" style={{ background: NAVY }}>
        <Store className="w-7 h-7 text-white" />
      </div>
      <h1 className="text-2xl font-bold text-slate-900 mb-1">DigitalBusayomi Tracker</h1>
      <p className="text-slate-500 text-sm mb-10">How is your business set up?</p>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 w-full max-w-lg">
        <button onClick={() => onPick("solo")} className="bg-white border border-slate-200 rounded-2xl p-6 text-left hover:border-slate-900 transition-colors shadow-sm">
          <User className="w-6 h-6 mb-3" style={{ color: NAVY }} />
          <div className="font-semibold text-slate-900">Just Me — One Shop</div>
          <div className="text-xs text-slate-500 mt-1">I run a single shop by myself and want my own tracker</div>
        </button>
        <button onClick={() => onPick("team")} className="bg-white border border-slate-200 rounded-2xl p-6 text-left hover:border-slate-900 transition-colors shadow-sm">
          <Users className="w-6 h-6 mb-3" style={{ color: ORANGE }} />
          <div className="font-semibold text-slate-900">My Team — Boss &amp; Workers</div>
          <div className="text-xs text-slate-500 mt-1">Multiple shops, a boss, and workers who log their own sales</div>
        </button>
      </div>
      <p className="text-[11px] text-slate-400 mt-10 max-w-sm text-center">
        Demo prototype — data is stored for anyone who opens this app link. A production version would use real accounts and a secure database.
      </p>
    </div>
  );
}

// ============================================================== SOLO MODE
function SoloLogin({ soloBiz, onBack, onSuccess, onCreate, onResetPin }) {
  const [mode, setMode] = useState("existing"); // existing | new | forgot
  const [name, setName] = useState("");
  const [pin, setPin] = useState("");
  const [recoveryInput, setRecoveryInput] = useState("");
  const [newPin, setNewPin] = useState("");
  const [err, setErr] = useState("");
  const [checking, setChecking] = useState(false);
  const [revealCode, setRevealCode] = useState(null); // { code, business } shown once after creation
  const [resetDone, setResetDone] = useState(false);

  async function handleExisting() {
    if (!name.trim()) return setErr("Enter your shop/business name.");
    setChecking(true);
    const b = soloBiz.find((x) => x.name.trim().toLowerCase() === name.trim().toLowerCase());
    if (!b) { setChecking(false); return setErr("No business found with that exact name."); }
    const enteredHash = await hashSecret(pin);
    setChecking(false);
    if (b.pinHash !== enteredHash) return setErr("Incorrect PIN.");
    onSuccess(b);
  }
  async function handleCreate() {
    if (!name.trim()) return setErr("Enter your shop/business name.");
    if (soloBiz.some((x) => x.name.trim().toLowerCase() === name.trim().toLowerCase())) {
      return setErr("That business name is already taken — try logging in instead, or pick another name.");
    }
    if (pin.length !== 4) return setErr("Choose a 4-digit PIN.");
    setChecking(true);
    const pinHash = await hashSecret(pin);
    const recoveryCode = genRecoveryCode();
    const recoveryHash = await hashSecret(recoveryCode);
    const b = { id: uid(), name: name.trim(), pinHash, recoveryHash };
    setChecking(false);
    await onCreate(b);
    setRevealCode({ code: recoveryCode, business: b });
  }
  async function handleForgotSubmit() {
    if (!name.trim() || !recoveryInput.trim()) return setErr("Fill in your business name and recovery code.");
    if (newPin.length !== 4) return setErr("Choose a new 4-digit PIN.");
    setChecking(true);
    const result = await onResetPin(name, recoveryInput.trim().toUpperCase(), newPin);
    setChecking(false);
    if (!result.ok) {
      return setErr(result.reason === "notfound" ? "No business found with that exact name." : "Incorrect recovery code.");
    }
    setResetDone(true);
  }

  // One-time reveal screen right after creating a new business
  if (revealCode) {
    return (
      <div className="min-h-[640px] flex flex-col items-center justify-center px-6">
        <ShieldCheck className="w-8 h-8 mb-3" style={{ color: NAVY }} />
        <h2 className="text-lg font-semibold text-slate-900 mb-1">Save your recovery code</h2>
        <p className="text-sm text-slate-500 mb-4 max-w-xs text-center">
          If you ever forget your PIN, this code is the only way back in. It won't be shown again.
        </p>
        <div className="bg-slate-100 rounded-lg px-6 py-3 text-xl font-mono font-bold tracking-wider text-slate-900 mb-6">
          {revealCode.code}
        </div>
        <button
          onClick={() => onSuccess(revealCode.business)}
          className="w-full max-w-xs text-white rounded-lg py-2 text-sm font-medium"
          style={{ background: NAVY }}
        >
          I've saved it — Continue
        </button>
      </div>
    );
  }

  if (mode === "forgot") {
    if (resetDone) {
      return (
        <div className="min-h-[640px] flex flex-col items-center justify-center px-6">
          <ShieldCheck className="w-8 h-8 mb-3" style={{ color: NAVY }} />
          <h2 className="text-lg font-semibold text-slate-900 mb-1">PIN reset</h2>
          <p className="text-sm text-slate-500 mb-6">Your PIN has been changed. You can log in with it now.</p>
          <button onClick={() => { setMode("existing"); setResetDone(false); setPin(""); }} className="w-full max-w-xs text-white rounded-lg py-2 text-sm font-medium" style={{ background: NAVY }}>Back to Login</button>
        </div>
      );
    }
    return (
      <div className="min-h-[640px] flex flex-col items-center justify-center px-6">
        <button onClick={() => { setMode("existing"); setErr(""); }} className="absolute top-6 left-6 flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800"><ArrowLeft className="w-4 h-4" /> Back</button>
        <Lock className="w-8 h-8 mb-3" style={{ color: ORANGE }} />
        <h2 className="text-lg font-semibold text-slate-900 mb-1">Reset your PIN</h2>
        <p className="text-sm text-slate-500 mb-6">Enter your business name and recovery code</p>
        <div className="w-full max-w-xs space-y-3">
          <input placeholder="Your exact shop / business name" value={name} onChange={(e) => { setName(e.target.value); setErr(""); }} className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm" />
          <input placeholder="Recovery code (e.g. AB2C3-4DEF6)" value={recoveryInput} onChange={(e) => { setRecoveryInput(e.target.value); setErr(""); }} className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm uppercase" />
          <input type="password" inputMode="numeric" maxLength={4} placeholder="New 4-digit PIN" value={newPin} onChange={(e) => setNewPin(e.target.value.replace(/\D/g, ""))} className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm" />
          {err && <p className="text-xs text-red-600">{err}</p>}
          <button disabled={checking} onClick={handleForgotSubmit} className="w-full text-white rounded-lg py-2 text-sm font-medium disabled:opacity-60" style={{ background: ORANGE }}>{checking ? "Checking…" : "Reset PIN"}</button>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-[640px] flex flex-col items-center justify-center px-6">
      <button onClick={onBack} className="absolute top-6 left-6 flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800">
        <ArrowLeft className="w-4 h-4" /> Back
      </button>
      <User className="w-8 h-8 mb-3" style={{ color: NAVY }} />
      <h2 className="text-lg font-semibold text-slate-900 mb-1">Your Shop</h2>
      <p className="text-sm text-slate-500 mb-6">Just you — no boss/worker split needed</p>

      <div className="w-full max-w-xs space-y-3">
        <div className="flex rounded-lg bg-slate-100 p-1 text-sm">
          <button onClick={() => { setMode("existing"); setErr(""); }} className={`flex-1 rounded-md py-1.5 ${mode === "existing" ? "bg-white shadow font-medium" : "text-slate-500"}`}>I have an account</button>
          <button onClick={() => { setMode("new"); setErr(""); }} className={`flex-1 rounded-md py-1.5 ${mode === "new" ? "bg-white shadow font-medium" : "text-slate-500"}`}>New shop</button>
        </div>

        {mode === "existing" ? (
          <>
            <input placeholder="Your exact shop / business name" value={name} onChange={(e) => { setName(e.target.value); setErr(""); }} className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm" />
            <input type="password" inputMode="numeric" maxLength={4} placeholder="4-digit PIN" value={pin} onChange={(e) => setPin(e.target.value.replace(/\D/g, ""))} className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm" />
            {err && <p className="text-xs text-red-600">{err}</p>}
            <button disabled={checking} onClick={handleExisting} className="w-full text-white rounded-lg py-2 text-sm font-medium disabled:opacity-60" style={{ background: NAVY }}>{checking ? "Checking…" : "Log In"}</button>
            <button onClick={() => { setMode("forgot"); setErr(""); }} className="w-full text-xs text-slate-400 hover:text-slate-700 text-center">Forgot PIN?</button>
          </>
        ) : (
          <>
            <input placeholder="Shop / business name" value={name} onChange={(e) => { setName(e.target.value); setErr(""); }} className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm" />
            <input type="password" inputMode="numeric" maxLength={4} placeholder="Create a 4-digit PIN" value={pin} onChange={(e) => setPin(e.target.value.replace(/\D/g, ""))} className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm" />
            {err && <p className="text-xs text-red-600">{err}</p>}
            <button disabled={checking} onClick={handleCreate} className="w-full text-white rounded-lg py-2 text-sm font-medium disabled:opacity-60" style={{ background: NAVY }}>{checking ? "Creating…" : "Create & Log In"}</button>
          </>
        )}
      </div>
    </div>
  );
}

function SoloDashboard({ business, sales, onLogout, onAddSale }) {
  const mySales = useMemo(
    () => sales.filter((s) => s.kind === "solo" && s.businessId === business.id).sort((a, b) => (a.date < b.date ? 1 : -1)),
    [sales, business.id]
  );
  const [form, setForm] = useState({ date: new Date().toISOString().slice(0, 10), product: "", category: "", qty: "", unitPrice: "", cost: "", paymentMethod: "Cash" });

  const totalToday = mySales.filter((s) => s.date === form.date).reduce((a, s) => a + s.total, 0);
  const totalAll = mySales.reduce((a, s) => a + s.total, 0);
  const profitAll = mySales.reduce((a, s) => a + s.profit, 0);

  function submit(e) {
    e.preventDefault();
    const qty = Number(form.qty) || 0;
    const unitPrice = Number(form.unitPrice) || 0;
    const cost = Number(form.cost) || 0;
    if (!form.product || qty <= 0 || unitPrice <= 0) return;
    const total = qty * unitPrice;
    const profit = total - cost * qty;
    onAddSale({
      id: uid(), kind: "solo", businessId: business.id, workerId: business.id, workerName: business.name,
      date: form.date, product: form.product, category: form.category || "General",
      qty, unitPrice, cost, paymentMethod: form.paymentMethod, total, profit,
    });
    setForm((f) => ({ ...f, product: "", category: "", qty: "", unitPrice: "", cost: "" }));
  }

  return (
    <div className="max-w-3xl mx-auto px-5 py-6">
      <div className="flex items-center justify-between mb-6">
        <div>
          <p className="text-xs text-slate-400">Solo shop</p>
          <h1 className="text-lg font-bold text-slate-900">{business.name}</h1>
        </div>
        <button onClick={onLogout} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800"><LogOut className="w-4 h-4" /> Log out</button>
      </div>

      <button
        onClick={() => exportToExcel(`${business.name}-sales.xlsx`, [{ name: "Sales", rows: mySales.map((s) => saleRow(s)) }])}
        className="mb-4 flex items-center gap-1.5 text-sm font-medium border border-slate-300 rounded-lg px-3 py-1.5 text-slate-700 hover:bg-slate-50"
      >
        <Download className="w-4 h-4" /> Download Excel
      </button>

      <div className="grid grid-cols-3 gap-3 mb-6">
        <div className="bg-white border border-slate-200 rounded-xl p-4">
          <p className="text-xs text-slate-400">Today</p>
          <p className="text-lg font-bold text-slate-900">{naira(totalToday)}</p>
        </div>
        <div className="bg-white border border-slate-200 rounded-xl p-4">
          <p className="text-xs text-slate-400">All-time Revenue</p>
          <p className="text-lg font-bold" style={{ color: NAVY }}>{naira(totalAll)}</p>
        </div>
        <div className="bg-white border border-slate-200 rounded-xl p-4">
          <p className="text-xs text-slate-400">All-time Profit</p>
          <p className="text-lg font-bold" style={{ color: ORANGE }}>{naira(profitAll)}</p>
        </div>
      </div>

      <form onSubmit={submit} className="bg-white border border-slate-200 rounded-xl p-4 mb-6">
        <p className="text-sm font-semibold text-slate-800 mb-3 flex items-center gap-1"><Plus className="w-4 h-4" /> Log a sale</p>
        <div className="grid grid-cols-2 gap-2 mb-2">
          <input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm col-span-2 sm:col-span-1" />
          <select value={form.paymentMethod} onChange={(e) => setForm({ ...form, paymentMethod: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm">
            <option>Cash</option><option>Transfer</option><option>POS</option>
          </select>
          <input placeholder="Product / item" value={form.product} onChange={(e) => setForm({ ...form, product: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm col-span-2" />
          <input placeholder="Category" value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm col-span-2" />
          <input type="number" placeholder="Qty sold" value={form.qty} onChange={(e) => setForm({ ...form, qty: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm" />
          <input type="number" placeholder="Unit price (₦)" value={form.unitPrice} onChange={(e) => setForm({ ...form, unitPrice: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm" />
          <input type="number" placeholder="Cost per item (₦)" value={form.cost} onChange={(e) => setForm({ ...form, cost: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm col-span-2" />
        </div>
        <button type="submit" className="w-full text-white rounded-lg py-2 text-sm font-medium mt-1" style={{ background: NAVY }}>Add Sale</button>
      </form>

      <p className="text-sm font-semibold text-slate-800 mb-2">Sales history</p>
      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-slate-500 text-xs">
            <tr><th className="text-left px-3 py-2">Date</th><th className="text-left px-3 py-2">Item</th><th className="text-right px-3 py-2">Qty</th><th className="text-right px-3 py-2">Total</th></tr>
          </thead>
          <tbody>
            {mySales.length === 0 && <tr><td colSpan={4} className="text-center text-slate-400 py-6">No sales logged yet.</td></tr>}
            {mySales.map((s) => (
              <tr key={s.id} className="border-t border-slate-100">
                <td className="px-3 py-2">{s.date}</td>
                <td className="px-3 py-2">{s.product}</td>
                <td className="px-3 py-2 text-right">{s.qty}</td>
                <td className="px-3 py-2 text-right font-medium">{naira(s.total)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ============================================================== TEAM MODE — ROLE / SHOP GATES
function RoleSelect({ onBack, onPick }) {
  return (
    <div className="min-h-[640px] flex flex-col items-center justify-center px-6 py-16">
      <button onClick={onBack} className="absolute top-6 left-6 flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800"><ArrowLeft className="w-4 h-4" /> Back</button>
      <div className="w-14 h-14 rounded-2xl flex items-center justify-center mb-5" style={{ background: NAVY }}>
        <Users className="w-7 h-7 text-white" />
      </div>
      <h1 className="text-xl font-bold text-slate-900 mb-1">Team Tracker</h1>
      <p className="text-slate-500 text-sm mb-10">Who's logging in?</p>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 w-full max-w-md">
        <button onClick={() => onPick("boss")} className="bg-white border border-slate-200 rounded-2xl p-6 text-left hover:border-slate-900 transition-colors shadow-sm">
          <ShieldCheck className="w-6 h-6 mb-3" style={{ color: NAVY }} />
          <div className="font-semibold text-slate-900">I'm the Boss</div>
          <div className="text-xs text-slate-500 mt-1">See every shop, every sale</div>
        </button>
        <button onClick={() => onPick("worker")} className="bg-white border border-slate-200 rounded-2xl p-6 text-left hover:border-slate-900 transition-colors shadow-sm">
          <User className="w-6 h-6 mb-3" style={{ color: ORANGE }} />
          <div className="font-semibold text-slate-900">I'm a Worker</div>
          <div className="text-xs text-slate-500 mt-1">Log sales, see only your own</div>
        </button>
      </div>
    </div>
  );
}

function ShopSelect({ shops, onBack, onPick }) {
  return (
    <div className="min-h-[640px] flex flex-col items-center justify-center px-6">
      <button onClick={onBack} className="absolute top-6 left-6 flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800"><ArrowLeft className="w-4 h-4" /> Back</button>
      <Store className="w-8 h-8 mb-3" style={{ color: ORANGE }} />
      <h2 className="text-lg font-semibold text-slate-900 mb-1">Which shop do you work at?</h2>
      <p className="text-sm text-slate-500 mb-6">You'll need your shop's portal code next</p>
      <div className="grid grid-cols-1 gap-3 w-full max-w-xs">
        {shops.map((s) => (
          <button key={s.id} onClick={() => onPick(s.id)} className="bg-white border border-slate-200 rounded-xl px-4 py-3 text-left hover:border-slate-900 transition-colors shadow-sm font-medium text-slate-800">
            {s.name}
          </button>
        ))}
      </div>
    </div>
  );
}

function ShopCodeGate({ shop, onBack, onSuccess }) {
  const [code, setCode] = useState("");
  const [err, setErr] = useState("");
  const [checking, setChecking] = useState(false);

  async function submit() {
    setChecking(true);
    const enteredHash = await hashSecret(code);
    setChecking(false);
    if (enteredHash === shop?.portalCodeHash) onSuccess();
    else setErr("Incorrect portal code for this shop.");
  }

  return (
    <div className="min-h-[640px] flex flex-col items-center justify-center px-6">
      <button onClick={onBack} className="absolute top-6 left-6 flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800"><ArrowLeft className="w-4 h-4" /> Back</button>
      <Lock className="w-8 h-8 mb-3" style={{ color: ORANGE }} />
      <h2 className="text-lg font-semibold text-slate-900 mb-1">{shop?.name}</h2>
      <p className="text-sm text-slate-500 mb-6">Enter this shop's portal code</p>
      <div className="w-full max-w-xs">
        <input type="password" inputMode="numeric" value={code} onChange={(e) => { setCode(e.target.value); setErr(""); }} placeholder="Shop portal code" className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm mb-3" />
        {err && <p className="text-xs text-red-600 mb-3">{err}</p>}
        <button
          disabled={checking}
          onClick={submit}
          className="w-full text-white rounded-lg py-2 text-sm font-medium disabled:opacity-60" style={{ background: ORANGE }}
        >
          {checking ? "Checking…" : "Unlock Shop Portal"}
        </button>
      </div>
    </div>
  );
}

function BossLogin({ hasCode, onBack, onSuccess, onVerify, onCreateCode, onResetCode }) {
  const [mode, setMode] = useState("login"); // login | forgot
  const [code, setCode] = useState("");
  const [confirmCode, setConfirmCode] = useState("");
  const [recoveryInput, setRecoveryInput] = useState("");
  const [newCode, setNewCode] = useState("");
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  const [revealCode, setRevealCode] = useState(null); // shown once right after creating the master code
  const [resetDone, setResetDone] = useState(false);

  async function submit() {
    setErr("");
    if (!hasCode) {
      if (code.length < 4) return setErr("Choose a master code with at least 4 characters.");
      if (code !== confirmCode) return setErr("Codes don't match — enter the same code twice.");
      setBusy(true);
      const recoveryCode = genRecoveryCode();
      await onCreateCode(code, recoveryCode);
      setBusy(false);
      setRevealCode(recoveryCode);
      return;
    }
    setBusy(true);
    const ok = await onVerify(code);
    setBusy(false);
    if (ok) onSuccess();
    else setErr("Incorrect code.");
  }
  async function submitReset() {
    if (!recoveryInput.trim()) return setErr("Enter your recovery code.");
    if (newCode.length < 4) return setErr("Choose a new master code with at least 4 characters.");
    setBusy(true);
    const ok = await onResetCode(recoveryInput.trim().toUpperCase(), newCode);
    setBusy(false);
    if (!ok) return setErr("Incorrect recovery code.");
    setResetDone(true);
  }

  if (revealCode) {
    return (
      <div className="min-h-[640px] flex flex-col items-center justify-center px-6">
        <ShieldCheck className="w-8 h-8 mb-3" style={{ color: NAVY }} />
        <h2 className="text-lg font-semibold text-slate-900 mb-1">Save your recovery code</h2>
        <p className="text-sm text-slate-500 mb-4 max-w-xs text-center">
          If you ever forget your master code, this is the only way back into the boss dashboard. It won't be shown again.
        </p>
        <div className="bg-slate-100 rounded-lg px-6 py-3 text-xl font-mono font-bold tracking-wider text-slate-900 mb-6">
          {revealCode}
        </div>
        <button onClick={onSuccess} className="w-full max-w-xs text-white rounded-lg py-2 text-sm font-medium" style={{ background: NAVY }}>
          I've saved it — Continue
        </button>
      </div>
    );
  }

  if (mode === "forgot") {
    if (resetDone) {
      return (
        <div className="min-h-[640px] flex flex-col items-center justify-center px-6">
          <ShieldCheck className="w-8 h-8 mb-3" style={{ color: NAVY }} />
          <h2 className="text-lg font-semibold text-slate-900 mb-1">Master code reset</h2>
          <p className="text-sm text-slate-500 mb-6">Your master code has been changed. You can log in with it now.</p>
          <button onClick={() => { setMode("login"); setResetDone(false); setCode(""); }} className="w-full max-w-xs text-white rounded-lg py-2 text-sm font-medium" style={{ background: NAVY }}>Back to Login</button>
        </div>
      );
    }
    return (
      <div className="min-h-[640px] flex flex-col items-center justify-center px-6">
        <button onClick={() => { setMode("login"); setErr(""); }} className="absolute top-6 left-6 flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800"><ArrowLeft className="w-4 h-4" /> Back</button>
        <Lock className="w-8 h-8 mb-3" style={{ color: ORANGE }} />
        <h2 className="text-lg font-semibold text-slate-900 mb-1">Reset master code</h2>
        <p className="text-sm text-slate-500 mb-6">Enter the recovery code you saved when you first set up boss access</p>
        <div className="w-full max-w-xs space-y-3">
          <input placeholder="Recovery code (e.g. AB2C3-4DEF6)" value={recoveryInput} onChange={(e) => { setRecoveryInput(e.target.value); setErr(""); }} className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm uppercase" />
          <input type="password" placeholder="New master code" value={newCode} onChange={(e) => setNewCode(e.target.value)} className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm" />
          {err && <p className="text-xs text-red-600">{err}</p>}
          <button disabled={busy} onClick={submitReset} className="w-full text-white rounded-lg py-2 text-sm font-medium disabled:opacity-60" style={{ background: ORANGE }}>{busy ? "Checking…" : "Reset Master Code"}</button>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-[640px] flex flex-col items-center justify-center px-6">
      <button onClick={onBack} className="absolute top-6 left-6 flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800"><ArrowLeft className="w-4 h-4" /> Back</button>
      <ShieldCheck className="w-8 h-8 mb-3" style={{ color: NAVY }} />
      <h2 className="text-lg font-semibold text-slate-900 mb-1">Boss Access</h2>
      <p className="text-sm text-slate-500 mb-6">{hasCode ? "Enter the master secret code" : "No master code is set yet — create one now"}</p>
      <div className="w-full max-w-xs">
        <input type="password" value={code} onChange={(e) => { setCode(e.target.value); setErr(""); }} placeholder={hasCode ? "Secret code" : "Create a master code"} className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm mb-3" />
        {!hasCode && (
          <input type="password" value={confirmCode} onChange={(e) => { setConfirmCode(e.target.value); setErr(""); }} placeholder="Confirm master code" className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm mb-3" />
        )}
        {err && <p className="text-xs text-red-600 mb-3">{err}</p>}
        <button disabled={busy} onClick={submit} className="w-full text-white rounded-lg py-2 text-sm font-medium disabled:opacity-60" style={{ background: NAVY }}>
          {busy ? "Please wait…" : hasCode ? "Enter Dashboard" : "Create Code & Enter Dashboard"}
        </button>
        {hasCode && (
          <button onClick={() => { setMode("forgot"); setErr(""); }} className="w-full text-xs text-slate-400 hover:text-slate-700 text-center mt-3">Forgot code?</button>
        )}
      </div>
    </div>
  );
}

function WorkerLogin({ shop, users, onBack, onSuccess, onCreateUser }) {
  const [mode, setMode] = useState("existing");
  const [selectedUserId, setSelectedUserId] = useState("");
  const [name, setName] = useState("");
  const [pin, setPin] = useState("");
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);

  async function handleExistingLogin() {
    const u = users.find((w) => w.id === selectedUserId);
    if (!u) return setErr("Select your name.");
    setBusy(true);
    const enteredHash = await hashSecret(pin);
    setBusy(false);
    if (u.pinHash !== enteredHash) return setErr("Incorrect PIN.");
    onSuccess(u);
  }
  async function handleCreate() {
    if (!name.trim()) return setErr("Enter your name.");
    if (pin.length !== 4) return setErr("Choose a 4-digit PIN.");
    setBusy(true);
    const pinHash = await hashSecret(pin);
    const newUser = { id: uid(), name: name.trim(), role: "worker", shopId: shop.id, pinHash };
    setBusy(false);
    await onCreateUser(newUser);
    onSuccess(newUser);
  }

  return (
    <div className="min-h-[640px] flex flex-col items-center justify-center px-6">
      <button onClick={onBack} className="absolute top-6 left-6 flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800"><ArrowLeft className="w-4 h-4" /> Back</button>
      <User className="w-8 h-8 mb-3" style={{ color: ORANGE }} />
      <h2 className="text-lg font-semibold text-slate-900 mb-1">{shop?.name}</h2>
      <p className="text-sm text-slate-500 mb-6">Now log in as yourself</p>

      <div className="w-full max-w-xs space-y-3">
        <div className="flex rounded-lg bg-slate-100 p-1 text-sm">
          <button onClick={() => { setMode("existing"); setErr(""); }} className={`flex-1 rounded-md py-1.5 ${mode === "existing" ? "bg-white shadow font-medium" : "text-slate-500"}`}>I have an account</button>
          <button onClick={() => { setMode("new"); setErr(""); }} className={`flex-1 rounded-md py-1.5 ${mode === "new" ? "bg-white shadow font-medium" : "text-slate-500"}`}>New worker</button>
        </div>

        {mode === "existing" ? (
          <>
            <select value={selectedUserId} onChange={(e) => setSelectedUserId(e.target.value)} className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm">
              <option value="">Select your name</option>
              {users.map((w) => <option key={w.id} value={w.id}>{w.name}</option>)}
            </select>
            <input type="password" inputMode="numeric" maxLength={4} placeholder="4-digit PIN" value={pin} onChange={(e) => setPin(e.target.value.replace(/\D/g, ""))} className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm" />
            {err && <p className="text-xs text-red-600">{err}</p>}
            <button disabled={busy} onClick={handleExistingLogin} className="w-full text-white rounded-lg py-2 text-sm font-medium disabled:opacity-60" style={{ background: ORANGE }}>{busy ? "Checking…" : "Log In"}</button>
          </>
        ) : (
          <>
            <input placeholder="Your name" value={name} onChange={(e) => setName(e.target.value)} className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm" />
            <input type="password" inputMode="numeric" maxLength={4} placeholder="Create a 4-digit PIN" value={pin} onChange={(e) => setPin(e.target.value.replace(/\D/g, ""))} className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm" />
            {err && <p className="text-xs text-red-600">{err}</p>}
            <button disabled={busy} onClick={handleCreate} className="w-full text-white rounded-lg py-2 text-sm font-medium disabled:opacity-60" style={{ background: ORANGE }}>{busy ? "Please wait…" : "Create Account & Log In"}</button>
          </>
        )}
      </div>
    </div>
  );
}

// ============================================================== WORKER DASHBOARD
function WorkerDashboard({ user, shop, sales, onLogout, onAddSale }) {
  const mySales = useMemo(
    () => sales.filter((s) => s.kind === "team" && s.workerId === user.id).sort((a, b) => (a.date < b.date ? 1 : -1)),
    [sales, user.id]
  );
  const [form, setForm] = useState({ date: new Date().toISOString().slice(0, 10), product: "", category: "", qty: "", unitPrice: "", cost: "", paymentMethod: "Cash" });
  const totalToday = mySales.filter((s) => s.date === form.date).reduce((a, s) => a + s.total, 0);

  function submit(e) {
    e.preventDefault();
    const qty = Number(form.qty) || 0;
    const unitPrice = Number(form.unitPrice) || 0;
    const cost = Number(form.cost) || 0;
    if (!form.product || qty <= 0 || unitPrice <= 0) return;
    const total = qty * unitPrice;
    const profit = total - cost * qty;
    onAddSale({
      id: uid(), kind: "team", shopId: shop.id, workerId: user.id, workerName: user.name,
      date: form.date, product: form.product, category: form.category || "General",
      qty, unitPrice, cost, paymentMethod: form.paymentMethod, total, profit,
    });
    setForm((f) => ({ ...f, product: "", category: "", qty: "", unitPrice: "", cost: "" }));
  }

  return (
    <div className="max-w-3xl mx-auto px-5 py-6">
      <div className="flex items-center justify-between mb-6">
        <div>
          <p className="text-xs text-slate-400">{shop?.name}</p>
          <h1 className="text-lg font-bold text-slate-900">Hi, {user.name}</h1>
        </div>
        <button onClick={onLogout} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800"><LogOut className="w-4 h-4" /> Log out</button>
      </div>

      <button
        onClick={() => exportToExcel(`${user.name}-sales.xlsx`, [{ name: "My Sales", rows: mySales.map((s) => saleRow(s)) }])}
        className="mb-4 flex items-center gap-1.5 text-sm font-medium border border-slate-300 rounded-lg px-3 py-1.5 text-slate-700 hover:bg-slate-50"
      >
        <Download className="w-4 h-4" /> Download Excel
      </button>

      <div className="grid grid-cols-2 gap-3 mb-6">
        <div className="bg-white border border-slate-200 rounded-xl p-4">
          <p className="text-xs text-slate-400">My sales today</p>
          <p className="text-xl font-bold text-slate-900">{naira(totalToday)}</p>
        </div>
        <div className="bg-white border border-slate-200 rounded-xl p-4">
          <p className="text-xs text-slate-400">My total records</p>
          <p className="text-xl font-bold text-slate-900">{mySales.length}</p>
        </div>
      </div>

      <form onSubmit={submit} className="bg-white border border-slate-200 rounded-xl p-4 mb-6">
        <p className="text-sm font-semibold text-slate-800 mb-3 flex items-center gap-1"><Plus className="w-4 h-4" /> Log a sale</p>
        <div className="grid grid-cols-2 gap-2 mb-2">
          <input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm col-span-2 sm:col-span-1" />
          <select value={form.paymentMethod} onChange={(e) => setForm({ ...form, paymentMethod: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm">
            <option>Cash</option><option>Transfer</option><option>POS</option>
          </select>
          <input placeholder="Product / item" value={form.product} onChange={(e) => setForm({ ...form, product: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm col-span-2" />
          <input placeholder="Category" value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm col-span-2" />
          <input type="number" placeholder="Qty sold" value={form.qty} onChange={(e) => setForm({ ...form, qty: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm" />
          <input type="number" placeholder="Unit price (₦)" value={form.unitPrice} onChange={(e) => setForm({ ...form, unitPrice: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm" />
          <input type="number" placeholder="Cost per item (₦)" value={form.cost} onChange={(e) => setForm({ ...form, cost: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm col-span-2" />
        </div>
        <button type="submit" className="w-full text-white rounded-lg py-2 text-sm font-medium mt-1" style={{ background: ORANGE }}>Add Sale</button>
      </form>

      <p className="text-sm font-semibold text-slate-800 mb-2">My sales history</p>
      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-slate-500 text-xs">
            <tr><th className="text-left px-3 py-2">Date</th><th className="text-left px-3 py-2">Item</th><th className="text-right px-3 py-2">Qty</th><th className="text-right px-3 py-2">Total</th></tr>
          </thead>
          <tbody>
            {mySales.length === 0 && <tr><td colSpan={4} className="text-center text-slate-400 py-6">No sales logged yet.</td></tr>}
            {mySales.map((s) => (
              <tr key={s.id} className="border-t border-slate-100">
                <td className="px-3 py-2">{s.date}</td>
                <td className="px-3 py-2">{s.product}</td>
                <td className="px-3 py-2 text-right">{s.qty}</td>
                <td className="px-3 py-2 text-right font-medium">{naira(s.total)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ============================================================== BOSS DASHBOARD
function BossDashboard({ shops, users, sales, onLogout, onUpdateShopCode, onAddShop, onRenameShop, onRemoveShop, onResetWorkerPin }) {
  const [shopFilter, setShopFilter] = useState("all");
  const [workerFilter, setWorkerFilter] = useState("all");
  const [editingCode, setEditingCode] = useState(null);
  const [codeDraft, setCodeDraft] = useState("");
  const [renamingShop, setRenamingShop] = useState(null);
  const [nameDraft, setNameDraft] = useState("");
  const [addingShop, setAddingShop] = useState(false);
  const [newShopName, setNewShopName] = useState("");
  const [revealShopCode, setRevealShopCode] = useState(null); // { name, code } shown once after adding a shop
  const [resettingWorker, setResettingWorker] = useState(null);
  const [revealWorkerPin, setRevealWorkerPin] = useState(null); // { name, pin } shown once after resetting

  async function handleAddShop() {
    if (!newShopName.trim()) return;
    const code = genPin();
    await onAddShop(newShopName.trim(), code);
    setRevealShopCode({ name: newShopName.trim(), code });
    setNewShopName("");
    setAddingShop(false);
  }
  async function handleResetWorkerPin(u) {
    const pin = genPin();
    await onResetWorkerPin(u.id, pin);
    setRevealWorkerPin({ name: u.name, pin });
    setResettingWorker(null);
  }

  const teamSales = sales.filter((s) => s.kind === "team");
  const workersInShop = shopFilter === "all" ? users : users.filter((u) => u.shopId === shopFilter);

  const filtered = teamSales.filter((s) => {
    if (shopFilter !== "all" && s.shopId !== shopFilter) return false;
    if (workerFilter !== "all" && s.workerId !== workerFilter) return false;
    return true;
  }).sort((a, b) => (a.date < b.date ? 1 : -1));

  const totalRevenue = filtered.reduce((a, s) => a + s.total, 0);
  const totalProfit = filtered.reduce((a, s) => a + s.profit, 0);
  const totalItems = filtered.reduce((a, s) => a + s.qty, 0);

  function shopName(id) { return shops.find((s) => s.id === id)?.name || id; }
  function shopStats(id) {
    const rows = teamSales.filter((s) => s.shopId === id);
    return {
      revenue: rows.reduce((a, s) => a + s.total, 0),
      profit: rows.reduce((a, s) => a + s.profit, 0),
      workers: users.filter((u) => u.shopId === id).length,
    };
  }

  return (
    <div className="max-w-5xl mx-auto px-5 py-6">
      <div className="flex items-center justify-between mb-6">
        <div>
          <p className="text-xs text-slate-400">All shops</p>
          <h1 className="text-lg font-bold text-slate-900 flex items-center gap-2"><ShieldCheck className="w-5 h-5" style={{ color: NAVY }} /> Boss Dashboard</h1>
        </div>
        <button onClick={onLogout} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800"><LogOut className="w-4 h-4" /> Log out</button>
      </div>

      <button
        onClick={() => {
          const summaryRows = shops.map((s) => {
            const st = shopStats(s.id);
            return { Shop: s.name, Workers: st.workers, "Revenue (₦)": st.revenue, "Profit (₦)": st.profit };
          });
          exportToExcel("boss-dashboard.xlsx", [
            { name: "Summary by Shop", rows: summaryRows },
            { name: "Sales (filtered)", rows: filtered.map((s) => saleRow(s, { shop: true, worker: true, shopName: shopName(s.shopId) })) },
          ]);
        }}
        className="mb-4 flex items-center gap-1.5 text-sm font-medium border border-slate-300 rounded-lg px-3 py-1.5 text-slate-700 hover:bg-slate-50"
      >
        <Download className="w-4 h-4" /> Download Excel
      </button>

      {revealShopCode && (
        <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-4 mb-4">
          <p className="text-sm font-semibold text-slate-800 mb-1">"{revealShopCode.name}" created</p>
          <p className="text-xs text-slate-500 mb-2">Give this portal code to that shop's team. It won't be shown again — you can only reset it afterward.</p>
          <div className="flex items-center gap-3">
            <span className="bg-white border border-emerald-300 rounded-lg px-4 py-1.5 font-mono font-bold text-slate-900">{revealShopCode.code}</span>
            <button onClick={() => setRevealShopCode(null)} className="text-xs text-slate-500 hover:text-slate-800">Dismiss</button>
          </div>
        </div>
      )}
      {revealWorkerPin && (
        <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-4 mb-4">
          <p className="text-sm font-semibold text-slate-800 mb-1">{revealWorkerPin.name}'s PIN was reset</p>
          <p className="text-xs text-slate-500 mb-2">Give them this new PIN. It won't be shown again.</p>
          <div className="flex items-center gap-3">
            <span className="bg-white border border-emerald-300 rounded-lg px-4 py-1.5 font-mono font-bold text-slate-900">{revealWorkerPin.pin}</span>
            <button onClick={() => setRevealWorkerPin(null)} className="text-xs text-slate-500 hover:text-slate-800">Dismiss</button>
          </div>
        </div>
      )}

      {/* Per-shop overview */}
      <div className="flex items-center justify-between mb-2">
        <p className="text-sm font-semibold text-slate-800">Overview by shop</p>
        {!addingShop && (
          <button onClick={() => setAddingShop(true)} className="flex items-center gap-1 text-xs font-medium text-white rounded-lg px-2.5 py-1" style={{ background: NAVY }}>
            <Plus className="w-3.5 h-3.5" /> Add Shop
          </button>
        )}
      </div>
      {addingShop && (
        <div className="bg-white border border-slate-200 rounded-xl p-4 mb-3 flex gap-2">
          <input
            autoFocus
            placeholder="New shop name (e.g. Main Branch)"
            value={newShopName}
            onChange={(e) => setNewShopName(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && handleAddShop()}
            className="flex-1 border border-slate-300 rounded-lg px-3 py-1.5 text-sm"
          />
          <button onClick={handleAddShop} className="text-sm px-3 py-1.5 rounded-lg text-white" style={{ background: NAVY }}>Create</button>
          <button onClick={() => { setAddingShop(false); setNewShopName(""); }} className="text-sm px-3 py-1.5 rounded-lg text-slate-500">Cancel</button>
        </div>
      )}
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 mb-6">
        {shops.length === 0 && !addingShop && (
          <div className="sm:col-span-3 text-center text-sm text-slate-400 py-6 border border-dashed border-slate-300 rounded-xl">
            No shops yet — click "Add Shop" to create your first one.
          </div>
        )}
        {shops.map((s) => {
          const st = shopStats(s.id);
          return (
            <div key={s.id} className="bg-white border border-slate-200 rounded-xl p-4">
              <div className="flex items-center justify-between mb-2">
                {renamingShop === s.id ? (
                  <div className="flex gap-1 flex-1">
                    <input value={nameDraft} onChange={(e) => setNameDraft(e.target.value)} className="border border-slate-300 rounded px-2 py-1 text-xs flex-1" autoFocus />
                    <button onClick={() => { if (nameDraft.trim()) { onRenameShop(s.id, nameDraft.trim()); setRenamingShop(null); } }} className="text-xs px-2 py-1 rounded text-white" style={{ background: NAVY }}>Save</button>
                  </div>
                ) : (
                  <button onClick={() => { setRenamingShop(s.id); setNameDraft(s.name); }} className="text-sm font-semibold text-slate-800 hover:underline text-left">{s.name}</button>
                )}
                <span className="text-[10px] text-slate-400 whitespace-nowrap ml-2">{st.workers} worker{st.workers === 1 ? "" : "s"}</span>
              </div>
              <p className="text-lg font-bold" style={{ color: NAVY }}>{naira(st.revenue)}</p>
              <p className="text-xs text-slate-400 mb-2">Profit: <span style={{ color: ORANGE }}>{naira(st.profit)}</span></p>
              {editingCode === s.id ? (
                <div className="flex gap-1 mt-2">
                  <input value={codeDraft} onChange={(e) => setCodeDraft(e.target.value)} className="border border-slate-300 rounded px-2 py-1 text-xs w-24" placeholder="New portal code" autoFocus />
                  <button
                    onClick={() => { if (codeDraft.trim()) { onUpdateShopCode(s.id, codeDraft.trim()); setEditingCode(null); } }}
                    className="text-xs px-2 py-1 rounded text-white" style={{ background: NAVY }}
                  >
                    Save
                  </button>
                  <button onClick={() => setEditingCode(null)} className="text-xs px-2 py-1 rounded text-slate-500">Cancel</button>
                </div>
              ) : (
                <div className="flex items-center justify-between">
                  <button onClick={() => { setEditingCode(s.id); setCodeDraft(""); }} className="text-[11px] text-slate-400 flex items-center gap-1 hover:text-slate-700">
                    <Lock className="w-3 h-3" /> Portal code set (change)
                  </button>
                  <button
                    onClick={() => { if (window.confirm(`Remove "${s.name}"? Its sales history stays, but workers there won't be able to log in.`)) onRemoveShop(s.id); }}
                    className="text-[11px] text-red-400 hover:text-red-600"
                  >
                    Remove
                  </button>
                </div>
              )}
            </div>
          );
        })}
      </div>

      {/* Worker PIN management */}
      {users.length > 0 && (
        <div className="mb-6">
          <p className="text-sm font-semibold text-slate-800 mb-2">Workers</p>
          <div className="bg-white border border-slate-200 rounded-xl divide-y divide-slate-100">
            {users.map((u) => (
              <div key={u.id} className="flex items-center justify-between px-4 py-2.5">
                <div>
                  <p className="text-sm font-medium text-slate-800">{u.name}</p>
                  <p className="text-xs text-slate-400">{shopName(u.shopId)}</p>
                </div>
                {resettingWorker === u.id ? (
                  <div className="flex gap-2 items-center">
                    <span className="text-xs text-slate-500">Reset this worker's PIN?</span>
                    <button onClick={() => handleResetWorkerPin(u)} className="text-xs px-2 py-1 rounded text-white" style={{ background: ORANGE }}>Confirm</button>
                    <button onClick={() => setResettingWorker(null)} className="text-xs px-2 py-1 rounded text-slate-500">Cancel</button>
                  </div>
                ) : (
                  <button onClick={() => setResettingWorker(u.id)} className="text-xs text-slate-400 hover:text-slate-700">Reset PIN</button>
                )}
              </div>
            ))}
          </div>
        </div>
      )}


      {/* Combined filters + totals */}
      <div className="flex flex-wrap gap-2 mb-4">
        <select value={shopFilter} onChange={(e) => { setShopFilter(e.target.value); setWorkerFilter("all"); }} className="border border-slate-300 rounded-lg px-3 py-1.5 text-sm">
          <option value="all">All Shops</option>
          {shops.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
        </select>
        <select value={workerFilter} onChange={(e) => setWorkerFilter(e.target.value)} className="border border-slate-300 rounded-lg px-3 py-1.5 text-sm">
          <option value="all">All Workers</option>
          {workersInShop.map((w) => <option key={w.id} value={w.id}>{w.name}</option>)}
        </select>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 mb-6">
        <div className="bg-white border border-slate-200 rounded-xl p-4">
          <p className="text-xs text-slate-400 flex items-center gap-1"><TrendingUp className="w-3.5 h-3.5" /> Filtered Revenue</p>
          <p className="text-xl font-bold" style={{ color: NAVY }}>{naira(totalRevenue)}</p>
        </div>
        <div className="bg-white border border-slate-200 rounded-xl p-4">
          <p className="text-xs text-slate-400 flex items-center gap-1"><Package className="w-3.5 h-3.5" /> Items Sold</p>
          <p className="text-xl font-bold text-slate-900">{totalItems}</p>
        </div>
        <div className="bg-white border border-slate-200 rounded-xl p-4">
          <p className="text-xs text-slate-400">Filtered Profit</p>
          <p className="text-xl font-bold" style={{ color: ORANGE }}>{naira(totalProfit)}</p>
        </div>
      </div>

      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-slate-500 text-xs">
            <tr>
              <th className="text-left px-3 py-2">Date</th>
              <th className="text-left px-3 py-2">Shop</th>
              <th className="text-left px-3 py-2">Worker</th>
              <th className="text-left px-3 py-2">Item</th>
              <th className="text-right px-3 py-2">Qty</th>
              <th className="text-right px-3 py-2">Total</th>
              <th className="text-right px-3 py-2">Profit</th>
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 && <tr><td colSpan={7} className="text-center text-slate-400 py-8">No sales recorded yet for this filter.</td></tr>}
            {filtered.map((s) => (
              <tr key={s.id} className="border-t border-slate-100">
                <td className="px-3 py-2">{s.date}</td>
                <td className="px-3 py-2">{shopName(s.shopId)}</td>
                <td className="px-3 py-2">{s.workerName}</td>
                <td className="px-3 py-2">{s.product}</td>
                <td className="px-3 py-2 text-right">{s.qty}</td>
                <td className="px-3 py-2 text-right font-medium">{naira(s.total)}</td>
                <td className="px-3 py-2 text-right text-emerald-700">{naira(s.profit)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
