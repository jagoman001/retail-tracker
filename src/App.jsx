import { useState, useEffect, useMemo } from "react";
import { Store, User, LogOut, Plus, TrendingUp, Package, ShieldCheck, ArrowLeft, Loader2, Lock, Users, Download } from "lucide-react";
import * as XLSX from "xlsx";
import { databases, DATABASE_ID, TABLE_ID } from "./appwriteClient";

// Ledger-inspired palette: deep ink (not generic corporate navy), a muted
// rust (not the terracotta every AI dashboard reaches for), warm naira-gold
// for emphasis, and a stamped-ink green for "paid"/positive states.
const NAVY = "#1B2A3A";    // ink — sidebar, primary text, primary actions
const ORANGE = "#9C3B2E";  // rust — secondary accent, worker-mode actions
const GOLD = "#B8863B";    // naira-gold — active nav marker, highlights
const GREEN = "#2F6E4F";   // ledger green — paid / positive figures
const PAPER = "#F6F1E7";   // warm paper background

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
// Opens a printable invoice for a single sale in a new tab (uses the browser's
// native print dialog — no PDF library needed, works everywhere).
function openInvoice(sale, businessName, shopName) {
  const win = window.open("", "_blank");
  if (!win) { alert("Please allow pop-ups to generate the invoice."); return; }
  const balance = sale.paymentMethod === "Credit" || sale.paymentMethod === "Part Payment"
    ? sale.total - (sale.amountPaidNow || 0) : 0;
  const paidNow = sale.paymentMethod === "Part Payment" ? (sale.amountPaidNow || 0) : (sale.paymentMethod === "Credit" ? 0 : sale.total);
  win.document.write(`
    <html>
      <head>
        <title>Invoice - ${sale.id.slice(0, 6).toUpperCase()}</title>
        <style>
          body { font-family: Arial, sans-serif; padding: 40px; color: #0B3D56; max-width: 600px; margin: 0 auto; }
          h1 { font-size: 20px; margin-bottom: 0; }
          .sub { color: #64748b; font-size: 13px; margin-top: 4px; }
          table { width: 100%; border-collapse: collapse; margin-top: 24px; }
          th, td { text-align: left; padding: 8px; border-bottom: 1px solid #e2e8f0; font-size: 14px; }
          th { color: #64748b; font-size: 12px; text-transform: uppercase; }
          .right { text-align: right; }
          .total-row td { font-weight: bold; font-size: 16px; border-top: 2px solid #0B3D56; border-bottom: none; }
          .meta { margin-top: 24px; font-size: 13px; color: #475569; }
          .balance { color: #C0501E; font-weight: bold; }
        </style>
      </head>
      <body onload="window.print()">
        <h1>${businessName}${shopName ? " — " + shopName : ""}</h1>
        <p class="sub">Invoice #${sale.id.slice(0, 6).toUpperCase()} &nbsp;•&nbsp; ${sale.date}</p>
        <p class="meta">Customer: <strong>${sale.customerName || "Walk-in"}</strong></p>
        <table>
          <thead><tr><th>Item</th><th class="right">Qty</th><th class="right">Unit Price</th><th class="right">Total</th></tr></thead>
          <tbody>
            <tr><td>${sale.product}</td><td class="right">${sale.qty}</td><td class="right">${naira(sale.unitPrice)}</td><td class="right">${naira(sale.total)}</td></tr>
            <tr class="total-row"><td colspan="3">TOTAL</td><td class="right">${naira(sale.total)}</td></tr>
          </tbody>
        </table>
        <p class="meta">Payment method: <strong>${sale.paymentMethod}</strong></p>
        <p class="meta">Amount paid: <strong>${naira(paidNow)}</strong></p>
        ${balance > 0 ? `<p class="meta balance">Balance due: ${naira(balance)}${sale.dueDate ? " (due " + sale.dueDate + ")" : ""}</p>` : ""}
        <p class="meta" style="margin-top:40px;color:#94a3b8;">Thank you for your business.</p>
      </body>
    </html>
  `);
  win.document.close();
}
// If a sale was logged against a tracked inventory item, reduce that item's
// stock by the quantity sold (never below zero).
function applySaleToInventory(sale, inventory) {
  if (!sale.inventoryItemId) return inventory;
  return inventory.map((it) => (it.id === sale.inventoryItemId ? { ...it, qty: Math.max(0, it.qty - sale.qty) } : it));
}
function receivableStatus(r) {
  const balance = r.invoiceAmt - r.amountPaid;
  if (balance <= 0) return "Paid";
  if (r.dueDate && new Date() > new Date(r.dueDate)) return "Overdue";
  return "Pending";
}
// Credit / Part Payment sales automatically create a receivable to track what's still owed.
function maybeCreateReceivable(sale) {
  if (sale.paymentMethod !== "Credit" && sale.paymentMethod !== "Part Payment") return null;
  const amountPaid = sale.paymentMethod === "Part Payment" ? sale.amountPaidNow || 0 : 0;
  return {
    id: uid(), kind: sale.kind, ownerId: sale.kind === "solo" ? sale.businessId : sale.shopId,
    date: sale.date, customerName: sale.customerName || "Walk-in", item: sale.product,
    invoiceAmt: sale.total, amountPaid, initialAmountPaid: amountPaid,
    dueDate: sale.dueDate || "", saleId: sale.id, workerName: sale.workerName,
  };
}

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
function monthKey(dateStr) { return dateStr ? dateStr.slice(0, 7) : ""; } // "YYYY-MM"
function monthLabel(dateStr) { return dateStr ? MONTHS[Number(dateStr.slice(5, 7)) - 1] + " " + dateStr.slice(0, 4) : ""; }
// Builds the Cash Flow / VAT / P&L figures for a given slice of sales/expenses/receivables,
// grouped by month — mirrors the structure of the original Excel tracker but fully automatic.
function buildFinancials(salesRows, expenseRows, receivableRows, vatRate = 0.075) {
  const months = new Set();
  salesRows.forEach((s) => months.add(monthKey(s.date)));
  expenseRows.forEach((e) => months.add(monthKey(e.date)));
  const sortedMonths = Array.from(months).filter(Boolean).sort();

  const perMonth = sortedMonths.map((mk) => {
    const monthSales = salesRows.filter((s) => monthKey(s.date) === mk);
    const monthExpenses = expenseRows.filter((e) => monthKey(e.date) === mk);
    const revenue = monthSales.reduce((a, s) => a + s.total, 0);
    const cogs = monthSales.reduce((a, s) => a + s.cost * s.qty, 0);
    const grossProfit = revenue - cogs;
    const expenseTotal = monthExpenses.reduce((a, e) => a + e.amount, 0);
    const netProfit = grossProfit - expenseTotal;
    // Cash actually collected: full total for Cash/POS/Transfer, only the amount paid so far for Credit/Part Payment.
    const cashCollected = monthSales.reduce((a, s) => a + (s.paymentMethod === "Credit" || s.paymentMethod === "Part Payment" ? s.amountPaidNow || 0 : s.total), 0);
    const purchases = monthExpenses.filter((e) => e.category === "Stock/Inventory").reduce((a, e) => a + e.amount, 0);
    const vatCollected = revenue * vatRate;
    const vatPaid = purchases * vatRate;
    return { month: monthLabel(mk + "-01"), mk, revenue, cogs, grossProfit, expenseTotal, netProfit, cashCollected, purchases, vatCollected, vatPaid, netVat: vatCollected - vatPaid };
  });

  const expenseByCategory = {};
  expenseRows.forEach((e) => { expenseByCategory[e.category || "Other"] = (expenseByCategory[e.category || "Other"] || 0) + e.amount; });

  const receivablesCollectedByMonth = {};
  // Payments recorded against receivables after the original sale count as "Receivables Collected" cash inflow.
  receivableRows.forEach((r) => {
    const extra = r.amountPaid - (r.initialAmountPaid || 0);
    if (extra > 0) {
      const mk = monthKey(r.lastPaymentDate || r.date);
      receivablesCollectedByMonth[mk] = (receivablesCollectedByMonth[mk] || 0) + extra;
    }
  });

  return { perMonth, expenseByCategory, receivablesCollectedByMonth };
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
  const [inventory, setInventory] = useState([]);
  const [expenses, setExpenses] = useState([]);
  const [receivables, setReceivables] = useState([]);
  const [soloBiz, setSoloBiz] = useState([]);
  const [bossCodeHash, setBossCodeHash] = useState(null);
  const [bossRecoveryHash, setBossRecoveryHash] = useState(null);
  const [screen, setScreen] = useState("mode");
  const [currentUser, setCurrentUser] = useState(null);
  const [activeShopId, setActiveShopId] = useState(null);

  useEffect(() => {
    (async () => {
      const [s, u, sl, sb, bch, brh, inv, exp, rec] = await Promise.all([
        loadList("shops", []),
        loadList("users", []),
        loadList("sales", []),
        loadList("solo_businesses", []),
        loadList("boss_code_hash", null),
        loadList("boss_recovery_hash", null),
        loadList("inventory", []),
        loadList("expenses", []),
        loadList("receivables", []),
      ]);
      setShops(s);
      setUsers(u);
      setSales(sl);
      setSoloBiz(sb);
      setBossCodeHash(bch);
      setBossRecoveryHash(brh);
      setInventory(inv);
      setExpenses(exp);
      setReceivables(rec);
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
          inventory={inventory.filter((it) => it.kind === "solo" && it.ownerId === currentUser.id)}
          expenses={expenses.filter((e) => e.kind === "solo" && e.ownerId === currentUser.id)}
          receivables={receivables.filter((r) => r.kind === "solo" && r.ownerId === currentUser.id)}
          onLogout={goHome}
          onAddSale={async (sale) => {
            const nextSales = [...sales, sale];
            setSales(nextSales);
            await saveList("sales", nextSales);
            if (sale.inventoryItemId) {
              const nextInv = applySaleToInventory(sale, inventory);
              setInventory(nextInv);
              await saveList("inventory", nextInv);
            }
            const receivable = maybeCreateReceivable(sale);
            if (receivable) {
              const nextRec = [...receivables, receivable];
              setReceivables(nextRec);
              await saveList("receivables", nextRec);
            }
          }}
          onAddItem={async (item) => {
            const next = [...inventory, item];
            setInventory(next);
            await saveList("inventory", next);
          }}
          onUpdateItem={async (itemId, patch) => {
            const next = inventory.map((it) => (it.id === itemId ? { ...it, ...patch } : it));
            setInventory(next);
            await saveList("inventory", next);
          }}
          onRemoveItem={async (itemId) => {
            const next = inventory.filter((it) => it.id !== itemId);
            setInventory(next);
            await saveList("inventory", next);
          }}
          onAddExpense={async (expense) => {
            const next = [...expenses, expense];
            setExpenses(next);
            await saveList("expenses", next);
          }}
          onRemoveExpense={async (expenseId) => {
            const next = expenses.filter((e) => e.id !== expenseId);
            setExpenses(next);
            await saveList("expenses", next);
          }}
          onRecordPayment={async (receivableId, amount) => {
            const next = receivables.map((r) => (r.id === receivableId ? { ...r, amountPaid: r.amountPaid + amount, lastPaymentDate: new Date().toISOString().slice(0, 10) } : r));
            setReceivables(next);
            await saveList("receivables", next);
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
          inventory={inventory.filter((it) => it.kind === "team" && it.ownerId === activeShopId)}
          expenses={expenses.filter((e) => e.kind === "team" && e.ownerId === activeShopId)}
          receivables={receivables.filter((r) => r.kind === "team" && r.ownerId === activeShopId)}
          onLogout={goHome}
          onAddSale={async (sale) => {
            const nextSales = [...sales, sale];
            setSales(nextSales);
            await saveList("sales", nextSales);
            if (sale.inventoryItemId) {
              const nextInv = applySaleToInventory(sale, inventory);
              setInventory(nextInv);
              await saveList("inventory", nextInv);
            }
            const receivable = maybeCreateReceivable(sale);
            if (receivable) {
              const nextRec = [...receivables, receivable];
              setReceivables(nextRec);
              await saveList("receivables", nextRec);
            }
          }}
          onAddItem={async (item) => {
            const next = [...inventory, item];
            setInventory(next);
            await saveList("inventory", next);
          }}
          onUpdateItem={async (itemId, patch) => {
            const next = inventory.map((it) => (it.id === itemId ? { ...it, ...patch } : it));
            setInventory(next);
            await saveList("inventory", next);
          }}
          onRemoveItem={async (itemId) => {
            const next = inventory.filter((it) => it.id !== itemId);
            setInventory(next);
            await saveList("inventory", next);
          }}
          onAddExpense={async (expense) => {
            const next = [...expenses, expense];
            setExpenses(next);
            await saveList("expenses", next);
          }}
          onRemoveExpense={async (expenseId) => {
            const next = expenses.filter((e) => e.id !== expenseId);
            setExpenses(next);
            await saveList("expenses", next);
          }}
          onRecordPayment={async (receivableId, amount) => {
            const next = receivables.map((r) => (r.id === receivableId ? { ...r, amountPaid: r.amountPaid + amount, lastPaymentDate: new Date().toISOString().slice(0, 10) } : r));
            setReceivables(next);
            await saveList("receivables", next);
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
          inventory={inventory}
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
          onAddItem={async (item) => {
            const next = [...inventory, item];
            setInventory(next);
            await saveList("inventory", next);
          }}
          onUpdateItem={async (itemId, patch) => {
            const next = inventory.map((it) => (it.id === itemId ? { ...it, ...patch } : it));
            setInventory(next);
            await saveList("inventory", next);
          }}
          onRemoveItem={async (itemId) => {
            const next = inventory.filter((it) => it.id !== itemId);
            setInventory(next);
            await saveList("inventory", next);
          }}
          expenses={expenses}
          receivables={receivables}
          onAddExpense={async (expense) => {
            const next = [...expenses, expense];
            setExpenses(next);
            await saveList("expenses", next);
          }}
          onRemoveExpense={async (expenseId) => {
            const next = expenses.filter((e) => e.id !== expenseId);
            setExpenses(next);
            await saveList("expenses", next);
          }}
          onRecordPayment={async (receivableId, amount) => {
            const next = receivables.map((r) => (r.id === receivableId ? { ...r, amountPaid: r.amountPaid + amount, lastPaymentDate: new Date().toISOString().slice(0, 10) } : r));
            setReceivables(next);
            await saveList("receivables", next);
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

// A shared shell: ink-navy sidebar with gold active-marker on desktop,
// collapsing to a top bar + horizontally scrollable tab strip on mobile.
function DashboardShell({ eyebrow, title, onLogout, tabs, active, onChange, actions, children }) {
  return (
    <div className="min-h-screen flex flex-col sm:flex-row" style={{ background: PAPER }}>
      <aside className="hidden sm:flex sm:flex-col w-56 shrink-0 text-white" style={{ background: NAVY }}>
        <div className="px-5 py-6">
          <p className="font-display text-lg font-semibold tracking-tight">Ledger</p>
          <p className="text-xs text-white/50 mt-0.5">{eyebrow}</p>
        </div>
        <nav className="flex-1 px-3 space-y-1">
          {tabs.map(([key, label]) => (
            <button
              key={key}
              onClick={() => onChange(key)}
              className={`w-full flex items-center gap-2 px-3 py-2 rounded-lg text-sm text-left transition-colors ${active === key ? "bg-white/10 font-medium" : "text-white/60 hover:text-white/90 hover:bg-white/5"}`}
            >
              <span className="w-1.5 h-1.5 rounded-full shrink-0" style={{ background: active === key ? GOLD : "transparent" }} />
              {label}
            </button>
          ))}
        </nav>
        <div className="px-5 py-4 border-t border-white/10">
          <button onClick={onLogout} className="flex items-center gap-1.5 text-xs text-white/60 hover:text-white">
            <LogOut className="w-3.5 h-3.5" /> Log out
          </button>
        </div>
      </aside>

      <div className="sm:hidden text-white" style={{ background: NAVY }}>
        <div className="flex items-center justify-between px-4 py-3">
          <div>
            <p className="font-display font-semibold">Ledger</p>
            <p className="text-[11px] text-white/50">{eyebrow}</p>
          </div>
          <button onClick={onLogout} className="text-xs text-white/60 flex items-center gap-1"><LogOut className="w-3.5 h-3.5" /> Log out</button>
        </div>
        <div className="flex gap-1 overflow-x-auto px-3 pb-2">
          {tabs.map(([key, label]) => (
            <button
              key={key}
              onClick={() => onChange(key)}
              className={`px-3 py-1.5 rounded-full text-xs whitespace-nowrap ${active === key ? "font-medium" : "text-white/50"}`}
              style={active === key ? { background: "rgba(255,255,255,0.15)", color: GOLD } : {}}
            >
              {label}
            </button>
          ))}
        </div>
      </div>

      <main className="flex-1 min-w-0">
        <div className="max-w-4xl mx-auto px-5 py-6">
          <div className="flex items-center justify-between mb-6 gap-3 flex-wrap">
            <h1 className="font-display text-xl font-semibold" style={{ color: NAVY }}>{title}</h1>
            {actions}
          </div>
          {children}
        </div>
      </main>
    </div>
  );
}

// The signature element: a rotated, dashed-border ink-stamp — evoking the
// rubber stamp on a physical receipt — used for status states (Paid/Overdue/Pending).
function StampBadge({ label, tone }) {
  const colors = { green: GREEN, rust: ORANGE, gold: GOLD };
  const c = colors[tone] || GREEN;
  return (
    <span
      className="inline-block px-2.5 py-0.5 text-[10px] font-bold uppercase tracking-wider border-2 border-dashed rounded-full -rotate-3"
      style={{ borderColor: c, color: c }}
    >
      {label}
    </span>
  );
}

function SoloDashboard({ business, sales, inventory, expenses, receivables, onLogout, onAddSale, onAddItem, onUpdateItem, onRemoveItem, onAddExpense, onRemoveExpense, onRecordPayment }) {
  const mySales = useMemo(
    () => sales.filter((s) => s.kind === "solo" && s.businessId === business.id).sort((a, b) => (a.date < b.date ? 1 : -1)),
    [sales, business.id]
  );
  const [form, setForm] = useState({ date: new Date().toISOString().slice(0, 10), itemId: "", product: "", category: "", qty: "", unitPrice: "", cost: "", customerName: "", paymentMethod: "Cash", amountPaidNow: "", dueDate: "" });

  const totalToday = mySales.filter((s) => s.date === form.date).reduce((a, s) => a + s.total, 0);
  const totalAll = mySales.reduce((a, s) => a + s.total, 0);
  const profitAll = mySales.reduce((a, s) => a + s.profit, 0);
  const expenseTotal = expenses.reduce((a, e) => a + e.amount, 0);
  const netProfit = profitAll - expenseTotal;
  const stockValue = inventory.reduce((a, it) => a + it.qty * it.cost, 0);
  const outstandingReceivables = receivables.reduce((a, r) => a + Math.max(0, r.invoiceAmt - r.amountPaid), 0);

  function pickInventoryItem(itemId) {
    if (itemId === "__custom__") {
      setForm((f) => ({ ...f, itemId: "", product: "", unitPrice: "", cost: "" }));
      return;
    }
    const it = inventory.find((x) => x.id === itemId);
    if (!it) return;
    setForm((f) => ({ ...f, itemId, product: it.name, unitPrice: String(it.price || ""), cost: String(it.cost || "") }));
  }

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
      qty, unitPrice, cost, customerName: form.customerName || "Walk-in", paymentMethod: form.paymentMethod, total, profit,
      inventoryItemId: form.itemId || null,
      amountPaidNow: form.paymentMethod === "Part Payment" ? Number(form.amountPaidNow) || 0 : undefined,
      dueDate: form.paymentMethod === "Credit" || form.paymentMethod === "Part Payment" ? form.dueDate : undefined,
    });
    setForm((f) => ({ ...f, itemId: "", product: "", category: "", qty: "", unitPrice: "", cost: "", customerName: "", amountPaidNow: "", dueDate: "" }));
  }

  const [tab, setTab] = useState("overview");

  return (
    <DashboardShell
      eyebrow="Solo shop"
      title={business.name}
      onLogout={onLogout}
      tabs={[["overview", "Overview"], ["inventory", "Inventory"], ["expenses", "Expenses"], ["receivables", "Receivables"], ["reports", "Reports"]]}
      active={tab}
      onChange={setTab}
      actions={
        <button
          onClick={() => exportToExcel(`${business.name}-sales.xlsx`, [
            { name: "Sales", rows: mySales.map((s) => saleRow(s)) },
            { name: "Inventory", rows: inventory.map((it) => ({ Item: it.name, "In Stock": it.qty, "Cost/unit (₦)": it.cost, "Price/unit (₦)": it.price })) },
            { name: "Expenses", rows: expenses.map((e) => ({ Date: e.date, Description: e.description, Category: e.category, "Amount (₦)": e.amount, "Paid By": e.paidBy, "Receipt #": e.receiptNo, Notes: e.notes })) },
            { name: "Receivables", rows: receivables.map((r) => ({ Date: r.date, Customer: r.customerName, Item: r.item, "Invoice Amt (₦)": r.invoiceAmt, "Amount Paid (₦)": r.amountPaid, "Balance (₦)": r.invoiceAmt - r.amountPaid, Status: receivableStatus(r) })) },
          ])}
          className="flex items-center gap-1.5 text-sm font-medium border border-slate-300 rounded-lg px-3 py-1.5 text-slate-700 hover:bg-white bg-white/60"
        >
          <Download className="w-4 h-4" /> Download Excel
        </button>
      }
    >
      {tab === "overview" && (
        <>
          <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 mb-6">
            <div className="bg-white border border-slate-200 rounded-xl p-4">
              <p className="text-xs text-slate-400">Today</p>
              <p className="text-lg font-semibold figure text-slate-900">{naira(totalToday)}</p>
            </div>
            <div className="bg-white border border-slate-200 rounded-xl p-4">
              <p className="text-xs text-slate-400">Total Revenue</p>
              <p className="text-lg font-semibold figure" style={{ color: NAVY }}>{naira(totalAll)}</p>
            </div>
            <div className="bg-white border border-slate-200 rounded-xl p-4">
              <p className="text-xs text-slate-400">Total Expenses</p>
              <p className="text-lg font-semibold figure text-red-600">{naira(expenseTotal)}</p>
            </div>
            <div className="bg-white border border-slate-200 rounded-xl p-4">
              <p className="text-xs text-slate-400">Net Profit</p>
              <p className="text-lg font-semibold figure" style={{ color: ORANGE }}>{naira(netProfit)}</p>
            </div>
            <div className="bg-white border border-slate-200 rounded-xl p-4">
              <p className="text-xs text-slate-400">Stock Value</p>
              <p className="text-lg font-semibold figure text-slate-900">{naira(stockValue)}</p>
            </div>
            <div className="bg-white border border-slate-200 rounded-xl p-4">
              <p className="text-xs text-slate-400">Outstanding Receivables</p>
              <p className="text-lg font-semibold figure" style={{ color: GOLD }}>{naira(outstandingReceivables)}</p>
            </div>
          </div>

          <form onSubmit={submit} className="bg-white border border-slate-200 rounded-xl p-4 mb-6">
            <p className="text-sm font-semibold text-slate-800 mb-3 flex items-center gap-1"><Plus className="w-4 h-4" /> Log a sale</p>
            <div className="grid grid-cols-2 gap-2 mb-2">
              <input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm col-span-2 sm:col-span-1" />
              <select value={form.paymentMethod} onChange={(e) => setForm({ ...form, paymentMethod: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm">
                <option>Cash</option><option>Transfer</option><option>POS</option><option>Credit</option><option>Part Payment</option>
              </select>
              <input placeholder="Customer name (optional)" value={form.customerName} onChange={(e) => setForm({ ...form, customerName: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm col-span-2" />
              <select value={form.itemId || "__custom__"} onChange={(e) => pickInventoryItem(e.target.value)} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm col-span-2">
                <option value="__custom__">— Custom item (not tracked in stock) —</option>
                {inventory.map((it) => <option key={it.id} value={it.id}>{it.name} ({it.qty} in stock)</option>)}
              </select>
              <input placeholder="Product / item" value={form.product} onChange={(e) => setForm({ ...form, product: e.target.value, itemId: "" })} disabled={!!form.itemId} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm col-span-2 disabled:bg-slate-50 disabled:text-slate-500" />
              <input placeholder="Category" value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm col-span-2" />
              <input type="number" placeholder="Qty sold" value={form.qty} onChange={(e) => setForm({ ...form, qty: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm" />
              <input type="number" placeholder="Unit price (₦)" value={form.unitPrice} onChange={(e) => setForm({ ...form, unitPrice: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm" />
              <input type="number" placeholder="Cost per item (₦)" value={form.cost} onChange={(e) => setForm({ ...form, cost: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm col-span-2" />
              {(form.paymentMethod === "Credit" || form.paymentMethod === "Part Payment") && (
                <>
                  {form.paymentMethod === "Part Payment" && (
                    <input type="number" placeholder="Amount paid now (₦)" value={form.amountPaidNow} onChange={(e) => setForm({ ...form, amountPaidNow: e.target.value })} className="border border-amber-300 rounded-lg px-2 py-1.5 text-sm" />
                  )}
                  <input type="date" placeholder="Due date" value={form.dueDate} onChange={(e) => setForm({ ...form, dueDate: e.target.value })} className={`border border-amber-300 rounded-lg px-2 py-1.5 text-sm ${form.paymentMethod === "Credit" ? "col-span-2" : ""}`} />
                </>
              )}
            </div>
            <button type="submit" className="w-full text-white rounded-lg py-2 text-sm font-medium mt-1" style={{ background: NAVY }}>Add Sale</button>
          </form>

          <p className="text-sm font-semibold text-slate-800 mb-2">Sales history</p>
          <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
            <table className="w-full text-sm">
              <thead className="bg-slate-50 text-slate-500 text-xs">
                <tr><th className="text-left px-3 py-2">Date</th><th className="text-left px-3 py-2">Item</th><th className="text-right px-3 py-2">Qty</th><th className="text-right px-3 py-2">Total</th><th className="px-3 py-2"></th></tr>
              </thead>
              <tbody>
                {mySales.length === 0 && <tr><td colSpan={5} className="text-center text-slate-400 py-6">No sales logged yet.</td></tr>}
                {mySales.map((s) => (
                  <tr key={s.id} className="border-t border-slate-100">
                    <td className="px-3 py-2">{s.date}</td>
                    <td className="px-3 py-2">{s.product}</td>
                    <td className="px-3 py-2 text-right">{s.qty}</td>
                    <td className="px-3 py-2 text-right font-medium">{naira(s.total)}</td>
                    <td className="px-3 py-2 text-right"><button onClick={() => openInvoice(s, business.name)} className="text-xs text-slate-400 hover:text-slate-700">Invoice</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}

      {tab === "inventory" && (
        <InventoryManager
          items={inventory}
          onAddItem={onAddItem}
          onUpdateItem={onUpdateItem}
          onRemoveItem={onRemoveItem}
          accent={NAVY}
          makeItem={(data) => ({ id: uid(), kind: "solo", ownerId: business.id, ...data })}
        />
      )}

      {tab === "expenses" && (
        <ExpenseManager
          expenses={expenses}
          onAddExpense={onAddExpense}
          onRemoveExpense={onRemoveExpense}
          accent={NAVY}
          makeExpense={(data) => ({ id: uid(), kind: "solo", ownerId: business.id, ...data })}
        />
      )}

      {tab === "receivables" && (
        <ReceivablesManager receivables={receivables} onRecordPayment={onRecordPayment} accent={NAVY} />
      )}

      {tab === "reports" && (
        <FinancialsView salesRows={mySales} expenseRows={expenses} receivableRows={receivables} accent={NAVY} />
      )}
    </DashboardShell>
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
function WorkerDashboard({ user, shop, sales, inventory, expenses, receivables, onLogout, onAddSale, onAddItem, onUpdateItem, onRemoveItem, onAddExpense, onRemoveExpense, onRecordPayment }) {
  const mySales = useMemo(
    () => sales.filter((s) => s.kind === "team" && s.workerId === user.id).sort((a, b) => (a.date < b.date ? 1 : -1)),
    [sales, user.id]
  );
  const [form, setForm] = useState({ date: new Date().toISOString().slice(0, 10), itemId: "", product: "", category: "", qty: "", unitPrice: "", cost: "", customerName: "", paymentMethod: "Cash", amountPaidNow: "", dueDate: "" });
  const totalToday = mySales.filter((s) => s.date === form.date).reduce((a, s) => a + s.total, 0);

  function pickInventoryItem(itemId) {
    if (itemId === "__custom__") {
      setForm((f) => ({ ...f, itemId: "", product: "", unitPrice: "", cost: "" }));
      return;
    }
    const it = inventory.find((x) => x.id === itemId);
    if (!it) return;
    setForm((f) => ({ ...f, itemId, product: it.name, unitPrice: String(it.price || ""), cost: String(it.cost || "") }));
  }

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
      qty, unitPrice, cost, customerName: form.customerName || "Walk-in", paymentMethod: form.paymentMethod, total, profit,
      inventoryItemId: form.itemId || null,
      amountPaidNow: form.paymentMethod === "Part Payment" ? Number(form.amountPaidNow) || 0 : undefined,
      dueDate: form.paymentMethod === "Credit" || form.paymentMethod === "Part Payment" ? form.dueDate : undefined,
    });
    setForm((f) => ({ ...f, itemId: "", product: "", category: "", qty: "", unitPrice: "", cost: "", customerName: "", amountPaidNow: "", dueDate: "" }));
  }

  const [tab, setTab] = useState("overview");

  return (
    <DashboardShell
      eyebrow={shop?.name || "Shop"}
      title={`Hi, ${user.name}`}
      onLogout={onLogout}
      tabs={[["overview", "Overview"], ["inventory", "Inventory"], ["expenses", "Expenses"], ["receivables", "Receivables"]]}
      active={tab}
      onChange={setTab}
      actions={
        <button
          onClick={() => exportToExcel(`${user.name}-sales.xlsx`, [
            { name: "My Sales", rows: mySales.map((s) => saleRow(s)) },
            { name: "Shop Inventory", rows: inventory.map((it) => ({ Item: it.name, "In Stock": it.qty, "Cost/unit (₦)": it.cost, "Price/unit (₦)": it.price })) },
            { name: "Shop Expenses", rows: expenses.map((e) => ({ Date: e.date, Description: e.description, Category: e.category, "Amount (₦)": e.amount, "Paid By": e.paidBy, "Receipt #": e.receiptNo, Notes: e.notes })) },
            { name: "Shop Receivables", rows: receivables.map((r) => ({ Date: r.date, Customer: r.customerName, Item: r.item, "Invoice Amt (₦)": r.invoiceAmt, "Amount Paid (₦)": r.amountPaid, "Balance (₦)": r.invoiceAmt - r.amountPaid, Status: receivableStatus(r) })) },
          ])}
          className="flex items-center gap-1.5 text-sm font-medium border border-slate-300 rounded-lg px-3 py-1.5 text-slate-700 hover:bg-white bg-white/60"
        >
          <Download className="w-4 h-4" /> Download Excel
        </button>
      }
    >
      {tab === "overview" && (
        <>
          <div className="grid grid-cols-2 gap-3 mb-6">
            <div className="bg-white border border-slate-200 rounded-xl p-4">
              <p className="text-xs text-slate-400">My sales today</p>
              <p className="text-xl font-semibold figure text-slate-900">{naira(totalToday)}</p>
            </div>
            <div className="bg-white border border-slate-200 rounded-xl p-4">
              <p className="text-xs text-slate-400">My total records</p>
              <p className="text-xl font-semibold figure text-slate-900">{mySales.length}</p>
            </div>
          </div>

          <form onSubmit={submit} className="bg-white border border-slate-200 rounded-xl p-4 mb-6">
            <p className="text-sm font-semibold text-slate-800 mb-3 flex items-center gap-1"><Plus className="w-4 h-4" /> Log a sale</p>
            <div className="grid grid-cols-2 gap-2 mb-2">
              <input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm col-span-2 sm:col-span-1" />
              <select value={form.paymentMethod} onChange={(e) => setForm({ ...form, paymentMethod: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm">
                <option>Cash</option><option>Transfer</option><option>POS</option><option>Credit</option><option>Part Payment</option>
              </select>
              <input placeholder="Customer name (optional)" value={form.customerName} onChange={(e) => setForm({ ...form, customerName: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm col-span-2" />
              <select value={form.itemId || "__custom__"} onChange={(e) => pickInventoryItem(e.target.value)} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm col-span-2">
                <option value="__custom__">— Custom item (not tracked in stock) —</option>
                {inventory.map((it) => <option key={it.id} value={it.id}>{it.name} ({it.qty} in stock)</option>)}
              </select>
              <input placeholder="Product / item" value={form.product} onChange={(e) => setForm({ ...form, product: e.target.value, itemId: "" })} disabled={!!form.itemId} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm col-span-2 disabled:bg-slate-50 disabled:text-slate-500" />
              <input placeholder="Category" value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm col-span-2" />
              <input type="number" placeholder="Qty sold" value={form.qty} onChange={(e) => setForm({ ...form, qty: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm" />
              <input type="number" placeholder="Unit price (₦)" value={form.unitPrice} onChange={(e) => setForm({ ...form, unitPrice: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm" />
              <input type="number" placeholder="Cost per item (₦)" value={form.cost} onChange={(e) => setForm({ ...form, cost: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm col-span-2" />
              {(form.paymentMethod === "Credit" || form.paymentMethod === "Part Payment") && (
                <>
                  {form.paymentMethod === "Part Payment" && (
                    <input type="number" placeholder="Amount paid now (₦)" value={form.amountPaidNow} onChange={(e) => setForm({ ...form, amountPaidNow: e.target.value })} className="border border-amber-300 rounded-lg px-2 py-1.5 text-sm" />
                  )}
                  <input type="date" placeholder="Due date" value={form.dueDate} onChange={(e) => setForm({ ...form, dueDate: e.target.value })} className={`border border-amber-300 rounded-lg px-2 py-1.5 text-sm ${form.paymentMethod === "Credit" ? "col-span-2" : ""}`} />
                </>
              )}
            </div>
            <button type="submit" className="w-full text-white rounded-lg py-2 text-sm font-medium mt-1" style={{ background: ORANGE }}>Add Sale</button>
          </form>

          <p className="text-sm font-semibold text-slate-800 mb-2">My sales history</p>
          <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
            <table className="w-full text-sm">
              <thead className="bg-slate-50 text-slate-500 text-xs">
                <tr><th className="text-left px-3 py-2">Date</th><th className="text-left px-3 py-2">Item</th><th className="text-right px-3 py-2">Qty</th><th className="text-right px-3 py-2">Total</th><th className="px-3 py-2"></th></tr>
              </thead>
              <tbody>
                {mySales.length === 0 && <tr><td colSpan={5} className="text-center text-slate-400 py-6">No sales logged yet.</td></tr>}
                {mySales.map((s) => (
                  <tr key={s.id} className="border-t border-slate-100">
                    <td className="px-3 py-2">{s.date}</td>
                    <td className="px-3 py-2">{s.product}</td>
                    <td className="px-3 py-2 text-right">{s.qty}</td>
                    <td className="px-3 py-2 text-right font-medium">{naira(s.total)}</td>
                    <td className="px-3 py-2 text-right"><button onClick={() => openInvoice(s, shop?.name || "", null)} className="text-xs text-slate-400 hover:text-slate-700">Invoice</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}

      {tab === "inventory" && (
        <InventoryManager
          items={inventory}
          onAddItem={onAddItem}
          onUpdateItem={onUpdateItem}
          onRemoveItem={onRemoveItem}
          accent={ORANGE}
          makeItem={(data) => ({ id: uid(), kind: "team", ownerId: shop.id, ...data })}
        />
      )}

      {tab === "expenses" && (
        <ExpenseManager
          expenses={expenses}
          onAddExpense={onAddExpense}
          onRemoveExpense={onRemoveExpense}
          accent={ORANGE}
          makeExpense={(data) => ({ id: uid(), kind: "team", ownerId: shop.id, ...data, createdBy: user.name })}
        />
      )}

      {tab === "receivables" && (
        <ReceivablesManager receivables={receivables} onRecordPayment={onRecordPayment} accent={ORANGE} />
      )}
    </DashboardShell>
  );
}

// ============================================================== INVENTORY
// Shared by SoloDashboard, WorkerDashboard, and (per shop) BossDashboard.
function InventoryManager({ items, onAddItem, onUpdateItem, onRemoveItem, accent, makeItem }) {
  const [adding, setAdding] = useState(false);
  const [form, setForm] = useState({ name: "", qty: "", cost: "", price: "" });
  const [editingQty, setEditingQty] = useState(null);
  const [qtyDraft, setQtyDraft] = useState("");

  function submit() {
    if (!form.name.trim()) return;
    const item = makeItem({
      name: form.name.trim(),
      qty: Number(form.qty) || 0,
      cost: Number(form.cost) || 0,
      price: Number(form.price) || 0,
    });
    onAddItem(item);
    setForm({ name: "", qty: "", cost: "", price: "" });
    setAdding(false);
  }

  return (
    <div className="mb-6">
      <div className="flex items-center justify-between mb-2">
        <p className="text-sm font-semibold text-slate-800">Inventory</p>
        {!adding && (
          <button onClick={() => setAdding(true)} className="flex items-center gap-1 text-xs font-medium text-white rounded-lg px-2.5 py-1" style={{ background: accent }}>
            <Plus className="w-3.5 h-3.5" /> Add Item
          </button>
        )}
      </div>
      {adding && (
        <div className="bg-white border border-slate-200 rounded-xl p-3 mb-3 grid grid-cols-2 gap-2">
          <input autoFocus placeholder="Item name" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm col-span-2" />
          <input type="number" placeholder="Starting stock (qty)" value={form.qty} onChange={(e) => setForm({ ...form, qty: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm" />
          <input type="number" placeholder="Cost/unit (₦)" value={form.cost} onChange={(e) => setForm({ ...form, cost: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm" />
          <input type="number" placeholder="Selling price/unit (₦)" value={form.price} onChange={(e) => setForm({ ...form, price: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm col-span-2" />
          <div className="col-span-2 flex gap-2">
            <button onClick={submit} className="flex-1 text-white rounded-lg py-1.5 text-sm font-medium" style={{ background: accent }}>Save Item</button>
            <button onClick={() => setAdding(false)} className="px-3 text-sm text-slate-500">Cancel</button>
          </div>
        </div>
      )}
      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        {items.length === 0 && !adding && (
          <p className="text-center text-sm text-slate-400 py-6">No items yet — click "Add Item" to start tracking stock.</p>
        )}
        {items.length > 0 && (
          <table className="w-full text-sm">
            <thead className="bg-slate-50 text-slate-500 text-xs">
              <tr><th className="text-left px-3 py-2">Item</th><th className="text-right px-3 py-2">In Stock</th><th className="text-right px-3 py-2">Price</th><th className="px-3 py-2"></th></tr>
            </thead>
            <tbody>
              {items.map((it) => (
                <tr key={it.id} className="border-t border-slate-100">
                  <td className="px-3 py-2">{it.name}</td>
                  <td className="px-3 py-2 text-right">
                    {editingQty === it.id ? (
                      <span className="inline-flex gap-1 items-center">
                        <input type="number" value={qtyDraft} onChange={(e) => setQtyDraft(e.target.value)} className="w-16 border border-slate-300 rounded px-1 py-0.5 text-xs text-right" autoFocus />
                        <button onClick={() => { onUpdateItem(it.id, { qty: Number(qtyDraft) || 0 }); setEditingQty(null); }} className="text-xs px-1.5 py-0.5 rounded text-white" style={{ background: accent }}>Save</button>
                      </span>
                    ) : (
                      <button onClick={() => { setEditingQty(it.id); setQtyDraft(String(it.qty)); }} className={`font-medium hover:underline ${it.qty <= 3 ? "text-red-600" : "text-slate-800"}`}>
                        {it.qty}{it.qty <= 3 ? " ⚠" : ""}
                      </button>
                    )}
                  </td>
                  <td className="px-3 py-2 text-right text-slate-500">{naira(it.price)}</td>
                  <td className="px-3 py-2 text-right">
                    <button onClick={() => { if (window.confirm(`Remove "${it.name}" from inventory?`)) onRemoveItem(it.id); }} className="text-xs text-red-400 hover:text-red-600">Remove</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}


// ============================================================== EXPENSES
const EXPENSE_CATEGORIES = ["Stock/Inventory", "Rent & Utilities", "Salaries & Wages", "Transport", "Marketing", "Miscellaneous"];
function ExpenseManager({ expenses, onAddExpense, onRemoveExpense, accent, makeExpense }) {
  const [adding, setAdding] = useState(false);
  const [form, setForm] = useState({ date: new Date().toISOString().slice(0, 10), description: "", category: EXPENSE_CATEGORIES[0], amount: "", paidBy: "Cash", receiptNo: "", notes: "" });
  const total = expenses.reduce((a, e) => a + e.amount, 0);

  function submit() {
    if (!form.description.trim() || !(Number(form.amount) > 0)) return;
    onAddExpense(makeExpense({ ...form, amount: Number(form.amount) }));
    setForm((f) => ({ ...f, description: "", amount: "", receiptNo: "", notes: "" }));
    setAdding(false);
  }

  return (
    <div className="mb-6">
      <div className="flex items-center justify-between mb-2">
        <p className="text-sm font-semibold text-slate-800">Expenses <span className="text-xs font-normal text-slate-400">({naira(total)} total)</span></p>
        {!adding && (
          <button onClick={() => setAdding(true)} className="flex items-center gap-1 text-xs font-medium text-white rounded-lg px-2.5 py-1" style={{ background: accent }}>
            <Plus className="w-3.5 h-3.5" /> Add Expense
          </button>
        )}
      </div>
      {adding && (
        <div className="bg-white border border-slate-200 rounded-xl p-3 mb-3 grid grid-cols-2 gap-2">
          <input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm" />
          <select value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm">
            {EXPENSE_CATEGORIES.map((c) => <option key={c}>{c}</option>)}
          </select>
          <input autoFocus placeholder="Description" value={form.description} onChange={(e) => setForm({ ...form, description: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm col-span-2" />
          <input type="number" placeholder="Amount (₦)" value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm" />
          <select value={form.paidBy} onChange={(e) => setForm({ ...form, paidBy: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm">
            <option>Cash</option><option>Transfer</option><option>POS</option>
          </select>
          <input placeholder="Receipt # (optional)" value={form.receiptNo} onChange={(e) => setForm({ ...form, receiptNo: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm" />
          <input placeholder="Notes (optional)" value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} className="border border-slate-300 rounded-lg px-2 py-1.5 text-sm" />
          <div className="col-span-2 flex gap-2">
            <button onClick={submit} className="flex-1 text-white rounded-lg py-1.5 text-sm font-medium" style={{ background: accent }}>Save Expense</button>
            <button onClick={() => setAdding(false)} className="px-3 text-sm text-slate-500">Cancel</button>
          </div>
        </div>
      )}
      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        {expenses.length === 0 && !adding && <p className="text-center text-sm text-slate-400 py-6">No expenses logged yet.</p>}
        {expenses.length > 0 && (
          <table className="w-full text-sm">
            <thead className="bg-slate-50 text-slate-500 text-xs">
              <tr><th className="text-left px-3 py-2">Date</th><th className="text-left px-3 py-2">Description</th><th className="text-left px-3 py-2">Category</th><th className="text-right px-3 py-2">Amount</th><th className="px-3 py-2"></th></tr>
            </thead>
            <tbody>
              {expenses.slice().sort((a, b) => (a.date < b.date ? 1 : -1)).map((e) => (
                <tr key={e.id} className="border-t border-slate-100">
                  <td className="px-3 py-2">{e.date}</td>
                  <td className="px-3 py-2">{e.description}</td>
                  <td className="px-3 py-2 text-slate-500">{e.category}</td>
                  <td className="px-3 py-2 text-right font-medium">{naira(e.amount)}</td>
                  <td className="px-3 py-2 text-right"><button onClick={() => onRemoveExpense(e.id)} className="text-xs text-red-400 hover:text-red-600">Remove</button></td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

// ============================================================== RECEIVABLES
function ReceivablesManager({ receivables, onRecordPayment, accent }) {
  const [payingId, setPayingId] = useState(null);
  const [amountDraft, setAmountDraft] = useState("");
  const totalOwed = receivables.reduce((a, r) => a + Math.max(0, r.invoiceAmt - r.amountPaid), 0);
  const badgeTone = { Paid: "green", Pending: "gold", Overdue: "rust" };

  return (
    <div className="mb-6">
      <p className="text-sm font-semibold text-slate-800 mb-2">Receivables (credit sales) <span className="text-xs font-normal text-slate-400">({naira(totalOwed)} outstanding)</span></p>
      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        {receivables.length === 0 && <p className="text-center text-sm text-slate-400 py-6">No credit sales yet — these appear automatically when a sale is logged as Credit or Part Payment.</p>}
        {receivables.length > 0 && (
          <table className="w-full text-sm">
            <thead className="bg-slate-50 text-slate-500 text-xs">
              <tr><th className="text-left px-3 py-2">Date</th><th className="text-left px-3 py-2">Customer</th><th className="text-left px-3 py-2">Item</th><th className="text-right px-3 py-2">Owed</th><th className="text-center px-3 py-2">Status</th><th className="px-3 py-2"></th></tr>
            </thead>
            <tbody>
              {receivables.slice().sort((a, b) => (a.date < b.date ? 1 : -1)).map((r) => {
                const status = receivableStatus(r);
                const balance = r.invoiceAmt - r.amountPaid;
                return (
                  <tr key={r.id} className="border-t border-slate-100">
                    <td className="px-3 py-2">{r.date}</td>
                    <td className="px-3 py-2">{r.customerName}</td>
                    <td className="px-3 py-2">{r.item}</td>
                    <td className="px-3 py-2 text-right font-medium">{naira(balance)}</td>
                    <td className="px-3 py-2 text-center"><StampBadge label={status} tone={badgeTone[status]} /></td>
                    <td className="px-3 py-2 text-right">
                      {status !== "Paid" && (
                        payingId === r.id ? (
                          <span className="inline-flex gap-1 items-center">
                            <input type="number" placeholder="₦ paid" value={amountDraft} onChange={(e) => setAmountDraft(e.target.value)} className="w-20 border border-slate-300 rounded px-1 py-0.5 text-xs" autoFocus />
                            <button onClick={() => { const amt = Number(amountDraft) || 0; if (amt > 0) { onRecordPayment(r.id, Math.min(amt, balance)); } setPayingId(null); setAmountDraft(""); }} className="text-xs px-1.5 py-0.5 rounded text-white" style={{ background: accent }}>Save</button>
                          </span>
                        ) : (
                          <button onClick={() => setPayingId(r.id)} className="text-xs text-slate-400 hover:text-slate-700">Record Payment</button>
                        )
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

// ============================================================== FINANCIALS (Cash Flow / VAT / P&L)
// Boss-level (and solo-owner-level) reporting — fully automatic from sales + expenses, no manual monthly entry needed.
function FinancialsView({ salesRows, expenseRows, receivableRows, accent }) {
  const [tab, setTab] = useState("cashflow");
  const { perMonth, expenseByCategory, receivablesCollectedByMonth } = useMemo(
    () => buildFinancials(salesRows, expenseRows, receivableRows),
    [salesRows, expenseRows, receivableRows]
  );

  const totalRevenue = salesRows.reduce((a, s) => a + s.total, 0);
  const totalCogs = salesRows.reduce((a, s) => a + s.cost * s.qty, 0);
  const grossProfit = totalRevenue - totalCogs;
  const totalOpEx = expenseRows.reduce((a, e) => a + e.amount, 0);
  const netProfit = grossProfit - totalOpEx;

  return (
    <div className="mb-6">
      <div className="flex items-center justify-between mb-2">
        <p className="text-sm font-semibold text-slate-800">Financial Reports</p>
        <div className="flex rounded-lg bg-slate-100 p-1 text-xs">
          {[["cashflow", "Cash Flow"], ["vat", "VAT Tracker"], ["pl", "P&L Statement"]].map(([key, label]) => (
            <button key={key} onClick={() => setTab(key)} className={`px-2.5 py-1 rounded-md ${tab === key ? "bg-white shadow font-medium" : "text-slate-500"}`}>{label}</button>
          ))}
        </div>
      </div>

      {perMonth.length === 0 ? (
        <div className="bg-white border border-slate-200 rounded-xl p-6 text-center text-sm text-slate-400">
          No dated sales or expenses yet — reports build automatically as you log sales and expenses.
        </div>
      ) : tab === "cashflow" ? (
        <div className="bg-white border border-slate-200 rounded-xl overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="bg-slate-50 text-slate-500 text-xs">
              <tr><th className="text-left px-3 py-2">Month</th><th className="text-right px-3 py-2">Sales Cash In</th><th className="text-right px-3 py-2">Receivables Collected</th><th className="text-right px-3 py-2">Total Inflow</th><th className="text-right px-3 py-2">Expenses (Outflow)</th><th className="text-right px-3 py-2">Net Cash Flow</th></tr>
            </thead>
            <tbody>
              {perMonth.map((m) => {
                const recCollected = receivablesCollectedByMonth[m.mk] || 0;
                const inflow = m.cashCollected + recCollected;
                return (
                  <tr key={m.mk} className="border-t border-slate-100">
                    <td className="px-3 py-2 font-medium">{m.month}</td>
                    <td className="px-3 py-2 text-right">{naira(m.cashCollected)}</td>
                    <td className="px-3 py-2 text-right">{naira(recCollected)}</td>
                    <td className="px-3 py-2 text-right font-medium" style={{ color: NAVY }}>{naira(inflow)}</td>
                    <td className="px-3 py-2 text-right text-red-600">{naira(m.expenseTotal)}</td>
                    <td className={`px-3 py-2 text-right font-semibold ${inflow - m.expenseTotal >= 0 ? "text-emerald-600" : "text-red-600"}`}>{naira(inflow - m.expenseTotal)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      ) : tab === "vat" ? (
        <div className="bg-white border border-slate-200 rounded-xl overflow-x-auto">
          <p className="text-xs text-slate-400 px-3 pt-2">Nigeria FIRS rate: 7.5%. "Purchases" = expenses categorized as Stock/Inventory.</p>
          <table className="w-full text-sm">
            <thead className="bg-slate-50 text-slate-500 text-xs">
              <tr><th className="text-left px-3 py-2">Month</th><th className="text-right px-3 py-2">Sales Revenue</th><th className="text-right px-3 py-2">VAT Collected</th><th className="text-right px-3 py-2">Purchases</th><th className="text-right px-3 py-2">VAT Paid</th><th className="text-right px-3 py-2">Net VAT Payable</th></tr>
            </thead>
            <tbody>
              {perMonth.map((m) => (
                <tr key={m.mk} className="border-t border-slate-100">
                  <td className="px-3 py-2 font-medium">{m.month}</td>
                  <td className="px-3 py-2 text-right">{naira(m.revenue)}</td>
                  <td className="px-3 py-2 text-right">{naira(m.vatCollected)}</td>
                  <td className="px-3 py-2 text-right">{naira(m.purchases)}</td>
                  <td className="px-3 py-2 text-right">{naira(m.vatPaid)}</td>
                  <td className="px-3 py-2 text-right font-semibold" style={{ color: accent }}>{naira(m.netVat)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : (
        <div className="bg-white border border-slate-200 rounded-xl p-4 text-sm">
          <div className="flex justify-between py-1"><span className="text-slate-500">Total Sales Revenue</span><span className="font-medium">{naira(totalRevenue)}</span></div>
          <div className="flex justify-between py-1"><span className="text-slate-500">Less: Cost of Goods Sold</span><span className="font-medium text-red-600">-{naira(totalCogs)}</span></div>
          <div className="flex justify-between py-1.5 border-t border-slate-200 font-semibold"><span>GROSS PROFIT</span><span style={{ color: NAVY }}>{naira(grossProfit)}</span></div>
          <p className="text-xs font-semibold text-slate-500 uppercase mt-3 mb-1">Operating Expenses</p>
          {Object.entries(expenseByCategory).length === 0 && <p className="text-xs text-slate-400 py-1">No expenses logged yet.</p>}
          {Object.entries(expenseByCategory).map(([cat, amt]) => (
            <div key={cat} className="flex justify-between py-1"><span className="text-slate-500">{cat}</span><span>{naira(amt)}</span></div>
          ))}
          <div className="flex justify-between py-1.5 border-t border-slate-200 font-semibold"><span>TOTAL OPERATING EXPENSES</span><span className="text-red-600">{naira(totalOpEx)}</span></div>
          <div className="flex justify-between py-2 border-t-2 border-slate-300 font-bold text-base mt-1"><span>NET PROFIT</span><span style={{ color: netProfit >= 0 ? "#059669" : "#dc2626" }}>{naira(netProfit)}</span></div>
        </div>
      )}
    </div>
  );
}

function BossDashboard({ shops, users, sales, inventory, expenses, receivables, onLogout, onUpdateShopCode, onAddShop, onRenameShop, onRemoveShop, onResetWorkerPin, onAddItem, onUpdateItem, onRemoveItem, onAddExpense, onRemoveExpense, onRecordPayment }) {
  const [viewingStockFor, setViewingStockFor] = useState(null);
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

  const scopedExpenses = shopFilter === "all" ? expenses : expenses.filter((e) => e.ownerId === shopFilter);
  const scopedReceivables = shopFilter === "all" ? receivables : receivables.filter((r) => r.ownerId === shopFilter);
  const scopedInventory = shopFilter === "all" ? inventory : inventory.filter((it) => it.ownerId === shopFilter);
  const expenseTotal = scopedExpenses.reduce((a, e) => a + e.amount, 0);
  const netProfit = totalProfit - expenseTotal;
  const stockValue = scopedInventory.reduce((a, it) => a + it.qty * it.cost, 0);
  const outstandingReceivables = scopedReceivables.reduce((a, r) => a + Math.max(0, r.invoiceAmt - r.amountPaid), 0);

  function shopName(id) { return shops.find((s) => s.id === id)?.name || id; }
  function shopStats(id) {
    const rows = teamSales.filter((s) => s.shopId === id);
    return {
      revenue: rows.reduce((a, s) => a + s.total, 0),
      profit: rows.reduce((a, s) => a + s.profit, 0),
      workers: users.filter((u) => u.shopId === id).length,
    };
  }

  const [tab, setTab] = useState("shops");

  return (
    <DashboardShell
      eyebrow="All shops"
      title="Boss Dashboard"
      onLogout={onLogout}
      tabs={[["shops", "Shops"], ["workers", "Workers"], ["sales", "Sales"], ["expenses", "Expenses"], ["receivables", "Receivables"], ["reports", "Reports"]]}
      active={tab}
      onChange={setTab}
      actions={
        <button
          onClick={() => {
            const summaryRows = shops.map((s) => {
              const st = shopStats(s.id);
              return { Shop: s.name, Workers: st.workers, "Revenue (₦)": st.revenue, "Profit (₦)": st.profit };
            });
            exportToExcel("boss-dashboard.xlsx", [
              { name: "Summary by Shop", rows: summaryRows },
              { name: "Sales (filtered)", rows: filtered.map((s) => saleRow(s, { shop: true, worker: true, shopName: shopName(s.shopId) })) },
              { name: "Inventory (all shops)", rows: inventory.filter((it) => it.kind === "team").map((it) => ({ Shop: shopName(it.ownerId), Item: it.name, "In Stock": it.qty, "Cost/unit (₦)": it.cost, "Price/unit (₦)": it.price })) },
              { name: "Expenses", rows: scopedExpenses.map((e) => ({ Shop: shopName(e.ownerId), Date: e.date, Description: e.description, Category: e.category, "Amount (₦)": e.amount, "Paid By": e.paidBy, "Receipt #": e.receiptNo })) },
              { name: "Receivables", rows: scopedReceivables.map((r) => ({ Shop: shopName(r.ownerId), Date: r.date, Customer: r.customerName, Item: r.item, "Invoice Amt (₦)": r.invoiceAmt, "Amount Paid (₦)": r.amountPaid, "Balance (₦)": r.invoiceAmt - r.amountPaid, Status: receivableStatus(r) })) },
            ]);
          }}
          className="flex items-center gap-1.5 text-sm font-medium border border-slate-300 rounded-lg px-3 py-1.5 text-slate-700 hover:bg-white bg-white/60"
        >
          <Download className="w-4 h-4" /> Download Excel
        </button>
      }
    >
      {revealShopCode && (
        <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-4 mb-4">
          <p className="text-sm font-semibold text-slate-800 mb-1">"{revealShopCode.name}" created</p>
          <p className="text-xs text-slate-500 mb-2">Give this portal code to that shop's team. It won't be shown again — you can only reset it afterward.</p>
          <div className="flex items-center gap-3">
            <span className="bg-white border border-emerald-300 rounded-lg px-4 py-1.5 figure font-semibold text-slate-900">{revealShopCode.code}</span>
            <button onClick={() => setRevealShopCode(null)} className="text-xs text-slate-500 hover:text-slate-800">Dismiss</button>
          </div>
        </div>
      )}
      {revealWorkerPin && (
        <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-4 mb-4">
          <p className="text-sm font-semibold text-slate-800 mb-1">{revealWorkerPin.name}'s PIN was reset</p>
          <p className="text-xs text-slate-500 mb-2">Give them this new PIN. It won't be shown again.</p>
          <div className="flex items-center gap-3">
            <span className="bg-white border border-emerald-300 rounded-lg px-4 py-1.5 figure font-semibold text-slate-900">{revealWorkerPin.pin}</span>
            <button onClick={() => setRevealWorkerPin(null)} className="text-xs text-slate-500 hover:text-slate-800">Dismiss</button>
          </div>
        </div>
      )}

      {tab === "shops" && (
        <>
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
                  <button onClick={() => setViewingStockFor(viewingStockFor === s.id ? null : s.id)} className="text-[11px] text-slate-400 hover:text-slate-700 mt-2 block">
                    {viewingStockFor === s.id ? "Hide stock ▲" : "Manage stock ▼"}
                  </button>
                  {viewingStockFor === s.id && (
                    <div className="mt-3 -mx-1">
                      <InventoryManager
                        items={inventory.filter((it) => it.kind === "team" && it.ownerId === s.id)}
                        onAddItem={onAddItem}
                        onUpdateItem={onUpdateItem}
                        onRemoveItem={onRemoveItem}
                        accent={NAVY}
                        makeItem={(data) => ({ id: uid(), kind: "team", ownerId: s.id, ...data })}
                      />
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </>
      )}

      {tab === "workers" && (
        <div>
          <p className="text-sm font-semibold text-slate-800 mb-2">Workers</p>
          {users.length === 0 ? (
            <div className="text-center text-sm text-slate-400 py-6 border border-dashed border-slate-300 rounded-xl">
              No workers yet — they'll appear here once they register at a shop.
            </div>
          ) : (
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
          )}
        </div>
      )}

      {tab === "sales" && (
        <>
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

          <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 mb-6">
            <div className="bg-white border border-slate-200 rounded-xl p-4">
              <p className="text-xs text-slate-400 flex items-center gap-1"><TrendingUp className="w-3.5 h-3.5" /> Filtered Revenue</p>
              <p className="text-xl font-semibold figure" style={{ color: NAVY }}>{naira(totalRevenue)}</p>
            </div>
            <div className="bg-white border border-slate-200 rounded-xl p-4">
              <p className="text-xs text-slate-400 flex items-center gap-1"><Package className="w-3.5 h-3.5" /> Items Sold</p>
              <p className="text-xl font-semibold figure text-slate-900">{totalItems}</p>
            </div>
            <div className="bg-white border border-slate-200 rounded-xl p-4">
              <p className="text-xs text-slate-400">Gross Profit</p>
              <p className="text-xl font-semibold figure" style={{ color: ORANGE }}>{naira(totalProfit)}</p>
            </div>
            <div className="bg-white border border-slate-200 rounded-xl p-4">
              <p className="text-xs text-slate-400">Total Expenses</p>
              <p className="text-xl font-semibold figure text-red-600">{naira(expenseTotal)}</p>
            </div>
            <div className="bg-white border border-slate-200 rounded-xl p-4">
              <p className="text-xs text-slate-400">Net Profit</p>
              <p className="text-xl font-semibold figure" style={{ color: netProfit >= 0 ? GREEN : ORANGE }}>{naira(netProfit)}</p>
            </div>
            <div className="bg-white border border-slate-200 rounded-xl p-4">
              <p className="text-xs text-slate-400">Stock Value</p>
              <p className="text-xl font-semibold figure text-slate-900">{naira(stockValue)}</p>
            </div>
            <div className="bg-white border border-slate-200 rounded-xl p-4 col-span-2 sm:col-span-3">
              <p className="text-xs text-slate-400">Outstanding Receivables {shopFilter !== "all" ? `(${shopName(shopFilter)})` : "(all shops)"}</p>
              <p className="text-xl font-semibold figure" style={{ color: GOLD }}>{naira(outstandingReceivables)}</p>
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
                  <th className="px-3 py-2"></th>
                </tr>
              </thead>
              <tbody>
                {filtered.length === 0 && <tr><td colSpan={8} className="text-center text-slate-400 py-8">No sales recorded yet for this filter.</td></tr>}
                {filtered.map((s) => (
                  <tr key={s.id} className="border-t border-slate-100">
                    <td className="px-3 py-2">{s.date}</td>
                    <td className="px-3 py-2">{shopName(s.shopId)}</td>
                    <td className="px-3 py-2">{s.workerName}</td>
                    <td className="px-3 py-2">{s.product}</td>
                    <td className="px-3 py-2 text-right">{s.qty}</td>
                    <td className="px-3 py-2 text-right font-medium">{naira(s.total)}</td>
                    <td className="px-3 py-2 text-right text-emerald-700">{naira(s.profit)}</td>
                    <td className="px-3 py-2 text-right"><button onClick={() => openInvoice(s, "", shopName(s.shopId))} className="text-xs text-slate-400 hover:text-slate-700">Invoice</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}

      {tab === "expenses" && (
        <ExpenseManager
          expenses={scopedExpenses}
          onAddExpense={onAddExpense}
          onRemoveExpense={onRemoveExpense}
          accent={NAVY}
          makeExpense={(data) => ({ id: uid(), kind: "team", ownerId: shopFilter === "all" ? (shops[0]?.id || "unassigned") : shopFilter, ...data, createdBy: "Boss" })}
        />
      )}

      {tab === "receivables" && (
        <ReceivablesManager receivables={scopedReceivables} onRecordPayment={onRecordPayment} accent={NAVY} />
      )}

      {tab === "reports" && (
        <FinancialsView salesRows={filtered} expenseRows={scopedExpenses} receivableRows={scopedReceivables} accent={NAVY} />
      )}
    </DashboardShell>
  );
}
