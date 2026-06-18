import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import {
  LayoutDashboard, Boxes, ArrowLeftRight, Database, Calculator,
  Plus, Trash2, Search, Package, Store, AlertTriangle, Download,
  TrendingUp, TrendingDown, Check, RefreshCw, FileText, Warehouse,
  Building2, Timer, Banknote, FileSignature, X, Pencil,
  Upload, ChevronRight, Users, Settings, ClipboardCheck,
} from "lucide-react";
import * as XLSX from "xlsx";
import {
  ResponsiveContainer, ComposedChart, Bar, Line, XAxis, YAxis,
  Tooltip, Legend, CartesianGrid,
} from "recharts";
import { supabase } from "./supabaseClient";

/* =========================================================================
   운용리스 관리 (리스제공자 / Lessor)
   계층: 고객사(Customer) → 매장(Store)
   계약(Contract): 고객사·계약번호·기간·과금방식(flat 월정액 | usage 사용량)
                   - usage: 품목코드별 일당단가표(rates)
   청구: 매장 → 고객사 → 그달 유효 계약 → 방식대로 산정
   회계(K-IFRS 1116 운용리스): 자산 보유→정액 감가상각, 리스료 정액 수익인식, 분개 자동
   저장: window.storage
   ========================================================================= */

const STORAGE_KEY = "lease-data-v4";
const WAREHOUSE = "__WH__";
const CWH_PREFIX = "cwh:";                        // 고객사 창고(보관) 위치 접두사
const isCwh = (loc) => typeof loc === "string" && loc.startsWith(CWH_PREFIX);
const cwhCust = (loc) => loc.slice(CWH_PREFIX.length);
const cwhId = (custId) => CWH_PREFIX + custId;

const fmt = (n) => (Math.round(Number(n) || 0)).toLocaleString("ko-KR");
const fmt1 = (n) => (Number(n) || 0).toLocaleString("ko-KR", { maximumFractionDigits: 1 });

const todayISO = () => new Date().toISOString().slice(0, 10);
const thisPeriod = () => new Date().toISOString().slice(0, 7);
const uid = () => Math.random().toString(36).slice(2, 10);
const parseD = (s) => { const [y, m, d] = s.slice(0, 10).split("-").map(Number); return new Date(y, m - 1, d); };

function monthsElapsed(startISO, period) {
  if (!startISO) return 0;
  const s = new Date(startISO.slice(0, 7) + "-01");
  const p = new Date(period + "-01");
  return (p.getFullYear() - s.getFullYear()) * 12 + (p.getMonth() - s.getMonth()) + 1;
}
function periodBounds(period) {
  const [y, m] = period.split("-").map(Number);
  const last = new Date(y, m, 0).getDate();
  return { start: `${period}-01`, end: `${period}-${String(last).padStart(2, "0")}` };
}

const DEFAULT_ACCOUNTS = {
  arLease: "미수리스료", leaseRevenue: "리스료수익",
  assetCost: "리스자산(취득원가)", payable: "미지급금",
  depExpense: "감가상각비", accDep: "감가상각누계액", disposalLoss: "유형자산처분손실",
  reimbAR: "미수금(변상)",
};
const DEFAULT_CURRENCIES = [
  { code: "KRW", name: "원", symbol: "₩", quoteUnit: 1, base: true },
  { code: "USD", name: "미국 달러", symbol: "$", quoteUnit: 1 },
  { code: "JPY", name: "일본 엔", symbol: "¥", quoteUnit: 100 },
  { code: "CNH", name: "위안화(역외)", symbol: "¥", quoteUnit: 1 },
];
function toKRW(amount, code, period, currencies, fxRates) {
  const cur = (currencies || []).find((c) => c.code === code);
  if (!cur || cur.base) return { krw: amount, ok: true, rate: 1 };
  const rate = Number(((fxRates || {})[code] || {})[period]);
  if (!rate || rate <= 0) return { krw: 0, ok: false, rate: 0 };
  return { krw: amount * rate / (Number(cur.quoteUnit) || 1), ok: true, rate };
}

const BILLING = { flat: "월정액", usage: "사용량" };
const TXN_LABEL = { issue: "불출", return: "반납", transfer: "이동", scrap: "폐기" };
const TXN_TONE = {
  issue: "text-emerald-700 bg-emerald-50 ring-emerald-600/20",
  return: "text-slate-700 bg-slate-100 ring-slate-500/20",
  transfer: "text-violet-700 bg-violet-50 ring-violet-600/20",
  scrap: "text-rose-700 bg-rose-50 ring-rose-600/20",
};

/* ---------- 위치별 재고 ---------- */
function computeInventory(items, transactions) {
  const inv = {};
  for (const it of items) inv[it.id] = { [WAREHOUSE]: Number(it.acquiredQty) || 0 };
  const sorted = [...transactions].sort((a, b) => (a.date < b.date ? -1 : 1));
  for (const t of sorted) {
    if (!inv[t.itemId]) inv[t.itemId] = { [WAREHOUSE]: 0 };
    const q = Number(t.qty) || 0, map = inv[t.itemId];
    const dec = (loc) => (map[loc] = (map[loc] || 0) - q);
    const inc = (loc) => (map[loc] = (map[loc] || 0) + q);
    if (t.type === "issue") { dec(t.fromStoreId || WAREHOUSE); inc(t.toStoreId); }
    else if (t.type === "return") { dec(t.fromStoreId); inc(t.toStoreId || WAREHOUSE); }
    else if (t.type === "transfer") { dec(t.fromStoreId); inc(t.toStoreId); }
    else if (t.type === "scrap") { dec(t.fromStoreId || WAREHOUSE); }
    else if (t.type === "adjust") { inc(t.toStoreId || WAREHOUSE); }
  }
  return inv;
}

/* ---------- tag-day: 고객사×아이템별 (매장분/고객사창고분 분리 적분) ---------- */
function tagDaysDetail(stores, transactions, period) {
  const [y, m] = period.split("-").map(Number);
  const monthStart = new Date(y, m - 1, 1);
  const monthEnd = new Date(y, m, 1);
  const DAY = 86400000;
  const storeMap = Object.fromEntries(stores.map((s) => [s.id, s]));
  const locOf = (loc) => {                          // 위치 → { customerId, kind } | null(우리창고·무효)
    if (!loc || loc === WAREHOUSE) return null;
    if (isCwh(loc)) return { customerId: cwhCust(loc), kind: "cwh" };
    const s = storeMap[loc];
    if (!s || !s.customerId) return null;
    return { customerId: s.customerId, kind: "store" };
  };
  const ev = {};
  const push = (loc, i, dateISO, delta) => {
    if (!loc || loc === WAREHOUSE) return;          // 우리 창고만 항상 제외
    const k = loc + "|" + i;
    (ev[k] = ev[k] || []).push({ d: parseD(dateISO), delta });
  };
  for (const t of transactions) {
    const q = Number(t.qty) || 0;
    if (t.type === "issue") push(t.toStoreId, t.itemId, t.date, +q);
    else if (t.type === "return") push(t.fromStoreId, t.itemId, t.date, -q);
    else if (t.type === "transfer") { push(t.fromStoreId, t.itemId, t.date, -q); push(t.toStoreId, t.itemId, t.date, +q); }
    else if (t.type === "scrap") push(t.fromStoreId, t.itemId, t.date, -q);
    else if (t.type === "adjust") push(t.toStoreId, t.itemId, t.date, +q);
  }
  const detail = {}; // { customerId: { itemId: { store, cwh } } }
  const byLoc = {};  // { customerId: { locId: { kind, items: { itemId: days } } } } — 매장별 명세용
  for (const k in ev) {
    const sep = k.lastIndexOf("|");
    const loc = k.slice(0, sep), itemId = k.slice(sep + 1);
    const lc = locOf(loc);
    if (!lc) continue;
    const events = ev[k].sort((a, b) => a.d - b.d);
    let bal = 0, i = 0;
    for (; i < events.length && events[i].d < monthStart; i++) bal += events[i].delta;
    let cursor = monthStart, days = 0;
    for (; i < events.length && events[i].d < monthEnd; i++) {
      days += bal * ((events[i].d - cursor) / DAY);
      cursor = events[i].d;
      bal += events[i].delta;
    }
    days += bal * ((monthEnd - cursor) / DAY);
    if (days === 0) continue;
    const c = lc.customerId;
    if (!detail[c]) detail[c] = {};
    if (!detail[c][itemId]) detail[c][itemId] = { store: 0, cwh: 0 };
    detail[c][itemId][lc.kind] += days;
    if (!byLoc[c]) byLoc[c] = {};
    if (!byLoc[c][loc]) byLoc[c][loc] = { kind: lc.kind, items: {} };
    byLoc[c][loc].items[itemId] = (byLoc[c][loc].items[itemId] || 0) + days;
  }
  return { detail, byLoc };
}

/* ---------- 그 달 유효 계약 ---------- */
function isActiveContract(c, period) {
  const { start, end } = periodBounds(period);
  const s = (c.startDate || "").slice(0, 10);
  const e = (c.endDate || "").slice(0, 10);
  if (s && s > end) return false;     // 시작이 월말보다 뒤 → 아직 미개시
  if (e && e < start) return false;   // 종료가 월초보다 앞 → 만료
  return true;
}

/* ---------- 회계 (운용리스) ---------- */
function periodAccounting(items, stores, customers, contracts, transactions, period, currencies = DEFAULT_CURRENCIES, fxRates = {}, claims = []) {
  let depExpense = 0, depAccum = 0, assetCost = 0, disposalLoss = 0, scrapCost = 0, scrapAccDep = 0, acquisitionCost = 0;
  const [py, pm] = period.split("-").map(Number);
  const pStart = new Date(py, pm - 1, 1), pEnd = new Date(py, pm, 1); // pEnd: 다음 달 1일(미포함)
  const scrapCum = {}, scrapThis = {};                                // 월말까지 누적 폐기 / 당월 폐기
  for (const t of transactions) {
    if (t.type !== "scrap") continue;
    const d = parseD(t.date), q = Number(t.qty) || 0;
    if (d < pEnd) scrapCum[t.itemId] = (scrapCum[t.itemId] || 0) + q;
    if (d >= pStart && d < pEnd) scrapThis[t.itemId] = (scrapThis[t.itemId] || 0) + q;
  }
  const itemBook = {};                                                // 아이템별 태그 1개당 현재 장부가액
  for (const it of items) {
    const N0 = Number(it.acquiredQty) || 0, unit = Number(it.unitCost) || 0, life = Number(it.usefulLifeMonths) || 0;
    const aliveEnd = Math.max(0, N0 - (scrapCum[it.id] || 0));        // 월말 살아있는 수량
    assetCost += aliveEnd * unit;                                     // 살아있는 자산의 취득원가
    if (it.acquiredDate) { const ad = parseD(it.acquiredDate); if (ad >= pStart && ad < pEnd) acquisitionCost += N0 * unit; } // 당월 신규취득
    const sThis = scrapThis[it.id] || 0;
    if (life > 0) {
      const perUnit = unit / life;
      const elapsed = monthsElapsed(it.acquiredDate, period);
      const capped = Math.max(0, Math.min(elapsed, life));
      if (elapsed >= 1 && elapsed <= life) depExpense += aliveEnd * perUnit;
      depAccum += aliveEnd * perUnit * capped;
      const bpu = Math.max(0, unit - perUnit * capped);
      itemBook[it.id] = bpu;
      if (sThis > 0) { disposalLoss += sThis * bpu; scrapCost += sThis * unit; scrapAccDep += sThis * (unit - bpu); } // 폐기: 처분손실·취득원가·누계액
    } else {
      itemBook[it.id] = unit;
      if (sThis > 0) { disposalLoss += sThis * unit; scrapCost += sThis * unit; }
    }
  }
  const bookValue = assetCost - depAccum;                            // 장부가액(순액)
  const { detail, byLoc } = tagDaysDetail(stores, transactions, period);
  const itemMap = Object.fromEntries(items.map((i) => [i.id, i]));
  const custMap = Object.fromEntries(customers.map((c) => [c.id, c]));
  const storeMap = Object.fromEntries(stores.map((s) => [s.id, s]));
  const billItems = []; // 계약 단위 청구
  let leaseIncome = 0;
  for (const ct of contracts) {
    if (!isActiveContract(ct, period)) continue;
    const cust = custMap[ct.customerId];
    if (!cust) continue;
    const currency = ct.currency || "KRW";
    if (ct.billing === "flat") {
      const amountFx = Number(ct.monthlyFee) || 0;
      const fx = toKRW(amountFx, currency, period, currencies, fxRates);
      billItems.push({ contract: ct, customer: cust, billing: "flat", currency, amountFx, amount: fx.krw, rate: fx.rate, missingFx: !fx.ok, tagDays: 0, lines: [], missingRate: false });
      leaseIncome += fx.krw;
    } else {
      const itemTd = detail[ct.customerId] || {};
      const excl = !!ct.excludeStorage;            // 고객사 창고 보관분 과금 제외 여부
      let amountFx = 0, td = 0, lines = [], missingRate = false, cwhExcluded = 0;
      for (const itemId in itemTd) {
        const it = itemMap[itemId]; if (!it) continue;
        const parts = itemTd[itemId];               // { store, cwh }
        const billDays = parts.store + (excl ? 0 : parts.cwh);
        if (excl) cwhExcluded += parts.cwh;
        if (billDays === 0) continue;
        const rate = Number((ct.rates || {})[itemId]) || 0; // 계약 단가표(p, 계약 통화)
        if (rate === 0 && billDays > 0) missingRate = true;
        const amt = billDays * rate;
        lines.push({ item: it, tagDays: billDays, storeDays: parts.store, cwhDays: parts.cwh, dailyRate: rate, amount: amt });
        td += billDays; amountFx += amt;
      }
      lines.sort((a, b) => b.amount - a.amount);
      const locTd = byLoc[ct.customerId] || {};       // 매장별 명세 (합계는 lines와 동일)
      const storeLines = [];
      for (const locId in locTd) {
        const lc = locTd[locId];
        if (excl && lc.kind === "cwh") continue;       // 창고 제외 계약이면 보관분 스킵
        const sName = lc.kind === "cwh" ? `${cust.name} 창고` : (storeMap[locId]?.name || locId);
        for (const itemId in lc.items) {
          const it = itemMap[itemId]; if (!it) continue;
          const days = lc.items[itemId];
          if (!days) continue;
          const rate = Number((ct.rates || {})[itemId]) || 0;
          storeLines.push({ storeId: locId, storeName: sName, kind: lc.kind, item: it, tagDays: days, dailyRate: rate, amount: days * rate });
        }
      }
      storeLines.sort((a, b) => a.storeName.localeCompare(b.storeName) || b.amount - a.amount);
      const fx = toKRW(amountFx, currency, period, currencies, fxRates);
      billItems.push({ contract: ct, customer: cust, billing: "usage", currency, amountFx, amount: fx.krw, rate: fx.rate, missingFx: !fx.ok, tagDays: td, lines, storeLines, missingRate, excludeStorage: excl, cwhExcluded });
      leaseIncome += fx.krw;
    }
  }
  billItems.sort((a, b) => b.amount - a.amount);

  // 고객사별 손익: 비용 = 당월 감가상각비를 사용량(tag-day) 비중으로 배분 (창고 보관분은 미배분)
  const itemDepMap = {};
  for (const it of items) {
    const life = Number(it.usefulLifeMonths) || 0;
    if (life <= 0) { itemDepMap[it.id] = 0; continue; }
    const aliveEnd = Math.max(0, (Number(it.acquiredQty) || 0) - (scrapCum[it.id] || 0));
    const monthly = aliveEnd * (Number(it.unitCost) || 0) / life;
    const el = monthsElapsed(it.acquiredDate, period);
    itemDepMap[it.id] = (el >= 1 && el <= life) ? monthly : 0;
  }
  const itemTotalTd = {};
  for (const cu in detail) for (const im in detail[cu]) itemTotalTd[im] = (itemTotalTd[im] || 0) + detail[cu][im].store;
  const custRev = {};
  billItems.forEach((b) => { custRev[b.customer.id] = (custRev[b.customer.id] || 0) + b.amount; });
  const custCost = {};
  for (const cu in detail) for (const im in detail[cu]) {
    const tot = itemTotalTd[im]; if (!tot) continue;
    custCost[cu] = (custCost[cu] || 0) + (itemDepMap[im] || 0) * (detail[cu][im].store / tot);
  }
  const custPL = customers.map((c) => {
    const revenue = custRev[c.id] || 0, cost = custCost[c.id] || 0;
    return { customer: c, revenue, cost, profit: revenue - cost };
  }).filter((r) => r.revenue > 0 || r.cost > 0).sort((a, b) => b.profit - a.profit);
  const allocatedCost = Object.values(custCost).reduce((a, b) => a + b, 0);
  const unallocatedCost = depExpense - allocatedCost;

  const reimbClaims = (claims || []).filter((cl) => (cl.date || "").slice(0, 7) === period);
  const reimbAR = reimbClaims.reduce((a, cl) => a + (Number(cl.total) || 0), 0);
  return { depExpense, depAccum, assetCost, bookValue, disposalLoss, scrapCost, scrapAccDep, acquisitionCost, itemBook, reimbAR, reimbClaims,
    leaseIncome, netIncome: leaseIncome - depExpense - disposalLoss + reimbAR,
    billItems, custPL, allocatedCost, unallocatedCost };
}

function buildJournal(acc, period, accounts) {
  const A = { ...DEFAULT_ACCOUNTS, ...(accounts || {}) };
  const lines = [];
  if (acc.leaseIncome > 0)
    lines.push({ memo: `리스료수익 인식 (운용리스)`,
      debit: { account: A.arLease, amount: acc.leaseIncome }, credit: { account: A.leaseRevenue, amount: acc.leaseIncome } });
  lines.push({ memo: `리스자산 신규취득 (당월 취득분)`,
    debit: { account: A.assetCost, amount: acc.acquisitionCost || 0 }, credit: { account: A.payable, amount: acc.acquisitionCost || 0 } });
  if (acc.depExpense > 0)
    lines.push({ memo: `리스자산 감가상각 (정액)`,
      debit: { account: A.depExpense, amount: acc.depExpense }, credit: { account: A.accDep, amount: acc.depExpense } });
  lines.push({ memo: `리스자산 폐기 — 감가상각누계액 제거`,
    debit: { account: A.accDep, amount: acc.scrapAccDep || 0 }, credit: { account: A.assetCost, amount: acc.scrapAccDep || 0 } });
  lines.push({ memo: `리스자산 폐기 — 처분손실 (잔존 장부가)`,
    debit: { account: A.disposalLoss, amount: acc.disposalLoss || 0 }, credit: { account: A.assetCost, amount: acc.disposalLoss || 0 } });
  if (acc.reimbAR > 0)
    lines.push({ memo: `리스자산 분실 변상청구 (잔존가 · 미수금)`,
      debit: { account: A.reimbAR, amount: acc.reimbAR }, credit: { account: A.disposalLoss, amount: acc.reimbAR } });
  return lines;
}

/* ---------- 샘플 ---------- */
function makeSample(custN = 10, storesPerCust = 20, itemN = 200) {
  const items = [];
  for (let i = 1; i <= itemN; i++) {
    items.push({ id: uid(), code: `ESL-${String(i).padStart(4, "0")}`, name: `전자가격표 모델 ${i}`,
      unitCost: [3000, 4500, 6000, 9000, 12000][i % 5], dailyRate: [20, 30, 40, 60, 80][i % 5],
      acquiredQty: 200 + (i % 7) * 50, acquiredDate: "2025-01-01", usefulLifeMonths: 36 });
  }
  const customers = [], stores = [], contracts = [];
  for (let c = 1; c <= custN; c++) {
    const cust = { id: uid(), code: `CUST-${String(c).padStart(3, "0")}`, name: `고객사 ${c}` };
    customers.push(cust);
    const usage = c % 2 === 0;
    const currency = c === 8 ? "USD" : c === 9 ? "JPY" : c === 10 ? "CNH" : "KRW";
    const ct = { id: uid(), customerId: cust.id, contractNo: `CT-2025-${String(c).padStart(3, "0")}`,
      startDate: "2025-01-01", endDate: "2026-12-31", billing: usage ? "usage" : "flat", currency,
      monthlyFee: usage ? 0 : 2000000 + (c % 4) * 500000,
      excludeStorage: c === 2,                       // 고객사 2: 폐점 보관분 과금 면제 계약(데모)
      recipientName: `고객사 ${c} 구매팀`, recipientEmail: `ap${c}@example.com`, recipientTel: `02-1234-${String(1000 + c).slice(-4)}`,
      rates: {} };
    if (usage) items.forEach((it) => { ct.rates[it.id] = it.dailyRate; });
    contracts.push(ct);
    for (let s = 1; s <= storesPerCust; s++) {
      stores.push({ id: uid(), customerId: cust.id,
        code: `ST-${String(c).padStart(2, "0")}${String(s).padStart(2, "0")}`, name: `${cust.name} ${s}호점`, active: true });
    }
  }
  const transactions = [];
  for (let s = 0; s < Math.min(60, stores.length); s++) {
    for (let k = 0; k < 4; k++) {
      const it = items[(s * 3 + k) % items.length];
      transactions.push({ id: uid(), date: "2025-01-05", type: "issue", itemId: it.id,
        qty: 30 + ((s + k) % 4) * 10, fromStoreId: WAREHOUSE, toStoreId: stores[s].id, memo: "초기 불출" });
    }
  }
  // 폐점 보관 데모: 고객사 2 1호점 태그를 고객사 창고로 이동 (해당 계약은 과금 제외)
  if (stores[20] && items[60]) {
    transactions.push({ id: uid(), date: "2025-02-01", type: "transfer", itemId: items[60].id, qty: 30,
      fromStoreId: stores[20].id, toStoreId: cwhId(customers[1].id), memo: "1호점 폐점 — 고객사 보관" });
  }
  const fxRates = { USD: {}, JPY: {}, CNH: {} };
  for (let i = 0; i < 24; i++) {
    const d = new Date(2025, i, 1); const p = d.toISOString().slice(0, 7);
    fxRates.USD[p] = 1300 + (i % 5) * 10;
    fxRates.JPY[p] = 880 + (i % 5) * 8;   // 100엔당
    fxRates.CNH[p] = 180 + (i % 4) * 2;
  }
  return { customers, items, stores, contracts, transactions, currencies: DEFAULT_CURRENCIES, fxRates, closedPeriods: [] };
}

/* ---------- UI 유틸 ---------- */
const Card = ({ children, className = "" }) => (
  <div className={`rounded-xl bg-white ring-1 ring-slate-200/80 shadow-sm ${className}`}>{children}</div>
);
const Btn = ({ children, onClick, tone = "default", className = "", ...p }) => {
  const tones = {
    default: "bg-white text-slate-700 ring-1 ring-slate-300 hover:bg-slate-50",
    primary: "bg-indigo-600 text-white hover:bg-indigo-700 shadow-sm",
    danger: "bg-white text-rose-600 ring-1 ring-rose-200 hover:bg-rose-50",
  };
  return (
    <button onClick={onClick}
      className={`inline-flex items-center gap-1.5 rounded-lg px-3 py-2 text-sm font-medium transition ${tones[tone]} ${className}`} {...p}>
      {children}
    </button>
  );
};
const BillBadge = ({ b }) => (
  <span className={`inline-flex items-center gap-1 rounded px-1.5 py-0.5 text-[11px] font-medium ring-1 ${
    b === "usage" ? "text-indigo-700 bg-indigo-50 ring-indigo-600/20" : "text-teal-700 bg-teal-50 ring-teal-600/20"}`}>
    {b === "usage" ? <Timer size={11} /> : <Banknote size={11} />}{BILLING[b]}
  </span>
);

/* =========================================================================
   메인
   ========================================================================= */
// ===== 사업성 분석 (HaaS 수익률 모델, 엑셀 v4 검증 완료) =====
const BIZ_FX_DEFAULT = { KRW: 1483, USD: 1, EUR: 0.86, CNY: 6.83, JPY: 159.8 };
const BIZ_DEFAULTS = {
  currency: "KRW",
  fxTable: { ...BIZ_FX_DEFAULT },
  models: [
    { name: '1.6"', kind: "HW", mat: 2.29, proc: 0.18, margin: 0.10, term: 36, stores: 2838, perStore: 2300, manualK: 0.075 },
    { name: '2.1"', kind: "HW", mat: 2.49, proc: 0.18, margin: 0.10, term: 36, stores: 2838, perStore: 2700, manualK: 0.084 },
    { name: '2.8"', kind: "HW", mat: 3.07, proc: 0.18, margin: 0.10, term: 36, stores: 2838, perStore: 10000, manualK: 0.097 },
    { name: '2.2" F', kind: "HW", mat: 3.58, proc: 0.18, margin: 0.10, term: 36, stores: 2838, perStore: 100, manualK: 0.125 },
    { name: '32"', kind: "HW", mat: 420, proc: 0.18, margin: 0.10, term: 36, stores: 2838, perStore: 5, manualK: 17.304 },
    { name: "SW", kind: "SW", mat: "", proc: "", margin: "", term: 36, stores: 2838, perStore: 15105, manualK: 0.001 },
  ],
  costRatios: { logistics: 0.005, finance: 0.02, baddebt: 0.01, handling: 0.005, maintenance: 0.005 },
  install: 0, maint: 0,
  rf: 0.03, beta: 0.51, erp: 0.055, kd: 0.045, tax: 0.22, eqW: 0.44, ibr: 0.06,
  pd: 0.015, lgd: 0.45,
};
const bnum = (v) => { const n = Number(v); return isFinite(n) ? n : 0; };
function bpmt(rate, nper, pv) { if (rate === 0) return -pv / nper; const f = Math.pow(1 + rate, nper); return -(pv * f) * rate / (f - 1); }
function bnpvAt(rate, cfs) { return cfs.reduce((s, c, i) => s + c / Math.pow(1 + rate, i), 0); }
function birr(cfs) {
  let lo = -0.9999, hi = 10, flo = bnpvAt(lo, cfs);
  if (!isFinite(flo)) return null;
  for (let k = 0; k < 200; k++) { const mid = (lo + hi) / 2, fm = bnpvAt(mid, cfs); if (flo * fm <= 0) hi = mid; else { lo = mid; flo = fm; } if (hi - lo < 1e-10) break; }
  const r = (lo + hi) / 2; return isFinite(r) ? r : null;
}
function computeBiz(inp) {
  const FX = bnum(inp.fxTable?.[inp.currency]) || 1;
  const cr = inp.costRatios || {};
  const CRS = bnum(cr.logistics) + bnum(cr.finance) + bnum(cr.baddebt) + bnum(cr.handling) + bnum(cr.maintenance);
  const Ke = bnum(inp.rf) + bnum(inp.beta) * bnum(inp.erp);
  const KdA = bnum(inp.kd) * (1 - bnum(inp.tax));
  const WACC = Ke * bnum(inp.eqW) + KdA * (1 - bnum(inp.eqW));
  const t = bnum(inp.tax), PD = bnum(inp.pd), LGD = bnum(inp.lgd);
  const INSTALL = bnum(inp.install), MAINT = bnum(inp.maint);
  const rows = (inp.models || []).map((m) => {
    const I = bnum(m.stores) * bnum(m.perStore);
    const isHW = m.kind === "HW";
    const K = m.manualK;
    const hasK = K !== "" && K !== null && K !== undefined;
    let J = null, L;
    if (isHW) { J = (bnum(m.mat) + bnum(m.proc)) / (1 - CRS - bnum(m.margin)); L = hasK ? bnum(K) : bpmt(WACC / 12, bnum(m.term), -J); }
    else { L = hasK ? bnum(K) : 0; }
    return { ...m, I, J, L, isHW };
  });
  const mm = (term, yr) => Math.max(0, Math.min(12, bnum(term) - yr * 12));
  const rev = [0, 0, 0, 0, 0], dep = [0, 0, 0, 0, 0];
  for (const r of rows) for (let yr = 0; yr < 5; yr++) {
    const m = mm(r.term, yr);
    rev[yr] += r.L * FX * r.I * m;
    if (r.isHW) dep[yr] += (bnum(r.mat) + bnum(r.proc)) * FX * r.I / bnum(r.term) * m;
  }
  const maxTerm = Math.max(1, ...rows.map((r) => bnum(r.term)));
  const install = [INSTALL, 0, 0, 0, 0];
  const maint = [0, 1, 2, 3, 4].map((yr) => MAINT / 12 * Math.max(0, Math.min(12, maxTerm - yr * 12)));
  const gp = rev.map((v, i) => v - dep[i] - install[i] - maint[i]);
  const tax = gp.map((v) => Math.max(v * t, 0));
  const ni = gp.map((v, i) => v - tax[i]);
  const sum = (a) => a.reduce((x, y) => x + y, 0);
  const revSum = sum(rev), depSum = sum(dep), gpSum = sum(gp), niSum = sum(ni);
  const npm = revSum ? niSum / revSum : 0, gpm = revSum ? gpSum / revSum : 0;
  const init = -(rows.filter((r) => r.isHW).reduce((s, r) => s + r.I * (bnum(r.mat) + bnum(r.proc)) * FX, 0)) - INSTALL;
  const ocf = [0, ...[0, 1, 2, 3, 4].map((yr) => ni[yr] + dep[yr])];
  const netcf = [init + ocf[0], ocf[1], ocf[2], ocf[3], ocf[4], ocf[5]];
  const cum = []; { let s = 0; for (const c of netcf) { s += c; cum.push(s); } }
  const NPV = netcf.reduce((s, c, i) => i === 0 ? s + c : s + c / Math.pow(1 + WACC, i), 0);
  const IRR = birr(netcf);
  let PB = null;
  if (cum[1] >= 0) PB = Math.abs(cum[0]) / netcf[1] * 12;
  else if (cum[2] >= 0) PB = 12 + Math.abs(cum[1]) / netcf[2] * 12;
  else if (cum[3] >= 0) PB = 24 + Math.abs(cum[2]) / netcf[3] * 12;
  else if (cum[4] >= 0) PB = 36 + Math.abs(cum[3]) / netcf[4] * 12;
  else if (cum[5] >= 0) PB = 48 + Math.abs(cum[4]) / netcf[5] * 12;
  const expo = rows.reduce((s, r) => s + r.I * r.L * FX * (bnum(r.term) / 12), 0);
  const ECL = expo * PD * LGD;
  const npvRatio = init !== 0 ? NPV / Math.abs(init) : 0;
  const verdict = !(NPV > 0) ? "NO-GO" : (IRR == null || IRR < WACC) ? "재검토" : (PB == null || PB > maxTerm) ? "재검토" : "GO";
  return { FX, CRS, Ke, KdA, WACC, rows, rev, dep, gp, ni, tax, revSum, depSum, gpSum, niSum, npm, gpm, init, ocf, netcf, cum, NPV, IRR, PB, ECL, expo, npvRatio, maxTerm, verdict };
}

export default function LeaseManager() {
  const [tab, setTab] = useState("dash");
  const [mode, setMode] = useState(() => { try { return localStorage.getItem("leaseMode") || null; } catch { return null; } });
  const pickMode = (m) => { setMode(m); try { localStorage.setItem("leaseMode", m); } catch {} };
  const [loading, setLoading] = useState(true);
  const [session, setSession] = useState(null);
  const [authReady, setAuthReady] = useState(false);
  const [data, setData] = useState({ customers: [], items: [], stores: [], contracts: [], transactions: [], currencies: DEFAULT_CURRENCIES, fxRates: {}, closedPeriods: [], supplier: {}, accounts: DEFAULT_ACCOUNTS });
  const [period, setPeriod] = useState(thisPeriod());
  const [toast, setToast] = useState(null);
  const flash = (m) => { setToast(m); setTimeout(() => setToast(null), 2200); };

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => { setSession(session); setAuthReady(true); });
    const { data: sub } = supabase.auth.onAuthStateChange((_e, s) => setSession(s));
    return () => sub.subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (!session) { setLoading(false); return; }
    let alive = true;
    (async () => {
      setLoading(true);
      try {
        const { data: row, error } = await supabase.from("app_state").select("data").eq("id", "main").single();
        if (!error && row && row.data && alive) {
          const d = row.data;
          d.customers = d.customers || []; d.items = d.items || []; d.stores = d.stores || [];
          d.contracts = (d.contracts || []).map((c) => ({ currency: "KRW", ...c }));
          d.transactions = d.transactions || [];
          d.currencies = (d.currencies && d.currencies.length) ? d.currencies : DEFAULT_CURRENCIES;
          d.fxRates = d.fxRates || {}; d.closedPeriods = d.closedPeriods || []; d.supplier = d.supplier || {};
          d.accounts = { ...DEFAULT_ACCOUNTS, ...(d.accounts || {}) };
          setData(d);
        }
      } catch (e) {}
      if (alive) setLoading(false);
    })();
    return () => { alive = false; };
  }, [session]);

  const persist = useCallback(async (next) => {
    setData(next);
    try {
      const { error } = await supabase.from("app_state").upsert({ id: "main", data: next, updated_at: new Date().toISOString() });
      if (error) flash("저장 실패 — 다시 시도해 주세요");
    } catch (e) { flash("저장 실패 — 네트워크를 확인해 주세요"); }
  }, []);

  const { customers, items, stores, contracts, transactions, currencies, fxRates, claims } = data;
  const storeMap = useMemo(() => Object.fromEntries(stores.map((s) => [s.id, s])), [stores]);
  const custMap = useMemo(() => Object.fromEntries(customers.map((c) => [c.id, c])), [customers]);
  const itemMap = useMemo(() => Object.fromEntries(items.map((i) => [i.id, i])), [items]);
  const inventory = useMemo(() => computeInventory(items, transactions), [items, transactions]);
  const acc = useMemo(() => periodAccounting(items, stores, customers, contracts, transactions, period, currencies, fxRates, claims || []),
    [items, stores, customers, contracts, transactions, period, currencies, fxRates, claims]);

  const totalTags = useMemo(() => items.reduce((s, i) => s + (Number(i.acquiredQty) || 0), 0), [items]);
  const issuedTags = useMemo(() => {
    let n = 0;
    for (const id in inventory) for (const loc in inventory[id]) if (loc !== WAREHOUSE) n += inventory[id][loc];
    return n;
  }, [inventory]);

  if (!authReady) return <div className="min-h-screen grid place-items-center bg-slate-50 text-slate-400">불러오는 중…</div>;
  if (!session) return <LoginScreen flash={flash} />;
  if (loading) return <div className="min-h-screen grid place-items-center bg-slate-50 text-slate-400">데이터 불러오는 중…</div>;
  if (!mode) return <ModeSelect email={session.user?.email} onPick={pickMode} onSignOut={() => supabase.auth.signOut()} />;
  if (mode === "biz") return <BizShell email={session.user?.email} onSwitch={() => pickMode("ops")} onSignOut={() => supabase.auth.signOut()} data={data} persist={persist} flash={flash} />;

  const NAV = [
    { id: "dash", label: "대시보드", icon: LayoutDashboard },
    { id: "stock", label: "재고현황", icon: Boxes },
    { id: "txn", label: "입출고", icon: ArrowLeftRight },
    { id: "stock", label: "재고실사", icon: ClipboardCheck },
    { id: "master", label: "마스터", icon: Database },
    { id: "acct", label: "회계·청구", icon: Calculator },
    { id: "settings", label: "설정", icon: Settings },
  ];

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800" style={{ fontFamily: "system-ui, sans-serif" }}>
      <header className="sticky top-0 z-20 border-b border-slate-200 bg-white/90 backdrop-blur">
        <div className="mx-auto max-w-6xl px-4 py-3 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="grid place-items-center h-8 w-8 rounded-lg bg-indigo-600 text-white"><Package size={18} /></div>
            <div>
              <h1 className="text-[15px] font-bold tracking-tight leading-none">운용리스 관리</h1>
              <p className="text-[11px] text-slate-400 mt-0.5">리스제공자 · 고객사 → 매장 · 계약 기반 청구</p>
            </div>
          </div>
          <div className="flex items-center gap-3">
            <PeriodPicker period={period} setPeriod={setPeriod} />
            <div className="flex items-center gap-2 text-xs text-slate-500">
              <button onClick={() => pickMode("biz")} className="rounded-lg ring-1 ring-indigo-200 bg-indigo-50 text-indigo-700 px-2 py-1 hover:bg-indigo-100 whitespace-nowrap font-medium">사업성 분석 →</button>
              <span className="hidden sm:inline max-w-[140px] truncate">{session.user?.email}</span>
              <button onClick={() => supabase.auth.signOut()} className="rounded-lg ring-1 ring-slate-300 px-2 py-1 hover:bg-slate-50 whitespace-nowrap">로그아웃</button>
            </div>
          </div>
        </div>
        <nav className="mx-auto max-w-6xl px-2 flex gap-1 overflow-x-auto">
          {NAV.map((n) => {
            const Icon = n.icon, on = tab === n.id;
            return (
              <button key={n.id} onClick={() => setTab(n.id)}
                className={`flex items-center gap-1.5 px-3 py-2.5 text-sm font-medium border-b-2 -mb-px whitespace-nowrap transition
                  ${on ? "border-indigo-600 text-indigo-700" : "border-transparent text-slate-500 hover:text-slate-700"}`}>
                <Icon size={16} /> {n.label}
              </button>
            );
          })}
        </nav>
      </header>

      <main className="mx-auto max-w-6xl px-4 py-5">
        {customers.length === 0 && items.length === 0 && tab === "dash" ? (
          <EmptyState onSample={() => persist(makeSample())} />
        ) : (
          <>
            {tab === "dash" && <Dashboard {...{ acc, period, totalTags, issuedTags, data, inventory, storeMap }} />}
            {tab === "stock" && <StockView {...{ items, stores, customers, inventory, storeMap, custMap }} />}
            {tab === "txn" && <TxnView {...{ data, persist, flash, storeMap, custMap, itemMap, inventory, userEmail: session.user?.email }} />}
            {tab === "stock" && <StocktakeView {...{ data, persist, flash, inventory, storeMap, custMap, userEmail: session.user?.email }} />}
            {tab === "master" && <MasterView {...{ data, persist, flash, custMap, period }} />}
            {tab === "acct" && <AcctView {...{ acc, period, setPeriod, data, persist, flash }} />}
            {tab === "settings" && <SettingsView {...{ data, persist, flash }} />}
          </>
        )}
      </main>

      {toast && (
        <div className="fixed bottom-5 left-1/2 -translate-x-1/2 z-50 rounded-lg bg-slate-900 text-white text-sm px-4 py-2 shadow-lg">{toast}</div>
      )}
    </div>
  );
}

function PeriodPicker({ period, setPeriod }) {
  return (
    <label className="flex items-center gap-2 text-sm">
      <span className="text-slate-400 text-xs">기준월</span>
      <input type="month" value={period} onChange={(e) => setPeriod(e.target.value)}
        className="rounded-lg ring-1 ring-slate-300 px-2.5 py-1.5 text-sm tabular-nums bg-white" />
    </label>
  );
}

function EmptyState({ onSample }) {
  return (
    <Card className="p-10 text-center max-w-xl mx-auto mt-8">
      <div className="grid place-items-center h-12 w-12 rounded-xl bg-indigo-50 text-indigo-600 mx-auto mb-4"><Database size={22} /></div>
      <h2 className="text-lg font-bold">데이터가 비어 있습니다</h2>
      <p className="text-sm text-slate-500 mt-1.5 leading-relaxed">
        마스터에서 고객사·계약·매장·아이템을 직접 등록하거나,<br />규모를 체감하려면 샘플(고객사 10 · 계약 10 · 매장 200 · 아이템 200)을 생성하세요.
      </p>
      <div className="mt-5 flex justify-center"><Btn tone="primary" onClick={onSample}><RefreshCw size={15} /> 샘플 생성</Btn></div>
    </Card>
  );
}

/* ---------- 대시보드 ---------- */
function Dashboard({ acc, period, totalTags, issuedTags, data, inventory, storeMap }) {
  const { items, stores, customers, contracts, transactions, currencies, fxRates, claims } = data;
  const warehouseTags = totalTags - issuedTags;
  const activeCt = contracts.filter((c) => isActiveContract(c, period)).length;
  const [chartMode, setChartMode] = useState("month");
  const [custSel, setCustSel] = useState("");

  const monthly = useMemo(() => {
    const arr = [];
    const base = new Date(period + "-01");
    for (let i = 11; i >= 0; i--) {
      const d = new Date(base.getFullYear(), base.getMonth() - i, 1);
      const p = d.toISOString().slice(0, 7);
      const a = periodAccounting(items, stores, customers, contracts, transactions, p, currencies, fxRates, claims || []);
      let rev, cost, prof;
      if (custSel) {
        const row = a.custPL.find((x) => x.customer.id === custSel);
        rev = row ? row.revenue : 0; cost = row ? row.cost : 0; prof = rev - cost;
      } else { rev = a.leaseIncome; cost = a.depExpense; prof = a.netIncome; }
      arr.push({ month: p.slice(2), 수익: Math.round(rev), 비용: Math.round(cost), 이익: Math.round(prof) });
    }
    let cr = 0, cc = 0;
    arr.forEach((r) => { cr += r.수익; cc += r.비용; r.누적수익 = cr; r.누적비용 = cc; r.누적이익 = cr - cc; });
    return arr;
  }, [items, stores, customers, contracts, transactions, period, currencies, fxRates, custSel]);

  const perStore = useMemo(() => {
    const m = {};
    for (const id in inventory) for (const loc in inventory[id])
      if (loc !== WAREHOUSE && inventory[id][loc] > 0) m[loc] = (m[loc] || 0) + inventory[id][loc];
    return Object.entries(m).map(([id, qty]) => ({ name: storeMap[id]?.name || id, qty }))
      .sort((a, b) => b.qty - a.qty).slice(0, 8);
  }, [inventory, storeMap]);
  const maxQty = Math.max(1, ...perStore.map((s) => s.qty));

  const custAsset = useMemo(() => {
    const m = {}, book = acc.itemBook || {};
    for (const itemId in inventory) {
      const bpu = book[itemId] || 0; if (!bpu) continue;
      for (const loc in inventory[itemId]) {
        if (loc === WAREHOUSE) continue;
        const qty = inventory[itemId][loc]; if (!qty) continue;
        const cid = isCwh(loc) ? cwhCust(loc) : storeMap[loc]?.customerId;
        if (!cid) continue;
        m[cid] = (m[cid] || 0) + qty * bpu;
      }
    }
    return m;
  }, [inventory, acc, storeMap]);
  const totalAsset = useMemo(() => Object.values(custAsset).reduce((a, b) => a + b, 0), [custAsset]);

  const kpis = [
    { label: "총 리스자산(취득원가)", value: `₩${fmt(acc.assetCost)}`, sub: `장부가 ₩${fmt(acc.bookValue)} · ${fmt(totalTags)} EA · 유효계약 ${activeCt}`, icon: Package, tone: "indigo" },
    { label: `${period} 리스료수익`, value: `₩${fmt(acc.leaseIncome)}`, sub: `고객사 ${customers.length} · 매장 ${stores.length}`, icon: TrendingUp, tone: "emerald" },
    { label: `${period} 감가상각비`, value: `₩${fmt(acc.depExpense)}`, sub: `누계 ₩${fmt(acc.depAccum)}`, icon: TrendingDown, tone: "rose" },
    { label: `${period} 리스손익`, value: `₩${fmt(acc.netIncome)}`, sub: acc.netIncome >= 0 ? "이익" : "손실", icon: Calculator, tone: acc.netIncome >= 0 ? "emerald" : "rose" },
  ];
  const toneBg = { indigo: "bg-indigo-50 text-indigo-600", emerald: "bg-emerald-50 text-emerald-600", rose: "bg-rose-50 text-rose-600" };

  return (
    <div className="space-y-5">
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        {kpis.map((k) => { const Icon = k.icon; return (
          <Card key={k.label} className="p-4">
            <div className={`inline-grid place-items-center h-8 w-8 rounded-lg mb-2.5 ${toneBg[k.tone]}`}><Icon size={16} /></div>
            <div className="text-[11px] text-slate-400 font-medium">{k.label}</div>
            <div className="text-xl font-bold tabular-nums mt-0.5">{k.value}</div>
            <div className="text-[11px] text-slate-400 mt-0.5">{k.sub}</div>
          </Card>
        ); })}
      </div>

      <Card className="p-5">
        <div className="flex flex-wrap items-center justify-center gap-x-5 gap-y-3 text-center">
          <PLItem label={`${period} 리스료수익`} v={acc.leaseIncome} tone="text-emerald-600" />
          <span className="text-slate-300 text-2xl font-light">−</span>
          <PLItem label="감가상각비" v={acc.depExpense} tone="text-rose-600" />
          {acc.disposalLoss > 0 && <>
            <span className="text-slate-300 text-2xl font-light">−</span>
            <PLItem label="폐기 처분손실" v={acc.disposalLoss} tone="text-rose-600" />
          </>}
          <span className="text-slate-300 text-2xl font-light">=</span>
          <PLItem label="리스손익" v={acc.netIncome} tone={acc.netIncome >= 0 ? "text-indigo-600" : "text-rose-600"} big />
        </div>
      </Card>

      <Card className="p-4">
        <div className="flex flex-wrap items-center justify-between gap-2 mb-3">
          <h3 className="text-sm font-bold flex items-center gap-1.5"><TrendingUp size={15} /> {chartMode === "month" ? "월별" : "누적"} 손익 추이 <span className="text-[11px] text-slate-400 font-normal">{custSel ? "· 비용은 배분 추정" : "(최근 12개월)"}</span></h3>
          <div className="flex items-center gap-2">
            <select value={custSel} onChange={(e) => setCustSel(e.target.value)} className="rounded-lg ring-1 ring-slate-300 px-2.5 py-1 text-xs bg-white max-w-[150px]">
              <option value="">전체 합계</option>
              {customers.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
            <div className="flex gap-1 text-xs">
              {[["month", "월별"], ["cum", "누적"]].map(([m, l]) => (
                <button key={m} onClick={() => setChartMode(m)} className={`px-2.5 py-1 rounded-md font-medium ${chartMode === m ? "bg-indigo-600 text-white" : "bg-slate-100 text-slate-500 hover:bg-slate-200"}`}>{l}</button>
              ))}
            </div>
          </div>
        </div>
        <ResponsiveContainer width="100%" height={260}>
          {chartMode === "month" ? (
            <ComposedChart data={monthly} margin={{ top: 6, right: 6, left: 0, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" vertical={false} />
              <XAxis dataKey="month" tick={{ fontSize: 11, fill: "#94a3b8" }} axisLine={{ stroke: "#e2e8f0" }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: "#94a3b8" }} axisLine={false} tickLine={false} tickFormatter={(v) => `${Math.round(v / 10000)}만`} width={44} />
              <Tooltip formatter={(v) => `₩${fmt(v)}`} contentStyle={{ fontSize: 12, borderRadius: 8, border: "1px solid #e2e8f0" }} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Bar dataKey="수익" fill="#10b981" radius={[3, 3, 0, 0]} maxBarSize={22} />
              <Bar dataKey="비용" fill="#fb7185" radius={[3, 3, 0, 0]} maxBarSize={22} />
              <Line type="monotone" dataKey="이익" stroke="#4f46e5" strokeWidth={2} dot={false} />
            </ComposedChart>
          ) : (
            <ComposedChart data={monthly} margin={{ top: 6, right: 6, left: 0, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" vertical={false} />
              <XAxis dataKey="month" tick={{ fontSize: 11, fill: "#94a3b8" }} axisLine={{ stroke: "#e2e8f0" }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: "#94a3b8" }} axisLine={false} tickLine={false} tickFormatter={(v) => `${Math.round(v / 10000)}만`} width={44} />
              <Tooltip formatter={(v) => `₩${fmt(v)}`} contentStyle={{ fontSize: 12, borderRadius: 8, border: "1px solid #e2e8f0" }} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Line type="monotone" dataKey="누적수익" stroke="#10b981" strokeWidth={2} dot={false} />
              <Line type="monotone" dataKey="누적비용" stroke="#fb7185" strokeWidth={2} dot={false} />
              <Line type="monotone" dataKey="누적이익" stroke="#4f46e5" strokeWidth={2.5} dot={false} />
            </ComposedChart>
          )}
        </ResponsiveContainer>
      </Card>

      <Card className="overflow-hidden">
        <div className="px-4 py-3 border-b border-slate-100 flex items-center justify-between">
          <h3 className="text-sm font-bold flex items-center gap-1.5"><Building2 size={15} /> {period} 고객사별 손익</h3>
          <span className="text-[11px] text-slate-400">비용=tag-day 비중 배분 · 자산=배치분 장부가(관리용 참고)</span>
        </div>
        <div className="overflow-x-auto max-h-[40vh]">
          <table className="w-full text-sm">
            <thead className="sticky top-0 bg-slate-50 text-slate-500 text-xs text-left">
              <tr><th className="px-4 py-2 font-medium">고객사</th><th className="px-4 py-2 font-medium text-right">수익</th>
                <th className="px-4 py-2 font-medium text-right">비용(배분)</th><th className="px-4 py-2 font-medium text-right">이익</th>
                <th className="px-4 py-2 font-medium text-right">배치 자산(장부)</th>
                <th className="px-4 py-2 font-medium text-right">이익률</th></tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {acc.custPL.map((r) => (
                <tr key={r.customer.id} className="hover:bg-slate-50/60">
                  <td className="px-4 py-2 font-medium">{r.customer.name}</td>
                  <td className="px-4 py-2 text-right tabular-nums text-emerald-600">₩{fmt(r.revenue)}</td>
                  <td className="px-4 py-2 text-right tabular-nums text-rose-500">₩{fmt(r.cost)}</td>
                  <td className={`px-4 py-2 text-right tabular-nums font-semibold ${r.profit >= 0 ? "text-slate-800" : "text-rose-600"}`}>₩{fmt(r.profit)}</td>
                  <td className="px-4 py-2 text-right tabular-nums text-slate-500">₩{fmt(custAsset[r.customer.id] || 0)}</td>
                  <td className="px-4 py-2 text-right tabular-nums text-slate-400">{r.revenue > 0 ? `${Math.round((r.profit / r.revenue) * 100)}%` : "—"}</td>
                </tr>
              ))}
              {acc.unallocatedCost > 0.5 && (
                <tr className="bg-slate-50/40 text-slate-500">
                  <td className="px-4 py-2 italic">미배분(창고 보관분)</td>
                  <td className="px-4 py-2 text-right tabular-nums">—</td>
                  <td className="px-4 py-2 text-right tabular-nums">₩{fmt(acc.unallocatedCost)}</td>
                  <td className="px-4 py-2 text-right tabular-nums">₩{fmt(-acc.unallocatedCost)}</td>
                  <td className="px-4 py-2 text-right tabular-nums">—</td>
                  <td className="px-4 py-2 text-right">—</td>
                </tr>
              )}
            </tbody>
            <tfoot>
              <tr className="border-t-2 border-slate-200 bg-slate-50/50 font-bold">
                <td className="px-4 py-2.5">합계</td>
                <td className="px-4 py-2.5 text-right tabular-nums text-emerald-700">₩{fmt(acc.leaseIncome)}</td>
                <td className="px-4 py-2.5 text-right tabular-nums text-rose-600">₩{fmt(acc.depExpense)}</td>
                <td className="px-4 py-2.5 text-right tabular-nums">₩{fmt(acc.netIncome)}</td>
                <td className="px-4 py-2.5 text-right tabular-nums text-slate-600">₩{fmt(totalAsset)}</td>
                <td className="px-4 py-2.5 text-right tabular-nums text-slate-400">{acc.leaseIncome > 0 ? `${Math.round((acc.netIncome / acc.leaseIncome) * 100)}%` : "—"}</td>
              </tr>
            </tfoot>
          </table>
          {acc.custPL.length === 0 && <p className="text-center text-sm text-slate-400 py-8">이 달 손익이 있는 고객사가 없습니다.</p>}
        </div>
      </Card>

      <div className="grid lg:grid-cols-3 gap-4">
        <Card className="p-4">
          <h3 className="text-sm font-bold mb-3 flex items-center gap-1.5"><Warehouse size={15} /> 태그 분포</h3>
          <DistRow label="창고 보관(미불출)" qty={warehouseTags} total={totalTags} tone="bg-slate-400" />
          <DistRow label="매장 불출" qty={issuedTags} total={totalTags} tone="bg-indigo-500" />
          <div className="mt-3 pt-3 border-t border-slate-100 grid grid-cols-4 gap-2 text-center">
            <Mini label="고객사" v={customers.length} />
            <Mini label="유효계약" v={activeCt} />
            <Mini label="매장" v={stores.length} />
            <Mini label="아이템" v={items.length} />
          </div>
        </Card>
        <Card className="p-4 lg:col-span-2">
          <h3 className="text-sm font-bold mb-3 flex items-center gap-1.5"><Store size={15} /> 매장별 보유 상위</h3>
          {perStore.length === 0 ? <p className="text-sm text-slate-400 py-6 text-center">아직 불출 내역이 없습니다.</p> : (
            <div className="space-y-2">
              {perStore.map((s) => (
                <div key={s.name} className="flex items-center gap-3">
                  <div className="w-28 text-xs text-slate-500 truncate">{s.name}</div>
                  <div className="flex-1 h-5 rounded bg-slate-100 overflow-hidden">
                    <div className="h-full bg-indigo-500/80 rounded" style={{ width: `${(s.qty / maxQty) * 100}%` }} />
                  </div>
                  <div className="w-14 text-right text-xs font-medium tabular-nums">{fmt(s.qty)}</div>
                </div>
              ))}
            </div>
          )}
        </Card>
      </div>
    </div>
  );
}
const PLItem = ({ label, v, tone, big }) => (
  <div>
    <div className="text-[11px] text-slate-400 font-medium mb-0.5">{label}</div>
    <div className={`tabular-nums font-bold ${tone} ${big ? "text-2xl sm:text-3xl" : "text-lg sm:text-xl"}`}>₩{fmt(v)}</div>
  </div>
);
const DistRow = ({ label, qty, total, tone }) => (
  <div className="mb-2.5">
    <div className="flex justify-between text-xs mb-1">
      <span className="text-slate-500">{label}</span>
      <span className="font-medium tabular-nums">{fmt(qty)} EA · {total ? Math.round((qty / total) * 100) : 0}%</span>
    </div>
    <div className="h-2 rounded-full bg-slate-100 overflow-hidden">
      <div className={`h-full rounded-full ${tone}`} style={{ width: `${total ? (qty / total) * 100 : 0}%` }} />
    </div>
  </div>
);
const Mini = ({ label, v }) => (
  <div><div className="text-base font-bold tabular-nums">{v}</div><div className="text-[10px] text-slate-400">{label}</div></div>
);

/* ---------- 재고현황 ---------- */
function StockView({ items, stores, customers, inventory, storeMap, custMap }) {
  const [view, setView] = useState("customer"); // customer | store | item
  const [q, setQ] = useState("");
  const [custFilter, setCustFilter] = useState("");
  const [storeFilter, setStoreFilter] = useState("");
  const [exp, setExp] = useState({});
  const itemMap = useMemo(() => Object.fromEntries(items.map((i) => [i.id, i])), [items]);
  const toggleExp = (id) => setExp((e) => ({ ...e, [id]: !e[id] }));
  const visibleStores = useMemo(() => stores.filter((s) => !custFilter || s.customerId === custFilter), [stores, custFilter]);

  const byStore = useMemo(() => {
    const m = {};
    for (const itemId in inventory) for (const loc in inventory[itemId]) {
      if (loc === WAREHOUSE) continue;
      const v = inventory[itemId][loc]; if (v === 0) continue;
      (m[loc] = m[loc] || {})[itemId] = v;
    }
    return m;
  }, [inventory]);

  const itemRows = useMemo(() => {
    const out = []; const allowed = new Set(visibleStores.map((s) => s.id));
    for (const it of items) {
      if (q && !`${it.code} ${it.name}`.toLowerCase().includes(q.toLowerCase())) continue;
      const locs = inventory[it.id] || {};
      for (const loc in locs) {
        if (locs[loc] === 0) continue;
        if (storeFilter) { if (loc !== storeFilter) continue; }
        else if (custFilter) {
          if (loc === WAREHOUSE) continue;
          if (isCwh(loc)) { if (cwhCust(loc) !== custFilter) continue; }
          else if (!allowed.has(loc)) continue;
        }
        out.push({ item: it, loc, qty: locs[loc] });
      }
    }
    return out;
  }, [items, inventory, q, custFilter, storeFilter, visibleStores]);

  const storeRows = useMemo(() => visibleStores
    .filter((s) => (!storeFilter || s.id === storeFilter) && (!q || `${s.code} ${s.name}`.toLowerCase().includes(q.toLowerCase())) && byStore[s.id])
    .map((s) => {
      const inv = byStore[s.id] || {};
      const lines = Object.entries(inv).map(([itemId, qty]) => ({ item: itemMap[itemId], qty })).filter((l) => l.item).sort((a, b) => b.qty - a.qty);
      return { store: s, lines, total: lines.reduce((a, b) => a + b.qty, 0), kinds: lines.length };
    }).sort((a, b) => b.total - a.total), [visibleStores, byStore, itemMap, q, storeFilter]);

  const custRows = useMemo(() => customers
    .filter((c) => (!custFilter || c.id === custFilter) && (!q || c.name.toLowerCase().includes(q.toLowerCase())))
    .map((c) => {
      const sLines = stores.filter((s) => s.customerId === c.id).map((s) => {
        const inv = byStore[s.id] || {};
        const itemLines = Object.entries(inv).map(([itemId, qty]) => ({ item: itemMap[itemId], qty })).filter((l) => l.item).sort((a, b) => b.qty - a.qty);
        return { store: s, total: Object.values(inv).reduce((a, b) => a + b, 0), kinds: itemLines.length, itemLines };
      }).filter((x) => x.total > 0).sort((a, b) => b.total - a.total);
      const cwhInv = byStore[cwhId(c.id)] || {};
      const cwhItemLines = Object.entries(cwhInv).map(([itemId, qty]) => ({ item: itemMap[itemId], qty })).filter((l) => l.item).sort((a, b) => b.qty - a.qty);
      const cwhTotal = cwhItemLines.reduce((a, b) => a + b.qty, 0);
      const allLines = cwhTotal > 0
        ? [...sLines, { store: { id: cwhId(c.id), name: "고객사 창고(보관)", _cwh: true }, total: cwhTotal, kinds: cwhItemLines.length, itemLines: cwhItemLines }]
        : sLines;
      return { customer: c, sLines: allLines, total: allLines.reduce((a, b) => a + b.total, 0), storeCount: sLines.length, cwhTotal };
    }).filter((r) => r.total > 0).sort((a, b) => b.total - a.total), [customers, stores, byStore, itemMap, q, custFilter]);

  const VIEWS = [["customer", "고객사별", Building2], ["store", "매장별", Store], ["item", "아이템별", Package]];
  const count = view === "item" ? itemRows.length : view === "store" ? storeRows.length : custRows.length;

  return (
    <div className="space-y-3">
      <div className="flex flex-wrap gap-2 items-center">
        <div className="flex gap-1 bg-slate-100 rounded-lg p-1">
          {VIEWS.map(([v, l, Icon]) => (
            <button key={v} onClick={() => { setView(v); setExp({}); }}
              className={`inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-md text-xs font-medium transition ${view === v ? "bg-white text-indigo-700 shadow-sm" : "text-slate-500 hover:text-slate-700"}`}>
              <Icon size={14} /> {l}
            </button>
          ))}
        </div>
        <div className="relative flex-1 min-w-[150px]">
          <Search size={15} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
          <input value={q} onChange={(e) => setQ(e.target.value)}
            placeholder={view === "item" ? "아이템코드/명" : view === "store" ? "매장코드/명" : "고객사명"}
            className="w-full rounded-lg ring-1 ring-slate-300 pl-8 pr-3 py-2 text-sm" />
        </div>
        {view !== "customer" && (
          <select value={custFilter} onChange={(e) => { setCustFilter(e.target.value); setStoreFilter(""); }}
            className="rounded-lg ring-1 ring-slate-300 px-3 py-2 text-sm bg-white">
            <option value="">전체 고객사</option>
            {customers.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
        )}
        {view === "item" && (
          <select value={storeFilter} onChange={(e) => setStoreFilter(e.target.value)}
            className="rounded-lg ring-1 ring-slate-300 px-3 py-2 text-sm bg-white">
            <option value="">전체 위치</option>
            <option value={WAREHOUSE}>창고(미불출)</option>
            {visibleStores.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
        )}
        {view === "customer" && (
          <select value={custFilter} onChange={(e) => setCustFilter(e.target.value)}
            className="rounded-lg ring-1 ring-slate-300 px-3 py-2 text-sm bg-white">
            <option value="">전체 고객사</option>
            {customers.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
        )}
        <span className="text-xs text-slate-400">{count}건</span>
      </div>

      {view === "item" && (
        <Card className="overflow-hidden">
          <div className="overflow-x-auto max-h-[64vh]">
            <table className="w-full text-sm">
              <thead className="sticky top-0 bg-slate-50 text-slate-500 text-xs text-left">
                <tr><th className="px-4 py-2.5 font-medium">아이템코드</th><th className="px-4 py-2.5 font-medium">명칭</th>
                  <th className="px-4 py-2.5 font-medium">고객사</th><th className="px-4 py-2.5 font-medium">현재 위치</th>
                  <th className="px-4 py-2.5 font-medium text-right">수량</th></tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {itemRows.slice(0, 500).map((r, i) => {
                  const st = storeMap[r.loc];
                  return (
                    <tr key={i} className="hover:bg-slate-50/60">
                      <td className="px-4 py-2 font-medium tabular-nums">{r.item.code}</td>
                      <td className="px-4 py-2 text-slate-600">{r.item.name}</td>
                      <td className="px-4 py-2 text-slate-500 text-xs">{r.loc === WAREHOUSE ? "—" : isCwh(r.loc) ? (custMap[cwhCust(r.loc)]?.name || "—") : (custMap[st?.customerId]?.name || "—")}</td>
                      <td className="px-4 py-2">
                        {r.loc === WAREHOUSE
                          ? <span className="inline-flex items-center gap-1 text-slate-500"><Warehouse size={13} /> 우리창고</span>
                          : isCwh(r.loc)
                          ? <span className="inline-flex items-center gap-1 text-amber-600"><Warehouse size={13} /> 고객사 창고</span>
                          : <span className="inline-flex items-center gap-1 text-indigo-700"><Store size={13} /> {st?.name || r.loc}</span>}
                      </td>
                      <td className="px-4 py-2 text-right font-medium tabular-nums">{fmt(r.qty)}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            {itemRows.length > 500 && <div className="px-4 py-2 text-xs text-slate-400 bg-slate-50">상위 500건 표시 — 필터로 좁혀 보세요.</div>}
          </div>
        </Card>
      )}

      {view === "store" && (
        <div className="space-y-2">
          {storeRows.slice(0, 200).map((r) => {
            const open = exp[r.store.id];
            return (
              <Card key={r.store.id} className="overflow-hidden">
                <button onClick={() => toggleExp(r.store.id)} className="w-full flex items-center gap-2 px-4 py-3 hover:bg-slate-50/60 text-left">
                  <ChevronRight size={15} className={`text-slate-400 transition ${open ? "rotate-90" : ""}`} />
                  <div className="flex-1">
                    <div className="text-sm font-semibold flex items-center gap-1.5"><Store size={14} className="text-indigo-600" /> {r.store.name}</div>
                    <div className="text-[11px] text-slate-400">{custMap[r.store.customerId]?.name} · {r.store.code}</div>
                  </div>
                  <div className="text-right"><div className="text-sm font-bold tabular-nums">{fmt(r.total)} EA</div><div className="text-[11px] text-slate-400">{r.kinds}종</div></div>
                </button>
                {open && (
                  <div className="border-t border-slate-100 overflow-x-auto">
                    <table className="w-full text-sm">
                      <tbody className="divide-y divide-slate-50">
                        {r.lines.map((l) => (
                          <tr key={l.item.id} className="bg-slate-50/30">
                            <td className="pl-11 pr-4 py-1.5 tabular-nums font-medium text-slate-600">{l.item.code}</td>
                            <td className="px-4 py-1.5 text-slate-500">{l.item.name}</td>
                            <td className="px-4 py-1.5 text-right tabular-nums">{fmt(l.qty)}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </Card>
            );
          })}
          {storeRows.length === 0 && <Card className="p-8 text-center text-sm text-slate-400">보유 중인 매장이 없습니다.</Card>}
        </div>
      )}

      {view === "customer" && (
        <div className="space-y-2">
          {custRows.map((r) => {
            const open = exp[r.customer.id];
            return (
              <Card key={r.customer.id} className="overflow-hidden">
                <button onClick={() => toggleExp(r.customer.id)} className="w-full flex items-center gap-2 px-4 py-3 hover:bg-slate-50/60 text-left">
                  <ChevronRight size={15} className={`text-slate-400 transition ${open ? "rotate-90" : ""}`} />
                  <div className="flex-1">
                    <div className="text-sm font-semibold flex items-center gap-1.5"><Building2 size={14} className="text-indigo-600" /> {r.customer.name}</div>
                    <div className="text-[11px] text-slate-400">{r.customer.code} · 보유 매장 {r.storeCount}{r.cwhTotal > 0 ? ` · 고객사창고 ${fmt(r.cwhTotal)} EA` : ""}</div>
                  </div>
                  <div className="text-right"><div className="text-sm font-bold tabular-nums">{fmt(r.total)} EA</div><div className="text-[11px] text-slate-400">{r.storeCount}개 매장</div></div>
                </button>
                {open && (
                  <div className="border-t border-slate-100">
                    {r.sLines.map((s) => {
                      const sOpen = exp[s.store.id];
                      return (
                        <div key={s.store.id} className="border-b border-slate-50 last:border-0">
                          <button onClick={() => toggleExp(s.store.id)} className="w-full flex items-center gap-2 pl-9 pr-4 py-2 hover:bg-slate-50/60 text-left">
                            <ChevronRight size={13} className={`text-slate-300 transition ${sOpen ? "rotate-90" : ""}`} />
                            <span className="flex-1 text-[13px] font-medium text-slate-600 flex items-center gap-1.5">
                              {s.store._cwh ? <Warehouse size={12} className="text-amber-500" /> : <Store size={12} className="text-slate-400" />}
                              {s.store.name}
                            </span>
                            <span className="text-xs text-slate-400">{s.kinds}종</span>
                            <span className="text-[13px] font-semibold tabular-nums w-24 text-right">{fmt(s.total)} EA</span>
                          </button>
                          {sOpen && (
                            <table className="w-full text-sm">
                              <tbody className="divide-y divide-slate-50">
                                {s.itemLines.map((l) => (
                                  <tr key={l.item.id} className="bg-slate-50/40">
                                    <td className="pl-16 pr-4 py-1.5 tabular-nums font-medium text-slate-500">{l.item.code}</td>
                                    <td className="px-4 py-1.5 text-slate-400">{l.item.name}</td>
                                    <td className="px-4 py-1.5 text-right tabular-nums w-24">{fmt(l.qty)}</td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </Card>
            );
          })}
          {custRows.length === 0 && <Card className="p-8 text-center text-sm text-slate-400">보유 중인 고객사가 없습니다.</Card>}
        </div>
      )}
    </div>
  );
}

/* ---------- 입출고 ---------- */
function TxnView({ data, persist, flash, storeMap, custMap, itemMap, inventory, userEmail }) {
  const { items, stores, transactions, customers } = data;
  const [form, setForm] = useState({ date: todayISO(), type: "issue", itemId: "", qty: "", fromStoreId: WAREHOUSE, toStoreId: "", memo: "" });
  const [showImport, setShowImport] = useState(false);
  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));
  const storeLabel = (s) => `${custMap[s.customerId]?.name || "?"} · ${s.name}`;

  const avail = useMemo(() => {
    if (!form.itemId) return null;
    const loc = form.type === "issue" || form.type === "scrap" ? (form.fromStoreId || WAREHOUSE) : form.fromStoreId;
    if (!loc) return null;
    return (inventory[form.itemId]?.[loc]) ?? 0;
  }, [form, inventory]);

  const submit = () => {
    const q = Number(form.qty);
    if (!form.itemId) return flash("아이템을 선택하세요");
    if (!q || q <= 0) return flash("수량을 입력하세요");
    if (form.type === "issue" && !form.toStoreId) return flash("불출 대상 매장을 선택하세요");
    if (form.type === "return" && !form.fromStoreId) return flash("반납 출발 매장을 선택하세요");
    if (form.type === "transfer" && (!form.fromStoreId || !form.toStoreId)) return flash("이동 출발/도착 매장을 선택하세요");
    if (form.type === "transfer" && form.fromStoreId === form.toStoreId) return flash("출발과 도착이 같습니다");
    if (avail != null && q > avail) return flash(`재고 부족 — 현재 ${fmt(avail)} EA`);
    const t = { id: uid(), date: form.date, type: form.type, itemId: form.itemId, qty: q,
      fromStoreId: form.type === "issue" ? WAREHOUSE : form.fromStoreId,
      toStoreId: form.type === "return" ? WAREHOUSE : form.toStoreId, memo: form.memo,
      createdBy: userEmail || "" };
    persist({ ...data, transactions: [t, ...transactions] });
    flash(`${TXN_LABEL[form.type]} 등록 완료`);
    setForm((f) => ({ ...f, qty: "", memo: "" }));
  };
  const del = (id) => persist({ ...data, transactions: transactions.filter((t) => t.id !== id) });
  const exportTxns = () => {
    if (!transactions.length) return flash("내려받을 이력이 없습니다");
    const loc = (id) => id === WAREHOUSE ? "우리창고" : isCwh(id) ? `${custMap[cwhCust(id)]?.name || "고객사"} 창고` : (storeMap[id]?.name || "");
    const rows = transactions.map((t) => ({
      일자: t.date, 유형: TXN_LABEL[t.type] || t.type,
      아이템코드: itemMap[t.itemId]?.code || "", 아이템명: itemMap[t.itemId]?.name || "",
      "출발 위치": loc(t.fromStoreId), "도착 위치": t.type === "scrap" ? "폐기" : loc(t.toStoreId),
      수량: Number(t.qty) || 0, 메모: t.memo || "", 작성자: t.createdBy || "",
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [{ wch: 11 }, { wch: 7 }, { wch: 14 }, { wch: 20 }, { wch: 16 }, { wch: 16 }, { wch: 8 }, { wch: 20 }, { wch: 10 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "입출고이력");
    XLSX.writeFile(wb, `입출고이력_${todayISO()}.xlsx`);
  };
  const needFrom = form.type === "return" || form.type === "transfer" || form.type === "scrap";
  const needTo = form.type === "issue" || form.type === "transfer";

  return (
    <div className="grid lg:grid-cols-5 gap-4">
      <Card className="p-4 lg:col-span-2 h-fit">
        <div className="flex items-center justify-between mb-3">
          <h3 className="text-sm font-bold">입출고 등록</h3>
          <button onClick={() => setShowImport(true)} className="inline-flex items-center gap-1 text-xs text-indigo-600 hover:underline font-medium"><Upload size={13} /> 엑셀 업로드</button>
        </div>
        <div className="space-y-3">
          <div className="grid grid-cols-4 gap-1.5">
            {["issue", "return", "transfer", "scrap"].map((t) => (
              <button key={t} onClick={() => set("type", t)}
                className={`py-2 rounded-lg text-xs font-medium ring-1 transition
                  ${form.type === t ? "bg-indigo-600 text-white ring-indigo-600" : "bg-white text-slate-600 ring-slate-200 hover:bg-slate-50"}`}>
                {TXN_LABEL[t]}
              </button>
            ))}
          </div>
          <Field label="일자"><input type="date" value={form.date} onChange={(e) => set("date", e.target.value)} className="inp" /></Field>
          <Field label="아이템코드">
            <select value={form.itemId} onChange={(e) => set("itemId", e.target.value)} className="inp">
              <option value="">선택…</option>
              {items.map((i) => <option key={i.id} value={i.id}>{i.code} · {i.name}</option>)}
            </select>
          </Field>
          {needFrom && (
            <Field label="출발 위치">
              <select value={form.fromStoreId} onChange={(e) => set("fromStoreId", e.target.value)} className="inp">
                <option value="">선택…</option>
                <optgroup label="매장">
                  {stores.map((s) => <option key={s.id} value={s.id}>{storeLabel(s)}</option>)}
                </optgroup>
                <optgroup label="고객사 창고(보관)">
                  {customers.map((c) => <option key={c.id} value={cwhId(c.id)}>{c.name} · 고객사 창고</option>)}
                </optgroup>
              </select>
            </Field>
          )}
          {needTo && (
            <Field label={form.type === "issue" ? "불출 위치" : "도착 위치"}>
              <select value={form.toStoreId} onChange={(e) => set("toStoreId", e.target.value)} className="inp">
                <option value="">선택…</option>
                <optgroup label="매장">
                  {stores.map((s) => <option key={s.id} value={s.id}>{storeLabel(s)}</option>)}
                </optgroup>
                <optgroup label="고객사 창고(보관)">
                  {customers.map((c) => <option key={c.id} value={cwhId(c.id)}>{c.name} · 고객사 창고</option>)}
                </optgroup>
              </select>
            </Field>
          )}
          <Field label={`수량${avail != null ? ` · 가용 ${fmt(avail)} EA` : ""}`}>
            <input type="number" value={form.qty} onChange={(e) => set("qty", e.target.value)} placeholder="0" className="inp tabular-nums" />
          </Field>
          <Field label="메모(선택)"><input value={form.memo} onChange={(e) => set("memo", e.target.value)} className="inp" /></Field>
          <Btn tone="primary" onClick={submit} className="w-full justify-center"><Check size={15} /> 등록</Btn>
        </div>
        <style>{`.inp{width:100%;border-radius:.5rem;outline:none;font-size:.875rem;padding:.5rem .625rem;background:#fff;box-shadow:inset 0 0 0 1px #cbd5e1}.inp:focus{box-shadow:inset 0 0 0 2px #4f46e5}`}</style>
      </Card>

      <Card className="lg:col-span-3 overflow-hidden">
        <div className="px-4 py-3 border-b border-slate-100 flex items-center justify-between">
          <h3 className="text-sm font-bold">입출고 이력</h3>
          <div className="flex items-center gap-3">
            <span className="text-xs text-slate-400">{transactions.length}건</span>
            <button onClick={exportTxns} className="inline-flex items-center gap-1 text-xs text-indigo-600 hover:underline font-medium"><Download size={13} /> 엑셀</button>
          </div>
        </div>
        <div className="overflow-x-auto max-h-[60vh]">
          <table className="w-full text-sm">
            <thead className="sticky top-0 bg-slate-50 text-slate-500 text-xs text-left">
              <tr>
                <th className="px-3 py-2 font-medium">일자</th><th className="px-3 py-2 font-medium">유형</th>
                <th className="px-3 py-2 font-medium">아이템</th><th className="px-3 py-2 font-medium">경로</th>
                <th className="px-3 py-2 font-medium">작성자</th>
                <th className="px-3 py-2 font-medium text-right">수량</th><th className="px-3 py-2"></th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {transactions.slice(0, 300).map((t) => {
                const nm = (id) => id === WAREHOUSE ? "우리창고"
                  : isCwh(id) ? `${custMap[cwhCust(id)]?.name || "고객사"} 창고`
                  : (storeMap[id]?.name || (t.type === "scrap" ? "폐기" : "—"));
                return (
                  <tr key={t.id} className="hover:bg-slate-50/60">
                    <td className="px-3 py-2 tabular-nums text-slate-500 whitespace-nowrap">{t.date}</td>
                    <td className="px-3 py-2"><span className={`inline-block rounded px-1.5 py-0.5 text-[11px] font-medium ring-1 ${TXN_TONE[t.type]}`}>{TXN_LABEL[t.type]}</span></td>
                    <td className="px-3 py-2 tabular-nums">{itemMap[t.itemId]?.code || "—"}</td>
                    <td className="px-3 py-2 text-xs text-slate-500 whitespace-nowrap">{nm(t.fromStoreId)} → {nm(t.toStoreId)}</td>
                    <td className="px-3 py-2 text-xs text-slate-400 whitespace-nowrap">{t.createdBy || "—"}</td>
                    <td className="px-3 py-2 text-right font-medium tabular-nums">{fmt(t.qty)}</td>
                    <td className="px-3 py-2 text-right"><button onClick={() => del(t.id)} className="text-slate-300 hover:text-rose-500"><Trash2 size={14} /></button></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          {transactions.length === 0 && <p className="text-center text-sm text-slate-400 py-10">아직 이력이 없습니다.</p>}
        </div>
      </Card>

      {showImport && <ImportModal data={data} persist={persist} flash={flash} onClose={() => setShowImport(false)} />}
    </div>
  );
}
const Field = ({ label, children }) => (
  <label className="block"><span className="block text-xs text-slate-500 mb-1">{label}</span>{children}</label>
);

function ImportModal({ data, persist, flash, onClose }) {
  const { items, stores, transactions } = data;
  const itemByCode = useMemo(() => Object.fromEntries(items.map((i) => [String(i.code).trim(), i])), [items]);
  const storeByCode = useMemo(() => Object.fromEntries(stores.map((s) => [String(s.code).trim(), s])), [stores]);
  const [rows, setRows] = useState([]);
  const [fileName, setFileName] = useState("");
  const typeMap = { "불출": "issue", "반납": "return", "이동": "transfer", "폐기": "scrap", issue: "issue", return: "return", transfer: "transfer", scrap: "scrap" };

  const downloadTemplate = () => {
    const aoa = [
      ["일자", "유형", "아이템코드", "수량", "출발매장", "도착매장", "메모"],
      ["2025-01-10", "불출", "ESL-0001", 50, "", "ST-0101", "예시: 창고→매장"],
      ["2025-01-12", "이동", "ESL-0001", 20, "ST-0101", "ST-0102", "매장↔매장"],
      ["2025-01-15", "반납", "ESL-0001", 10, "ST-0102", "", "매장→창고"],
      ["2025-01-20", "폐기", "ESL-0001", 5, "ST-0102", "", "분실/파손"],
    ];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 12 }, { wch: 6 }, { wch: 12 }, { wch: 8 }, { wch: 12 }, { wch: 12 }, { wch: 18 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "수불");
    XLSX.writeFile(wb, "수불_업로드_양식.xlsx");
  };

  const onFile = async (e) => {
    const file = e.target.files?.[0]; if (!file) return;
    setFileName(file.name);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" });
      const body = json.slice(1).filter((r) => r.some((c) => String(c).trim() !== ""));
      const parsed = body.map((r, idx) => {
        const [dateRaw, typeRaw, itemCode, qtyRaw, fromCode, toCode, memo] = r;
        const date = String(dateRaw).slice(0, 10);
        const type = typeMap[String(typeRaw).trim()];
        const item = itemByCode[String(itemCode).trim()];
        const qty = Number(qtyRaw);
        const fcode = String(fromCode).trim(), tcode = String(toCode).trim();
        const from = fcode ? storeByCode[fcode] : null;
        const to = tcode ? storeByCode[tcode] : null;
        const errs = [];
        if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) errs.push("일자형식");
        if (!type) errs.push("유형");
        if (!item) errs.push("아이템코드");
        if (!qty || qty <= 0) errs.push("수량");
        let fromStoreId, toStoreId;
        if (type === "issue") { fromStoreId = WAREHOUSE; if (!to) errs.push("도착매장"); toStoreId = to?.id; }
        else if (type === "return") { if (!from) errs.push("출발매장"); fromStoreId = from?.id; toStoreId = WAREHOUSE; }
        else if (type === "transfer") { if (!from) errs.push("출발매장"); if (!to) errs.push("도착매장"); fromStoreId = from?.id; toStoreId = to?.id; if (from && to && from.id === to.id) errs.push("출발=도착"); }
        else if (type === "scrap") { if (!from) errs.push("출발매장"); fromStoreId = from?.id; toStoreId = WAREHOUSE; }
        return { line: idx + 2, date, type, typeLabel: String(typeRaw).trim(), item, itemCode: String(itemCode).trim(), qty: qty || 0, fromStoreId, toStoreId, fromCode: fcode, toCode: tcode, memo: String(memo).trim(), errs };
      });
      setRows(parsed);
      if (parsed.length === 0) flash("데이터 행이 없습니다");
    } catch (err) { flash("파일을 읽지 못했습니다 — 양식을 확인하세요"); }
  };

  const valid = rows.filter((r) => r.errs.length === 0);
  const invalid = rows.length - valid.length;
  const doImport = () => {
    if (valid.length === 0) return flash("등록할 유효한 행이 없습니다");
    const txns = valid.map((r) => ({ id: uid(), date: r.date, type: r.type, itemId: r.item.id, qty: r.qty, fromStoreId: r.fromStoreId, toStoreId: r.toStoreId, memo: r.memo || "엑셀 업로드" }));
    persist({ ...data, transactions: [...txns, ...transactions] });
    flash(`${txns.length}건 등록 완료`);
    onClose();
  };

  const locName = (id) => id === WAREHOUSE ? "창고" : (stores.find((s) => s.id === id)?.name || "?");

  return (
    <div className="fixed inset-0 z-40 bg-slate-900/40 grid place-items-center p-4" onClick={onClose}>
      <div className="bg-white rounded-2xl w-full max-w-3xl max-h-[88vh] flex flex-col shadow-2xl" onClick={(e) => e.stopPropagation()}>
        <div className="px-5 py-3.5 border-b border-slate-100 flex items-center justify-between">
          <div>
            <h3 className="text-sm font-bold flex items-center gap-1.5"><Upload size={15} /> 수불 엑셀 업로드</h3>
            <p className="text-[11px] text-slate-400 mt-0.5">컬럼: 일자 · 유형(불출/반납/이동/폐기) · 아이템코드 · 수량 · 출발매장 · 도착매장 · 메모</p>
          </div>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-700"><X size={18} /></button>
        </div>

        <div className="px-5 py-3 border-b border-slate-100 flex flex-wrap gap-2 items-center">
          <Btn onClick={downloadTemplate}><Download size={14} /> 양식 다운로드</Btn>
          <label className="inline-flex items-center gap-1.5 rounded-lg px-3 py-2 text-sm font-medium bg-indigo-600 text-white hover:bg-indigo-700 cursor-pointer">
            <Upload size={14} /> 파일 선택(.xlsx/.csv)
            <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={onFile} />
          </label>
          {fileName && <span className="text-xs text-slate-500 truncate">{fileName}</span>}
          {rows.length > 0 && (
            <span className="ml-auto text-xs">
              <span className="text-emerald-600 font-medium">유효 {valid.length}</span>
              {invalid > 0 && <span className="text-rose-500 font-medium ml-2">오류 {invalid}</span>}
            </span>
          )}
        </div>

        <div className="overflow-y-auto flex-1">
          {rows.length === 0 ? (
            <div className="p-10 text-center text-sm text-slate-400">
              양식을 받아 채운 뒤 파일을 선택하면 미리보기가 표시됩니다.<br />
              출발/도착 매장은 <b className="text-slate-500">매장코드</b>로 입력하고, 창고는 비워 두세요.
            </div>
          ) : (
            <table className="w-full text-sm">
              <thead className="sticky top-0 bg-slate-50 text-slate-500 text-xs text-left">
                <tr><th className="px-4 py-2 font-medium">행</th><th className="px-4 py-2 font-medium">일자</th><th className="px-4 py-2 font-medium">유형</th>
                  <th className="px-4 py-2 font-medium">아이템</th><th className="px-4 py-2 font-medium text-right">수량</th>
                  <th className="px-4 py-2 font-medium">경로</th><th className="px-4 py-2 font-medium">상태</th></tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {rows.slice(0, 300).map((r, i) => (
                  <tr key={i} className={r.errs.length ? "bg-rose-50/40" : "hover:bg-slate-50/60"}>
                    <td className="px-4 py-1.5 text-slate-400 tabular-nums">{r.line}</td>
                    <td className="px-4 py-1.5 tabular-nums text-slate-500">{r.date}</td>
                    <td className="px-4 py-1.5">{r.type ? <span className={`inline-block rounded px-1.5 py-0.5 text-[11px] font-medium ring-1 ${TXN_TONE[r.type]}`}>{TXN_LABEL[r.type]}</span> : <span className="text-rose-500 text-xs">{r.typeLabel || "—"}</span>}</td>
                    <td className="px-4 py-1.5 tabular-nums">{r.itemCode}</td>
                    <td className="px-4 py-1.5 text-right tabular-nums">{fmt(r.qty)}</td>
                    <td className="px-4 py-1.5 text-xs text-slate-500">
                      {r.type ? `${r.fromStoreId ? locName(r.fromStoreId) : (r.fromCode || "?")} → ${r.toStoreId ? locName(r.toStoreId) : (r.toCode || "?")}` : "—"}
                    </td>
                    <td className="px-4 py-1.5">
                      {r.errs.length === 0
                        ? <span className="inline-flex items-center gap-1 text-emerald-600 text-xs"><Check size={12} /> 정상</span>
                        : <span className="text-rose-600 text-[11px]">{r.errs.join(", ")}</span>}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
          {rows.length > 300 && <div className="px-4 py-2 text-xs text-slate-400">상위 300행 미리보기 (등록은 전체 유효행 기준)</div>}
        </div>

        <div className="px-5 py-3 border-t border-slate-100 flex items-center justify-between">
          <span className="text-[11px] text-slate-400">{invalid > 0 ? "오류 행은 제외하고 등록됩니다. 재고 부족 여부는 등록 후 재고현황에서 확인하세요." : "\u00a0"}</span>
          <div className="flex gap-2">
            <Btn onClick={onClose}>취소</Btn>
            <Btn tone="primary" onClick={doImport}><Check size={15} /> 유효 {valid.length}건 등록</Btn>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ---------- 마스터 ---------- */
function MasterView({ data, persist, flash, custMap, period }) {
  const [sub, setSub] = useState("cust");
  return (
    <div className="space-y-3">
      <div className="flex gap-1.5 flex-wrap">
        <Btn tone={sub === "cust" ? "primary" : "default"} onClick={() => setSub("cust")}><Building2 size={15} /> 고객사</Btn>
        <Btn tone={sub === "contract" ? "primary" : "default"} onClick={() => setSub("contract")}><FileSignature size={15} /> 계약</Btn>
        <Btn tone={sub === "store" ? "primary" : "default"} onClick={() => setSub("store")}><Store size={15} /> 매장</Btn>
        <Btn tone={sub === "item" ? "primary" : "default"} onClick={() => setSub("item")}><Package size={15} /> 아이템코드</Btn>
        <Btn tone={sub === "fx" ? "primary" : "default"} onClick={() => setSub("fx")}><Banknote size={15} /> 통화·환율</Btn>
        <Btn tone={sub === "acct" ? "primary" : "default"} onClick={() => setSub("acct")}><Calculator size={15} /> 계정과목</Btn>
      </div>
      {sub === "cust" && <CustomerMaster {...{ data, persist, flash }} />}
      {sub === "contract" && <ContractMaster {...{ data, persist, flash, custMap, period }} />}
      {sub === "store" && <StoreMaster {...{ data, persist, flash, custMap }} />}
      {sub === "item" && <ItemMaster {...{ data, persist, flash }} />}
      {sub === "fx" && <CurrencyMaster {...{ data, persist, flash, period }} />}
      {sub === "acct" && <AccountMaster {...{ data, persist, flash }} />}
    </div>
  );
}

function AccountMaster({ data, persist, flash }) {
  const A = { ...DEFAULT_ACCOUNTS, ...(data.accounts || {}) };
  const [f, setF] = useState(A);
  const set = (k, v) => setF((p) => ({ ...p, [k]: v }));
  const save = () => { persist({ ...data, accounts: f }); flash("계정과목 저장됨"); };
  const FIELDS = [
    ["arLease", "리스료수익 인식 · 차변"],
    ["leaseRevenue", "리스료수익 인식 · 대변"],
    ["assetCost", "신규취득 차변 / 폐기 대변 (자산)"],
    ["payable", "신규취득 · 대변"],
    ["depExpense", "감가상각 · 차변"],
    ["accDep", "감가상각 대변 / 폐기 차변 (누계액)"],
    ["disposalLoss", "폐기 · 처분손실 차변"],
  ];
  return (
    <Card className="p-4">
      <div className="flex items-center justify-between mb-1">
        <h3 className="text-sm font-bold">계정과목</h3>
        <button onClick={() => setF(DEFAULT_ACCOUNTS)} className="text-xs text-slate-400 underline hover:text-slate-600">기본값 복원</button>
      </div>
      <p className="text-[11px] text-slate-400 mb-4">회계·청구 탭의 자동 분개에 쓰이는 계정명입니다. 회사 계정체계에 맞게 수정하세요.</p>
      <div className="space-y-2.5 max-w-xl">
        {FIELDS.map(([k, desc]) => (
          <BizField key={k} label={desc}>
            <input value={f[k]} onChange={(e) => set(k, e.target.value)} className="rounded-md ring-1 ring-slate-200 px-2 py-1.5 text-sm w-48 focus:ring-2 focus:ring-indigo-400 focus:outline-none" />
          </BizField>
        ))}
      </div>
      <p className="text-[10px] text-slate-400 mt-3 max-w-xl">※ '자산(취득원가)'은 신규취득 차변과 폐기 대변에, '누계액'은 감가상각 대변과 폐기 차변에 공통으로 쓰입니다.</p>
      <div className="mt-4"><Btn tone="primary" onClick={save}><Check size={15} /> 저장</Btn></div>
    </Card>
  );
}

function CustomerMaster({ data, persist, flash }) {
  const { customers, stores, contracts } = data;
  const [f, setF] = useState({ code: "", name: "" });
  const set = (k, v) => setF((p) => ({ ...p, [k]: v }));
  const add = () => {
    if (!f.code.trim()) return flash("고객사 코드를 입력하세요");
    if (customers.some((c) => c.code === f.code.trim())) return flash("이미 존재하는 코드입니다");
    persist({ ...data, customers: [{ id: uid(), code: f.code.trim(), name: f.name.trim() || f.code.trim() }, ...customers] });
    flash("고객사 등록 완료");
    setF({ code: "", name: "" });
  };
  const del = (id) => {
    if (stores.some((s) => s.customerId === id)) return flash("소속 매장이 있어 삭제할 수 없습니다");
    if (contracts.some((c) => c.customerId === id)) return flash("연결된 계약이 있어 삭제할 수 없습니다");
    persist({ ...data, customers: customers.filter((c) => c.id !== id) });
  };

  return (
    <div className="space-y-3">
      <Card className="p-4">
        <div className="grid md:grid-cols-4 gap-2">
          <input placeholder="고객사 코드" value={f.code} onChange={(e) => set("code", e.target.value)} className="mi" />
          <input placeholder="고객사명" value={f.name} onChange={(e) => set("name", e.target.value)} className="mi md:col-span-2" />
          <Btn tone="primary" onClick={add} className="justify-center"><Plus size={15} /> 추가</Btn>
        </div>
        <p className="text-[11px] text-slate-400 mt-2">과금방식·단가는 '계약' 탭에서 설정합니다. 한 고객사에 여러 계약(갱신·추가도입)을 둘 수 있습니다.</p>
        <style>{`.mi{border-radius:.5rem;font-size:.875rem;padding:.5rem .625rem;background:#fff;box-shadow:inset 0 0 0 1px #cbd5e1;outline:none}.mi:focus{box-shadow:inset 0 0 0 2px #4f46e5}`}</style>
      </Card>
      <Card className="overflow-hidden">
        <div className="overflow-x-auto max-h-[58vh]">
          <table className="w-full text-sm">
            <thead className="sticky top-0 bg-slate-50 text-slate-500 text-xs text-left">
              <tr><th className="px-3 py-2 font-medium">코드</th><th className="px-3 py-2 font-medium">고객사명</th>
                <th className="px-3 py-2 font-medium text-right">계약수</th><th className="px-3 py-2 font-medium text-right">매장수</th><th className="px-3 py-2"></th></tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {customers.map((c) => (
                <tr key={c.id} className="hover:bg-slate-50/60">
                  <td className="px-3 py-2 font-medium">{c.code}</td>
                  <td className="px-3 py-2 text-slate-600">{c.name}</td>
                  <td className="px-3 py-2 text-right tabular-nums text-slate-500">{contracts.filter((x) => x.customerId === c.id).length}</td>
                  <td className="px-3 py-2 text-right tabular-nums text-slate-500">{stores.filter((s) => s.customerId === c.id).length}</td>
                  <td className="px-3 py-2 text-right"><button onClick={() => del(c.id)} className="text-slate-300 hover:text-rose-500"><Trash2 size={14} /></button></td>
                </tr>
              ))}
            </tbody>
          </table>
          {customers.length === 0 && <p className="text-center text-sm text-slate-400 py-10">등록된 고객사가 없습니다.</p>}
        </div>
      </Card>
    </div>
  );
}

function ContractMaster({ data, persist, flash, custMap, period }) {
  const { contracts, customers, items } = data;
  const [f, setF] = useState({ customerId: "", contractNo: "", startDate: todayISO(), endDate: "", billing: "flat", monthlyFee: "", currency: "KRW", excludeStorage: false, recipientName: "", recipientEmail: "", recipientTel: "" });
  const [editing, setEditing] = useState(null); // 단가표 편집 대상 계약 id
  const set = (k, v) => setF((p) => ({ ...p, [k]: v }));

  const add = () => {
    if (!f.customerId) return flash("고객사를 선택하세요");
    if (!f.contractNo.trim()) return flash("계약번호를 입력하세요");
    const ct = { id: uid(), customerId: f.customerId, contractNo: f.contractNo.trim(),
      startDate: f.startDate, endDate: f.endDate || "", billing: f.billing, currency: f.currency,
      monthlyFee: f.billing === "flat" ? (Number(f.monthlyFee) || 0) : 0,
      excludeStorage: !!f.excludeStorage,
      recipientName: f.recipientName.trim(), recipientEmail: f.recipientEmail.trim(), recipientTel: f.recipientTel.trim(),
      rates: {} };
    persist({ ...data, contracts: [ct, ...contracts] });
    flash("계약 등록 완료");
    setF({ customerId: f.customerId, contractNo: "", startDate: todayISO(), endDate: "", billing: "flat", monthlyFee: "", currency: f.currency, excludeStorage: f.excludeStorage, recipientName: "", recipientEmail: "", recipientTel: "" });
    if (ct.billing === "usage") setEditing(ct.id); // 사용량이면 바로 단가표 편집 유도
  };
  const del = (id) => persist({ ...data, contracts: contracts.filter((c) => c.id !== id) });
  const editContract = contracts.find((c) => c.id === editing);

  return (
    <div className="space-y-3">
      <Card className="p-4">
        <div className="grid md:grid-cols-7 gap-2">
          <select value={f.customerId} onChange={(e) => set("customerId", e.target.value)} className="mi md:col-span-2">
            <option value="">고객사 선택…</option>
            {customers.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
          <input placeholder="계약번호" value={f.contractNo} onChange={(e) => set("contractNo", e.target.value)} className="mi" />
          <select value={f.billing} onChange={(e) => set("billing", e.target.value)} className="mi">
            <option value="flat">월정액</option><option value="usage">사용량(tag-day)</option>
          </select>
          <select value={f.currency} onChange={(e) => set("currency", e.target.value)} className="mi">
            {(data.currencies || []).map((c) => <option key={c.code} value={c.code}>{c.code}</option>)}
          </select>
          {f.billing === "flat"
            ? <input type="number" placeholder="월정액" value={f.monthlyFee} onChange={(e) => set("monthlyFee", e.target.value)} className="mi tabular-nums" />
            : <div className="mi grid place-items-center text-[11px] text-slate-400">단가표는 등록 후 입력</div>}
          <Btn tone="primary" onClick={add} className="justify-center"><Plus size={15} /> 추가</Btn>
        </div>
        <div className="grid md:grid-cols-2 gap-2 mt-2">
          <label className="text-xs text-slate-500 flex items-center gap-2">계약 시작일
            <input type="date" value={f.startDate} onChange={(e) => set("startDate", e.target.value)} className="mi flex-1" /></label>
          <label className="text-xs text-slate-500 flex items-center gap-2">종료일(선택)
            <input type="date" value={f.endDate} onChange={(e) => set("endDate", e.target.value)} className="mi flex-1" /></label>
        </div>
        <div className="grid md:grid-cols-3 gap-2 mt-2">
          <input placeholder="청구서 수신인" value={f.recipientName} onChange={(e) => set("recipientName", e.target.value)} className="mi" />
          <input placeholder="수신 이메일" value={f.recipientEmail} onChange={(e) => set("recipientEmail", e.target.value)} className="mi" />
          <input placeholder="연락처" value={f.recipientTel} onChange={(e) => set("recipientTel", e.target.value)} className="mi" />
        </div>
        <label className="flex items-center gap-2 mt-2 text-[13px] text-slate-600 cursor-pointer select-none">
          <input type="checkbox" checked={f.excludeStorage} onChange={(e) => set("excludeStorage", e.target.checked)} className="w-4 h-4 accent-indigo-600" />
          고객사 창고 보관분은 과금에서 제외 <span className="text-slate-400">(폐점 후 고객사가 보관만 하는 기간 면제)</span>
        </label>
        <p className="text-[11px] text-slate-400 mt-1">월정액: 계약당 단일 정액 · 사용량: 품목코드별 일당단가(p) × 수량(q) × 일수. 종료일을 비우면 무기한 계약입니다. 수신인 정보는 PDF 청구서에 표시됩니다.</p>
        <style>{`.mi{border-radius:.5rem;font-size:.875rem;padding:.5rem .625rem;background:#fff;box-shadow:inset 0 0 0 1px #cbd5e1;outline:none}.mi:focus{box-shadow:inset 0 0 0 2px #4f46e5}`}</style>
      </Card>

      <Card className="overflow-hidden">
        <div className="overflow-x-auto max-h-[56vh]">
          <table className="w-full text-sm">
            <thead className="sticky top-0 bg-slate-50 text-slate-500 text-xs text-left">
              <tr><th className="px-3 py-2 font-medium">계약번호</th><th className="px-3 py-2 font-medium">고객사</th>
                <th className="px-3 py-2 font-medium">기간</th><th className="px-3 py-2 font-medium">방식</th>
                <th className="px-3 py-2 font-medium">통화</th>
                <th className="px-3 py-2 font-medium text-right">월정액/단가</th><th className="px-3 py-2 font-medium text-center">상태</th><th className="px-3 py-2"></th></tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {contracts.map((c) => {
                const active = isActiveContract(c, period);
                const rateN = Object.values(c.rates || {}).filter((v) => Number(v) > 0).length;
                return (
                  <tr key={c.id} className="hover:bg-slate-50/60">
                    <td className="px-3 py-2 font-medium">{c.contractNo}
                      {c.excludeStorage && <span className="ml-1.5 text-[10px] text-amber-700 bg-amber-50 rounded px-1 py-0.5">창고면제</span>}
                    </td>
                    <td className="px-3 py-2 text-slate-600">{custMap[c.customerId]?.name || "—"}</td>
                    <td className="px-3 py-2 text-slate-500 text-xs tabular-nums">{c.startDate} ~ {c.endDate || "무기한"}</td>
                    <td className="px-3 py-2"><BillBadge b={c.billing} /></td>
                    <td className="px-3 py-2 text-xs font-medium text-slate-500">{c.currency || "KRW"}</td>
                    <td className="px-3 py-2 text-right tabular-nums">
                      {c.billing === "flat"
                        ? `${c.currency || "KRW"} ${fmt(c.monthlyFee)}`
                        : <button onClick={() => setEditing(c.id)} className="inline-flex items-center gap-1 text-indigo-600 hover:underline">
                            <Pencil size={12} /> 단가표 {rateN > 0 ? `(${rateN}종)` : "(미입력)"}
                          </button>}
                    </td>
                    <td className="px-3 py-2 text-center">
                      <span className={`rounded-full px-2 py-0.5 text-[11px] font-medium ${active ? "bg-emerald-50 text-emerald-700" : "bg-slate-100 text-slate-400"}`}>
                        {active ? "유효" : "기간외"}
                      </span>
                    </td>
                    <td className="px-3 py-2 text-right"><button onClick={() => del(c.id)} className="text-slate-300 hover:text-rose-500"><Trash2 size={14} /></button></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          {contracts.length === 0 && <p className="text-center text-sm text-slate-400 py-10">등록된 계약이 없습니다.</p>}
        </div>
      </Card>

      {editContract && (
        <RateEditor contract={editContract} items={items} custName={custMap[editContract.customerId]?.name}
          onClose={() => setEditing(null)}
          onSave={(rates) => { persist({ ...data, contracts: contracts.map((c) => c.id === editContract.id ? { ...c, rates } : c) }); setEditing(null); flash("단가표 저장 완료"); }} />
      )}
    </div>
  );
}

function RateEditor({ contract, items, custName, onClose, onSave }) {
  const [rates, setRates] = useState(() => ({ ...(contract.rates || {}) }));
  const [q, setQ] = useState("");
  const [bulk, setBulk] = useState("");
  const setRate = (id, v) => setRates((r) => ({ ...r, [id]: v === "" ? "" : Number(v) }));
  const filtered = items.filter((i) => !q || `${i.code} ${i.name}`.toLowerCase().includes(q.toLowerCase()));
  const applyStandard = () => { const r = { ...rates }; items.forEach((i) => { if (!r[i.id]) r[i.id] = Number(i.dailyRate) || 0; }); setRates(r); };
  const applyBulk = () => { const v = Number(bulk) || 0; const r = { ...rates }; filtered.forEach((i) => { r[i.id] = v; }); setRates(r); };
  const clean = () => { const r = {}; for (const k in rates) { const v = Number(rates[k]); if (v > 0) r[k] = v; } return r; };
  const setCount = Object.values(rates).filter((v) => Number(v) > 0).length;

  return (
    <div className="fixed inset-0 z-40 bg-slate-900/40 grid place-items-center p-4" onClick={onClose}>
      <div className="bg-white rounded-2xl w-full max-w-2xl max-h-[85vh] flex flex-col shadow-2xl" onClick={(e) => e.stopPropagation()}>
        <div className="px-5 py-3.5 border-b border-slate-100 flex items-center justify-between">
          <div>
            <h3 className="text-sm font-bold flex items-center gap-1.5"><FileSignature size={15} /> 품목별 일당단가표</h3>
            <p className="text-[11px] text-slate-400 mt-0.5">{custName} · {contract.contractNo} · <b className="text-slate-500">{contract.currency || "KRW"}</b> 기준 · 단가 입력 {setCount}종</p>
          </div>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-700"><X size={18} /></button>
        </div>
        <div className="px-5 py-3 border-b border-slate-100 flex flex-wrap gap-2 items-center">
          <div className="relative flex-1 min-w-[150px]">
            <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
            <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="아이템 검색"
              className="w-full rounded-lg ring-1 ring-slate-300 pl-8 pr-3 py-1.5 text-sm" />
          </div>
          <Btn onClick={applyStandard} className="text-xs">표준단가 일괄</Btn>
          <div className="flex items-center gap-1">
            <input type="number" value={bulk} onChange={(e) => setBulk(e.target.value)} placeholder="일괄값"
              className="w-20 rounded-lg ring-1 ring-slate-300 px-2 py-1.5 text-sm tabular-nums" />
            <Btn onClick={applyBulk} className="text-xs">검색결과 적용</Btn>
          </div>
        </div>
        <div className="overflow-y-auto flex-1">
          <table className="w-full text-sm">
            <thead className="sticky top-0 bg-slate-50 text-slate-500 text-xs text-left">
              <tr><th className="px-5 py-2 font-medium">아이템코드</th><th className="px-5 py-2 font-medium">명칭</th>
                <th className="px-5 py-2 font-medium text-right">표준단가</th><th className="px-5 py-2 font-medium text-right">계약 일당단가(p)</th></tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {filtered.slice(0, 300).map((i) => (
                <tr key={i.id} className="hover:bg-slate-50/60">
                  <td className="px-5 py-1.5 font-medium tabular-nums">{i.code}</td>
                  <td className="px-5 py-1.5 text-slate-600">{i.name}</td>
                  <td className="px-5 py-1.5 text-right tabular-nums text-slate-400">{fmt(i.dailyRate)}</td>
                  <td className="px-5 py-1.5 text-right">
                    <input type="number" value={rates[i.id] ?? ""} onChange={(e) => setRate(i.id, e.target.value)} placeholder="0"
                      className="w-24 text-right rounded ring-1 ring-slate-200 px-2 py-1 tabular-nums focus:ring-2 focus:ring-indigo-500 outline-none" />
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          {filtered.length > 300 && <div className="px-5 py-2 text-xs text-slate-400">상위 300건 표시 — 검색으로 좁혀 보세요.</div>}
        </div>
        <div className="px-5 py-3 border-t border-slate-100 flex justify-end gap-2">
          <Btn onClick={onClose}>취소</Btn>
          <Btn tone="primary" onClick={() => onSave(clean())}><Check size={15} /> 단가표 저장</Btn>
        </div>
      </div>
    </div>
  );
}

function StoreMaster({ data, persist, flash, custMap }) {
  const { stores, customers } = data;
  const [f, setF] = useState({ customerId: "", code: "", name: "" });
  const set = (k, v) => setF((p) => ({ ...p, [k]: v }));
  const add = () => {
    if (!f.customerId) return flash("소속 고객사를 선택하세요");
    if (!f.code.trim()) return flash("매장코드를 입력하세요");
    if (stores.some((s) => s.code === f.code.trim())) return flash("이미 존재하는 코드입니다");
    persist({ ...data, stores: [{ id: uid(), customerId: f.customerId, code: f.code.trim(), name: f.name.trim() || f.code.trim(), active: true }, ...stores] });
    flash("매장 등록 완료");
    setF({ customerId: f.customerId, code: "", name: "" });
  };
  const del = (id) => persist({ ...data, stores: stores.filter((s) => s.id !== id) });
  const toggle = (id) => persist({ ...data, stores: stores.map((s) => s.id === id ? { ...s, active: !s.active } : s) });

  return (
    <div className="space-y-3">
      <Card className="p-4">
        <div className="grid md:grid-cols-5 gap-2">
          <select value={f.customerId} onChange={(e) => set("customerId", e.target.value)} className="mi md:col-span-2">
            <option value="">소속 고객사…</option>
            {customers.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
          <input placeholder="매장코드" value={f.code} onChange={(e) => set("code", e.target.value)} className="mi" />
          <input placeholder="매장명" value={f.name} onChange={(e) => set("name", e.target.value)} className="mi" />
          <Btn tone="primary" onClick={add} className="justify-center"><Plus size={15} /> 추가</Btn>
        </div>
        <p className="text-[11px] text-slate-400 mt-2">매장은 위치 식별만 합니다. 리스기간·요금은 '계약'에서 관리됩니다.</p>
        <style>{`.mi{border-radius:.5rem;font-size:.875rem;padding:.5rem .625rem;background:#fff;box-shadow:inset 0 0 0 1px #cbd5e1;outline:none}.mi:focus{box-shadow:inset 0 0 0 2px #4f46e5}`}</style>
      </Card>
      <Card className="overflow-hidden">
        <div className="overflow-x-auto max-h-[58vh]">
          <table className="w-full text-sm">
            <thead className="sticky top-0 bg-slate-50 text-slate-500 text-xs text-left">
              <tr><th className="px-3 py-2 font-medium">코드</th><th className="px-3 py-2 font-medium">매장명</th>
                <th className="px-3 py-2 font-medium">고객사</th><th className="px-3 py-2 font-medium text-center">상태</th><th className="px-3 py-2"></th></tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {stores.map((s) => (
                <tr key={s.id} className="hover:bg-slate-50/60">
                  <td className="px-3 py-2 font-medium">{s.code}</td>
                  <td className="px-3 py-2 text-slate-600">{s.name}</td>
                  <td className="px-3 py-2 text-slate-500 text-xs">{custMap[s.customerId]?.name || "—"}</td>
                  <td className="px-3 py-2 text-center">
                    <button onClick={() => toggle(s.id)} className={`rounded-full px-2 py-0.5 text-[11px] font-medium ${s.active ? "bg-emerald-50 text-emerald-700" : "bg-slate-100 text-slate-400"}`}>{s.active ? "활성" : "중지"}</button>
                  </td>
                  <td className="px-3 py-2 text-right"><button onClick={() => del(s.id)} className="text-slate-300 hover:text-rose-500"><Trash2 size={14} /></button></td>
                </tr>
              ))}
            </tbody>
          </table>
          {stores.length === 0 && <p className="text-center text-sm text-slate-400 py-10">등록된 매장이 없습니다.</p>}
        </div>
      </Card>
    </div>
  );
}

function ItemMaster({ data, persist, flash }) {
  const { items } = data;
  const [f, setF] = useState({ code: "", name: "", unitCost: "", dailyRate: "", acquiredQty: "", acquiredDate: todayISO(), usefulLifeMonths: "36" });
  const set = (k, v) => setF((p) => ({ ...p, [k]: v }));
  const add = () => {
    if (!f.code.trim()) return flash("아이템코드를 입력하세요");
    if (items.some((i) => i.code === f.code.trim())) return flash("이미 존재하는 코드입니다");
    const it = { id: uid(), code: f.code.trim(), name: f.name.trim() || f.code.trim(),
      unitCost: Number(f.unitCost) || 0, dailyRate: Number(f.dailyRate) || 0, acquiredQty: Number(f.acquiredQty) || 0,
      acquiredDate: f.acquiredDate, usefulLifeMonths: Number(f.usefulLifeMonths) || 36 };
    persist({ ...data, items: [it, ...items] });
    flash("아이템 등록 완료");
    setF({ code: "", name: "", unitCost: "", dailyRate: "", acquiredQty: "", acquiredDate: todayISO(), usefulLifeMonths: "36" });
  };
  const del = (id) => persist({ ...data, items: items.filter((i) => i.id !== id) });

  return (
    <div className="space-y-3">
      <Card className="p-4">
        <div className="grid md:grid-cols-7 gap-2">
          <input placeholder="코드" value={f.code} onChange={(e) => set("code", e.target.value)} className="mi" />
          <input placeholder="명칭" value={f.name} onChange={(e) => set("name", e.target.value)} className="mi md:col-span-2" />
          <input type="number" placeholder="단가" value={f.unitCost} onChange={(e) => set("unitCost", e.target.value)} className="mi tabular-nums" />
          <input type="number" placeholder="표준일당단가" value={f.dailyRate} onChange={(e) => set("dailyRate", e.target.value)} className="mi tabular-nums" />
          <input type="number" placeholder="취득수량" value={f.acquiredQty} onChange={(e) => set("acquiredQty", e.target.value)} className="mi tabular-nums" />
          <Btn tone="primary" onClick={add} className="justify-center"><Plus size={15} /> 추가</Btn>
        </div>
        <div className="grid md:grid-cols-6 gap-2 mt-2">
          <label className="text-xs text-slate-500 md:col-span-3 flex items-center gap-2">취득일
            <input type="date" value={f.acquiredDate} onChange={(e) => set("acquiredDate", e.target.value)} className="mi flex-1" /></label>
          <label className="text-xs text-slate-500 md:col-span-3 flex items-center gap-2">내용연수(월)
            <input type="number" value={f.usefulLifeMonths} onChange={(e) => set("usefulLifeMonths", e.target.value)} className="mi w-24 tabular-nums" /></label>
        </div>
        <p className="text-[11px] text-slate-400 mt-1">표준일당단가는 계약 단가표의 기본값으로 쓰입니다('표준단가 일괄' 버튼). 실제 청구는 각 계약의 단가표를 따릅니다.</p>
        <style>{`.mi{border-radius:.5rem;font-size:.875rem;padding:.5rem .625rem;background:#fff;box-shadow:inset 0 0 0 1px #cbd5e1;outline:none}.mi:focus{box-shadow:inset 0 0 0 2px #4f46e5}`}</style>
      </Card>
      <Card className="overflow-hidden">
        <div className="overflow-x-auto max-h-[56vh]">
          <table className="w-full text-sm">
            <thead className="sticky top-0 bg-slate-50 text-slate-500 text-xs text-left">
              <tr><th className="px-3 py-2 font-medium">코드</th><th className="px-3 py-2 font-medium">명칭</th>
                <th className="px-3 py-2 font-medium text-right">단가</th><th className="px-3 py-2 font-medium text-right">표준일당단가</th>
                <th className="px-3 py-2 font-medium text-right">취득수량</th>
                <th className="px-3 py-2 font-medium text-right">취득원가</th><th className="px-3 py-2 font-medium">취득일</th>
                <th className="px-3 py-2 font-medium text-right">내용연수</th><th className="px-3 py-2"></th></tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {items.map((i) => (
                <tr key={i.id} className="hover:bg-slate-50/60">
                  <td className="px-3 py-2 font-medium">{i.code}</td>
                  <td className="px-3 py-2 text-slate-600">{i.name}</td>
                  <td className="px-3 py-2 text-right tabular-nums">{fmt(i.unitCost)}</td>
                  <td className="px-3 py-2 text-right tabular-nums text-indigo-700">{fmt(i.dailyRate)}</td>
                  <td className="px-3 py-2 text-right tabular-nums">{fmt(i.acquiredQty)}</td>
                  <td className="px-3 py-2 text-right tabular-nums font-medium">{fmt(i.unitCost * i.acquiredQty)}</td>
                  <td className="px-3 py-2 text-slate-500 tabular-nums">{i.acquiredDate}</td>
                  <td className="px-3 py-2 text-right tabular-nums">{i.usefulLifeMonths}월</td>
                  <td className="px-3 py-2 text-right"><button onClick={() => del(i.id)} className="text-slate-300 hover:text-rose-500"><Trash2 size={14} /></button></td>
                </tr>
              ))}
            </tbody>
          </table>
          {items.length === 0 && <p className="text-center text-sm text-slate-400 py-10">등록된 아이템이 없습니다.</p>}
        </div>
      </Card>
    </div>
  );
}

function CurrencyMaster({ data, persist, flash, period }) {
  const { currencies, fxRates } = data;
  const [f, setF] = useState({ code: "", name: "", symbol: "", quoteUnit: "1" });
  const set = (k, v) => setF((p) => ({ ...p, [k]: v }));
  const foreign = currencies.filter((c) => !c.base);

  const addCur = () => {
    const code = f.code.trim().toUpperCase();
    if (!code) return flash("통화 코드를 입력하세요");
    if (currencies.some((c) => c.code === code)) return flash("이미 있는 통화입니다");
    persist({ ...data, currencies: [...currencies, { code, name: f.name.trim() || code, symbol: f.symbol.trim() || code, quoteUnit: Number(f.quoteUnit) || 1 }] });
    flash("통화 추가 완료"); setF({ code: "", name: "", symbol: "", quoteUnit: "1" });
  };
  const delCur = (code) => {
    if ((data.contracts || []).some((c) => (c.currency || "KRW") === code)) return flash("이 통화를 쓰는 계약이 있어 삭제할 수 없습니다");
    persist({ ...data, currencies: currencies.filter((c) => c.code !== code) });
  };
  const setRate = (code, p, v) => {
    const fr = { ...(fxRates || {}) }; fr[code] = { ...(fr[code] || {}) };
    if (v === "") delete fr[code][p]; else fr[code][p] = Number(v);
    persist({ ...data, fxRates: fr });
  };
  const months = useMemo(() => {
    const s = new Set([period]);
    foreign.forEach((c) => Object.keys(fxRates[c.code] || {}).forEach((m) => s.add(m)));
    return [...s].sort().reverse();
  }, [fxRates, foreign, period]);

  return (
    <div className="space-y-3">
      <Card className="p-4">
        <h3 className="text-sm font-bold mb-2">통화</h3>
        <div className="grid md:grid-cols-6 gap-2">
          <input placeholder="코드(예 USD)" value={f.code} onChange={(e) => set("code", e.target.value)} className="mi" />
          <input placeholder="명칭" value={f.name} onChange={(e) => set("name", e.target.value)} className="mi md:col-span-2" />
          <input placeholder="기호" value={f.symbol} onChange={(e) => set("symbol", e.target.value)} className="mi" />
          <input type="number" placeholder="고시단위" value={f.quoteUnit} onChange={(e) => set("quoteUnit", e.target.value)} className="mi tabular-nums" />
          <Btn tone="primary" onClick={addCur} className="justify-center"><Plus size={15} /> 추가</Btn>
        </div>
        <div className="flex flex-wrap gap-2 mt-3">
          {currencies.map((c) => (
            <span key={c.code} className="inline-flex items-center gap-1.5 rounded-lg ring-1 ring-slate-200 px-2.5 py-1 text-xs">
              <b>{c.code}</b> <span className="text-slate-400">{c.name}{c.quoteUnit > 1 ? ` ·${c.quoteUnit}단위` : ""}</span>
              {c.base ? <span className="text-[10px] text-teal-600 font-medium">기준</span> : <button onClick={() => delCur(c.code)} className="text-slate-300 hover:text-rose-500"><X size={12} /></button>}
            </span>
          ))}
        </div>
        <p className="text-[11px] text-slate-400 mt-2">기준통화는 KRW입니다. 고시단위는 환율 고시 기준(예: JPY 100). 환산 = 외화 × 환율 ÷ 고시단위.</p>
        <style>{`.mi{border-radius:.5rem;font-size:.875rem;padding:.5rem .625rem;background:#fff;box-shadow:inset 0 0 0 1px #cbd5e1;outline:none}.mi:focus{box-shadow:inset 0 0 0 2px #4f46e5}`}</style>
      </Card>

      <Card className="overflow-hidden">
        <div className="px-4 py-3 border-b border-slate-100">
          <h3 className="text-sm font-bold">월평균 환율 (KRW 환산용)</h3>
          <p className="text-[11px] text-slate-400 mt-0.5">고시단위당 원. 기준월 {period} 행에 입력하면 그 달 청구가 환산됩니다.</p>
        </div>
        {foreign.length === 0 ? (
          <p className="p-6 text-center text-sm text-slate-400">외화 통화를 추가하면 환율을 입력할 수 있습니다.</p>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead className="bg-slate-50 text-slate-500 text-xs text-left">
                <tr><th className="px-4 py-2 font-medium">월</th>
                  {foreign.map((c) => <th key={c.code} className="px-4 py-2 font-medium text-right">{c.code}{c.quoteUnit > 1 ? `(${c.quoteUnit})` : ""}</th>)}</tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {months.map((m) => (
                  <tr key={m} className={m === period ? "bg-indigo-50/40" : ""}>
                    <td className="px-4 py-1.5 tabular-nums font-medium">{m}{m === period && <span className="ml-1 text-[10px] text-indigo-600">기준</span>}</td>
                    {foreign.map((c) => (
                      <td key={c.code} className="px-4 py-1.5 text-right">
                        <input type="number" value={(fxRates[c.code] || {})[m] ?? ""} onChange={(e) => setRate(c.code, m, e.target.value)} placeholder="0"
                          className="w-24 text-right rounded ring-1 ring-slate-200 px-2 py-1 tabular-nums focus:ring-2 focus:ring-indigo-500 outline-none" />
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>
    </div>
  );
}

/* ---------- 회계·청구 ---------- */
function StocktakeView({ data, persist, flash, inventory, storeMap, custMap, userEmail }) {
  const items = data.items || [];
  const stores = data.stores || [];
  const customers = data.customers || [];
  const locOptions = useMemo(() => {
    const opts = [{ id: WAREHOUSE, name: "우리 창고" }];
    stores.forEach((s) => opts.push({ id: s.id, name: s.name }));
    customers.forEach((c) => opts.push({ id: cwhId(c.id), name: `${c.name} 창고(보관)` }));
    return opts;
  }, [stores, customers]);

  const [loc, setLoc] = useState(WAREHOUSE);
  const [date, setDate] = useState(todayISO());
  const [counts, setCounts] = useState({});
  const [viewing, setViewing] = useState(null); // 이력 조회 중인 실사
  const [confirmOpen, setConfirmOpen] = useState(false);

  const bookRows = useMemo(() => {
    const rows = [];
    items.forEach((it) => { const b = (inventory[it.id]?.[loc]) || 0; if (b > 0) rows.push({ item: it, book: b }); });
    return rows.sort((a, b) => b.book - a.book);
  }, [items, inventory, loc]);

  useEffect(() => {
    const init = {};
    items.forEach((it) => { const b = (inventory[it.id]?.[loc]) || 0; if (b > 0) init[it.id] = String(b); });
    setCounts(init);
  }, [loc]); // eslint-disable-line react-hooks/exhaustive-deps

  const setCount = (id, v) => setCounts((p) => ({ ...p, [id]: v }));
  const locName = locOptions.find((o) => o.id === loc)?.name || loc;

  const diffs = bookRows.map((r) => {
    const raw = counts[r.item.id];
    const counted = raw === "" || raw == null ? null : Number(raw);
    const diff = counted == null ? null : counted - r.book;
    return { ...r, counted, diff };
  });
  const shortage = diffs.filter((d) => d.diff != null && d.diff < 0);
  const surplus = diffs.filter((d) => d.diff != null && d.diff > 0);
  const shortQty = shortage.reduce((a, d) => a - d.diff, 0);
  const surpQty = surplus.reduce((a, d) => a + d.diff, 0);

  const save = () => {
    const lines = bookRows.map((r) => ({ itemId: r.item.id, book: r.book, counted: counts[r.item.id] === "" || counts[r.item.id] == null ? null : Number(counts[r.item.id]) }));
    const rec = { id: `st_${Date.now()}`, date, locId: loc, locName, lines, status: "draft", createdBy: userEmail || "", createdAt: new Date().toISOString() };
    persist({ ...data, stocktakes: [rec, ...(data.stocktakes || [])] });
    flash("재고실사 저장됨 (임시저장)");
  };

  const custId = loc === WAREHOUSE ? null : isCwh(loc) ? cwhCust(loc) : (storeMap[loc]?.customerId || null);
  const doConfirm = () => {
    const pm = (date || todayISO()).slice(0, 7);
    const newTxns = [], claimLines = [];
    diffs.forEach((d) => {
      if (d.diff == null || d.diff === 0) return;
      if (d.diff < 0) {
        const qty = -d.diff;
        newTxns.push({ id: uid(), type: "scrap", date, itemId: d.item.id, qty, fromStoreId: loc, memo: `재고실사 부족 (${locName})`, createdBy: userEmail || "" });
        if (custId) {
          const unit = Number(d.item.unitCost) || 0, life = Number(d.item.usefulLifeMonths) || 0;
          const el = monthsElapsed(d.item.acquiredDate, pm);
          const bpu = life > 0 ? Math.max(0, unit - (unit / life) * Math.min(el, life)) : unit;
          claimLines.push({ itemId: d.item.id, code: d.item.code, name: d.item.name, qty, unitBook: bpu, amount: bpu * qty });
        }
      } else {
        newTxns.push({ id: uid(), type: "adjust", date, itemId: d.item.id, qty: d.diff, toStoreId: loc, memo: `재고실사 과잉 (${locName})`, createdBy: userEmail || "" });
      }
    });
    const stId = `st_${Date.now()}`;
    const claimsArr = [...(data.claims || [])];
    let claimRec = null;
    if (custId && claimLines.length) {
      claimRec = { id: `cl_${Date.now()}`, date, customerId: custId, customerName: custMap[custId]?.name || "", locName, stocktakeId: stId, lines: claimLines, total: claimLines.reduce((a, l) => a + l.amount, 0) };
      claimsArr.unshift(claimRec);
    }
    const lines = bookRows.map((r) => ({ itemId: r.item.id, book: r.book, counted: counts[r.item.id] === "" || counts[r.item.id] == null ? null : Number(counts[r.item.id]) }));
    const stRec = { id: stId, date, locId: loc, locName, lines, status: "confirmed", claimId: claimRec?.id || null, createdBy: userEmail || "", createdAt: new Date().toISOString(), confirmedAt: new Date().toISOString() };
    persist({ ...data, transactions: [...(data.transactions || []), ...newTxns], claims: claimsArr, stocktakes: [stRec, ...(data.stocktakes || [])] });
    setConfirmOpen(false);
    flash("실사 확정 — 폐기·변상 반영됨");
  };

  const downloadTemplate = () => {
    const aoa = [["재고실사 양식"], ["위치", locName], ["실사일자", date], [], ["아이템코드", "아이템명", "장부수량", "실사수량"]];
    bookRows.forEach((r) => aoa.push([r.item.code, r.item.name, r.book, ""]));
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 16 }, { wch: 24 }, { wch: 10 }, { wch: 10 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "실사");
    XLSX.writeFile(wb, `재고실사_${locName}_${date}.xlsx`);
  };

  const uploadTemplate = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const aoa = XLSX.utils.sheet_to_json(ws, { header: 1 });
        const hi = aoa.findIndex((r) => String(r[0] || "").includes("아이템코드"));
        if (hi < 0) { flash("양식을 찾을 수 없습니다"); return; }
        const next = { ...counts }; let n = 0;
        for (let i = hi + 1; i < aoa.length; i++) {
          const code = String(aoa[i][0] || "").trim(); const cnt = aoa[i][3];
          if (!code) continue;
          const it = items.find((x) => x.code === code);
          if (it && cnt !== "" && cnt != null && !isNaN(Number(cnt))) { next[it.id] = String(Number(cnt)); n++; }
        }
        setCounts(next);
        flash(`실사수량 ${n}건 불러옴`);
      } catch (err) { flash("엑셀 읽기 실패"); }
    };
    reader.readAsArrayBuffer(file);
  };

  const fmtDate = (s) => { try { return new Date(s).toLocaleDateString("ko-KR", { year: "2-digit", month: "2-digit", day: "2-digit" }); } catch { return s; } };
  const history = data.stocktakes || [];

  return (
    <div className="space-y-4">
      <Card className="p-4">
        <h2 className="text-base font-bold mb-1">재고실사</h2>
        <p className="text-[11px] text-slate-400 mb-4">위치별 장부수량과 실물(실사)을 대조합니다. 실사수량을 직접 입력하거나 엑셀 양식으로 일괄 입력하세요.</p>
        <div className="flex flex-wrap items-end gap-3">
          <label className="text-xs text-slate-500">위치
            <select value={loc} onChange={(e) => { setLoc(e.target.value); setViewing(null); }} className="mt-1 block rounded-lg ring-1 ring-slate-300 px-2 py-1.5 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-indigo-400 max-w-[200px]">
              {locOptions.map((o) => <option key={o.id} value={o.id}>{o.name}</option>)}
            </select>
          </label>
          <label className="text-xs text-slate-500">실사일자
            <input type="date" value={date} onChange={(e) => setDate(e.target.value)} className="mt-1 block rounded-lg ring-1 ring-slate-300 px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-400" />
          </label>
          <Btn onClick={downloadTemplate}><Download size={15} /> 엑셀 양식</Btn>
          <label className="inline-flex items-center gap-1.5 rounded-lg ring-1 ring-slate-300 px-3 py-1.5 text-sm cursor-pointer hover:bg-slate-50">
            <Upload size={15} /> 엑셀 업로드
            <input type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => { uploadTemplate(e.target.files[0]); e.target.value = ""; }} />
          </label>
        </div>
      </Card>

      <div className="grid grid-cols-3 gap-3">
        <BizStat label="실사 품목" value={`${fmt(bookRows.length)}종`} tone="slate" />
        <BizStat label="부족 (분실·파손)" value={`${fmt(shortQty)}개`} tone="rose" />
        <BizStat label="과잉" value={`${fmt(surpQty)}개`} tone="amber" />
      </div>

      <Card className="overflow-hidden">
        <div className="px-4 py-3 border-b border-slate-100 flex items-center justify-between">
          <h3 className="text-sm font-bold">{locName} · 장부 vs 실사</h3>
          <div className="flex gap-2">
            <Btn onClick={save}>임시저장</Btn>
            <Btn tone="primary" onClick={() => setConfirmOpen(true)} disabled={shortage.length === 0 && surplus.length === 0}><Check size={15} /> 실사 확정</Btn>
          </div>
        </div>
        {bookRows.length === 0 ? (
          <p className="text-sm text-slate-400 text-center py-10">이 위치에 장부상 재고가 없습니다.</p>
        ) : (
          <div className="overflow-x-auto max-h-[55vh]">
            <table className="w-full text-sm">
              <thead className="sticky top-0 bg-slate-50 text-slate-500 text-xs text-left">
                <tr><th className="px-4 py-2 font-medium">아이템</th><th className="px-4 py-2 font-medium text-right">장부</th>
                  <th className="px-4 py-2 font-medium text-right">실사</th><th className="px-4 py-2 font-medium text-right">차이</th></tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {diffs.map((d) => (
                  <tr key={d.item.id} className={d.diff ? (d.diff < 0 ? "bg-rose-50/40" : "bg-amber-50/40") : ""}>
                    <td className="px-4 py-1.5"><span className="font-medium">{d.item.code}</span> <span className="text-slate-500 text-xs">{d.item.name}</span></td>
                    <td className="px-4 py-1.5 text-right tabular-nums text-slate-500">{fmt(d.book)}</td>
                    <td className="px-4 py-1.5 text-right">
                      <input value={counts[d.item.id] ?? ""} onChange={(e) => setCount(d.item.id, e.target.value)} inputMode="numeric"
                        className="w-20 rounded-md ring-1 ring-slate-200 px-2 py-1 text-sm text-right tabular-nums focus:ring-2 focus:ring-indigo-400 focus:outline-none" />
                    </td>
                    <td className={`px-4 py-1.5 text-right tabular-nums font-semibold ${d.diff == null ? "text-slate-300" : d.diff < 0 ? "text-rose-600" : d.diff > 0 ? "text-amber-600" : "text-slate-400"}`}>
                      {d.diff == null ? "—" : d.diff > 0 ? `+${fmt(d.diff)}` : fmt(d.diff)}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
        {shortage.length > 0 && (
          <div className="px-4 py-2.5 bg-rose-50 border-t border-rose-100 text-[11px] text-rose-600">
            부족 {shortage.length}종 {fmt(shortQty)}개 — 확정 시 폐기 처리 및 고객사 변상청구 대상입니다. (다음 단계에서 처리)
          </div>
        )}
      </Card>

      {history.length > 0 && (
        <Card className="overflow-hidden">
          <div className="px-4 py-3 border-b border-slate-100"><h3 className="text-sm font-bold">실사 이력</h3></div>
          <div className="divide-y divide-slate-100">
            {history.map((h) => {
              const sh = (h.lines || []).filter((l) => l.counted != null && l.counted < l.book);
              const shq = sh.reduce((a, l) => a + (l.book - l.counted), 0);
              return (
                <div key={h.id} className="px-4 py-2.5 flex items-center justify-between text-sm">
                  <div>
                    <span className="font-medium">{h.locName}</span>
                    <span className="text-slate-400 text-xs ml-2">{fmtDate(h.date)}</span>
                    {h.status === "draft" && <span className="text-[10px] text-amber-600 ml-2">임시저장</span>}
                  </div>
                  <div className="text-xs text-slate-500 tabular-nums">
                    {sh.length > 0 ? <span className="text-rose-600">부족 {fmt(shq)}개</span> : <span className="text-slate-400">차이 없음</span>}
                  </div>
                </div>
              );
            })}
          </div>
        </Card>
      )}
      <p className="text-[11px] text-slate-400 px-1">※ 확정 시 부족분은 폐기(자산 제거·처분손실), 매장·고객사창고 부족분은 잔존가액으로 변상청구(미수금)됩니다. 과잉분은 장부 수량이 증가 조정됩니다.</p>

      {confirmOpen && (
        <div className="fixed inset-0 z-50 bg-black/30 grid place-items-center px-4" onClick={() => setConfirmOpen(false)}>
          <div className="bg-white rounded-2xl p-5 w-full max-w-md shadow-xl" onClick={(e) => e.stopPropagation()}>
            <h3 className="font-bold mb-1">재고실사 확정</h3>
            <p className="text-sm text-slate-500 mb-3">{locName} · {date}</p>
            <div className="space-y-1.5 text-sm bg-slate-50 rounded-lg p-3 mb-3">
              {shortage.length > 0 && <div className="flex justify-between"><span className="text-rose-600">부족 → 폐기</span><span className="tabular-nums">{shortage.length}종 {fmt(shortQty)}개</span></div>}
              {surplus.length > 0 && <div className="flex justify-between"><span className="text-amber-600">과잉 → 장부 증가</span><span className="tabular-nums">{surplus.length}종 {fmt(surpQty)}개</span></div>}
              {shortage.length === 0 && surplus.length === 0 && <div className="text-slate-400">차이 없음</div>}
            </div>
            <p className="text-[11px] text-slate-500 mb-1">확정하면:</p>
            <ul className="text-[11px] text-slate-500 list-disc pl-4 space-y-0.5 mb-3">
              <li>부족분은 <b>폐기 거래</b>로 자산에서 제거되고 처분손실이 인식됩니다</li>
              {custId ? <li>이 위치 부족분은 <b>잔존가액으로 변상청구</b>(미수금)됩니다</li> : <li className="text-slate-400">우리 창고 분실은 변상청구 대상이 아닙니다</li>}
              <li>과잉분은 장부 수량이 증가 조정됩니다</li>
            </ul>
            <p className="text-[11px] text-rose-500 mb-4">※ 확정 후에는 거래로 반영되어, 되돌리려면 입출고에서 수동 조정해야 합니다.</p>
            <div className="flex justify-end gap-2">
              <button onClick={() => setConfirmOpen(false)} className="text-sm rounded-lg ring-1 ring-slate-300 px-3 py-1.5 hover:bg-slate-50">취소</button>
              <button onClick={doConfirm} className="text-sm rounded-lg bg-rose-600 text-white px-3 py-1.5 font-semibold hover:bg-rose-700">확정</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function AcctView({ acc, period, setPeriod, data, persist, flash }) {
  const [expanded, setExpanded] = useState({});
  const toggleExp = (id) => setExpanded((e) => ({ ...e, [id]: !e[id] }));
  const journal = buildJournal(acc, period, data.accounts);
  const [pdfCust, setPdfCust] = useState("");
  const billCusts = useMemo(() => {
    const seen = {}, arr = [];
    acc.billItems.forEach((b) => { if (!seen[b.customer.id]) { seen[b.customer.id] = 1; arr.push(b.customer); } });
    return arr;
  }, [acc.billItems]);
  const closed = (data.closedPeriods || []).includes(period);
  const close = () => { if (closed) return; persist({ ...data, closedPeriods: [...(data.closedPeriods || []), period] }); flash(`${period} 마감 완료`); };
  const reopen = () => persist({ ...data, closedPeriods: (data.closedPeriods || []).filter((p) => p !== period) });

  const [claimUnpaidOnly, setClaimUnpaidOnly] = useState(true);
  const toggleClaimPaid = (id) => {
    const claims = (data.claims || []).map((c) => c.id === id ? { ...c, paid: !c.paid, paidDate: !c.paid ? todayISO() : null } : c);
    persist({ ...data, claims });
  };

  const exportBilling = () => {
    const rows = [["고객사", "계약번호", "통화", "과금방식", "태그코드", "tag-day(q×일수)", "일당단가(외화)", "금액(외화)", "금액(KRW)"]];
    acc.billItems.forEach((b) => {
      if (b.billing === "usage" && b.lines.length) {
        b.lines.forEach((ln) => rows.push([b.customer.name, b.contract.contractNo, b.currency, "사용량", ln.item.code, Math.round(ln.tagDays * 10) / 10, Math.round(ln.dailyRate), Math.round(ln.amount), ""]));
        rows.push([b.customer.name, b.contract.contractNo, b.currency, "소계", "", "", "", Math.round(b.amountFx), Math.round(b.amount)]);
      } else {
        rows.push([b.customer.name, b.contract.contractNo, b.currency, BILLING[b.billing], "", "", "", Math.round(b.amountFx), Math.round(b.amount)]);
      }
    });
    const csv = rows.map((r) => r.join(",")).join("\n");
    const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = `청구명세_${period}.csv`; a.click();
  };

  const escH = (s) => String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const wonH = (n) => "₩" + fmt(Math.round(n));
  const invoicePage = (cust, bills) => {
    const sup = data.supplier || {};
    const totalKRW = bills.reduce((a, b) => a + b.amount, 0);
    const rc = (bills.find((b) => b.contract.recipientName) || {}).contract || {};
    const cwhExcl = bills.reduce((a, b) => a + (b.cwhExcluded || 0), 0);
    const supLines = [
      sup.name ? `<div class="supn">${escH(sup.name)}</div>` : "",
      sup.bizNo ? `<div>사업자등록번호 ${escH(sup.bizNo)}</div>` : "",
      sup.ceo ? `<div>대표자 ${escH(sup.ceo)}</div>` : "",
      sup.addr ? `<div>${escH(sup.addr)}</div>` : "",
      sup.tel ? `<div>TEL ${escH(sup.tel)}</div>` : "",
    ].join("");
    const cunit = (c, n) => (c === "KRW" ? "₩" : escH(c) + " ") + fmt(Math.round(n));
    const blocks = bills.map((b) => {
      if (b.billing === "flat") {
        return `<tr class="grp"><td colspan="4">${escH(b.contract.contractNo)} · 월정액</td><td class="r b">${cunit(b.currency, b.amountFx)}${b.missingFx ? ' <span class="warn">환율미입력</span>' : ""}</td></tr>`;
      }
      const sl = b.storeLines || [];
      const head = `<tr class="grp"><td colspan="5">${escH(b.contract.contractNo)} · 사용량 (${escH(b.currency)})${b.missingRate ? ' <span class="warn">단가 미설정 태그</span>' : ""}</td></tr>`;
      const rows = sl.length
        ? sl.map((ln) => `<tr><td>${escH(ln.storeName)}</td><td>${escH(ln.item.code)} ${escH(ln.item.name)}</td><td class="r">${fmt1(ln.tagDays)}</td><td class="r">${ln.dailyRate > 0 ? cunit(b.currency, ln.dailyRate) : '<span class="warn">미설정</span>'}</td><td class="r">${cunit(b.currency, ln.amount)}</td></tr>`).join("")
        : `<tr><td colspan="5" class="sub2">사용 내역 없음</td></tr>`;
      const sub = `<tr class="subr"><td colspan="4" class="r">소계</td><td class="r b">${cunit(b.currency, b.amountFx)}${b.currency !== "KRW" ? ` <span class="sub2">→ ${wonH(b.amount)}</span>` : ""}</td></tr>`;
      return head + rows + sub;
    }).join("");
    return `<section class="inv">
      <div class="head"><div class="ttl">청 구 서</div><div class="per">${escH(period)}</div></div>
      <div class="pty">
        <div class="to"><div class="lbl">받는 분</div><div class="cust">${escH(rc.recipientName || cust.name)} 귀하</div>
          <div class="sub">${escH(cust.name)}</div>
          ${rc.recipientEmail ? `<div class="sub">${escH(rc.recipientEmail)}</div>` : ""}
          ${rc.recipientTel ? `<div class="sub">${escH(rc.recipientTel)}</div>` : ""}</div>
        <div class="from">${supLines || '<div class="sub">공급자 정보 미입력 (설정 탭에서 입력)</div>'}</div>
      </div>
      <table class="tbl"><thead><tr><th>매장</th><th>아이템</th><th class="r">tag-day</th><th class="r">일당단가</th><th class="r">금액</th></tr></thead>
        <tbody>${blocks}</tbody>
        <tfoot><tr><td colspan="4" class="r b">합계 (KRW)</td><td class="r b">${wonH(totalKRW)}</td></tr></tfoot></table>
      ${cwhExcl > 0 ? `<div class="note">※ 고객사 창고 보관분 ${fmt1(cwhExcl)} tag-day는 계약 조건에 따라 청구에서 제외되었습니다.</div>` : ""}
      <div class="foot">본 청구서는 ${escH(period)} 운용리스 이용분에 대한 청구입니다. (tag-day = 수량 × 일수)</div>
    </section>`;
  };
  const printInvoices = (custId) => {
    const byCust = {};
    acc.billItems.forEach((b) => { if (custId && b.customer.id !== custId) return; (byCust[b.customer.id] = byCust[b.customer.id] || { customer: b.customer, bills: [] }).bills.push(b); });
    const groups = Object.values(byCust);
    if (!groups.length) return flash("청구할 내역이 없습니다");
    const pages = groups.map((g) => invoicePage(g.customer, g.bills)).join('<div class="brk"></div>');
    const html = `<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>청구서 ${escH(period)}</title>
      <style>
        *{box-sizing:border-box} body{font-family:-apple-system,'Malgun Gothic','Apple SD Gothic Neo',sans-serif;color:#1e293b;margin:0;padding:24px}
        .inv{max-width:720px;margin:0 auto 24px} .brk{page-break-after:always}
        .head{display:flex;justify-content:space-between;align-items:baseline;border-bottom:3px solid #4f46e5;padding-bottom:10px}
        .ttl{font-size:28px;font-weight:800;letter-spacing:8px} .per{font-size:15px;color:#64748b}
        .pty{display:flex;justify-content:space-between;gap:24px;margin:20px 0}
        .to,.from{font-size:12px;line-height:1.6} .lbl{font-size:11px;color:#94a3b8;margin-bottom:2px}
        .cust{font-size:16px;font-weight:700} .from{text-align:right} .supn{font-size:15px;font-weight:700;margin-bottom:2px}
        .sub{color:#64748b;font-size:11px} .warn{color:#e11d48}
        .tbl{width:100%;border-collapse:collapse;margin-top:8px;font-size:13px}
        .tbl th{background:#f1f5f9;text-align:left;padding:8px 10px;font-weight:600;font-size:12px}
        .tbl td{padding:7px 10px;border-bottom:1px solid #eef2f6} .tbl .r{text-align:right} .tbl .b{font-weight:700}
        .tbl .grp td{background:#eef2ff;font-weight:700;font-size:11px;border-bottom:1px solid #c7d2fe;color:#3730a3}
        .tbl .subr td{border-top:1px solid #cbd5e1;border-bottom:2px solid #e2e8f0;font-size:12px;background:#f8fafc}
        .sub2{color:#94a3b8;font-weight:400;font-size:10px}
        .tbl tfoot td{border-top:2px solid #cbd5e1;border-bottom:none;font-size:14px;padding-top:10px}
        .note{margin-top:12px;font-size:11px;color:#b45309;background:#fffbeb;border:1px solid #fde68a;border-radius:6px;padding:8px 10px}
        .foot{margin-top:24px;font-size:11px;color:#94a3b8;text-align:center}
        @media print{body{padding:0}}
      </style></head><body>${pages}</body></html>`;
    const w = window.open("", "_blank");
    if (!w) return flash("팝업이 차단되었습니다. 브라우저에서 팝업을 허용한 뒤 다시 시도하세요.");
    w.document.write(html); w.document.close();
    setTimeout(() => { w.focus(); w.print(); }, 350);
  };

  const printReimburse = () => {
    const cls = acc.reimbClaims || [];
    if (!cls.length) return flash("이 달 변상청구 내역이 없습니다");
    const won = (n) => "₩" + fmt(n);
    const blocks = cls.map((cl) => {
      const rows = (cl.lines || []).map((l) => `<tr><td>${escH(l.code)} <span class="sub2">${escH(l.name)}</span></td><td class="r">${fmt(l.qty)}</td><td class="r">${won(l.unitBook)}</td><td class="r">${won(l.amount)}</td></tr>`).join("");
      return `<table class="tbl"><thead><tr class="grp"><td colspan="4">${escH(cl.customerName)} · ${escH(cl.locName)} <span class="sub2">(${escH(cl.date)})</span></td></tr><tr><th>아이템</th><th class="r">수량</th><th class="r">잔존가(단가)</th><th class="r">변상액</th></tr></thead><tbody>${rows}</tbody><tfoot><tr><td colspan="3" class="r b">소계</td><td class="r b">${won(cl.total)}</td></tr></tfoot></table>`;
    }).join('<div style="height:14px"></div>');
    const total = cls.reduce((a, cl) => a + (cl.total || 0), 0);
    const sup = data.supplier || {};
    const supLines = [sup.name ? `<div class="supn">${escH(sup.name)}</div>` : "", sup.bizNo ? `<div>사업자등록번호 ${escH(sup.bizNo)}</div>` : "", sup.tel ? `<div>TEL ${escH(sup.tel)}</div>` : ""].join("");
    const html = `<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>변상청구서 ${escH(period)}</title>
      <style>
        *{box-sizing:border-box} body{font-family:-apple-system,'Malgun Gothic','Apple SD Gothic Neo',sans-serif;color:#1e293b;margin:0;padding:24px}
        .inv{max-width:720px;margin:0 auto}
        .head{display:flex;justify-content:space-between;align-items:baseline;border-bottom:3px solid #e11d48;padding-bottom:10px}
        .ttl{font-size:26px;font-weight:800;letter-spacing:6px} .per{font-size:15px;color:#64748b}
        .from{text-align:right;font-size:12px;line-height:1.6;margin:16px 0} .supn{font-size:15px;font-weight:700}
        .desc{font-size:12px;color:#64748b;margin:14px 0;line-height:1.6}
        .tbl{width:100%;border-collapse:collapse;font-size:13px} .tbl .r{text-align:right} .tbl .b{font-weight:700}
        .tbl th{background:#f1f5f9;text-align:left;padding:8px 10px;font-weight:600;font-size:12px}
        .tbl td{padding:7px 10px;border-bottom:1px solid #eef2f6}
        .tbl .grp td{background:#fff1f2;font-weight:700;font-size:12px;color:#9f1239;border-bottom:1px solid #fecdd3}
        .sub2{color:#94a3b8;font-weight:400;font-size:10px}
        .tbl tfoot td{border-top:2px solid #cbd5e1;font-size:13px}
        .tot{margin-top:18px;text-align:right;font-size:16px;font-weight:800}
        .foot{margin-top:24px;font-size:11px;color:#94a3b8;text-align:center;line-height:1.6}
        @media print{body{padding:0}}
      </style></head><body><div class="inv">
        <div class="head"><div class="ttl">변상 청구서</div><div class="per">${escH(period)}</div></div>
        <div class="from">${supLines}</div>
        <div class="desc">아래는 재고실사 결과 분실·미반환으로 확인된 리스자산에 대한 변상 청구입니다. 변상액은 분실 시점의 잔존 장부가액 기준입니다.</div>
        ${blocks}
        <div class="tot">변상 청구 합계 &nbsp; ₩${fmt(total)}</div>
        <div class="foot">본 청구서는 ${escH(period)} 재고실사 기준 분실 리스자산 변상 청구입니다.<br/>변상액 = 분실 수량 × 잔존 장부가액</div>
      </div></body></html>`;
    const w = window.open("", "_blank");
    if (!w) return flash("팝업이 차단되었습니다. 팝업을 허용한 뒤 다시 시도하세요.");
    w.document.write(html); w.document.close();
    setTimeout(() => { w.focus(); w.print(); }, 350);
  };

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div className="flex items-center gap-2">
          <PeriodPicker period={period} setPeriod={setPeriod} />
          {closed ? <span className="rounded-full bg-slate-800 text-white px-2.5 py-1 text-[11px] font-medium">마감됨</span>
            : <span className="rounded-full bg-amber-50 text-amber-700 px-2.5 py-1 text-[11px] font-medium ring-1 ring-amber-200">미마감</span>}
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <select value={pdfCust} onChange={(e) => setPdfCust(e.target.value)} className="rounded-lg ring-1 ring-slate-300 px-2 py-1.5 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-indigo-400 max-w-[150px]">
            <option value="">전체 고객사</option>
            {billCusts.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
          <Btn tone="primary" onClick={() => printInvoices(pdfCust)}><FileText size={15} /> PDF 청구서</Btn>
          {acc.reimbAR > 0 && <Btn onClick={printReimburse}><FileText size={15} /> 변상 청구서</Btn>}
          <Btn onClick={exportBilling}><Download size={15} /> 청구명세 CSV</Btn>
          {closed ? <Btn onClick={reopen}>마감 취소</Btn> : <Btn tone="primary" onClick={close}><Check size={15} /> 월 마감</Btn>}
        </div>
      </div>

      <div className="grid grid-cols-2 lg:grid-cols-3 gap-3">
        <Stat label="리스료수익" v={acc.leaseIncome} tone="emerald" />
        <Stat label="감가상각비" v={acc.depExpense} tone="rose" />
        <Stat label="폐기 처분손실" v={acc.disposalLoss} tone="rose" />
        <Stat label="분실 변상(미수금)" v={acc.reimbAR} tone="slate" />
        <Stat label="리스손익" v={acc.netIncome} tone={acc.netIncome >= 0 ? "emerald" : "rose"} />
        <Stat label="자산 장부가액(순액)" v={acc.bookValue} tone="slate" />
        <Stat label="감가상각누계액" v={acc.depAccum} tone="slate" />
      </div>

      <Card className="overflow-hidden">
        <div className="px-4 py-3 border-b border-slate-100"><h3 className="text-sm font-bold flex items-center gap-1.5"><FileSignature size={15} /> {period} 계약별 청구 명세</h3></div>
        <div className="overflow-x-auto max-h-[44vh]">
          <table className="w-full text-sm">
            <thead className="sticky top-0 bg-slate-50 text-slate-500 text-xs text-left">
              <tr><th className="px-4 py-2 font-medium">고객사 · 계약</th><th className="px-4 py-2 font-medium">방식</th>
                <th className="px-4 py-2 font-medium text-right">tag-day</th><th className="px-4 py-2 font-medium text-right">산식</th>
                <th className="px-4 py-2 font-medium text-right">청구금액</th></tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {acc.billItems.map((b) => {
                const open = expanded[b.contract.id];
                const canExpand = b.billing === "usage" && b.lines.length > 0;
                return (
                  <React.Fragment key={b.contract.id}>
                    <tr className={`hover:bg-slate-50/60 ${canExpand ? "cursor-pointer" : ""}`} onClick={() => canExpand && toggleExp(b.contract.id)}>
                      <td className="px-4 py-2 font-medium">
                        {canExpand && <span className="inline-block w-3 text-slate-400">{open ? "▾" : "▸"}</span>} {b.customer.name}
                        <span className="text-[11px] text-slate-400 ml-1.5">{b.contract.contractNo}</span>
                        {b.missingRate && <span className="ml-1.5 text-[10px] text-rose-600 bg-rose-50 rounded px-1 py-0.5">단가 미설정 태그</span>}
                      </td>
                      <td className="px-4 py-2"><BillBadge b={b.billing} /> <span className="text-[10px] text-slate-400 font-medium ml-1">{b.currency}</span></td>
                      <td className="px-4 py-2 text-right tabular-nums text-slate-500">{b.billing === "usage" ? fmt1(b.tagDays) : "—"}</td>
                      <td className="px-4 py-2 text-right tabular-nums text-[11px] text-slate-400">
                        {b.billing === "usage" ? `Σ 태그별 (p×q×일수) · ${b.lines.length}종` : "계약 월정액"}
                      </td>
                      <td className="px-4 py-2 text-right">
                        {b.missingFx
                          ? <span className="text-rose-600 text-[11px]">환율 미입력</span>
                          : <>
                              <div className="tabular-nums font-semibold">₩{fmt(b.amount)}</div>
                              {b.currency !== "KRW" && <div className="text-[10px] text-slate-400 tabular-nums">{b.currency} {fmt(b.amountFx)}</div>}
                            </>}
                      </td>
                    </tr>
                    {open && b.lines.map((ln) => (
                      <tr key={b.contract.id + ln.item.id} className="bg-slate-50/40 text-[12px]">
                        <td className="pl-10 pr-4 py-1.5 text-slate-500">{ln.item.code} · {ln.item.name}</td>
                        <td className="px-4 py-1.5"></td>
                        <td className="px-4 py-1.5 text-right tabular-nums text-slate-500">{fmt1(ln.tagDays)}</td>
                        <td className="px-4 py-1.5 text-right tabular-nums text-[11px]">
                          {ln.dailyRate > 0
                            ? <span className="text-slate-400">{fmt1(ln.tagDays)} × {b.currency === "KRW" ? "₩" : b.currency + " "}{fmt(ln.dailyRate)}</span>
                            : <span className="text-rose-500">단가 미설정</span>}
                        </td>
                        <td className="px-4 py-1.5 text-right tabular-nums text-slate-600">{b.currency === "KRW" ? "₩" : b.currency + " "}{fmt(ln.amount)}</td>
                      </tr>
                    ))}
                  </React.Fragment>
                );
              })}
            </tbody>
            <tfoot>
              <tr className="border-t-2 border-slate-200 bg-slate-50/50">
                <td className="px-4 py-2.5 font-bold" colSpan={4}>합계 (= 리스료수익)</td>
                <td className="px-4 py-2.5 text-right tabular-nums font-bold">₩{fmt(acc.leaseIncome)}</td>
              </tr>
            </tfoot>
          </table>
          {acc.billItems.length === 0 && <p className="text-center text-sm text-slate-400 py-8">이 달 유효한 계약이 없습니다.</p>}
        </div>
      </Card>

      <Card className="p-4">
        <h3 className="text-sm font-bold mb-1 flex items-center gap-1.5"><FileText size={15} /> {period} 자동 분개</h3>
        <p className="text-[11px] text-slate-400 mb-3">운용리스 · 리스제공자 (정액 수익인식 + 정액 감가상각)</p>
        <JournalTable lines={journal} />
      </Card>

      {(data.claims || []).length > 0 && (() => {
        const claims = data.claims || [];
        const unpaidTotal = claims.filter((c) => !c.paid).reduce((a, c) => a + (c.total || 0), 0);
        const paidTotal = claims.filter((c) => c.paid).reduce((a, c) => a + (c.total || 0), 0);
        const shown = claimUnpaidOnly ? claims.filter((c) => !c.paid) : claims;
        const fmtD = (s) => { try { return new Date(s).toLocaleDateString("ko-KR", { year: "2-digit", month: "2-digit", day: "2-digit" }); } catch { return s; } };
        return (
          <Card className="overflow-hidden">
            <div className="px-4 py-3 border-b border-slate-100 flex items-center justify-between gap-2 flex-wrap">
              <h3 className="text-sm font-bold flex items-center gap-1.5"><Banknote size={15} /> 분실 변상 청구 현황 <span className="text-[11px] text-slate-400 font-normal">(전체 기간)</span></h3>
              <label className="flex items-center gap-1.5 text-xs text-slate-500 cursor-pointer">
                <input type="checkbox" checked={claimUnpaidOnly} onChange={(e) => setClaimUnpaidOnly(e.target.checked)} /> 미수만 보기
              </label>
            </div>
            <div className="grid grid-cols-2 gap-3 p-4">
              <BizStat label="미수 잔액" value={`₩${fmt(unpaidTotal)}`} tone="rose" />
              <BizStat label="수금 완료" value={`₩${fmt(paidTotal)}`} tone="emerald" />
            </div>
            {shown.length === 0 ? (
              <p className="text-sm text-slate-400 text-center py-8">{claimUnpaidOnly ? "미수 변상청구가 없습니다." : "변상청구 내역이 없습니다."}</p>
            ) : (
              <div className="divide-y divide-slate-100">
                {shown.map((c) => (
                  <div key={c.id} className="px-4 py-2.5 flex items-center justify-between gap-3 text-sm">
                    <div className="min-w-0">
                      <span className="font-medium">{c.customerName || "—"}</span>
                      <span className="text-slate-400 text-xs ml-2">{c.locName} · {fmtD(c.date)}</span>
                      {c.paid && c.paidDate && <span className="text-[10px] text-emerald-600 ml-2">수금 {fmtD(c.paidDate)}</span>}
                    </div>
                    <div className="flex items-center gap-3 shrink-0">
                      <span className="tabular-nums font-semibold">₩{fmt(c.total)}</span>
                      <button onClick={() => toggleClaimPaid(c.id)}
                        className={`text-xs rounded-lg px-2.5 py-1 font-medium ring-1 ${c.paid ? "bg-emerald-50 text-emerald-700 ring-emerald-200" : "bg-white text-slate-600 ring-slate-300 hover:bg-slate-50"}`}>
                        {c.paid ? "✓ 수금 완료" : "수금 체크"}
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}
            <p className="px-4 py-2.5 text-[11px] text-slate-400 border-t border-slate-100">※ 변상청구 발생액 표시·수금 관리용입니다. 회계 미수금 분개는 청구 발생 기준 그대로이며, 수금 체크는 회계에 반영되지 않습니다.</p>
          </Card>
        );
      })()}

      <Card className="p-4 bg-slate-50/60">
        <div className="flex gap-2 text-slate-500 text-xs leading-relaxed">
          <AlertTriangle size={15} className="shrink-0 mt-0.5 text-amber-500" />
          <p><b className="text-slate-700">검증 필요:</b> 청구는 그 달 '유효 계약'을 기준으로 산정됩니다 — 사용량은 계약 단가표(p) × 수량(q) × 일수, 월정액은 계약 단일 정액(창고 보관분 제외).
          한 고객사에 같은 달 유효 계약이 둘 이상이면 각각 합산되니, 갱신 시 이전 계약 종료일을 꼭 지정하세요.
          실제 적용 전 단가표·tag-day 경계(시작일 포함·종료일 제외)·개설직접원가·중도해지·잔존가치, ERP 본계정 대사를 확인하세요.</p>
        </div>
      </Card>
    </div>
  );
}
const Stat = ({ label, v, tone }) => {
  const map = { emerald: "text-emerald-600", rose: "text-rose-600", slate: "text-slate-600" };
  return (
    <Card className="p-4">
      <div className="text-[11px] text-slate-400 font-medium">{label}</div>
      <div className={`text-xl font-bold tabular-nums mt-1 ${map[tone]}`}>₩{fmt(v)}</div>
    </Card>
  );
};

function JournalTable({ lines }) {
  if (!lines || lines.length === 0) return <p className="text-sm text-slate-400 py-4 text-center">해당 월 인식할 항목이 없습니다.</p>;
  return (
    <div className="space-y-2.5">
      {lines.map((l, i) => (
        <div key={i} className="rounded-lg ring-1 ring-slate-200 overflow-hidden">
          <div className="px-4 py-2.5 bg-slate-50 border-b border-slate-100 text-[13px] font-bold text-slate-800">{l.memo}</div>
          <div className="px-4 py-3 space-y-2 text-sm">
            <div className="flex items-center justify-between">
              <span className="text-slate-700">{l.debit.account}</span>
              <span className="tabular-nums font-semibold">₩{fmt(l.debit.amount)}</span>
            </div>
            <div className="flex items-center justify-between text-slate-500">
              <span className="pl-6">{l.credit.account}</span>
              <span className="tabular-nums">₩{fmt(l.credit.amount)}</span>
            </div>
          </div>
        </div>
      ))}
    </div>
  );
}

/* ---------- 설정 ---------- */
function SettingsView({ data, persist, flash }) {
  const sup = data.supplier || {};
  const [f, setF] = useState({ name: sup.name || "", bizNo: sup.bizNo || "", ceo: sup.ceo || "", addr: sup.addr || "", tel: sup.tel || "" });
  const set = (k, v) => setF((p) => ({ ...p, [k]: v }));
  const [confirm, setConfirm] = useState(null); // 'reset' | 'wipe'

  const saveSupplier = () => { persist({ ...data, supplier: f }); flash("공급자 정보를 저장했습니다"); };
  const doReset = () => { persist(makeSample()); setConfirm(null); flash("샘플 데이터로 초기화했습니다"); };
  const doWipe = () => {
    persist({ customers: [], items: [], stores: [], contracts: [], transactions: [], currencies: DEFAULT_CURRENCIES, fxRates: {}, closedPeriods: [], supplier: {} });
    setConfirm(null); flash("모든 데이터를 삭제했습니다");
  };

  const counts = `고객사 ${data.customers.length} · 매장 ${data.stores.length} · 아이템 ${data.items.length} · 계약 ${data.contracts.length} · 수불 ${data.transactions.length}건`;

  return (
    <div className="space-y-3 max-w-3xl">
      <Card className="p-4">
        <h3 className="text-sm font-bold flex items-center gap-1.5"><FileText size={15} /> 공급자(청구서 발신) 정보</h3>
        <p className="text-[11px] text-slate-400 mt-0.5 mb-3">PDF 청구서 상단에 표시됩니다. 비워두면 해당 항목은 생략됩니다.</p>
        <div className="grid md:grid-cols-2 gap-2">
          <input placeholder="상호 (예: ㈜솔루엠)" value={f.name} onChange={(e) => set("name", e.target.value)} className="mi" />
          <input placeholder="사업자등록번호" value={f.bizNo} onChange={(e) => set("bizNo", e.target.value)} className="mi" />
          <input placeholder="대표자" value={f.ceo} onChange={(e) => set("ceo", e.target.value)} className="mi" />
          <input placeholder="연락처" value={f.tel} onChange={(e) => set("tel", e.target.value)} className="mi" />
          <input placeholder="주소" value={f.addr} onChange={(e) => set("addr", e.target.value)} className="mi md:col-span-2" />
        </div>
        <div className="mt-3"><Btn tone="primary" onClick={saveSupplier}><Check size={15} /> 저장</Btn></div>
        <style>{`.mi{border-radius:.5rem;font-size:.875rem;padding:.5rem .625rem;background:#fff;box-shadow:inset 0 0 0 1px #cbd5e1;outline:none}.mi:focus{box-shadow:inset 0 0 0 2px #4f46e5}`}</style>
      </Card>

      <Card className="p-4">
        <h3 className="text-sm font-bold flex items-center gap-1.5"><Database size={15} /> 데이터 관리</h3>
        <p className="text-[11px] text-slate-400 mt-0.5 mb-3">현재 {counts}</p>
        <div className="flex flex-wrap gap-2">
          <Btn onClick={() => setConfirm("reset")}><RefreshCw size={15} /> 샘플로 초기화</Btn>
          <Btn tone="danger" onClick={() => setConfirm("wipe")}><Trash2 size={15} /> 모든 데이터 삭제</Btn>
        </div>
        <p className="text-[11px] text-slate-400 mt-2">'샘플로 초기화'는 기존 데이터를 지우고 데모 샘플로 덮어씁니다. '모든 데이터 삭제'는 전부 비웁니다(통화 기본값만 유지).</p>
      </Card>

      {confirm && (
        <div className="fixed inset-0 z-50 grid place-items-center bg-black/40 p-4" onClick={() => setConfirm(null)}>
          <div className="bg-white rounded-2xl shadow-xl max-w-sm w-full p-5" onClick={(e) => e.stopPropagation()}>
            <div className="flex items-center gap-2 mb-2">
              <AlertTriangle size={20} className={confirm === "wipe" ? "text-rose-600" : "text-amber-500"} />
              <h4 className="font-bold">{confirm === "wipe" ? "모든 데이터를 삭제할까요?" : "샘플로 초기화할까요?"}</h4>
            </div>
            <p className="text-sm text-slate-500 mb-1">
              {confirm === "wipe"
                ? "고객사·매장·아이템·계약·수불·환율이 모두 삭제됩니다. 이 작업은 되돌릴 수 없습니다."
                : "현재 입력된 데이터가 모두 사라지고 데모 샘플로 덮어써집니다."}
            </p>
            <p className="text-[11px] text-slate-400 mb-4">현재 {counts}</p>
            <div className="flex justify-end gap-2">
              <Btn onClick={() => setConfirm(null)}>취소</Btn>
              <Btn tone={confirm === "wipe" ? "danger" : "primary"} onClick={confirm === "wipe" ? doWipe : doReset}>
                {confirm === "wipe" ? <><Trash2 size={15} /> 영구 삭제</> : <><RefreshCw size={15} /> 초기화</>}
              </Btn>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------- 로그인 ---------- */
function ModeSelect({ email, onPick, onSignOut }) {
  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 grid place-items-center px-4" style={{ fontFamily: "system-ui, sans-serif" }}>
      <div className="w-full max-w-lg">
        <div className="text-center mb-7">
          <div className="inline-grid place-items-center h-11 w-11 rounded-xl bg-indigo-600 text-white mb-3"><Package size={22} /></div>
          <h1 className="text-lg font-bold">SoluM 리스 플랫폼</h1>
          <p className="text-sm text-slate-400 mt-1">사용할 모드를 선택하세요</p>
        </div>
        <div className="grid sm:grid-cols-2 gap-3">
          <button onClick={() => onPick("ops")} className="text-left rounded-2xl bg-white ring-1 ring-slate-200 p-5 hover:ring-indigo-400 hover:shadow-md transition">
            <div className="inline-grid place-items-center h-10 w-10 rounded-lg bg-indigo-50 text-indigo-600 mb-3"><Boxes size={20} /></div>
            <div className="font-bold">운영 관리</div>
            <p className="text-[13px] text-slate-500 mt-1 leading-relaxed">재고·입출고·청구·회계. ESL 태그 운용리스의 실제 운영을 관리합니다.</p>
          </button>
          <button onClick={() => onPick("biz")} className="text-left rounded-2xl bg-white ring-1 ring-slate-200 p-5 hover:ring-emerald-400 hover:shadow-md transition">
            <div className="inline-grid place-items-center h-10 w-10 rounded-lg bg-emerald-50 text-emerald-600 mb-3"><TrendingUp size={20} /></div>
            <div className="font-bold">사업성 분석</div>
            <p className="text-[13px] text-slate-500 mt-1 leading-relaxed">HaaS 구독 사업의 NPV·IRR·투자회수 분석과 의사결정 리포트.</p>
          </button>
        </div>
        <div className="text-center mt-6 text-xs text-slate-400">
          <span className="mr-2">{email}</span>
          <button onClick={onSignOut} className="underline hover:text-slate-600">로그아웃</button>
        </div>
      </div>
    </div>
  );
}

function BizShell({ email, onSwitch, onSignOut, flash }) {
  const [bizTab, setBizTab] = useState("analyze");
  const [inp, setInp] = useState(BIZ_DEFAULTS);
  const [reports, setReports] = useState([]);
  const [loadingR, setLoadingR] = useState(false);
  const [editing, setEditing] = useState(null);
  const res = useMemo(() => computeBiz(inp), [inp]);

  const loadReports = useCallback(async () => {
    setLoadingR(true);
    try {
      const { data, error } = await supabase.from("biz_reports").select("*").order("updated_at", { ascending: false });
      if (!error && data) setReports(data);
      else if (error) flash("리포트 목록 불러오기 실패 — 테이블 확인");
    } catch (e) { flash("네트워크 오류"); }
    setLoadingR(false);
  }, [flash]);
  useEffect(() => { loadReports(); }, [loadReports]);

  const saveReport = async (name, memo) => {
    const s = res;
    const summary = { NPV: s.NPV, IRR: s.IRR, PB: s.PB, npm: s.npm, gpm: s.gpm, WACC: s.WACC, verdict: s.verdict, maxTerm: s.maxTerm, init: s.init, currency: inp.currency };
    try {
      if (editing?.id) {
        const { error } = await supabase.from("biz_reports").update({ name, memo, input: inp, summary, updated_at: new Date().toISOString() }).eq("id", editing.id);
        if (error) return flash("저장 실패");
        setEditing({ id: editing.id, name, memo }); flash("리포트 수정됨");
      } else {
        const { data, error } = await supabase.from("biz_reports").insert({ name, memo, input: inp, summary, created_by: email || "" }).select().single();
        if (error || !data) return flash("저장 실패");
        setEditing({ id: data.id, name, memo }); flash("리포트 저장됨");
      }
      loadReports();
    } catch (e) { flash("저장 실패 — 네트워크 확인"); }
  };
  const openReport = (r) => { setInp({ ...BIZ_DEFAULTS, ...(r.input || {}) }); setEditing({ id: r.id, name: r.name, memo: r.memo || "" }); setBizTab("analyze"); };
  const newReport = () => { setInp(BIZ_DEFAULTS); setEditing(null); setBizTab("analyze"); };
  const deleteReport = async (id) => {
    try { const { error } = await supabase.from("biz_reports").delete().eq("id", id); if (error) return flash("삭제 실패"); if (editing?.id === id) setEditing(null); loadReports(); flash("삭제됨"); }
    catch (e) { flash("삭제 실패"); }
  };

  const BTAB = [{ id: "analyze", label: "새 분석", icon: TrendingUp }, { id: "reports", label: "저장된 리포트", icon: FileText }];
  return (
    <div className="min-h-screen bg-slate-50 text-slate-800" style={{ fontFamily: "system-ui, sans-serif" }}>
      <header className="sticky top-0 z-20 border-b border-slate-200 bg-white/90 backdrop-blur">
        <div className="mx-auto max-w-6xl px-4 py-3 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="grid place-items-center h-8 w-8 rounded-lg bg-emerald-600 text-white"><TrendingUp size={18} /></div>
            <div>
              <h1 className="text-[15px] font-bold tracking-tight leading-none">사업성 분석</h1>
              <p className="text-[11px] text-slate-400 mt-0.5">HaaS 구독 사업 — NPV · IRR · 투자회수</p>
            </div>
          </div>
          <div className="flex items-center gap-2 text-xs text-slate-500">
            <button onClick={onSwitch} className="rounded-lg ring-1 ring-slate-300 px-2 py-1 hover:bg-slate-50 whitespace-nowrap font-medium">← 운영 관리</button>
            <span className="hidden sm:inline max-w-[140px] truncate">{email}</span>
            <button onClick={onSignOut} className="rounded-lg ring-1 ring-slate-300 px-2 py-1 hover:bg-slate-50 whitespace-nowrap">로그아웃</button>
          </div>
        </div>
        <nav className="mx-auto max-w-6xl px-2 flex gap-1 overflow-x-auto">
          {BTAB.map((n) => {
            const Icon = n.icon, on = bizTab === n.id;
            return (
              <button key={n.id} onClick={() => setBizTab(n.id)}
                className={`flex items-center gap-1.5 px-3 py-2.5 text-sm font-medium border-b-2 -mb-px whitespace-nowrap transition ${on ? "border-emerald-600 text-emerald-700" : "border-transparent text-slate-500 hover:text-slate-700"}`}>
                <Icon size={16} /> {n.label}{n.id === "reports" && reports.length ? ` (${reports.length})` : ""}
              </button>
            );
          })}
        </nav>
      </header>
      <main className="mx-auto max-w-6xl px-4 py-5">
        {bizTab === "analyze" && <BizAnalyze {...{ inp, setInp, res, editing, onSave: saveReport, onNew: newReport }} />}
        {bizTab === "reports" && <BizReports {...{ reports, loading: loadingR, onOpen: openReport, onDelete: deleteReport, editingId: editing?.id }} />}
      </main>
    </div>
  );
}

function BizReports({ reports, loading, onOpen, onDelete, editingId }) {
  if (loading) return <div className="text-slate-400 text-sm py-10 text-center">불러오는 중…</div>;
  if (!reports.length) return <Card className="p-10 text-center text-slate-400 text-sm">저장된 리포트가 없습니다.<br />'새 분석'에서 분석 후 <b>리포트로 저장</b>해 보세요.</Card>;
  const fmtDate = (s) => { try { return new Date(s).toLocaleDateString("ko-KR", { year: "2-digit", month: "2-digit", day: "2-digit" }); } catch { return ""; } };
  const vB = (v) => v === "GO" ? "bg-emerald-100 text-emerald-700" : v === "재검토" ? "bg-amber-100 text-amber-700" : "bg-rose-100 text-rose-700";
  return (
    <div className="space-y-3">
      {reports.map((r) => {
        const s = r.summary || {};
        const c = s.currency || "KRW";
        return (
          <Card key={r.id} className={`p-4 ${editingId === r.id ? "ring-2 ring-emerald-400" : ""}`}>
            <div className="flex items-start justify-between gap-3">
              <button onClick={() => onOpen(r)} className="text-left flex-1 min-w-0">
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="font-bold truncate">{r.name}</span>
                  <span className={`px-2 py-0.5 rounded-full text-[11px] font-semibold ${vB(s.verdict)}`}>{s.verdict || "—"}</span>
                  {editingId === r.id && <span className="text-[10px] text-emerald-600 font-medium">편집 중</span>}
                </div>
                {r.memo && <p className="text-xs text-slate-500 mt-1 line-clamp-2">{r.memo}</p>}
                <div className="flex flex-wrap gap-x-3 gap-y-0.5 mt-2 text-[11px] text-slate-400 tabular-nums">
                  <span>NPV {c} {fmt(Math.round(s.NPV || 0))}</span>
                  <span>IRR {s.IRR == null ? "—" : (s.IRR * 100).toFixed(1) + "%"}</span>
                  <span>회수 {s.PB == null ? "—" : s.PB.toFixed(0) + "개월"}</span>
                  <span>순이익률 {s.npm == null ? "—" : (s.npm * 100).toFixed(1) + "%"}</span>
                  <span className="text-slate-300">·</span>
                  <span>{fmtDate(r.updated_at)}</span>
                </div>
              </button>
              <button onClick={() => { if (confirm(`'${r.name}' 리포트를 삭제할까요?`)) onDelete(r.id); }} className="text-slate-300 hover:text-rose-500 p-1 shrink-0"><Trash2 size={16} /></button>
            </div>
          </Card>
        );
      })}
    </div>
  );
}

function BizAnalyze({ inp, setInp, res, editing, onSave, onNew }) {
  const [saveOpen, setSaveOpen] = useState(false);
  const [sName, setSName] = useState("");
  const [sMemo, setSMemo] = useState("");
  const [showAdv, setShowAdv] = useState(false);
  const setF = (k, v) => setInp((p) => ({ ...p, [k]: v }));
  const setModel = (i, k, v) => setInp((p) => ({ ...p, models: p.models.map((m, j) => j === i ? { ...m, [k]: v } : m) }));
  const setCR = (k, v) => setInp((p) => ({ ...p, costRatios: { ...p.costRatios, [k]: v } }));
  const setFx = (c, v) => setInp((p) => ({ ...p, fxTable: { ...p.fxTable, [c]: v } }));
  const cur = inp.currency;
  const inCls = "w-full rounded-md ring-1 ring-slate-200 px-2 py-1 text-sm text-right tabular-nums focus:ring-2 focus:ring-emerald-400 focus:outline-none";
  const vBadge = res.verdict === "GO" ? "bg-emerald-100 text-emerald-700" : res.verdict === "재검토" ? "bg-amber-100 text-amber-700" : "bg-rose-100 text-rose-700";
  const pct = (x) => `${(x * 100).toFixed(2)}%`;
  const money = (x) => `${cur} ${fmt(Math.round(x))}`;
  const years = ["1년차", "2년차", "3년차", "4년차", "5년차"];

  return (
    <div className="space-y-5">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div className="text-sm text-slate-500">
          {editing ? <>편집 중: <span className="font-semibold text-slate-700">{editing.name}</span></> : <span className="text-slate-400">새 분석 <span className="text-[11px]">(미저장)</span></span>}
        </div>
        <div className="flex items-center gap-2">
          {editing && <button onClick={onNew} className="text-xs rounded-lg ring-1 ring-slate-300 px-3 py-1.5 hover:bg-slate-50">새 분석 시작</button>}
          <button onClick={() => { setSName(editing?.name || ""); setSMemo(editing?.memo || ""); setSaveOpen(true); }} className="text-sm rounded-lg bg-emerald-600 text-white px-3 py-1.5 font-semibold hover:bg-emerald-700">{editing ? "재저장" : "리포트로 저장"}</button>
        </div>
      </div>
      <Card className="p-5">
        <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
          <div className="flex items-center gap-3">
            <span className={`px-3 py-1 rounded-full text-sm font-bold ${vBadge}`}>{res.verdict === "GO" ? "GO ✓" : res.verdict === "재검토" ? "재검토 △" : "NO-GO ✗"}</span>
            <span className="text-xs text-slate-400 hidden sm:inline">NPV&gt;0 · IRR≥WACC · 회수≤계약기간</span>
          </div>
          <button onClick={() => setInp(BIZ_DEFAULTS)} className="text-xs text-slate-400 underline hover:text-slate-600">기본값으로 초기화</button>
        </div>
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
          <BizStat label="NPV" value={money(res.NPV)} tone={res.NPV > 0 ? "emerald" : "rose"} />
          <BizStat label="IRR" value={res.IRR == null ? "—" : pct(res.IRR)} sub={`WACC ${pct(res.WACC)}`} tone={res.IRR != null && res.IRR >= res.WACC ? "emerald" : "rose"} />
          <BizStat label="투자회수기간" value={res.PB == null ? "회수 불가" : `${res.PB.toFixed(1)}개월`} sub={`계약 ${res.maxTerm}개월`} tone={res.PB != null && res.PB <= res.maxTerm ? "emerald" : "amber"} />
          <BizStat label="5년 순이익률" value={pct(res.npm)} tone={res.npm > 0.15 ? "emerald" : res.npm > 0 ? "amber" : "rose"} />
        </div>
      </Card>

      <BizSectionLabel>입력</BizSectionLabel>

      <Card className="p-4">
        <h3 className="text-sm font-bold">분석 통화 · 환율</h3>
        <p className="text-[11px] text-slate-400 mb-3">재료비·가공비(USD)를 이 통화로 환산합니다</p>
        <div className="flex flex-wrap items-end gap-3">
          <label className="text-xs text-slate-500">분석 통화
            <select value={cur} onChange={(e) => setF("currency", e.target.value)} className="mt-1 block rounded-md ring-1 ring-slate-200 px-2 py-1 text-sm focus:ring-2 focus:ring-emerald-400 focus:outline-none">
              {Object.keys(inp.fxTable).map((c) => <option key={c}>{c}</option>)}
            </select>
          </label>
          <label className="text-xs text-slate-500">1 USD = {cur}
            <input value={inp.fxTable[cur] ?? ""} onChange={(e) => setFx(cur, e.target.value)} className={`mt-1 ${inCls} w-28`} />
          </label>
        </div>
      </Card>

      <Card className="overflow-hidden">
        <div className="px-4 py-3 border-b border-slate-100">
          <h3 className="text-sm font-bold">모델별 원가</h3>
          <p className="text-[11px] text-slate-400">월판가를 비우면 원가·마진으로 자동 계산(PMT)합니다</p>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-sm whitespace-nowrap">
            <thead className="bg-slate-50 text-slate-500 text-xs">
              <tr>
                <th className="px-2 py-2 text-left">모델</th><th className="px-2 py-2">구분</th>
                <th className="px-2 py-2">재료비$</th><th className="px-2 py-2">가공비$</th><th className="px-2 py-2">마진율</th>
                <th className="px-2 py-2">기간(월)</th><th className="px-2 py-2">매장수</th><th className="px-2 py-2">매장당</th><th className="px-2 py-2">월판가$(manual)</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {inp.models.map((m, i) => {
                const hw = m.kind === "HW";
                const dim = hw ? "" : "bg-slate-50 text-slate-300";
                return (
                  <tr key={i}>
                    <td className="px-2 py-1.5"><input value={m.name} onChange={(e) => setModel(i, "name", e.target.value)} className={`${inCls} w-20 text-left`} /></td>
                    <td className="px-2 py-1.5 text-center"><select value={m.kind} onChange={(e) => setModel(i, "kind", e.target.value)} className="rounded-md ring-1 ring-slate-200 px-1 py-1 text-xs focus:outline-none"><option>HW</option><option>SW</option></select></td>
                    <td className="px-2 py-1.5"><input disabled={!hw} value={hw ? m.mat : ""} onChange={(e) => setModel(i, "mat", e.target.value)} className={`${inCls} w-20 ${dim}`} /></td>
                    <td className="px-2 py-1.5"><input disabled={!hw} value={hw ? m.proc : ""} onChange={(e) => setModel(i, "proc", e.target.value)} className={`${inCls} w-20 ${dim}`} /></td>
                    <td className="px-2 py-1.5">{hw ? <PctInput value={m.margin} onChange={(v) => setModel(i, "margin", v)} w="w-16" /> : <span className="text-slate-300 text-xs">—</span>}</td>
                    <td className="px-2 py-1.5"><input value={m.term} onChange={(e) => setModel(i, "term", e.target.value)} className={`${inCls} w-16`} /></td>
                    <td className="px-2 py-1.5"><input value={m.stores} onChange={(e) => setModel(i, "stores", e.target.value)} className={`${inCls} w-20`} /></td>
                    <td className="px-2 py-1.5"><input value={m.perStore} onChange={(e) => setModel(i, "perStore", e.target.value)} className={`${inCls} w-20`} /></td>
                    <td className="px-2 py-1.5"><input value={m.manualK} onChange={(e) => setModel(i, "manualK", e.target.value)} className={`${inCls} w-24`} /></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>

      <Card className="overflow-hidden">
        <button onClick={() => setShowAdv((v) => !v)} className="w-full flex items-center justify-between px-4 py-3 hover:bg-slate-50 text-left">
          <div>
            <h3 className="text-sm font-bold">고급 가정 <span className="text-[11px] text-slate-400 font-normal">WACC · 비용비율 · ECL</span></h3>
            <p className="text-[11px] text-slate-400">잘 모르면 그대로 두세요 · 현재 WACC(할인율) {pct(res.WACC)}</p>
          </div>
          <ChevronRight size={18} className={`text-slate-400 transition-transform shrink-0 ${showAdv ? "rotate-90" : ""}`} />
        </button>
        {showAdv && (
          <div className="border-t border-slate-100 p-4 grid lg:grid-cols-3 gap-x-6 gap-y-5">
            <div>
              <h4 className="text-xs font-bold text-slate-600 mb-2">판가 기반 비용비율 <span className="text-[10px] text-slate-400 font-normal">판가에서 차감</span></h4>
              <div className="space-y-1.5">
                <BizField label="물류비" desc="운송·보관"><PctInput value={inp.costRatios.logistics} onChange={(v) => setCR("logistics", v)} /></BizField>
                <BizField label="금융비용" desc="조달 이자 부담"><PctInput value={inp.costRatios.finance} onChange={(v) => setCR("finance", v)} /></BizField>
                <BizField label="대손충당" desc="미회수 대비"><PctInput value={inp.costRatios.baddebt} onChange={(v) => setCR("baddebt", v)} /></BizField>
                <BizField label="핸들링비" desc="입출고·설치"><PctInput value={inp.costRatios.handling} onChange={(v) => setCR("handling", v)} /></BizField>
                <BizField label="유지보수료" desc="A/S 비용"><PctInput value={inp.costRatios.maintenance} onChange={(v) => setCR("maintenance", v)} /></BizField>
                <div className="flex items-center justify-between pt-2 border-t border-slate-100 text-xs"><span className="text-slate-400">합계</span><span className="font-semibold tabular-nums">{pct(res.CRS)}</span></div>
              </div>
            </div>
            <div>
              <h4 className="text-xs font-bold text-slate-600 mb-2">WACC <span className="text-[10px] text-slate-400 font-normal">자본조달 비용 = 할인율</span></h4>
              <div className="space-y-1.5">
                <BizField label="무위험수익률 Rf" desc="국고채 금리"><PctInput value={inp.rf} onChange={(v) => setF("rf", v)} /></BizField>
                <BizField label="베타 β" desc="시장 대비 민감도(배수)"><NumIn value={inp.beta} onChange={(v) => setF("beta", v)} w="w-20" /></BizField>
                <BizField label="시장위험프리미엄" desc="주식 초과수익(ERP)"><PctInput value={inp.erp} onChange={(v) => setF("erp", v)} /></BizField>
                <BizField label="차입이자율 Kd" desc="대출 금리"><PctInput value={inp.kd} onChange={(v) => setF("kd", v)} /></BizField>
                <BizField label="법인세율" desc="세율"><PctInput value={inp.tax} onChange={(v) => setF("tax", v)} /></BizField>
                <BizField label="자기자본 비중" desc="자본÷(자본+부채)"><PctInput value={inp.eqW} onChange={(v) => setF("eqW", v)} /></BizField>
                <BizField label="할부 할인율 IBR" desc="고객 차입 금리"><PctInput value={inp.ibr} onChange={(v) => setF("ibr", v)} /></BizField>
                <div className="flex items-center justify-between pt-2 border-t border-slate-100 text-xs"><span className="text-slate-400">WACC</span><span className="font-semibold tabular-nums text-emerald-700">{pct(res.WACC)}</span></div>
              </div>
            </div>
            <div>
              <h4 className="text-xs font-bold text-slate-600 mb-2">ECL · 사업 가정 <span className="text-[10px] text-slate-400 font-normal">대손 예상</span></h4>
              <div className="space-y-1.5">
                <BizField label="부도확률 PD" desc="1년 내 부도 가능성"><PctInput value={inp.pd} onChange={(v) => setF("pd", v)} /></BizField>
                <BizField label="부도시 손실률 LGD" desc="회수 못 하는 비율"><PctInput value={inp.lgd} onChange={(v) => setF("lgd", v)} /></BizField>
                <BizField label={`설치원가 (${cur})`} desc="총액, 1회성"><NumIn value={inp.install} onChange={(v) => setF("install", v)} /></BizField>
                <BizField label={`연간 유지보수 (${cur})`} desc="연 단위 비용"><NumIn value={inp.maint} onChange={(v) => setF("maint", v)} /></BizField>
                <div className="flex items-center justify-between pt-2 border-t border-slate-100 text-xs"><span className="text-slate-400">예상 대손(Day1 ECL)</span><span className="font-semibold tabular-nums text-rose-600">{money(res.ECL)}</span></div>
              </div>
            </div>
          </div>
        )}
      </Card>

      <BizSectionLabel>결과 상세</BizSectionLabel>

      <Card className="overflow-hidden">
        <div className="px-4 py-3 border-b border-slate-100"><h3 className="text-sm font-bold">모델별 단가 · 마진 (계산 결과)</h3></div>
        <div className="overflow-x-auto">
          <table className="w-full text-sm whitespace-nowrap">
            <thead className="bg-slate-50 text-slate-500 text-xs">
              <tr><th className="px-3 py-2 text-left">모델</th><th className="px-3 py-2 text-right">총수량</th><th className="px-3 py-2 text-right">기준판가$</th><th className="px-3 py-2 text-right">월판가$</th><th className="px-3 py-2 text-right">실효마진</th><th className="px-3 py-2 text-left">단가 근거</th></tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {res.rows.map((r, i) => {
                const hasK = r.manualK !== "" && r.manualK != null;
                const eff = !r.isHW ? null : (hasK && bnum(r.manualK) * bnum(r.term) ? (bnum(r.manualK) * bnum(r.term) - (bnum(r.mat) + bnum(r.proc)) - bnum(r.manualK) * bnum(r.term) * res.CRS) / (bnum(r.manualK) * bnum(r.term)) : bnum(r.margin));
                return (
                  <tr key={i}>
                    <td className="px-3 py-2 font-medium">{r.name}<span className="ml-1 text-[10px] text-slate-400">{r.kind}</span></td>
                    <td className="px-3 py-2 text-right tabular-nums">{fmt(r.I)}</td>
                    <td className="px-3 py-2 text-right tabular-nums text-slate-500">{r.J == null ? "—" : r.J.toFixed(4)}</td>
                    <td className="px-3 py-2 text-right tabular-nums font-semibold">{r.L.toFixed(4)}</td>
                    <td className="px-3 py-2 text-right tabular-nums">{eff == null ? "—" : pct(eff)}</td>
                    <td className="px-3 py-2 text-left text-[11px] text-slate-400">{!r.isHW ? "SW 단가" : hasK ? "월판가 manual" : "목표마진 PMT"}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>

      <Card className="overflow-hidden">
        <div className="px-4 py-3 border-b border-slate-100"><h3 className="text-sm font-bold">5개년 손익 ({cur})</h3></div>
        <div className="overflow-x-auto">
          <table className="w-full text-sm whitespace-nowrap">
            <thead className="bg-slate-50 text-slate-500 text-xs text-right">
              <tr><th className="px-3 py-2 text-left">항목</th>{years.map((y) => <th key={y} className="px-3 py-2">{y}</th>)}<th className="px-3 py-2">누계</th></tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-right tabular-nums">
              <BizRow label="매출(구독료)" arr={res.rev} sum={res.revSum} tone="text-emerald-600" />
              <BizRow label="감가상각비" arr={res.dep} sum={res.depSum} tone="text-rose-500" neg />
              <BizRow label="매출총이익" arr={res.gp} sum={res.gpSum} bold />
              <BizRow label="법인세" arr={res.tax} sum={res.tax.reduce((a, b) => a + b, 0)} tone="text-rose-500" neg />
              <BizRow label="순이익" arr={res.ni} sum={res.niSum} bold tone="text-indigo-700" />
            </tbody>
            <tfoot>
              <tr className="border-t-2 border-slate-200 bg-slate-50/50 text-xs text-slate-500 text-right">
                <td className="px-3 py-2 text-left font-medium">순이익률</td>
                {res.rev.map((rv, i) => <td key={i} className="px-3 py-2 tabular-nums">{rv ? pct(res.ni[i] / rv) : "—"}</td>)}
                <td className="px-3 py-2 tabular-nums font-semibold">{pct(res.npm)}</td>
              </tr>
            </tfoot>
          </table>
        </div>
        <div className="px-4 py-2.5 text-[11px] text-slate-400 border-t border-slate-100 flex flex-wrap gap-x-4 gap-y-1">
          <span>초기투자 {money(res.init)}</span><span>NPV/투자 {res.npvRatio.toFixed(3)}</span><span>매출총이익률 {pct(res.gpm)}</span>
        </div>
      </Card>

      <p className="text-[11px] text-slate-400 px-1">※ 엑셀 모델 v4와 동일 로직(검증 완료). Day1 ECL은 표준식(익스포저×PD×LGD)으로 계산합니다. 실적 비교는 다음 단계에서 추가됩니다.</p>

      {saveOpen && (
        <div className="fixed inset-0 z-50 bg-black/30 grid place-items-center px-4" onClick={() => setSaveOpen(false)}>
          <div className="bg-white rounded-2xl p-5 w-full max-w-sm" onClick={(e) => e.stopPropagation()}>
            <h3 className="font-bold mb-3">{editing ? "리포트 재저장" : "리포트로 저장"}</h3>
            <label className="text-xs text-slate-500 block">리포트 이름 *
              <input value={sName} onChange={(e) => setSName(e.target.value)} placeholder="예: A고객 2.8인치 5천대 제안" autoFocus className="mt-1 w-full rounded-md ring-1 ring-slate-200 px-2 py-1.5 text-sm focus:ring-2 focus:ring-emerald-400 focus:outline-none" />
            </label>
            <label className="text-xs text-slate-500 block mt-3">메모 (선택)
              <textarea value={sMemo} onChange={(e) => setSMemo(e.target.value)} rows={2} className="mt-1 w-full rounded-md ring-1 ring-slate-200 px-2 py-1.5 text-sm focus:ring-2 focus:ring-emerald-400 focus:outline-none resize-none" />
            </label>
            <div className="flex justify-end gap-2 mt-4">
              <button onClick={() => setSaveOpen(false)} className="text-sm rounded-lg ring-1 ring-slate-300 px-3 py-1.5 hover:bg-slate-50">취소</button>
              <button disabled={!sName.trim()} onClick={() => { onSave(sName.trim(), sMemo.trim()); setSaveOpen(false); }} className="text-sm rounded-lg bg-emerald-600 text-white px-3 py-1.5 font-semibold hover:bg-emerald-700 disabled:opacity-40">{editing ? "재저장" : "저장"}</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function PctInput({ value, onChange, w = "w-24" }) {
  const toText = (v) => { const n = Number(v); return isFinite(n) ? String(+(n * 100).toFixed(4)) : ""; };
  const ref = useRef(value);
  const [t, setT] = useState(toText(value));
  useEffect(() => { if (value !== ref.current) { setT(toText(value)); ref.current = value; } }, [value]);
  const commit = (s) => { setT(s); const n = parseFloat(s); const dec = isFinite(n) ? n / 100 : 0; ref.current = dec; onChange(dec); };
  return (
    <div className="relative shrink-0">
      <input value={t} inputMode="decimal" onChange={(e) => commit(e.target.value)}
        className={`rounded-md ring-1 ring-slate-200 pl-2 pr-5 py-1 text-sm text-right tabular-nums focus:ring-2 focus:ring-emerald-400 focus:outline-none ${w}`} />
      <span className="absolute right-1.5 top-1/2 -translate-y-1/2 text-[11px] text-slate-400 pointer-events-none">%</span>
    </div>
  );
}
function NumIn({ value, onChange, w = "w-24", disabled, align = "text-right" }) {
  return <input disabled={disabled} value={value} onChange={(e) => onChange(e.target.value)}
    className={`rounded-md ring-1 ring-slate-200 px-2 py-1 text-sm ${align} tabular-nums focus:ring-2 focus:ring-emerald-400 focus:outline-none ${w} ${disabled ? "bg-slate-50 text-slate-300" : ""}`} />;
}
function BizField({ label, desc, children }) {
  return (
    <div className="flex items-center justify-between gap-3 py-0.5">
      <div className="min-w-0">
        <div className="text-xs text-slate-700">{label}</div>
        {desc && <div className="text-[10px] text-slate-400 leading-tight mt-0.5">{desc}</div>}
      </div>
      {children}
    </div>
  );
}
function BizSectionLabel({ children }) {
  return (
    <div className="flex items-center gap-2 pt-1">
      <span className="text-xs font-bold text-slate-400 tracking-wide">{children}</span>
      <div className="flex-1 h-px bg-slate-200" />
    </div>
  );
}

function BizStat({ label, value, sub, tone = "slate" }) {
  const c = { emerald: "text-emerald-600", rose: "text-rose-600", amber: "text-amber-600", slate: "text-slate-800" }[tone];
  return (
    <div className="rounded-xl ring-1 ring-slate-200 p-3">
      <div className="text-[11px] text-slate-400 font-medium">{label}</div>
      <div className={`text-lg font-bold tabular-nums mt-0.5 ${c}`}>{value}</div>
      {sub && <div className="text-[11px] text-slate-400 mt-0.5">{sub}</div>}
    </div>
  );
}
function BizRow({ label, arr, sum, tone = "", bold, neg }) {
  return (
    <tr>
      <td className={`px-3 py-2 text-left ${bold ? "font-semibold" : ""}`}>{label}</td>
      {arr.map((v, i) => <td key={i} className={`px-3 py-2 ${tone} ${bold ? "font-semibold" : ""}`}>{neg && Math.round(v) ? "−" : ""}{fmt(Math.round(Math.abs(v)))}</td>)}
      <td className={`px-3 py-2 ${tone} font-semibold`}>{neg && Math.round(sum) ? "−" : ""}{fmt(Math.round(Math.abs(sum)))}</td>
    </tr>
  );
}

function LoginScreen({ flash }) {
  const [email, setEmail] = useState("");
  const [pw, setPw] = useState("");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const login = async () => {
    if (!email || !pw) { setErr("이메일과 비밀번호를 입력하세요"); return; }
    setBusy(true); setErr("");
    try {
      const { error } = await supabase.auth.signInWithPassword({ email: email.trim(), password: pw });
      if (error) setErr("로그인 실패 — 이메일 또는 비밀번호를 확인하세요");
    } catch (e) { setErr("로그인 중 오류가 발생했습니다. 네트워크를 확인해 주세요."); }
    setBusy(false);
  };
  return (
    <div className="min-h-screen grid place-items-center bg-slate-50 px-4" style={{ fontFamily: "system-ui, sans-serif" }}>
      <div className="w-full max-w-sm">
        <div className="flex items-center gap-2.5 justify-center mb-6">
          <div className="grid place-items-center h-9 w-9 rounded-lg bg-indigo-600 text-white"><Package size={20} /></div>
          <div><h1 className="text-base font-bold leading-none">운용리스 관리</h1><p className="text-[11px] text-slate-400 mt-1">리스제공자 콘솔</p></div>
        </div>
        <div className="bg-white rounded-2xl shadow-sm ring-1 ring-slate-200 p-5 space-y-3">
          <div>
            <label className="block text-xs text-slate-500 mb-1">이메일</label>
            <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} onKeyDown={(e) => e.key === "Enter" && login()}
              className="w-full rounded-lg ring-1 ring-slate-300 px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 outline-none" placeholder="you@company.com" />
          </div>
          <div>
            <label className="block text-xs text-slate-500 mb-1">비밀번호</label>
            <input type="password" value={pw} onChange={(e) => setPw(e.target.value)} onKeyDown={(e) => e.key === "Enter" && login()}
              className="w-full rounded-lg ring-1 ring-slate-300 px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 outline-none" placeholder="••••••••" />
          </div>
          {err && <p className="text-xs text-rose-600">{err}</p>}
          <button onClick={login} disabled={busy}
            className="w-full rounded-lg bg-indigo-600 text-white py-2.5 text-sm font-semibold hover:bg-indigo-700 disabled:opacity-50 transition">
            {busy ? "로그인 중…" : "로그인"}
          </button>
          <p className="text-[11px] text-slate-400 text-center">계정은 관리자(Supabase 콘솔)에서 발급됩니다.</p>
        </div>
      </div>
    </div>
  );
}
