import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { createClient } from '@supabase/supabase-js';

const supabaseUrl = 'https://ehvtgvljtsfgfnaouhmd.supabase.co';
const supabaseKey = 'sb_publishable_2wHnHf6T-A_PdcLLZoCF-w_8CAicB81';
const supabase = createClient(supabaseUrl, supabaseKey);

// ── Helpers ──────────────────────────────────────────────────────────────────

const clean = (s) => String(s || "").trim();
const toStrict = (s) => clean(s).replace(/[^a-z0-9]/gi, '').toLowerCase();

const daysDiff = (d1, d2) => {
  if (!d1 || !d2) return null;
  const t1 = toDate(d1), t2 = toDate(d2);
  if (!t1 || !t2) return null;
  return Math.round((t1 - t2) / (1000 * 60 * 60 * 24));
};

const toDate = (val) => {
  if (!val) return null;
  if (val instanceof Date) return val;
  if (typeof val === 'number') return new Date(Math.round((val - 25569) * 86400 * 1000));
  const s = String(val).trim();
  const parts = s.split(/[-./]/);
  if (parts.length === 3) {
    if (parts[2].length === 4) return new Date(+parts[2], +parts[1] - 1, +parts[0]);
    if (parts[0].length === 4) return new Date(+parts[0], +parts[1] - 1, +parts[2]);
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
};

const fmtDate = (val) => {
  const d = toDate(val);
  return d ? d.toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' }) : '-';
};

const fmtNum = (val, decimals = 0) => {
  const n = Number(val);
  if (isNaN(n)) return '-';
  return n.toLocaleString('en-IN', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
};

const getV = (r, keys) => {
  if (!r) return '';
  for (const k of keys) { if (r[k] !== undefined && r[k] !== null && r[k] !== '') return r[k]; }
  const nKeys = keys.map(k => k.toLowerCase().replace(/[^a-z0-9]/g, ''));
  for (const rk in r) {
    const rnk = rk.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (nKeys.includes(rnk) && r[rk] !== undefined && r[rk] !== null && r[rk] !== '') return r[rk];
  }
  for (const k of keys) {
    const lk = k.toLowerCase();
    for (const rk in r) {
      const rlk = rk.toLowerCase();
      if ((rlk.includes(lk) || lk.includes(rlk)) && r[rk] !== undefined && r[rk] !== null && r[rk] !== '') {
        if (lk.includes('date') && !rlk.includes('date')) continue;
        return r[rk];
      }
    }
  }
  return '';
};

const isNumericCol = (key) => {
  const kl = String(key).toLowerCase();
  if (kl.includes(' vs ')) return true;
  if (kl.includes('date') || kl.includes('no') || kl.includes('party') || kl.includes('name') || 
      kl.includes('status') || kl.includes('action') || kl.includes('category') || kl.includes('due') || 
      kl.includes('on') || kl.includes('so') || kl.includes('details') || kl.includes('po') || kl.includes('vpo')) {
    if (!kl.includes('qty') && !kl.includes('balance') && !kl.includes('ordered') && !kl.includes('value') && !kl.includes('open')) return false;
  }
  return kl.includes('rate') || kl.includes('value') || kl.includes('qty') ||
    kl.includes('quantity') || kl.includes('balance') || kl.includes('amount') ||
    kl.includes('price') || kl.includes('ordered') || kl.includes('stock') ||
    kl.includes('allocated') || kl.includes('discount') || kl.includes('open');
};

const isDateCol = (key) => {
  const kl = String(key).toLowerCase();
  if (kl.includes(' vs ')) return false;
  return kl.includes('date') || kl.includes('due') || kl.includes('on') || kl.includes('mad') || kl.includes('requested');
};

const fmtCell = (key, val) => {
  if (val === undefined || val === null || val === '') return '-';
  if (typeof val === 'object' && !(val instanceof Date) && !React.isValidElement(val)) return JSON.stringify(val);
  const kl = String(key).toLowerCase();
  if (kl.includes(' vs ') && val !== null && val !== '') {
    const n = Number(val);
    if (!isNaN(n)) return n + " Days";
  }
  if (isDateCol(key)) { const d = toDate(val); if (d) return fmtDate(d); }
  if (key === 'Category') {
    const isDue = String(val) === 'Due' || String(val) === 'Due Order';
    return <span style={{ background: isDue ? '#fdf2f8' : '#f0fdf4', color: isDue ? '#9d174d' : '#15803d', padding: '1px 5px', borderRadius: '3px', fontWeight: 'bold', fontSize: '9px', border: `1px solid ${isDue ? '#fbcfe8' : '#bbf7d0'}`, display: 'inline-block' }}>{String(val)}</span>;
  }
  if (key === 'Status') return <strong style={{ color: String(val) === 'Available' ? '#15803d' : '#b91c1c' }}>{String(val)}</strong>;
  if (key === 'Action') {
    const col = String(val) === 'Covered' ? '#15803d' : (String(val).includes('Partial') ? '#9a3412' : '#b91c1c');
    return <strong style={{ color: col }}>{String(val)}</strong>;
  }
  if ((kl.includes('rate') || kl.includes('value') || kl.includes('amount') || kl.includes('price') || kl.includes('discount')) && !isNaN(Number(val)) && val !== '') return fmtNum(val, 2);
  return String(val);
};

// ── Cloud Persistence ───────────────────────────────────────────────────────

const TABLE_MAP = { stock: 'stock', so: 'sales_orders', po: 'purchase_orders', oo: 'vendor_orders' };

const saveData = async (type, list, headers = []) => {
  if (!type || !TABLE_MAP[type]) return;
  try {
    await supabase.from(TABLE_MAP[type]).delete().neq('id', -1); 
    if (list && list.length > 0) {
      const payload = list.map(r => ({ data: r, headers }));
      const chunkSize = 500;
      for (let i = 0; i < payload.length; i += chunkSize) {
        await supabase.from(TABLE_MAP[type]).insert(payload.slice(i, i + chunkSize));
      }
    }
  } catch (e) { console.error(`Cloud Save failed for ${type}`, e); }
};

const fetchFull = async (table) => {
  let allData = []; let from = 0; const step = 1000; let headers = [];
  while (true) {
    const { data, error } = await supabase.from(table).select('data, headers').range(from, from + step - 1);
    if (error) throw error;
    if (!data || data.length === 0) break;
    if (data[0].headers) headers = data[0].headers;
    allData = allData.concat(data.map(r => r.data));
    if (data.length < step) break;
    from += step;
  }
  return { list: allData, headers };
};

const loadData = async () => {
  try {
    const [stock, so, po, oo] = await Promise.all([fetchFull('stock'), fetchFull('sales_orders'), fetchFull('purchase_orders'), fetchFull('vendor_orders')]);
    return { stock: stock.list, so: so.list, po: po.list, oo: oo.list, headers: { stock: stock.headers, so: so.headers, po: po.headers, oo: oo.headers } };
  } catch (e) { console.error('Cloud Load failed', e); return { stock: [], so: [], po: [], oo: [], headers: {} }; }
};

// ── Logic: Build Master Report ──────────────────────────────────────────────

const buildMasterReport = (data) => {
  if (!data || !data.so || !data.so.length) return [];
  const TODAY = new Date(); TODAY.setHours(0, 0, 0, 0);
  const DUE_CUTOFF = new Date(TODAY); DUE_CUTOFF.setDate(DUE_CUTOFF.getDate() + 30);

  const stockMap = {};
  data.stock.forEach(s => {
    const desc = clean(getV(s, ['Description', 'Name of Item', 'Item Name', 'Material'])).toLowerCase();
    const part = toStrict(getV(s, ['Part No', 'Material', 'Material Code', 'Code', 'Part Number']));
    const qty = Number(getV(s, ['Quantity', 'Qty', 'Closing Stock'])) || 0;
    const entry = { total: qty, remaining: qty };
    if (desc) stockMap[desc] = entry;
    if (part) stockMap[part] = entry;
  });

  const poGroups = {};
  data.po.forEach(p => {
    const desc = clean(getV(p, ['Name of Item', 'Item Name', 'Description', 'Material'])).toLowerCase();
    const part = toStrict(getV(p, ['Part No', 'Material', 'Material Code', 'Code', 'Part Number']));
    const entry = {
      date: getV(p, ['Date', 'PO Date', 'Order Date', 'Document Date', 'Created On']),
      order: getV(p, ['Order No', 'Order', 'PO No', 'Document No', 'PO Number', 'Reference']),
      party: getV(p, ['Party Name', "Party's Name", 'Vendor', 'Supplier', 'Name']),
      ordered: Number(getV(p, ['Ordered', 'Quantity', 'Order Qty', 'Qty'])) || 0,
      balance: Number(getV(p, ['Balance', 'Pending Qty', 'Open Qty', 'Remaining'])) || 0,
      remaining: Number(getV(p, ['Balance', 'Pending Qty', 'Open Qty', 'Remaining'])) || 0,
      dueOn: toDate(getV(p, ['Due on', 'Delivery Date', 'Est Date', 'Due Date', 'Promised Date']))
    };
    if (desc) { if (!poGroups[desc]) poGroups[desc] = []; poGroups[desc].push(entry); }
    if (part && part !== desc) { if (!poGroups[part]) poGroups[part] = []; poGroups[part].push(entry); }
  });
  Object.values(poGroups).forEach(g => g.sort((a, b) => (a.dueOn || 0) - (b.dueOn || 0)));

  const ooGroups = {};
  data.oo.forEach(o => {
    const desc = clean(getV(o, ['Description', 'Material', 'Name of Item', 'Item Name', 'Text'])).toLowerCase();
    const part = toStrict(getV(o, ['Material No', 'Part No', 'Material', 'Material Code', 'Code', 'Part Number']));
    const entry = {
      soNo: clean(getV(o, ['Sales Order No', 'SO No', 'Order No', 'Document No', 'Sales Order Number', 'Sales Order', 'SO. No'])),
      soDate: getV(o, ['SO Created Date', 'Order Date', 'Date', 'Sales Order Date', 'SO. Date']),
      lineDate: getV(o, ['Line Item Create Date', 'Line Date', 'Item Date', 'Line Created']),
      line: getV(o, ['SO Line', 'Line', 'Item No', 'Position', 'Item', 'Row']),
      custPO: getV(o, ['Cust PO No', 'Customer PO No', 'PO No', 'Ref', 'Customer Ref', 'Reference', 'Customer PO Number']),
      custPODate: getV(o, ['Cust PO Date', 'Customer PO Date', 'PO Date', 'Ref Date']),
      importBy: getV(o, ['Import By', 'Imported By', 'Mode', 'Shipment Mode']),
      ordered: Number(getV(o, ['Order Qty', 'Ordered', 'Quantity', 'Qty', 'Ordered Quantity'])) || 0,
      open: Number(getV(o, ['Open Qty', 'Balance', 'Pending', 'Open Quantity', 'Outstanding'])) || 0,
      remaining: Number(getV(o, ['Open Qty', 'Balance', 'Pending', 'Open Quantity', 'Outstanding'])) || 0,
      reqDate: getV(o, ['Cust Req Date', 'Customer Requested Date', 'Requested Date', 'Request Date', 'Req Date', 'Requested Delivery Date']),
      mad: getV(o, ['Estimated M.A.D.', 'M.A.D.', 'Availability Date', 'MAD', 'Est Date', 'Estimated Availability', 'Sch. Date', 'Commit Date']),
      whName: getV(o, ['Warehouse Name', 'Warehouse', 'WH Name', 'Storage Location', 'Whse'])
    };
    if (desc) { if (!ooGroups[desc]) ooGroups[desc] = []; ooGroups[desc].push(entry); }
    if (part && part !== desc) { if (!ooGroups[part]) ooGroups[part] = []; ooGroups[part].push(entry); }
  });
  Object.values(ooGroups).forEach(g => g.sort((a, b) => (toDate(a.mad) || 0) - (toDate(b.mad) || 0)));

  const overdue = data.so.map(r => {
    const dueOn = toDate(getV(r, ['Due on', 'Delivery Date', 'MAD', 'M.A.D.', 'Due Date']));
    const balance = Number(getV(r, ['Balance', 'Pending Qty', 'Open Qty', 'Open Qty.'])) || 0;
    const name = clean(getV(r, ['Name of Item', 'Description', 'Material', 'Item Name'])).toLowerCase();
    const partNo = clean(getV(r, ['Part No', 'Material Code', 'Material', 'Code', 'Part Number']));
    return {
      _dueDate: dueOn, _balance: balance, _category: (dueOn && dueOn < DUE_CUTOFF) ? 'Due Order' : 'Schedule Order', _name: name, _partStrict: toStrict(partNo),
      date: getV(r, ['Date', 'Order Date', 'Document Date']), order: getV(r, ['Order No', 'Order', 'Sales Order']), partyName: getV(r, ['Party Name', "Party's Name", 'Customer', 'Name']), nameOfItem: clean(getV(r, ['Name of Item', 'Description'])), partNo, ordered: Number(getV(r, ['Ordered', 'Quantity', 'Qty'])) || 0, balance, value: Number(getV(r, ['Value', 'Amount', 'Net Value'])) || 0, dueOn, stockQty: 0, allocatedQty: 0, poAllocated: 0, vpoAllocated: 0, stockStatus: '', action: '',
      poDetails: { order: '', party: '', date: '', ordered: 0, balance: 0, dueOn: null }, vpoDetails: { soNo: '', soDate: '', lineDate: '', line: '', custPO: '', custPODate: '', importBy: '', ordered: 0, open: 0, reqDate: '', mad: '', whName: '' }
    };
  });
  overdue.forEach(r => {
    const n = r._name;
    const p = r._partStrict;
    const n1 = n.endsWith(' 1') ? n.slice(0, -2) : null;
    const p1 = p.endsWith('1') ? p.slice(0, -1) : null;

    const s = stockMap[n] || stockMap[p] || (n1 && stockMap[n1]) || (p1 && stockMap[p1]);
    r.stockQty = s ? s.total : 0;
    if (s && s.remaining > 0) { const take = Math.min(s.remaining, r._balance); s.remaining -= take; r.allocatedQty = take; }
    r.stockStatus = r.allocatedQty >= r._balance ? 'Available' : 'Need to Arrange';
    
    // Step 1: PO Allocation
    let remPO = r._balance - r.allocatedQty;
    if (remPO > 0) {
      const pos = poGroups[n] || poGroups[p] || (n1 && poGroups[n1]) || (p1 && poGroups[p1]);
      if (pos) { for (const p of pos) { if (p.remaining <= 0) continue; const take = Math.min(p.remaining, remPO); p.remaining -= take; r.poAllocated += take; if (!r.poDetails.order) r.poDetails = { ...p }; remPO -= take; if (remPO <= 0) break; } }
    }
    
    // Step 2: Vendor PO Allocation (ONLY IF NOT COVERED BY STOCK)
    let remV = r._balance - r.allocatedQty;
    const vpos = ooGroups[n] || ooGroups[p] || 
                 ooGroups[n + " 100"] || ooGroups[p + "100"] ||
                 (n1 && (ooGroups[n1] || ooGroups[n1 + " 100"])) ||
                 (p1 && (ooGroups[p1] || ooGroups[p1 + "100"]));
    if (remV > 0 && vpos) {
      for (const v of vpos) {
        if (v.remaining <= 0) continue;
        const take = Math.min(v.remaining, remV);
        v.remaining -= take;
        r.vpoAllocated += take;
        if (!r.vpoDetails.soNo) r.vpoDetails = { ...v };
        remV -= take;
        if (remV <= 0) break;
      }
    }
    
    // Step 3: Action Logic (Follow Priority Rule: Stock+PO OR Stock+VPO if PO nil)
    let effectiveCoverage = 0;
    if (r.poAllocated > 0) {
      effectiveCoverage = r.allocatedQty + r.poAllocated;
    } else {
      effectiveCoverage = r.allocatedQty + r.vpoAllocated;
    }
    
    if (r.balance <= 0 || effectiveCoverage >= (r.balance - 0.001)) r.action = 'Covered';
    else if (effectiveCoverage > 0) r.action = 'Partial Qty ordered need to make order';
    else r.action = 'Make PO need to raise';
  });
  return overdue;
};

// ── Components ───────────────────────────────────────────────────────────────

const DataTable = ({ list, headers = [], columnGroups = [], rowStyle = null, hideSearch = false, onExport = null, searchTerm = '' }) => {
  const [sortCol, setSortCol] = useState(null);
  const [sortDir, setSortDir] = useState('asc');
  const [filter, setFilter] = useState('');
  const [colFilters, setColFilters] = useState({});
  const [page, setPage] = useState(1);
  const pageSize = 100;
  useEffect(() => setPage(1), [filter, colFilters, sortCol, sortDir, searchTerm]);
  const cols = useMemo(() => { if (!list || !list.length) return []; const keys = headers.length > 0 ? headers : Object.keys(list[0]); return keys.filter(k => { const v = list[0][k]; return v === null || v === undefined || typeof v !== 'object' || v instanceof Date; }); }, [list, headers]);
  const filtered = useMemo(() => {
    if (!list || !list.length) return [];
    const finalSearch = (filter || searchTerm).toLowerCase();
    let res = list;
    if (finalSearch) res = res.filter(r => Object.values(r).some(v => String(v || '').toLowerCase().includes(finalSearch)));
    Object.keys(colFilters).forEach(c => { const fv = colFilters[c].toLowerCase(); if (fv) res = res.filter(r => String(r[c] || '').toLowerCase().includes(fv)); });
    if (sortCol) {
      res = [...res].sort((a, b) => {
        const av = a[sortCol], bv = b[sortCol];
        if (isNumericCol(sortCol)) return sortDir === 'asc' ? (Number(av) || 0) - (Number(bv) || 0) : (Number(bv) || 0) - (Number(av) || 0);
        if (isDateCol(sortCol)) {
          const da = toDate(av)?.getTime() || 0;
          const db = toDate(bv)?.getTime() || 0;
          return sortDir === 'asc' ? da - db : db - da;
        }
        return sortDir === 'asc' ? String(av || '').localeCompare(String(bv || '')) : String(bv || '').localeCompare(String(av || ''));
      });
    }
    return res;
  }, [list, filter, searchTerm, colFilters, sortCol, sortDir]);
  const totals = useMemo(() => { const t = {}; cols.forEach(c => { if (isNumericCol(c)) t[c] = filtered.reduce((s, r) => s + (Number(r[c]) || 0), 0); }); return t; }, [filtered, cols]);
  if (!list || !list.length) return <div className="empty">No data to display.</div>;
  return (
    <div>
      <div className="tbl-toolbar">
        {!hideSearch && <input className="tbl-search" placeholder="Search..." value={filter} onChange={e => setFilter(e.target.value)} />}
        <span className="tbl-count">{filtered.length} rows</span>
        {onExport && <button onClick={() => onExport(filtered)} className="tbl-btn-export">Export</button>}
      </div>
      <div className="table-container">
        <table>
          <thead>
            {columnGroups.length > 0 && <tr>{columnGroups.map((g, i) => <th key={i} colSpan={g.span} style={{ background: g.bg, color: '#fff', padding: '3px' }}>{g.label}</th>)}</tr>}
            <tr>{cols.map(c => <th key={c} onClick={() => { if (sortCol === c) setSortDir(d => d === 'asc' ? 'desc' : 'asc'); else { setSortCol(c); setSortDir('asc'); } }} style={{ cursor: 'pointer', textAlign: isNumericCol(c) ? 'right' : 'left' }}>
              <div style={{display:'flex', alignItems:'center', justifyContent: isNumericCol(c) ? 'flex-end' : 'flex-start', gap:'4px'}}>{c} {sortCol === c ? (sortDir === 'asc' ? '▲' : '▼') : ''}</div>
              {totals[c] !== undefined && <div style={{ fontSize: '8px', color: '#1e40af' }}>{c.toLowerCase().includes('value') ? fmtNum(totals[c], 2) : fmtNum(totals[c])}</div>}
              <input className="col-filter" placeholder="Filter..." value={colFilters[c] || ''} onClick={e => e.stopPropagation()} onChange={e => setColFilters(prev => ({ ...prev, [c]: e.target.value }))} />
            </th>)}</tr>
          </thead>
          <tbody>
            {filtered.slice((page - 1) * pageSize, page * pageSize).map((r, i) => <tr key={i} style={rowStyle ? rowStyle(r) : {}}>{cols.map(c => <td key={c} style={{ textAlign: isNumericCol(c) ? 'right' : 'left' }}>{fmtCell(c, r[c])}</td>)}</tr>)}
          </tbody>
        </table>
      </div>
      {filtered.length > pageSize && <div className="tbl-toolbar" style={{ marginTop: 5 }}><button disabled={page === 1} onClick={() => setPage(p => p - 1)} className="tbl-btn">Prev</button><span>Page {page} of {Math.ceil(filtered.length / pageSize)}</span><button disabled={page >= Math.ceil(filtered.length / pageSize)} onClick={() => setPage(p => p + 1)} className="tbl-btn">Next</button></div>}
    </div>
  );
};

const Dashboard = ({ data, searchTerm }) => {
  const rawRows = useMemo(() => buildMasterReport(data), [data]);
  const dashboardRows = useMemo(() => rawRows.filter(r => r.poDetails?.order || r.vpoDetails?.soNo), [rawRows]);
  const display = useMemo(() => dashboardRows.map(r => ({
    'Order No': r.order,
    'Party Name': r.partyName,
    'Name of Item': r.nameOfItem,
    'Order Qty': r.ordered,
    'Balance Qty': r.balance,
    'Value': r.value,
    'Date (SO)': fmtDate(r.date),
    'PO No': r.poDetails?.order || '-',
    'PO Date': fmtDate(r.poDetails?.date),
    'PO Vs SO': daysDiff(r.poDetails?.date, r.date),
    'Vendor SO No': r.vpoDetails?.soNo || '-',
    'Vendor SO Date': fmtDate(r.vpoDetails?.soDate),
    'SO Date vs PO Date': daysDiff(r.vpoDetails?.soDate, r.poDetails?.date),
    'VSO Vs Pending SO': daysDiff(r.vpoDetails?.soDate, r.date),
    'Vendor MAD': fmtDate(r.vpoDetails?.mad),
    'MAD vs VSO Date': daysDiff(r.vpoDetails?.mad, r.vpoDetails?.soDate),
    'MAD vs SO': daysDiff(r.vpoDetails?.mad, r.date)
  })), [dashboardRows]);
  return <div className="card"><div className="stat-pill" style={{marginBottom:10, display:'inline-block', borderColor:'#1e40af', color:'#1e40af'}}>Active Procurement (PO/VSO Exist): <strong>{dashboardRows.length}</strong></div><DataTable list={display} hideSearch={true} searchTerm={searchTerm} /></div>;
};

const MasterReport = ({ data, searchTerm }) => {
  const rawRows = useMemo(() => buildMasterReport(data), [data]);
  const stats = useMemo(() => {
    const t = { ordered: 0, balance: 0, value: 0, stk: 0, alc: 0, due: 0, sch: 0 };
    rawRows.forEach(r => { t.ordered += Number(r.ordered) || 0; t.balance += Number(r.balance) || 0; t.value += Number(r.value) || 0; t.stk += Number(r.stockQty) || 0; t.alc += Number(r.allocatedQty) || 0; if (r._category === 'Due Order') t.due++; else t.sch++; });
    return t;
  }, [rawRows]);
  const display = useMemo(() => rawRows.map(r => ({
    'Category': r._category, 'Date': fmtDate(r.date), 'Order No': r.order, 'Party Name': r.partyName, 'Name of Item': r.nameOfItem, 'Part No': r.partNo, 'Ordered': r.ordered, 'Balance': r.balance, 'Value': r.value, 'Due on': fmtDate(r.dueOn),
    'Stock Qty': r.stockQty, 'Allocated Qty': r.allocatedQty, 'Status': r.stockStatus,
    'PO Order': r.poDetails?.order || '-', 'PO Party': r.poDetails?.party || '-', 'PO Date': fmtDate(r.poDetails?.date), 'PO Ordered': r.poDetails?.ordered || 0, 'PO Balance': r.poDetails?.balance || 0, 'PO Due': fmtDate(r.poDetails?.dueOn),
    'VPO SO': r.vpoDetails?.soNo || '-', 'VPO SO Date': fmtDate(r.vpoDetails?.soDate), 'VPO Line Date': fmtDate(r.vpoDetails?.lineDate), 'VPO Line': r.vpoDetails?.line || '-', 'VPO CustPO': r.vpoDetails?.custPO || '-', 'VPO CustPODate': fmtDate(r.vpoDetails?.custPODate), 'VPO ImportBy': r.vpoDetails?.importBy || '-', 'VPO Ordered': r.vpoDetails?.ordered || 0, 'VPO Open': r.vpoDetails?.open || 0, 'VPO ReqDate': fmtDate(r.vpoDetails?.reqDate), 'VPO MAD': fmtDate(r.vpoDetails?.mad), 'VPO Warehouse': r.vpoDetails?.whName || '-',
    'Action': r.action
  })), [rawRows]);
  return (
    <div className="card">
      <div style={{ display: 'flex', gap: 10, marginBottom: 10, flexWrap: 'wrap' }}>
        <div className="stat-pill" style={{ borderColor: '#e11d48' }}>Due: <strong>{stats.due}</strong></div><div className="stat-pill" style={{ borderColor: '#10b981' }}>Scheduled: <strong>{stats.sch}</strong></div><div className="stat-pill" style={{ borderColor: '#2563eb' }}>Ordered: <strong>{fmtNum(stats.ordered)}</strong></div><div className="stat-pill" style={{ borderColor: '#f59e0b' }}>Balance: <strong>{fmtNum(stats.balance)}</strong></div><div className="stat-pill" style={{ borderColor: '#0ea5e9' }}>Value: <strong>{fmtNum(stats.value, 2)}</strong></div>
      </div>
      <DataTable list={display} hideSearch={true} searchTerm={searchTerm} columnGroups={[{ label: 'SO Details', span: 10, bg: '#1e40af' }, { label: 'Stock', span: 3, bg: '#0369a1' }, { label: 'PO', span: 6, bg: '#0891b2' }, { label: 'Vendor PO', span: 12, bg: '#0d9488' }, { label: 'Action', span: 1, bg: '#111827' }]} rowStyle={r => (r['Category'] === 'Due Order' ? { background: '#fff1f2' } : {})} />
    </div>
  );
};

// ── App ──────────────────────────────────────────────────────────────────────

const App = () => {
  const [data, setData] = useState({ stock: [], so: [], po: [], oo: [] });
  const [headers, setHeaders] = useState({ stock: [], so: [], po: [], oo: [] });
  const [activeTab, setActiveTab] = useState('dashboard');
  const [searchTerm, setSearchTerm] = useState('');
  const [showSidebar, setShowSidebar] = useState(true);
  const [loading, setLoading] = useState(true);
  useEffect(() => { loadData().then(saved => { if (saved) { setData(saved); setHeaders(saved.headers); } setLoading(false); }); }, []);
  const handleUpload = (e, type) => {
    const file = e.target.files[0]; if (!file) return; setLoading(true);
    const reader = new FileReader(); reader.onload = (evt) => { 
      const wb = XLSX.read(evt.target.result, { type: 'binary', cellDates: true }); 
      const sheet = wb.Sheets[wb.SheetNames[0]]; const raw = XLSX.utils.sheet_to_json(sheet, { defval: "" }); 
      const range = XLSX.utils.decode_range(sheet['!ref']); const h = []; for (let C = range.s.c; C <= range.e.c; ++C) { const cell = sheet[XLSX.utils.encode_cell({ r: range.s.r, c: C })]; if (cell) h.push(clean(cell.v)); }
      setData(prev => ({ ...prev, [type]: raw })); setHeaders(prev => ({ ...prev, [type]: h })); saveData(type, raw, h); setLoading(false); 
    }; reader.readAsBinaryString(file);
  };
  const TAB_LABELS = { dashboard: 'DASHBOARD', master: 'MASTER REPORT', stock: 'STOCK', so: 'PENDING SO', po: 'PENDING PO', oo: 'VENDOR OPEN PO' };
  const masterCount = useMemo(() => buildMasterReport(data).length, [data]);
  const counts = {
    dashboard: buildMasterReport(data).filter(r => r.poDetails?.order || r.vpoDetails?.soNo).length,
    master: masterCount,
    stock: data.stock?.length || 0,
    so: data.so?.length || 0,
    po: data.po?.length || 0,
    oo: data.oo?.length || 0
  };

  return (
    <div className="app">
      {showSidebar && (
        <aside className="sidebar">
          <div className="sidebar-header">Upload Files</div>
          {['stock', 'so', 'po', 'oo'].map(t => (
            <div key={t} className="upload-box">
              <label>{t.toUpperCase()}</label>
              <input type="file" onChange={e => handleUpload(e, t)} />
            </div>
          ))}
          <div style={{marginTop:'auto'}}>
            <button onClick={() => { if (window.confirm('Delete all?')) { Promise.all(Object.values(TABLE_MAP).map(t => supabase.from(t).delete().neq('id', -1))).then(() => window.location.reload()); } }} className="tbl-btn-danger" style={{width:'100%'}}>Reset All</button>
          </div>
        </aside>
      )}

      <div className="main-layout">
        <header className="header">
          <button className="sidebar-toggle" onClick={() => setShowSidebar(!showSidebar)} title={showSidebar ? "Hide Sidebar" : "Show Sidebar"}>
            {showSidebar ? "◀" : "▶"}
          </button>
          <div className="header-titles">
            <h1 className="header-main">Siddhi Kabel Corporation Pvt Ltd</h1>
            <p className="header-sub">Pending SO review report</p>
          </div>
          <div className="header-search-container">
            <span className="header-search-label">UNIVERSAL SEARCH:</span>
            <input className="header-search" placeholder="Search across all columns..." value={searchTerm} onChange={e => setSearchTerm(e.target.value)} />
          </div>
        </header>

        <nav className="tabs">
          {Object.keys(TAB_LABELS).map(t => (
            <button key={t} className={activeTab === t ? 'active' : ''} onClick={() => setActiveTab(t)}>
              {TAB_LABELS[t]} {counts[t] > 0 ? `(${counts[t]})` : ''}
            </button>
          ))}
        </nav>

        <main className="content">
          {loading ? <div className="empty">Loading...</div> : (
            activeTab === 'dashboard' ? <Dashboard data={data} searchTerm={searchTerm} /> :
            activeTab === 'master' ? <MasterReport data={data} searchTerm={searchTerm} /> :
            <DataTable list={data[activeTab]} headers={headers[activeTab]} searchTerm={searchTerm} />
          )}
        </main>
      </div>
      <style>{css}</style>
    </div>
  );
};

const css = `
  :root { --bg: #f8fafc; --surface: #ffffff; --border: #e2e8f0; --accent: #2563eb; --text: #1e293b; --muted: #64748b; --font: 'Cambria', serif; }
  body { margin: 0; font-family: var(--font); background: var(--bg); color: var(--text); font-size: 13px; }
  
  .app { display: flex; height: 100vh; overflow: hidden; background: var(--bg); }
  
  .sidebar { width: 180px; background: #0f172a; color: #f1f5f9; padding: 15px; display: flex; flex-direction: column; gap: 12px; border-right: 1px solid var(--border); flex-shrink: 0; }
  .sidebar-header { font-size: 11px; font-weight: bold; color: #38bdf8; text-transform: uppercase; letter-spacing: 1.5px; margin-bottom: 5px; opacity: 0.8; }
  .upload-box { background: #1e293b; padding: 8px; border-radius: 6px; display: flex; flex-direction: column; gap: 4px; border: 1px solid #334155; transition: all 0.2s; }
  .upload-box:hover { border-color: #38bdf8; background: #1a2436; }
  .upload-box label { font-size: 9px; color: #94a3b8; font-weight: bold; }
  .upload-box input { font-size: 9px; color: #cbd5e1; width: 100%; cursor: pointer; }
  
  .main-layout { flex: 1; display: flex; flex-direction: column; overflow: hidden; }
  
  .header { background: #fff; padding: 12px 30px; border-bottom: 1px solid var(--border); display: flex; align-items: center; justify-content: space-between; gap: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
  .header-titles { display: flex; flex-direction: column; }
  .header-main { margin: 0; font-size: 24px; color: #1e3a8a; font-weight: 900; letter-spacing: -0.8px; line-height: 1; }
  .header-sub { margin: 2px 0 0; font-size: 11px; color: var(--muted); font-weight: 600; text-transform: uppercase; letter-spacing: 1px; }
  
  .sidebar-toggle { background: #f1f5f9; border: 1px solid var(--border); padding: 5px 10px; border-radius: 6px; cursor: pointer; color: var(--accent); font-weight: bold; transition: all 0.2s; margin-right: 10px; }
  .sidebar-toggle:hover { background: #e2e8f0; transform: scale(1.05); }
  
  .header-search-container { flex: 1; display: flex; justify-content: center; align-items: center; gap: 15px; }
  .header-search-label { font-size: 10px; font-weight: 900; color: #475569; white-space: nowrap; letter-spacing: 0.5px; }
  .header-search { width: 400px; padding: 10px 25px; border: 2px solid #e2e8f0; border-radius: 30px; outline: none; font-size: 13px; font-family: var(--font); transition: all 0.2s; background: #f8fafc; }
  .header-search:focus { border-color: var(--accent); background: #fff; box-shadow: 0 0 0 4px rgba(37, 99, 235, 0.1); width: 450px; }
  
  .tabs { display: flex; background: #fff; border-bottom: 1px solid var(--border); padding: 0 30px; gap: 25px; flex-shrink: 0; }
  .tabs button { padding: 15px 5px; border: none; background: none; cursor: pointer; border-bottom: 4px solid transparent; font-size: 11px; font-family: var(--font); color: var(--muted); font-weight: 700; transition: all 0.2s; text-transform: uppercase; letter-spacing: 0.5px; }
  .tabs button.active { border-bottom-color: var(--accent); color: var(--accent); }
  .tabs button:hover:not(.active) { color: var(--text); border-bottom-color: #cbd5e1; }
  
  .content { flex: 1; overflow: auto; padding: 20px 30px; }
  .table-container { background: #fff; border: 1px solid var(--border); overflow: auto; max-height: 72vh; border-radius: 10px; box-shadow: 0 4px 15px -3px rgba(0,0,0,0.07); }
  table { width: 100%; border-collapse: collapse; font-size: 11px; }
  th { background: #f1f5f9; padding: 8px 10px; border: 1px solid var(--border); position: sticky; top: 0; z-index: 10; white-space: nowrap; font-weight: 800; color: #334155; text-transform: uppercase; font-size: 10px; }
  td { padding: 5px 10px; border: 1px solid #f1f5f9; white-space: nowrap; border-right: 1px solid #e2e8f0; line-height: 1.4; color: #475569; }
  tr:nth-child(even) { background: #f8fafc; }
  tr:hover { background: #eff6ff; }
  
  .tbl-toolbar { display: flex; gap: 10px; margin-bottom: 8px; align-items: center; font-size: 11px; }
  .tbl-search { padding: 4px 10px; border: 1px solid var(--border); border-radius: 6px; width: 180px; outline: none; font-size: 11px; font-family: var(--font); background: #f8fafc; }
  .tbl-btn { padding: 4px 12px; border: 1px solid var(--border); background: #fff; border-radius: 6px; cursor: pointer; font-size: 10px; font-family: var(--font); font-weight: bold; transition: all 0.2s; }
  .tbl-btn:hover { background: #f1f5f9; }
  .tbl-btn-export { padding: 4px 15px; border: 1px solid #bbf7d0; background: #f0fdf4; color: #15803d; border-radius: 6px; cursor: pointer; font-size: 10px; font-weight: bold; }
  .tbl-btn-danger { padding: 8px 15px; border: 1px solid #fecaca; background: #fef2f2; color: #b91c1c; border-radius: 8px; cursor: pointer; font-size: 10px; font-weight: bold; transition: all 0.2s; }
  .tbl-btn-danger:hover { background: #fee2e2; transform: translateY(-1px); }
  
  .col-filter { width: 100%; box-sizing: border-box; margin-top: 4px; padding: 2px 4px; border: 1px solid var(--border); border-radius: 4px; font-size: 9px; font-family: var(--font); outline: none; font-weight: normal; background: #fff; transition: all 0.2s; }
  .col-filter:focus { border-color: var(--accent); box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.1); }
  
  .stat-pill { background: #fff; padding: 5px 15px; border: 1px solid var(--border); border-radius: 25px; font-size: 11px; font-weight: 700; box-shadow: 0 1px 2px rgba(0,0,0,0.05); }
  .card { background: #fff; padding: 20px; border: 1px solid var(--border); border-radius: 14px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
  .empty { padding: 60px; text-align: center; color: var(--muted); font-size: 15px; font-weight: 500; }
`;

class ErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { hasError: false }; }
  static getDerivedStateFromError() { return { hasError: true }; }
  render() { if (this.state.hasError) return <div style={{ padding: 20, textAlign: 'center' }}><h2>Error.</h2><button onClick={() => { localStorage.clear(); window.location.reload(); }}>Reset</button></div>; return this.props.children; }
}

const WrappedApp = () => <ErrorBoundary><App /></ErrorBoundary>;
export default WrappedApp;
