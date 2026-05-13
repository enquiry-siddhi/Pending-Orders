import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { createClient } from '@supabase/supabase-js';

const supabaseUrl = 'https://ehvtgvljtsfgfnaouhmd.supabase.co';
const supabaseKey = 'sb_publishable_2wHnHf6T-A_PdcLLZoCF-w_8CAicB81';
const supabase = createClient(supabaseUrl, supabaseKey);

// ── Helpers ──────────────────────────────────────────────────────────────────

const clean = (s) => String(s || "").trim();
const toStrict = (s) => clean(s).replace(/[^a-z0-9]/gi, '').toLowerCase();

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
  return kl.includes('date') || kl.includes('due') || kl.includes('on') || kl.includes('mad') || kl.includes('requested');
};

const fmtCell = (key, val) => {
  if (val === undefined || val === null || val === '') return '-';
  if (typeof val === 'object' && !(val instanceof Date) && !React.isValidElement(val)) return JSON.stringify(val);
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
  const kl = String(key).toLowerCase();
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
    const part = toStrict(getV(s, ['Part No', 'Material', 'Material Code', 'Code']));
    const qty = Number(getV(s, ['Quantity', 'Qty', 'Closing Stock'])) || 0;
    const entry = { total: qty, remaining: qty };
    if (desc) stockMap[desc] = entry;
    if (part) stockMap[part] = entry;
  });

  const poGroups = {};
  data.po.forEach(p => {
    const desc = clean(getV(p, ['Name of Item', 'Item Name', 'Description', 'Material'])).toLowerCase();
    const part = toStrict(getV(p, ['Part No', 'Material', 'Material Code', 'Code']));
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
      custPODate: getV(o, ['Cust PO Date', 'Customer PO Date', 'PO Date', 'Ref Date', 'PO Date']),
      ordered: Number(getV(o, ['Order Qty', 'Ordered', 'Quantity', 'Qty', 'Ordered Quantity'])) || 0,
      open: Number(getV(o, ['Open Qty', 'Balance', 'Pending', 'Open Quantity', 'Outstanding'])) || 0,
      remaining: Number(getV(o, ['Open Qty', 'Balance', 'Pending', 'Open Quantity', 'Outstanding'])) || 0,
      reqDate: getV(o, ['Cust Req Date', 'Customer Requested Date', 'Requested Date', 'Request Date', 'Req Date', 'Requested Delivery Date']),
      mad: getV(o, ['Estimated M.A.D.', 'M.A.D.', 'Availability Date', 'MAD', 'Est Date', 'Estimated Availability', 'Sch. Date', 'Commit Date'])
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
      poDetails: { order: '', party: '', date: '', ordered: 0, balance: 0, dueOn: null }, vpoDetails: { soNo: '', soDate: '', lineDate: '', line: '', custPO: '', custPODate: '', ordered: 0, open: 0, reqDate: '', mad: '' }
    };
  });
  overdue.forEach(r => {
    const s = stockMap[r._name] || stockMap[r._partStrict]; r.stockQty = s ? s.total : 0;
    if (s && s.remaining > 0) { const take = Math.min(s.remaining, r._balance); s.remaining -= take; r.allocatedQty = take; }
    r.stockStatus = r.allocatedQty >= r._balance ? 'Available' : 'Need to Arrange';
    let rem = r._balance - r.allocatedQty;
    if (rem > 0) {
      const pos = poGroups[r._name] || poGroups[r._partStrict];
      if (pos) { for (const p of pos) { if (p.remaining <= 0) continue; const take = Math.min(p.remaining, rem); p.remaining -= take; r.poAllocated += take; if (!r.poDetails.order) r.poDetails = { ...p }; rem -= take; if (rem <= 0) break; } }
    }
    if (r.poAllocated === 0) {
      let remV = r._balance - r.allocatedQty;
      if (remV > 0) {
        const vpos = ooGroups[r._name] || ooGroups[r._partStrict];
        if (vpos) { for (const v of vpos) { if (v.remaining <= 0) continue; const take = Math.min(v.remaining, remV); v.remaining -= take; r.vpoAllocated += take; if (!r.vpoDetails.soNo) r.vpoDetails = { ...v }; remV -= take; if (remV <= 0) break; } }
      }
    }
    const totalCovered = r.allocatedQty + r.poAllocated + r.vpoAllocated;
    if (r.balance <= 0 || totalCovered >= (r.balance - 0.001)) r.action = 'Covered';
    else if (totalCovered > 0) r.action = 'Partial Qty ordered need to make order';
    else r.action = 'Make PO need to raise';
  });
  return overdue;
};

// ── Components ───────────────────────────────────────────────────────────────

const DataTable = ({ list, headers = [], columnGroups = [], rowStyle = null, hideSearch = false, onExport = null }) => {
  const [sortCol, setSortCol] = useState(null);
  const [sortDir, setSortDir] = useState('asc');
  const [filter, setFilter] = useState('');
  const [page, setPage] = useState(1);
  const pageSize = 100;
  useEffect(() => setPage(1), [filter, sortCol, sortDir]);
  const cols = useMemo(() => { if (!list || !list.length) return []; const keys = headers.length > 0 ? headers : Object.keys(list[0]); return keys.filter(k => { const v = list[0][k]; return v === null || v === undefined || typeof v !== 'object' || v instanceof Date; }); }, [list, headers]);
  const filtered = useMemo(() => {
    if (!list || !list.length) return [];
    let res = filter ? list.filter(r => Object.values(r).some(v => String(v || '').toLowerCase().includes(filter.toLowerCase()))) : list;
    if (sortCol) { res = [...res].sort((a, b) => { const av = a[sortCol], bv = b[sortCol]; if (isNumericCol(sortCol)) return sortDir === 'asc' ? (Number(av)||0) - (Number(bv)||0) : (Number(bv)||0) - (Number(av)||0); return sortDir === 'asc' ? String(av || '').localeCompare(String(bv || '')) : String(bv || '').localeCompare(String(av || '')); }); }
    return res;
  }, [list, filter, sortCol, sortDir]);
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
            <tr>{cols.map(c => <th key={c} onClick={() => { if (sortCol === c) setSortDir(d => d === 'asc' ? 'desc' : 'asc'); else { setSortCol(c); setSortDir('asc'); } }} style={{ cursor: 'pointer', textAlign: isNumericCol(c) ? 'right' : 'left' }}>{c} {sortCol === c ? (sortDir === 'asc' ? '▲' : '▼') : ''}{totals[c] !== undefined && <div style={{ fontSize: '8px', color: '#1e40af' }}>{c.toLowerCase().includes('value') ? fmtNum(totals[c], 2) : fmtNum(totals[c])}</div>}</th>)}</tr>
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

const MasterReport = ({ data }) => {
  const [search, setSearch] = useState('');
  const rawRows = useMemo(() => buildMasterReport(data), [data]);
  const filtered = useMemo(() => (!search.trim() ? rawRows : rawRows.filter(r => Object.values(r).some(v => String(v || '').toLowerCase().includes(search.trim().toLowerCase())))), [rawRows, search]);
  const stats = useMemo(() => {
    const t = { ordered: 0, balance: 0, value: 0, stk: 0, alc: 0, due: 0, sch: 0 };
    filtered.forEach(r => { t.ordered += Number(r.ordered) || 0; t.balance += Number(r.balance) || 0; t.value += Number(r.value) || 0; t.stk += Number(r.stockQty) || 0; t.alc += Number(r.allocatedQty) || 0; if (r._category === 'Due Order') t.due++; else t.sch++; });
    return t;
  }, [filtered]);
  const display = useMemo(() => filtered.map(r => ({
    'Category': r._category, 'Date': fmtDate(r.date), 'Order No': r.order, 'Party Name': r.partyName, 'Name of Item': r.nameOfItem, 'Part No': r.partNo, 'Ordered': r.ordered, 'Balance': r.balance, 'Value': r.value, 'Due on': fmtDate(r.dueOn),
    'Stock Qty': r.stockQty, 'Allocated Qty': r.allocatedQty, 'Status': r.stockStatus,
    'PO Order': r.poDetails?.order || '-', 'PO Party': r.poDetails?.party || '-', 'PO Date': fmtDate(r.poDetails?.date), 'PO Ordered': r.poDetails?.ordered || 0, 'PO Balance': r.poDetails?.balance || 0, 'PO Due': fmtDate(r.poDetails?.dueOn),
    'VPO SO': r.vpoDetails?.soNo || '-', 'VPO SO Date': fmtDate(r.vpoDetails?.soDate), 'VPO Line Date': fmtDate(r.vpoDetails?.lineDate), 'VPO Line': r.vpoDetails?.line || '-', 'VPO CustPO': r.vpoDetails?.custPO || '-', 'VPO CustPODate': fmtDate(r.vpoDetails?.custPODate), 'VPO Ordered': r.vpoDetails?.ordered || 0, 'VPO Open': r.vpoDetails?.open || 0, 'VPO ReqDate': fmtDate(r.vpoDetails?.reqDate), 'VPO MAD': fmtDate(r.vpoDetails?.mad),
    'Action': r.action
  })), [filtered]);
  return (
    <div className="card">
      <div className="tbl-toolbar"><input className="tbl-search" placeholder="Search..." value={search} onChange={e => setSearch(e.target.value)} style={{ width: 220 }} /></div>
      <div style={{ display: 'flex', gap: 10, marginBottom: 10, flexWrap: 'wrap' }}>
        <div className="stat-pill" style={{ borderColor: '#e11d48' }}>Due: <strong>{stats.due}</strong></div><div className="stat-pill" style={{ borderColor: '#10b981' }}>Scheduled: <strong>{stats.sch}</strong></div><div className="stat-pill" style={{ borderColor: '#2563eb' }}>Ordered: <strong>{fmtNum(stats.ordered)}</strong></div><div className="stat-pill" style={{ borderColor: '#f59e0b' }}>Balance: <strong>{fmtNum(stats.balance)}</strong></div><div className="stat-pill" style={{ borderColor: '#0ea5e9' }}>Value: <strong>{fmtNum(stats.value, 2)}</strong></div>
      </div>
      <DataTable list={display} hideSearch={true} columnGroups={[{ label: 'SO Details', span: 10, bg: '#1e40af' }, { label: 'Stock', span: 3, bg: '#0369a1' }, { label: 'PO', span: 6, bg: '#0891b2' }, { label: 'Vendor PO', span: 10, bg: '#0d9488' }, { label: 'Action', span: 1, bg: '#111827' }]} rowStyle={r => (r['Category'] === 'Due Order' ? { background: '#fff1f2' } : {})} />
    </div>
  );
};

// ── App ──────────────────────────────────────────────────────────────────────

const App = () => {
  const [data, setData] = useState({ stock: [], so: [], po: [], oo: [] });
  const [headers, setHeaders] = useState({ stock: [], so: [], po: [], oo: [] });
  const [activeTab, setActiveTab] = useState('master');
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
  const masterCount = useMemo(() => buildMasterReport(data).length, [data]);
  const TAB_LABELS = { master: 'MASTER', stock: 'STOCK', so: 'SALES', po: 'PO', oo: 'VENDOR PO' };
  return (
    <div className="app">
      <header className="header"><div className="header-logo">SKC</div><div className="header-sub">Intelligence</div><button onClick={() => { if (window.confirm('Delete all?')) { Promise.all(Object.values(TABLE_MAP).map(t => supabase.from(t).delete().neq('id', -1))).then(() => window.location.reload()); } }} className="tbl-btn-danger">Reset All</button></header>
      <div className="upload-section">{['stock', 'so', 'po', 'oo'].map(t => (<div key={t} className="upload-mini-card"><label>{t.toUpperCase()}</label><input type="file" onChange={e => handleUpload(e, t)} /></div>))}</div>
      <div className="tabs">{Object.keys(TAB_LABELS).map(t => (<button key={t} className={activeTab === t ? 'active' : ''} onClick={() => setActiveTab(t)}>{TAB_LABELS[t]} { (t === 'master' ? masterCount : data[t]?.length || 0) > 0 ? `(${t === 'master' ? masterCount : data[t].length})` : ''}</button>))}</div>
      <main className="content">{loading ? <div className="empty">Loading...</div> : (activeTab === 'master' ? <MasterReport data={data} /> : <DataTable list={data[activeTab]} headers={headers[activeTab]} />)}</main>
      <style>{css}</style>
    </div>
  );
};

const css = `
  :root { --bg: #f1f5f9; --surface: #fff; --border: #cbd5e1; --accent: #1e40af; --text: #0f172a; --muted: #475569; --font: 'Cambria', serif; }
  body { margin: 0; font-family: var(--font); background: var(--bg); color: var(--text); font-size: 13px; }
  .header { background: #fff; padding: 5px 15px; border-bottom: 1px solid var(--border); display: flex; align-items: center; gap: 8px; }
  .header-logo { font-size: 18px; font-weight: bold; color: var(--accent); }
  .header-sub { font-size: 11px; color: var(--muted); }
  .upload-section { display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px; padding: 8px 15px; background: #fff; border-bottom: 1px solid var(--border); }
  .upload-mini-card { display: flex; align-items: center; gap: 8px; font-size: 10px; border-right: 1px solid #e2e8f0; padding-right: 8px; }
  .upload-mini-card:last-child { border-right: none; }
  .upload-mini-card label { font-weight: bold; color: var(--muted); }
  .upload-mini-card input { font-size: 9px; width: 120px; }
  .tabs { display: flex; background: #fff; border-bottom: 1px solid var(--border); padding: 0 15px; }
  .tabs button { padding: 5px 12px; border: none; background: none; cursor: pointer; border-bottom: 2px solid transparent; font-size: 11px; font-family: var(--font); }
  .tabs button.active { border-bottom-color: var(--accent); color: var(--accent); font-weight: bold; }
  .content { padding: 10px; }
  .table-container { background: #fff; border: 1px solid var(--border); overflow: auto; max-height: 80vh; }
  table { width: 100%; border-collapse: collapse; font-size: 11px; }
  th { background: #f8fafc; padding: 4px 6px; border: 1px solid var(--border); position: sticky; top: 0; z-index: 10; white-space: nowrap; font-weight: bold; }
  td { padding: 3px 6px; border: 1px solid #f1f5f9; white-space: nowrap; border-right: 1px solid #e2e8f0; line-height: 1.2; }
  .tbl-toolbar { display: flex; gap: 10px; margin-bottom: 5px; align-items: center; font-size: 11px; }
  .tbl-search { padding: 3px 8px; border: 1px solid var(--border); border-radius: 4px; width: 180px; outline: none; font-size: 11px; }
  .tbl-btn { padding: 2px 8px; border: 1px solid var(--border); background: #fff; border-radius: 3px; cursor: pointer; font-size: 10px; }
  .tbl-btn-export { padding: 2px 10px; border: 1px solid #bbf7d0; background: #f0fdf4; color: #15803d; border-radius: 3px; cursor: pointer; font-size: 10px; font-weight: bold; }
  .tbl-btn-danger { padding: 2px 8px; border: 1px solid #fecaca; background: #fef2f2; color: #b91c1c; border-radius: 3px; cursor: pointer; font-size: 10px; margin-left: auto; }
  .stat-pill { background: #fff; padding: 2px 10px; border: 1px solid var(--border); border-radius: 15px; font-size: 11px; }
  .card { background: #fff; padding: 10px; border: 1px solid var(--border); border-radius: 8px; box-shadow: 0 1px 2px rgba(0,0,0,0.05); }
  .empty { padding: 20px; text-align: center; color: var(--muted); }
`;

class ErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { hasError: false }; }
  static getDerivedStateFromError() { return { hasError: true }; }
  render() { if (this.state.hasError) return <div style={{ padding: 20, textAlign: 'center' }}><h2>Error.</h2><button onClick={() => { localStorage.clear(); window.location.reload(); }}>Reset</button></div>; return this.props.children; }
}

const WrappedApp = () => <ErrorBoundary><App /></ErrorBoundary>;
export default WrappedApp;
