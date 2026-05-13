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
  for (const k of keys) {
    if (r[k] !== undefined && r[k] !== null && r[k] !== '') return r[k];
    const nk = k.toLowerCase().replace(/[\s_.']/g, '');
    for (const rk in r) {
      if (rk.toLowerCase().replace(/[\s_.']/g, '') === nk && r[rk] !== undefined && r[rk] !== null && r[rk] !== '') return r[rk];
    }
  }
  return '';
};

const isNumericCol = (key) => {
  const kl = String(key).toLowerCase();
  if (kl.includes('date') || kl.includes('no') || kl.includes('party') || kl.includes('name') || 
      kl.includes('status') || kl.includes('action') || kl.includes('category') || kl.includes('due') || 
      kl.includes('on') || kl.includes('so') || kl.includes('details')) {
    if (!kl.includes('qty') && !kl.includes('balance') && !kl.includes('ordered') && !kl.includes('value')) return false;
  }
  return kl.includes('rate') || kl.includes('value') || kl.includes('qty') ||
    kl.includes('quantity') || kl.includes('balance') || kl.includes('amount') ||
    kl.includes('price') || kl.includes('ordered') || kl.includes('stock') ||
    kl.includes('allocated') || kl.includes('discount');
};

const isDateCol = (key) => {
  const kl = String(key).toLowerCase();
  return kl.includes('date') || kl.includes('due') || kl.includes('on') || kl.includes('mad');
};

const fmtCell = (key, val) => {
  if (val === undefined || val === null || val === '') return '-';
  if (typeof val === 'object' && !(val instanceof Date) && !React.isValidElement(val)) return JSON.stringify(val);
  if (isDateCol(key)) { const d = toDate(val); if (d) return fmtDate(d); }
  if (key === 'Category') {
    const isDue = String(val) === 'Due' || String(val) === 'Due Order';
    return <span style={{ background: isDue ? '#fdf2f8' : '#f0fdf4', color: isDue ? '#9d174d' : '#15803d', padding: '2px 8px', borderRadius: '4px', fontWeight: 'bold', fontSize: '9px', border: `1px solid ${isDue ? '#fbcfe8' : '#bbf7d0'}`, display: 'inline-block' }}>{String(val)}</span>;
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

const saveData = async (type, list) => {
  if (!type || !TABLE_MAP[type]) return;
  try {
    await supabase.from(TABLE_MAP[type]).delete().neq('id', -1); 
    if (list && list.length > 0) {
      const chunkSize = 500;
      for (let i = 0; i < list.length; i += chunkSize) {
        const chunk = list.slice(i, i + chunkSize);
        await supabase.from(TABLE_MAP[type]).insert(chunk.map(r => ({ data: r })));
      }
    }
  } catch (e) { console.error(`Cloud Save failed for ${type}`, e); }
};

const fetchFull = async (table) => {
  let allData = []; let from = 0; const step = 1000;
  while (true) {
    const { data, error } = await supabase.from(table).select('data').range(from, from + step - 1);
    if (error) throw error;
    if (!data || data.length === 0) break;
    allData = allData.concat(data.map(r => r.data));
    if (data.length < step) break;
    from += step;
  }
  return allData;
};

const loadData = async () => {
  try {
    const [stock, so, po, oo] = await Promise.all([fetchFull('stock'), fetchFull('sales_orders'), fetchFull('purchase_orders'), fetchFull('vendor_orders')]);
    return { stock, so, po, oo };
  } catch (e) { console.error('Cloud Load failed', e); return { stock: [], so: [], po: [], oo: [] }; }
};

// ── Logic: Build Master Report ──────────────────────────────────────────────

const buildMasterReport = (data) => {
  if (!data || !data.so || !data.so.length) return [];
  const TODAY = new Date(); TODAY.setHours(0, 0, 0, 0);
  const DUE_CUTOFF = new Date(TODAY); DUE_CUTOFF.setDate(DUE_CUTOFF.getDate() + 30);

  // 1. Stock Preparation
  const stockMap = {};
  data.stock.forEach(s => {
    const desc = clean(getV(s, ['Description', 'Name of Item', 'Item Name'])).toLowerCase();
    const part = toStrict(getV(s, ['Part No', 'Material', 'Material Code']));
    const qty = Number(getV(s, ['Quantity', 'Qty', 'Closing Stock'])) || 0;
    const entry = { total: qty, remaining: qty };
    if (desc) stockMap[desc] = entry;
    if (part) stockMap[part] = entry; // Match on part number too
  });

  // 2. PO Preparation
  const poGroups = {};
  data.po.forEach(p => {
    const desc = clean(getV(p, ['Name of Item', 'Item Name'])).toLowerCase();
    const part = toStrict(getV(p, ['Part No', 'Material', 'Material Code']));
    const entry = {
      date: getV(p, ['Date', 'PO Date']),
      order: getV(p, ['Order No', 'Order', 'PO No']),
      party: getV(p, ['Party Name', "Party's Name", 'Vendor']),
      ordered: Number(getV(p, ['Ordered', 'Quantity'])) || 0,
      balance: Number(getV(p, ['Balance', 'Pending Qty', 'Open Qty'])) || 0,
      remaining: Number(getV(p, ['Balance', 'Pending Qty', 'Open Qty'])) || 0,
      dueOn: toDate(getV(p, ['Due on', 'Delivery Date']))
    };
    if (desc) { if (!poGroups[desc]) poGroups[desc] = []; poGroups[desc].push(entry); }
    if (part && part !== desc) { if (!poGroups[part]) poGroups[part] = []; poGroups[part].push(entry); }
  });
  Object.values(poGroups).forEach(g => g.sort((a, b) => (a.dueOn || 0) - (b.dueOn || 0)));

  // 3. OO Preparation (Vendor Orders)
  const ooEntries = data.oo.map(o => ({
    matNoStrict: toStrict(getV(o, ['Material No', 'Part No', 'Material'])),
    soNo: clean(getV(o, ['Sales Order No', 'SO No', 'Order No'])),
    open: Number(getV(o, ['Open Qty', 'Balance'])) || 0,
    remaining: Number(getV(o, ['Open Qty', 'Balance'])) || 0,
    mad: getV(o, ['Estimated M.A.D.', 'M.A.D.', 'Availability Date'])
  }));

  // 4. Map SO
  const overdue = data.so.map(r => {
    const dueOn = toDate(getV(r, ['Due on', 'Delivery Date']));
    const balance = Number(getV(r, ['Balance', 'Pending Qty', 'Open Qty'])) || 0;
    const name = clean(getV(r, ['Name of Item', 'Description'])).toLowerCase();
    const partNo = clean(getV(r, ['Part No', 'Material Code', 'Material']));
    const partStrict = toStrict(partNo);
    return {
      _dueDate: dueOn, _balance: balance, _category: (dueOn && dueOn < DUE_CUTOFF) ? 'Due Order' : 'Schedule Order', _name: name, _partStrict: partStrict,
      date: getV(r, ['Date', 'Order Date']), order: getV(r, ['Order No', 'Order']), partyName: getV(r, ['Party Name', "Party's Name"]), nameOfItem: clean(getV(r, ['Name of Item'])), partNo, ordered: Number(getV(r, ['Ordered'])) || 0, balance, value: Number(getV(r, ['Value', 'Amount'])) || 0, dueOn, stockQty: 0, allocatedQty: 0, poAllocated: 0, vpoAllocated: 0, stockStatus: '', action: '',
      poDetails: { order: '', party: '', date: '', ordered: 0, balance: 0, dueOn: null }, vpoDetails: { soNo: '', open: 0, mad: '' }
    };
  });

  overdue.sort((a, b) => { if (a._category !== b._category) return a._category === 'Due Order' ? -1 : 1; return (a._dueDate || 0) - (b._dueDate || 0); });

  overdue.forEach(r => {
    // PASS 1: Stock (Match name OR part no)
    const s = stockMap[r._name] || stockMap[r._partStrict];
    r.stockQty = s ? s.total : 0;
    if (s && s.remaining > 0) { const take = Math.min(s.remaining, r._balance); s.remaining -= take; r.allocatedQty = take; }
    r.stockStatus = r.allocatedQty >= r._balance ? 'Available' : 'Need to Arrange';
    
    // PASS 2: PO (Match name OR part no)
    let rem = r._balance - r.allocatedQty;
    if (rem > 0) {
      const pos = poGroups[r._name] || poGroups[r._partStrict];
      if (pos) { for (const p of pos) { if (p.remaining <= 0) continue; const take = Math.min(p.remaining, rem); p.remaining -= take; r.poAllocated += take; if (!r.poDetails.order) r.poDetails = { ...p }; rem -= take; if (rem <= 0) break; } }
    }
    
    // PASS 3: Vendor Open Orders (VPO)
    let remV = r._balance - r.allocatedQty - r.poAllocated;
    if (remV > 0) {
      const vpos = ooEntries.filter(o => o.matNoStrict === r._partStrict && o.remaining > 0);
      for (const v of vpos) { const take = Math.min(v.remaining, remV); v.remaining -= take; r.vpoAllocated += take; if (!r.vpoDetails.soNo) r.vpoDetails = { soNo: v.soNo, open: v.open, mad: v.mad }; remV -= take; if (remV <= 0) break; }
    }
    
    const covered = r.allocatedQty + r.vpoAllocated + r.poAllocated;
    if (r.balance <= 0 || r.allocatedQty >= r.balance) r.action = 'Covered';
    else if (covered > 0) r.action = 'Partial Qty ordered need to make order';
    else r.action = 'Make PO need to raise';
  });
  return overdue;
};

// ── Components ───────────────────────────────────────────────────────────────

const DataTable = ({ list, columnGroups = [], rowStyle = null, hideSearch = false }) => {
  const [sortCol, setSortCol] = useState(null);
  const [sortDir, setSortDir] = useState('asc');
  const [filter, setFilter] = useState('');
  const [page, setPage] = useState(1);
  const pageSize = 50;
  useEffect(() => setPage(1), [filter, sortCol, sortDir]);
  const cols = useMemo(() => { if (!list || !list.length) return []; const first = list[0]; return Object.keys(first).filter(k => { const v = first[k]; return v === null || v === undefined || typeof v !== 'object' || v instanceof Date; }); }, [list]);
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
      {!hideSearch && <div className="tbl-toolbar"><input className="tbl-search" placeholder="🔍 Search..." value={filter} onChange={e => setFilter(e.target.value)} /><span className="tbl-count">{filtered.length} rows</span></div>}
      <div className="table-container">
        <table>
          <thead>
            {columnGroups.length > 0 && <tr>{columnGroups.map((g, i) => <th key={i} colSpan={g.span} style={{ background: g.bg, color: '#fff' }}>{g.label}</th>)}</tr>}
            <tr>{cols.map(c => <th key={c} onClick={() => { if (sortCol === c) setSortDir(d => d === 'asc' ? 'desc' : 'asc'); else { setSortCol(c); setSortDir('asc'); } }} style={{ cursor: 'pointer', textAlign: isNumericCol(c) ? 'right' : 'left' }}>{c} {sortCol === c ? (sortDir === 'asc' ? '▲' : '▼') : ''}{totals[c] !== undefined && <div style={{ fontSize: '9px', color: '#1e40af' }}>{c.toLowerCase().includes('value') ? fmtNum(totals[c], 2) : fmtNum(totals[c])}</div>}</th>)}</tr>
          </thead>
          <tbody>
            {filtered.slice((page - 1) * pageSize, page * pageSize).map((r, i) => <tr key={i} style={rowStyle ? rowStyle(r) : {}}>{cols.map(c => <td key={c} style={{ textAlign: isNumericCol(c) ? 'right' : 'left' }}>{fmtCell(c, r[c])}</td>)}</tr>)}
          </tbody>
        </table>
      </div>
      {filtered.length > pageSize && <div className="tbl-toolbar" style={{ marginTop: 10 }}><button disabled={page === 1} onClick={() => setPage(p => p - 1)} className="tbl-clear">Prev</button><span>Page {page} of {Math.ceil(filtered.length / pageSize)}</span><button disabled={page >= Math.ceil(filtered.length / pageSize)} onClick={() => setPage(p => p + 1)} className="tbl-clear">Next</button></div>}
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
  const display = useMemo(() => filtered.map(r => ({ 'Category': r._category, 'Date': fmtDate(r.date), 'Order No': r.order, 'Party Name': r.partyName, 'Name of Item': r.nameOfItem, 'Part No': r.partNo, 'Ordered': r.ordered, 'Balance': r.balance, 'Value': r.value, 'Due on': fmtDate(r.dueOn), 'Stock Qty': r.stockQty, 'Allocated Qty': r.allocatedQty, 'Status': r.stockStatus, 'PO Order': r.poDetails?.order || '-', 'PO Balance': r.poDetails?.balance || 0, 'VPO SO': r.vpoDetails?.soNo || '-', 'VPO MAD': fmtDate(r.vpoDetails?.mad), 'Action': r.action })), [filtered]);
  return (
    <div className="card">
      <div className="tbl-toolbar"><input className="tbl-search" placeholder="🔍 Global Search..." value={search} onChange={e => setSearch(e.target.value)} style={{ width: 300 }} /><span className="tbl-count">{filtered.length} rows</span></div>
      <div style={{ display: 'flex', gap: 15, marginBottom: 20, flexWrap: 'wrap' }}>
        <div className="upload-card" style={{ borderLeft: '4px solid #e11d48' }}>Due: <strong>{stats.due}</strong></div><div className="upload-card" style={{ borderLeft: '4px solid #10b981' }}>Scheduled: <strong>{stats.sch}</strong></div><div className="upload-card" style={{ borderLeft: '4px solid #2563eb' }}>Ordered: <strong>{fmtNum(stats.ordered)}</strong></div><div className="upload-card" style={{ borderLeft: '4px solid #f59e0b' }}>Balance: <strong>{fmtNum(stats.balance)}</strong></div><div className="upload-card" style={{ borderLeft: '4px solid #0ea5e9' }}>Value: <strong>{fmtNum(stats.value, 2)}</strong></div>
      </div>
      <DataTable list={display} hideSearch={true} columnGroups={[{ label: 'Sales Order Details', span: 10, bg: '#1e40af' }, { label: 'Inventory & PO Status', span: 7, bg: '#0369a1' }, { label: 'Action', span: 1, bg: '#111827' }]} rowStyle={r => (r['Category'] === 'Due Order' ? { background: '#fff1f2' } : {})} />
    </div>
  );
};

// ── App ──────────────────────────────────────────────────────────────────────

const App = () => {
  const [data, setData] = useState({ stock: [], so: [], po: [], oo: [] });
  const [activeTab, setActiveTab] = useState('master');
  const [loading, setLoading] = useState(true);
  useEffect(() => { loadData().then(saved => { if (saved) setData(saved); setLoading(false); }); }, []);
  const handleUpload = (e, type) => {
    const file = e.target.files[0]; if (!file) return; setLoading(true);
    const reader = new FileReader(); reader.onload = (evt) => { const wb = XLSX.read(evt.target.result, { type: 'binary' }); const raw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" }); setData(prev => ({ ...prev, [type]: raw })); saveData(type, raw); setLoading(false); }; reader.readAsBinaryString(file);
  };
  const TAB_LABELS = { master: 'MASTER REPORT', stock: 'STOCK', so: 'SALES ORDERS', po: 'PURCHASE ORDERS', oo: 'OPEN VENDOR ORDERS' };
  return (
    <div className="app">
      <header className="header"><div className="header-logo">SKC</div><div className="header-sub">Pending Order Intelligence</div><button onClick={() => { if (window.confirm('Delete all?')) { Promise.all(Object.values(TABLE_MAP).map(t => supabase.from(t).delete().neq('id', -1))).then(() => window.location.reload()); } }} className="tbl-clear" style={{ marginLeft: 'auto', background: '#fee2e2', color: '#b91c1c' }}>Reset All Data</button></header>
      <div className="upload-section">{['stock', 'so', 'po', 'oo'].map(t => (<div key={t} className="upload-card"><label>{t === 'oo' ? 'VENDOR ORDERS' : t.toUpperCase()}</label><input type="file" onChange={e => handleUpload(e, t)} />{data[t]?.length > 0 && <span className="badge">{data[t].length} rows</span>}</div>))}</div>
      <div className="tabs">{Object.keys(TAB_LABELS).map(t => (<button key={t} className={activeTab === t ? 'active' : ''} onClick={() => setActiveTab(t)}>{TAB_LABELS[t]} {data[t]?.length > 0 ? `(${data[t].length})` : ''}</button>))}</div>
      <main className="content">{loading ? <div className="empty">Loading...</div> : (activeTab === 'master' ? <MasterReport data={data} /> : <DataTable list={data[activeTab]} />)}</main>
      <style>{css}</style>
    </div>
  );
};

const css = `
  :root { --bg: #f8fafc; --surface: #fff; --border: #cbd5e1; --accent: #1e40af; --text: #0f172a; --muted: #475569; --font: 'Cambria', serif; }
  body { margin: 0; font-family: var(--font); background: var(--bg); color: var(--text); font-size: 10px; }
  .header { background: #fff; padding: 10px 20px; border-bottom: 1px solid var(--border); display: flex; align-items: center; gap: 10px; }
  .header-logo { font-size: 22px; font-weight: bold; color: var(--accent); }
  .upload-section { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; padding: 15px 20px; }
  .upload-card { background: #fff; padding: 10px; border: 1px solid var(--border); border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
  .tabs { display: flex; background: #fff; border-bottom: 1px solid var(--border); padding: 0 20px; flex-wrap: wrap; }
  .tabs button { padding: 10px 20px; border: none; background: none; cursor: pointer; border-bottom: 3px solid transparent; transition: 0.2s; font-size: 11px; }
  .tabs button.active { border-bottom-color: var(--accent); color: var(--accent); font-weight: bold; }
  .content { padding: 20px; }
  .table-container { background: #fff; border: 1px solid var(--border); border-radius: 8px; overflow: auto; max-height: 75vh; }
  table { width: 100%; border-collapse: collapse; }
  th { background: #f1f5f9; padding: 8px; border: 1px solid var(--border); position: sticky; top: 0; z-index: 10; white-space: nowrap; }
  td { padding: 6px 8px; border: 1px solid #f1f5f9; white-space: nowrap; border-right: 1px solid #e2e8f0; }
  .tbl-toolbar { display: flex; gap: 15px; margin-bottom: 12px; align-items: center; }
  .tbl-search { padding: 6px 12px; border: 1px solid var(--border); border-radius: 6px; width: 280px; outline: none; }
  .tbl-clear { padding: 6px 12px; border: 1px solid var(--border); background: #fff; border-radius: 4px; cursor: pointer; }
  .card { background: #fff; padding: 20px; border: 1px solid var(--border); border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); }
  .badge { background: #dbeafe; color: #1e40af; padding: 2px 6px; border-radius: 4px; font-size: 9px; margin-top: 5px; display: inline-block; }
  .empty { padding: 40px; text-align: center; color: var(--muted); }
`;

class ErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { hasError: false }; }
  static getDerivedStateFromError() { return { hasError: true }; }
  render() { if (this.state.hasError) return <div style={{ padding: 40, textAlign: 'center' }}><h2>Something went wrong.</h2><button onClick={() => { localStorage.clear(); window.location.reload(); }}>Clear Cache & Restart</button></div>; return this.props.children; }
}

const WrappedApp = () => <ErrorBoundary><App /></ErrorBoundary>;
export default WrappedApp;
