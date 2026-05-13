import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { createClient } from '@supabase/supabase-js';

const supabaseUrl = 'https://ehvtgvljtsfgfnaouhmd.supabase.co';
const supabaseKey = 'sb_publishable_2wHnHf6T-A_PdcLLZoCF-w_8CAicB81';
const supabase = createClient(supabaseUrl, supabaseKey);

const clean = (s) => String(s || "").trim();

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

// Cloud persistence mapping
const TABLE_MAP = {
  stock: 'stock',
  so: 'sales_orders',
  po: 'purchase_orders',
  oo: 'vendor_orders'
};

const saveData = async (type, list) => {
  if (!type || !TABLE_MAP[type]) return;
  try {
    // We clear the table and re-insert the fresh batch for that type
    await supabase.from(TABLE_MAP[type]).delete().neq('id', -1); 
    if (list && list.length > 0) {
      // Supabase can handle bulk insert of objects
      const { error } = await supabase.from(TABLE_MAP[type]).insert(list.map(r => ({ data: r })));
      if (error) throw error;
    }
  } catch (e) { console.error(`Cloud Save failed for ${type}`, e); }
};

const loadData = async () => {
  try {
    const results = await Promise.all([
      supabase.from('stock').select('data'),
      supabase.from('sales_orders').select('data'),
      supabase.from('purchase_orders').select('data'),
      supabase.from('vendor_orders').select('data')
    ]);

    return {
      stock: results[0].data?.map(r => r.data) || [],
      so:    results[1].data?.map(r => r.data) || [],
      po:    results[2].data?.map(r => r.data) || [],
      oo:    results[3].data?.map(r => r.data) || []
    };
  } catch (e) { 
    console.error('Cloud Load failed', e); 
    return { stock: [], so: [], po: [], oo: [] };
  }
};

class ErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { hasError: false, error: null }; }
  static getDerivedStateFromError(error) { return { hasError: true, error }; }
  render() {
    if (this.state.hasError) {
      return (
        <div style={{ padding: 40, textAlign: 'center', fontFamily: 'sans-serif' }}>
          <h2>Oops! Something went wrong.</h2>
          <p style={{ color: '#b91c1c' }}>{this.state.error?.message}</p>
          <button onClick={() => { localStorage.clear(); location.reload(); }} style={{ padding: '10px 20px', cursor: 'pointer' }}>
            Clear Cache & Restart
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}

const App = () => {
  const [data, setData] = useState({ stock: [], so: [], po: [], oo: [] });
  const [activeTab, setActiveTab] = useState('master');
  const [loading, setLoading] = useState(true);

  // Load data once on mount
  useEffect(() => {
    loadData().then(saved => {
      if (saved) setData(saved);
      setLoading(false);
    });
  }, []);

  // Save data on changes - DEPRECATED for granular saves
  /*
  useEffect(() => {
    if (!loading) saveData(data);
  }, [data, loading]);
  */

  const clearAllData = (e) => {
    if (e) {
      e.preventDefault();
      e.stopPropagation();
    }
    // Use a small delay to ensure the event loop is clear
    setTimeout(async () => {
      if (window.confirm('This will PERMANENTLY delete all uploaded files from the cloud and reset the app. Continue?')) {
        localStorage.clear();
        try {
          await Promise.all([
            supabase.from('stock').delete().neq('id', -1),
            supabase.from('sales_orders').delete().neq('id', -1),
            supabase.from('purchase_orders').delete().neq('id', -1),
            supabase.from('vendor_orders').delete().neq('id', -1)
          ]);
          window.location.reload();
        } catch (e) {
          console.error('Cloud Delete failed', e);
          window.location.reload();
        }
      }
    }, 50);
  };

  const handleUpload = (e, type) => {
    const file = e.target.files[0];
    if (!file) return;
    setLoading(true);
    const reader = new FileReader();
    reader.onload = (evt) => {
      const wb = XLSX.read(evt.target.result, { type: 'binary' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      // defval: "" ensures all columns are captured even if the first row is empty
      const raw = XLSX.utils.sheet_to_json(ws, { defval: "" });
      
      // Update local state
      setData(prev => ({ ...prev, [type]: raw }));
      
      // Update cloud storage (granular save)
      saveData(type, raw);
      
      setLoading(false);
    };
    reader.readAsBinaryString(file);
  };

  const downloadTemplate = (type) => {
    let headers = [];
    let filename = "";
    if (type === 'stock') {
      headers = ["Description", "Quantity"];
      filename = "Stock_Template.xlsx";
    } else if (type === 'so') {
      headers = ["Date", "Order No", "Party Name", "Name of Item", "Part No", "Ordered", "Balance", "Value", "Due on"];
      filename = "Sales_Order_Template.xlsx";
    } else if (type === 'po') {
      headers = ["Date", "Order No", "Party Name", "Name of Item", "Part No", "Ordered", "Balance", "Value", "Due on"];
      filename = "Purchase_Order_Template.xlsx";
    } else if (type === 'oo') {
      headers = ["Material No", "Sales Order No", "Sales Order Date", "Line No", "Customer PO No", "Order Qty", "Open Qty", "Cust Req Date", "Estimated M.A.D.", "Import by", "Warehouse Name"];
      filename = "Vendor_Order_Template.xlsx";
    }

    const ws = XLSX.utils.aoa_to_sheet([headers]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    XLSX.writeFile(wb, filename);
  };

  // Flexible field getter
  const getV = (r, keys) => {
    for (const k of keys) {
      if (r[k] !== undefined && r[k] !== null && r[k] !== '') return r[k];
      const nk = k.toLowerCase().replace(/[\s_.]/g, '');
      for (const rk in r) {
        if (rk.toLowerCase().replace(/[\s_.]/g, '') === nk && r[rk] !== undefined && r[rk] !== null && r[rk] !== '') return r[rk];
      }
    }
    return '';
  };

  // ── Master Report Builder ────────────────────────────────────────────────────
  const buildMasterReport = () => {
    if (!data.so.length) return [];

    const TODAY = new Date(); TODAY.setHours(0, 0, 0, 0);
    const DUE_CUTOFF = new Date(TODAY); DUE_CUTOFF.setDate(DUE_CUTOFF.getDate() + 30);

    // Build stock index keyed by FULL Description (lowercase)
    // Also keep inner name (inside parentheses) as a secondary key
    const stockEntries = data.stock.map(s => {
      const desc = clean(s['Description'] || s['Item Name'] || s['Name of Item'] || '');
      const qty  = Number(s['Quantity'] || s['Qty'] || s['Closing Stock'] || 0);
      // Extract inner name from "PartNo (ItemName)" format
      const pOpen = desc.indexOf('(');
      const pClose = desc.lastIndexOf(')');
      const innerName = pOpen > 0 && pClose > pOpen
        ? desc.slice(pOpen + 1, pClose).trim()
        : desc;
      return { fullDesc: desc, innerName, qty };
    });

    // Consolidated stockMap: key = innerName lowercase, value = { total, remaining, fullDesc }
    const stockMap = {};
    stockEntries.forEach(e => {
      const key = e.innerName.toLowerCase();
      if (!stockMap[key]) stockMap[key] = { total: 0, remaining: 0 };
      stockMap[key].total    += e.qty;
      stockMap[key].remaining += e.qty;
    });

    // Lookup helper: match SO nameOfItem → best stock key
    // 1) Exact match on innerName
    // 2) Stock innerName contains SO nameOfItem
    // 3) SO nameOfItem contains stock innerName
    const findStockKey = (nameOfItem) => {
      const soKey = nameOfItem.toLowerCase().trim();
      if (!soKey) return null;
      if (stockMap[soKey]) return soKey;                                    // exact
      const contained = Object.keys(stockMap).find(k => k.includes(soKey)); // SO inside stock
      if (contained) return contained;
      const contains = Object.keys(stockMap).find(k => soKey.includes(k) && k.length > 4); // stock inside SO
      return contains || null;
    };

    // Map SO rows
    const rows = data.so.map(r => {
      const dueOn   = toDate(getV(r, ['Due on', 'Delivery Date', 'Due Date', 'Promised Date']));
      const balance = Number(getV(r, ['Balance', 'Pending Qty', 'Open Qty', 'Outstanding Qty'])) || 0;
      const nameOfItem = clean(getV(r, ['Name of Item', 'Item Name', 'Description']));
      // Due Order: due date is before today + 30 days; Schedule Order: 30+ days away or missing date
      const category = (dueOn && dueOn < DUE_CUTOFF) ? 'Due Order' : 'Schedule Order';

      return {
        _dueDate:   dueOn,
        _balance:   balance,
        _category:  category,
        _nameOfItem: nameOfItem,
        date:       getV(r, ['Date', 'Order Date', 'Document Date']),
        order:      clean(getV(r, ['Order No', 'Order', 'Reference', 'Sales Order', 'Document No'])),
        partyName:  clean(getV(r, ["Party's Name", 'Customer', 'Party', 'Client', 'Sold-to party'])),
        nameOfItem,
        partNo:     clean(getV(r, ['Part No', 'Part Number', 'Material Code', 'Material', 'Item Code'])),
        ordered:    Number(getV(r, ['Ordered', 'Order Qty', 'Quantity'])) || 0,
        balance,
        value:      Number(getV(r, ['Value', 'Amount', 'Total Amount'])) || 0,
        dueOn,
        // Initialize all status fields to prevent "undefined" crashes
        stockQty:   0,
        allocatedQty: 0,
        stockStatus: '',
        poAllocated: 0,
        poDetails: { date: '', order: '', party: '', ordered: 0, balance: 0, dueOn: null },
        vpoAllocated: 0,
        vpoDetails: { soNo: '', soDate: '', lineDate: '', line: '', custPO: '', ordered: 0, open: 0, reqDate: '', importBy: '', wh: '' },
        action: ''
      };
    });

    // Show ALL rows from SO file
    const overdue = [...rows];

    // Sort FIFO: Due Order first → then Schedule, within each group by date asc
    overdue.sort((a, b) => {
      if (a._category !== b._category) return a._category === 'Due Order' ? -1 : 1;
      if (!a._dueDate && !b._dueDate) return 0;
      if (!a._dueDate) return 1;
      if (!b._dueDate) return -1;
      return a._dueDate - b._dueDate;
    });

    // ── PASS 2: Pending PO Allocation ──────────────────────────────────────────
    const poEntries = data.po.map(p => {
      const nameOfItem = clean(getV(p, ['Name of Item', 'Item Name', 'Description']));
      const qty = Number(getV(p, ['Balance', 'Pending Qty', 'Open Qty', 'Outstanding Qty'])) || 0;
      return {
        _nameOfItem: nameOfItem,
        date: getV(p, ['Date', 'Order Date', 'PO Date']),
        order: getV(p, ['Order No', 'Order', 'PO No', 'Reference']),
        partyName: getV(p, ["Party's Name", 'Vendor', 'Supplier', 'Party']),
        ordered: Number(getV(p, ['Ordered', 'Order Qty', 'Quantity'])) || 0,
        balance: qty,
        remaining: qty,
        dueOn: toDate(getV(p, ['Due on', 'Delivery Date', 'Due Date', 'Promised Date']))
      };
    });

    // Group POs by item name (fuzzy)
    const poGroups = {};
    poEntries.forEach(p => {
      const key = findStockKey(p._nameOfItem) || p._nameOfItem.toLowerCase().trim();
      if (!poGroups[key]) poGroups[key] = [];
      poGroups[key].push(p);
    });
    // Sort POs in each group by due date (FIFO)
    Object.values(poGroups).forEach(group => {
      group.sort((a, b) => (a.dueOn || 0) - (b.dueOn || 0));
    });

    // Lookup stock qty per row (total, before allocation)
    overdue.forEach(r => {
      const matchKey = findStockKey(r._nameOfItem);
      const s = matchKey ? stockMap[matchKey] : null;
      r.stockQty = s ? s.total : 0;
    });

    // FIFO allocation from STOCK
    overdue.forEach(r => {
      const matchKey = findStockKey(r._nameOfItem);
      const s = matchKey ? stockMap[matchKey] : null;
      if (!s || s.remaining <= 0) { r.allocatedQty = 0; return; }
      const take = Math.min(s.remaining, r._balance);
      s.remaining -= take;
      r.allocatedQty = take;
    });

    // FIFO allocation from PO
    overdue.forEach(r => {
      const shortage = r._balance - r.allocatedQty;
      r.stockStatus = r.allocatedQty >= r._balance ? 'Available' : 'Need to Arrange';
      
      if (shortage <= 0) return;

      const matchKey = findStockKey(r._nameOfItem) || r._nameOfItem.toLowerCase().trim();
      const pos = poGroups[matchKey];
      if (!pos || pos.length === 0) return;

      let remainingShortage = shortage;
      for (const p of pos) {
        if (p.remaining <= 0) continue;
        const take = Math.min(p.remaining, remainingShortage);
        p.remaining -= take;
        r.poAllocated += take;
        remainingShortage -= take;
        
        // Take details from the first PO that provides quantity
        if (!r.poDetails.order) {
          r.poDetails = {
            date: p.date,
            order: p.order,
            party: p.partyName,
            ordered: p.ordered,
            balance: p.balance,
            dueOn: p.dueOn
          };
        }
        if (remainingShortage <= 0) break;
      }
    });

    // PASS 3: Open Vendor Orders Allocation (VPO)
    if (!data.oo || data.oo.length === 0) return overdue;

    // Helper to find column name in a row object by multiple keywords
    const findCol = (row, keywords) => {
      const rKeys = Object.keys(row);
      for (const k of keywords) {
        const found = rKeys.find(rk => rk.toLowerCase().replace(/[\s_.-]/g, '').includes(k.toLowerCase().replace(/[\s_.-]/g, '')));
        if (found) return found;
      }
      return null;
    };

    // Pre-detect columns in OO
    const sampleOO = data.oo[0];
    const ooCols = {
      mat: findCol(sampleOO, ['Material No', 'Material', 'Part No', 'Item No', 'Article Number', 'Product Code']),
      so: findCol(sampleOO, ['Sales Order No', 'Sales Order', 'SO No', 'Order No']),
      soDate: findCol(sampleOO, ['So Created Date', 'SO Date', 'Created Date']),
      lineDate: findCol(sampleOO, ['Line Item Created Date', 'Line Date']),
      line: findCol(sampleOO, ['Line No', 'Line Item', 'Pos', 'Position', 'Item']),
      custPO: findCol(sampleOO, ['Customer PO No', 'Customer PO', 'PO No', 'Ref No', 'Cust PO']),
      ordQty: findCol(sampleOO, ['Oder Qty', 'Order Qty', 'Ordered', 'Qty']),
      opnQty: findCol(sampleOO, ['Open Qty', 'Balance', 'Outstanding', 'Open']),
      reqDate: findCol(sampleOO, ['Customer Requested Date', 'Requested Date', 'Req Date', 'Delivery Date']),
      mad: findCol(sampleOO, ['Estimated M.A.D.', 'M.A.D.', 'Estimated MAD', 'MAD', 'Estimated Date', 'Est Date', 'Availability Date', 'Est MAD', 'EST MAD', 'MAD Date', 'Exp MAD', 'Expected MAD']),
      impBy: findCol(sampleOO, ['Import by', 'Buyer', 'Importer']),
      wh: findCol(sampleOO, ['Warehouse Name', 'Warehouse', 'Location', 'Whse']),
    };

    const ooEntries = data.oo.map(o => {
      const matVal = ooCols.mat ? o[ooCols.mat] : '';
      const matNo = clean(matVal);
      const openQty = Number(ooCols.opnQty ? o[ooCols.opnQty] : 0) || 0;
      return {
        _matNo: matNo,
        _matNoStrict: matNo.replace(/[^a-z0-9]/gi, '').toLowerCase(), // Only alphanumeric
        soNo: clean(ooCols.so ? o[ooCols.so] : ''),
        soDate: ooCols.soDate ? o[ooCols.soDate] : '',
        lineDate: ooCols.lineDate ? o[ooCols.lineDate] : '',
        line: clean(ooCols.line ? o[ooCols.line] : ''),
        custPONo: clean(ooCols.custPO ? o[ooCols.custPO] : ''),
        orderQty: Number(ooCols.ordQty ? o[ooCols.ordQty] : 0) || 0,
        openQty: openQty,
        remaining: openQty,
        reqDate: ooCols.reqDate ? o[ooCols.reqDate] : '',
        mad: ooCols.mad ? o[ooCols.mad] : '',
        importBy: ooCols.impBy ? o[ooCols.impBy] : '',
        whName: ooCols.wh ? o[ooCols.wh] : '',
      };
    });

    // Part No matching helper
    const getPartNoCandidates = (soPartNo) => {
      const p = clean(soPartNo);
      if (!p) return [];
      const c = [p];
      if (p.includes('U')) c.push(p + '100');
      if (p.endsWith('1')) c.push(p.slice(0, -1));
      return c;
    };

    // Pre-index OO by Part Number for O(1) lookups in Pass 3
    const ooIndex = {};
    const ooIndexStrict = {};
    ooEntries.forEach(o => {
      const key = (o._matNo || '').toLowerCase();
      if (!ooIndex[key]) ooIndex[key] = [];
      ooIndex[key].push(o);

      const keyS = o._matNoStrict;
      if (!ooIndexStrict[keyS]) ooIndexStrict[keyS] = [];
      ooIndexStrict[keyS].push(o);
    });

    // Allocate VPO to remaining shortage
    overdue.forEach(r => {
      const shortageToFill = r._balance - r.allocatedQty - r.poAllocated;
      
      const candidates = getPartNoCandidates(r.partNo);
      const candidatesStrict = candidates.map(x => x.replace(/[^a-z0-9]/gi, '').toLowerCase());
      
      // Find matching OO lines using indexes
      const matchesSet = new Set();
      candidates.forEach(c => (ooIndex[c.toLowerCase()] || []).forEach(o => matchesSet.add(o)));
      candidatesStrict.forEach(cs => (ooIndexStrict[cs] || []).forEach(o => matchesSet.add(o)));
      let matches = Array.from(matchesSet);

      // Fallback: Match by Item Name if Part No fails
      if (matches.length === 0 && r.nameOfItem) {
        const itemKey = r.nameOfItem.toLowerCase().replace(/[^a-z0-9]/gi, '');
        matches = ooEntries.filter(o => {
          const ooItem = (o._matNo || '').toLowerCase().replace(/[^a-z0-9]/gi, '');
          return ooItem && (ooItem.includes(itemKey) || itemKey.includes(ooItem));
        });
      }
      
      if (matches.length === 0) return;

      // Priority 1: Match Pending PO No with OO Customer PO No
      const poNo = clean(r.poDetails.order).replace(/[^a-z0-9]/gi, '').toLowerCase();
      const priorityMatches = poNo ? matches.filter(o => {
        const cpo = o.custPONo.replace(/[^a-z0-9]/gi, '').toLowerCase();
        return cpo.includes(poNo) || poNo.includes(cpo);
      }) : [];
      
      const sortedMatches = [...priorityMatches, ...matches.filter(m => !priorityMatches.includes(m))];
      
      // We calculate shortage after STOCK only, so we can see how much the VENDOR covers
      const bal = Number(r._balance) || 0;
      const stk = Number(r.allocatedQty) || 0;
      const shortageAfterStock = bal - stk;
      
      if (shortageAfterStock > 0) {
        let rem = shortageAfterStock;
        for (const m of sortedMatches) {
          const avail = Number(m.remaining) || 0;
          if (avail <= 0) continue;
          
          const take = Math.min(avail, rem);
          
          // Only populate details if this PO actually provides quantity to this SO line
          if (!r.vpoDetails.soNo && take > 0) {
            r.vpoDetails = {
              soNo: m.soNo,
              soDate: m.soDate,
              lineDate: m.lineDate,
              line: m.line,
              custPO: m.custPONo,
              ordered: m.orderQty,
              open: m.openQty,
              reqDate: m.reqDate,
              mad: m.mad,
              importBy: m.importBy,
              wh: m.whName
            };
          }

          m.remaining = (m.remaining || 0) - take;
          r.vpoAllocated = (r.vpoAllocated || 0) + take;
          rem -= take;
          if (rem <= 0) break;
        }
      }
    });

    // Final Action status
    overdue.forEach(r => {
      const vendorCovered = r.allocatedQty + r.vpoAllocated;
      if (r._balance <= 0) {
        r.action = 'Covered';
      } else if (vendorCovered >= r._balance) {
        r.action = 'Covered';
      } else if (vendorCovered > 0) {
        r.action = 'Partial Qty ordered need to make order';
      } else {
        r.action = 'Make PO need to raise';
      }
    });

    return overdue;
  };

  // ── Generic table renderer (all columns from data) ───────────────────────────
  const isNumericCol = (key) => {
    const kl = String(key).toLowerCase();
    return kl.includes('rate') || kl.includes('value') || kl.includes('qty') ||
      kl.includes('quantity') || kl.includes('balance') || kl.includes('amount') ||
      kl.includes('price') || kl.includes('ordered') || kl.includes('stock') ||
      kl.includes('allocated') || kl.includes('po');
  };
  const isDateCol = (key) => {
    const kl = String(key).toLowerCase();
    return kl.includes('date') || kl.includes('due') || kl.includes('delivery') || kl.includes('mad');
  };
  const fmtCell = (key, val) => {
    if (val === undefined || val === null || val === '') return '-';
    if (isDateCol(key)) { const d = toDate(val); if (d) return fmtDate(d); }
    
    if (key === 'Category') {
      const isDue = val === 'Due';
      const bg = isDue ? '#fdf2f8' : '#f0fdf4';
      const fg = isDue ? '#9d174d' : '#15803d';
      const bd = isDue ? '#fbcfe8' : '#bbf7d0';
      return (
        <span style={{ 
          background: bg, color: fg, padding: '2px 8px', borderRadius: '4px', 
          fontWeight: 'bold', fontSize: '9px', border: `1px solid ${bd}`,
          display: 'inline-block'
        }}>
          {val}
        </span>
      );
    }
    if (key === 'Status') {
      if (val === 'Available') return <strong style={{ color: '#15803d' }}>{val}</strong>;
      if (val === 'Need to Arrange') return <strong style={{ color: '#b91c1c' }}>{val}</strong>;
    }
    if (key === 'Action') {
      if (val === 'Covered') return <strong style={{ color: '#15803d' }}>{val}</strong>;
      if (val === 'Make PO need to raise') return <strong style={{ color: '#b91c1c' }}>{val}</strong>;
      if (val === 'Partial Qty ordered need to make order') return <strong style={{ color: '#9a3412' }}>{val}</strong>;
    }

    const kl = String(key).toLowerCase();
    if ((kl.includes('rate') || kl.includes('value') || kl.includes('amount') || kl.includes('price')) && !isNaN(Number(val)) && val !== '')
      return fmtNum(val, 2);
    return String(val);
  };

  // ── DataTable: Sort + Filter ─────────────────────────────────────────────────
  // columnGroups: [{label, span, bg?, color?}] — renders a group header row above columns
  // rowStyle: (row) => CSSProperties — applies per-row inline styles
  const exportToExcel = (rows, filename = 'Report.xlsx', columnGroups = []) => {
    if (!rows || rows.length === 0) return;
    
    // Create a new workbook and worksheet
    const wb = XLSX.utils.book_new();
    
    // Convert data to sheet
    const ws = XLSX.utils.json_to_sheet(rows);

    // If we have column groups, we can't easily merge cells with json_to_sheet 
    // without complex manipulation. We'll stick to a standard clean export for now
    // which is the "Correct format" users expect.
    
    XLSX.utils.book_append_sheet(wb, ws, "Report");
    
    // Generate and download
    XLSX.writeFile(wb, filename);
  };

  const DataTable = ({ list, upToCol = null, extraCols = [], columnGroups = [], rowStyle = null, columnTotals = null, onExport = null }) => {
    const [sortCol, setSortCol] = React.useState(null);
    const [sortDir, setSortDir] = React.useState('asc');
    const [inputValue, setInputValue] = React.useState('');
    const [filter, setFilter] = React.useState('');
    const [page, setPage] = React.useState(1);
    const pageSize = 50;

    // Reset page on search/sort
    React.useEffect(() => setPage(1), [filter, sortCol, sortDir]);

    // Debounce search input
    React.useEffect(() => {
      const t = setTimeout(() => setFilter(inputValue), 250);
      return () => clearTimeout(t);
    }, [inputValue]);

    if (!list || list.length === 0) return <div className="empty">No data uploaded yet.</div>;

    // Memoize columns to avoid recalculating on every keystroke
    const cols = React.useMemo(() => {
      let c = Object.keys(list[0]);
      if (upToCol) {
        const cut = c.findIndex(x => x.toLowerCase() === upToCol.toLowerCase());
        if (cut !== -1) c = c.slice(0, cut + 1);
      }
      extraCols.forEach(ec => { if (!c.includes(ec.key)) c.push(ec.key); });
      return c;
    }, [list, upToCol, extraCols]);

    // Memoize and optimize filtering
    const filteredRows = React.useMemo(() => {
      const q = filter.trim().toLowerCase();
      if (!q) return list;
      return list.filter(r => {
        for (const c of cols) {
          const val = r[c];
          if (val !== undefined && val !== null && String(val).toLowerCase().includes(q)) return true;
        }
        return false;
      });
    }, [list, filter, cols]);

    // Memoize and optimize sorting
    const rows = React.useMemo(() => {
      if (!sortCol) return filteredRows;
      return [...filteredRows].sort((a, b) => {
        const av = a[sortCol] ?? '', bv = b[sortCol] ?? '';
        const an = Number(av), bn = Number(bv);
        const cmp = !isNaN(an) && !isNaN(bn) ? an - bn : String(av).localeCompare(String(bv));
        return sortDir === 'asc' ? cmp : -cmp;
      });
    }, [filteredRows, sortCol, sortDir]);

    const handleSort = (col) => {
      if (sortCol === col) setSortDir(d => d === 'asc' ? 'desc' : 'asc');
      else { setSortCol(col); setSortDir('asc'); }
    };

    return (
      <div>
        <div className="tbl-toolbar">
          <input className="tbl-search" placeholder="🔍 Search / Filter..." value={inputValue} onChange={e => setInputValue(e.target.value)} />
          {inputValue && <button className="tbl-clear" onClick={() => { setInputValue(''); setFilter(''); }}>✕ Clear</button>}
          <span className="tbl-count">{rows.length} rows{filter ? ` (filtered from ${list.length})` : ''}</span>
          {onExport && (
            <button onClick={() => onExport(rows, 'Master_Report.xlsx', columnGroups)} style={{ marginLeft:'auto', background:'#15803d', color:'#fff', border:'none', padding:'6px 12px', borderRadius:'4px', cursor:'pointer', fontSize:'11px', fontWeight:'bold', display:'flex', alignItems:'center', gap:'6px' }}>
              📊 Export {filter ? `Filtered (${rows.length})` : `All (${rows.length})`} to Excel
            </button>
          )}
        </div>

        {filteredRows.length > pageSize && (
          <div style={{ display:'flex', alignItems:'center', gap:'10px', padding:'10px 0', borderBottom:'1px solid #e2e8f0', marginBottom:'10px' }}>
            <button disabled={page === 1} onClick={() => setPage(p => p - 1)} style={{ padding:'4px 10px', cursor: page === 1 ? 'default' : 'pointer', opacity: page === 1 ? 0.5 : 1 }}>◀ Prev</button>
            <span style={{ fontSize:'11px' }}>Page <strong>{page}</strong> of {Math.ceil(filteredRows.length / pageSize)}</span>
            <button disabled={page >= Math.ceil(filteredRows.length / pageSize)} onClick={() => setPage(p => p + 1)} style={{ padding:'4px 10px', cursor: page >= Math.ceil(filteredRows.length / pageSize) ? 'default' : 'pointer', opacity: page >= Math.ceil(filteredRows.length / pageSize) ? 0.5 : 1 }}>Next ▶</button>
          </div>
        )}

        <div className="table-container">
          <table>
            <thead>
              {columnGroups.length > 0 && (
                <tr>
                  {columnGroups.map((g, i) => (
                    <th key={i} colSpan={g.span}
                      style={{ background: g.bg || '#1e40af', color: g.color || '#fff', textAlign: 'center', fontSize: '10px', padding: '3px 6px' }}>
                      {g.label}
                    </th>
                  ))}
                </tr>
              )}
              <tr>
                {cols.map(c => (
                  <th key={c}
                    style={{ textAlign: isNumericCol(c) ? 'right' : 'left', cursor: 'pointer', userSelect: 'none' }}
                    onClick={() => handleSort(c)}>
                    <div style={{ display:'flex', flexDirection:'column', alignItems: isNumericCol(c) ? 'flex-end' : 'flex-start' }}>
                      <div>{c}{sortCol === c ? (sortDir === 'asc' ? ' ▲' : ' ▼') : ' ⇅'}</div>
                      {columnTotals && columnTotals[c] !== undefined && (
                        <div style={{ fontSize: '9px', color: '#1e40af', marginTop: '2px', fontWeight: '800' }}>
                          {isNumericCol(c) ? (c.toLowerCase().includes('value') || c.toLowerCase().includes('amount') ? fmtNum(columnTotals[c], 2) : fmtNum(columnTotals[c])) : ''}
                        </div>
                      )}
                    </div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.slice((page - 1) * pageSize, page * pageSize).map((r, i) => (
                <tr key={i} style={rowStyle ? rowStyle(r) : {}}>
                  {cols.map(c => {
                    const ec = extraCols.find(x => x.key === c);
                    const val = ec ? ec.render(r) : r[c];
                    return (
                      <td key={c} style={{ textAlign: isNumericCol(c) ? 'right' : 'left' }}>
                        {ec ? val : fmtCell(c, val)}
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    );
  };

  const renderTable = (list, upToCol = null) => <DataTable list={list} upToCol={upToCol} />;

  // ── Master Report renderer ───────────────────────────────────────────────────
  const renderMaster = () => {
    const rows = React.useMemo(() => buildMasterReport(), [data]);
    if (!rows.length) return <div className="empty">Upload Sales Order and Stock files to generate the Master Report.</div>;

    // Calculate totals for summary and column headers
    const totals = {
      ordered: 0, balance: 0, value: 0, stockQty: 0, allocatedQty: 0,
      poOrdered: 0, poBalance: 0, poAllocated: 0,
      vpoOrdered: 0, vpoOpen: 0, vpoAllocated: 0,
      due: 0, scheduled: 0
    };

    rows.forEach(r => {
      totals.ordered += (Number(r.ordered) || 0);
      totals.balance += (Number(r.balance) || 0);
      totals.value += (Number(r.value) || 0);
      totals.stockQty += (Number(r.stockQty) || 0);
      totals.allocatedQty += (Number(r.allocatedQty) || 0);
      if (r.poDetails) {
        totals.poOrdered += (Number(r.poDetails.ordered) || 0);
        totals.poBalance += (Number(r.poDetails.balance) || 0);
      }
      totals.poAllocated += (Number(r.poAllocated) || 0);
      if (r.vpoDetails) {
        totals.vpoOrdered += (Number(r.vpoDetails.ordered) || 0);
        totals.vpoOpen += (Number(r.vpoDetails.open) || 0);
      }
      totals.vpoAllocated += (Number(r.vpoAllocated) || 0);
      if (r._category === 'Due Order') totals.due++;
      else totals.scheduled++;
    });

    const display = rows.map(r => ({
      'Category':      r._category === 'Due Order' ? 'Due' : 'Scheduled',
      'Date':          fmtDate(r.date),
      'Order No':      r.order,
      'Party Name':    r.partyName,
      'Name of Item':  r.nameOfItem,
      'Part No':       r.partNo,
      'Ordered':       r.ordered,
      'Balance':       r.balance,
      'Value':         r.value,
      'Due on':        fmtDate(r.dueOn),
      'Stock Qty':     r.stockQty,
      'Allocated Qty': r.allocatedQty || 0,
      'Status':        r.stockStatus,
      'PO Date':       fmtDate(r?.poDetails?.date),
      'PO Order':      r?.poDetails?.order || '-',
      'PO Party':      r?.poDetails?.party || '-',
      'PO Ordered':    r?.poDetails?.ordered || 0,
      'PO Balance':    r?.poDetails?.balance || 0,
      'PO Allocated':  r.poAllocated || 0,
      'PO Due on':     fmtDate(r?.poDetails?.dueOn),
      'VPO SO No':     r?.vpoDetails?.soNo || '-',
      'VPO Date':      fmtDate(r?.vpoDetails?.soDate),
      'VPO Line':      r?.vpoDetails?.line || '-',
      'VPO Cust PO':   r?.vpoDetails?.custPO || '-',
      'VPO Ordered':   r?.vpoDetails?.ordered || 0,
      'VPO Open':      r?.vpoDetails?.open || 0,
      'VPO Allocated': r.vpoAllocated || 0,
      'Cust Req Date': fmtDate(r?.vpoDetails?.reqDate),
      'Estimated M.A.D.': fmtDate(r?.vpoDetails?.mad),
      'VPO Import':    r?.vpoDetails?.importBy || '-',
      'VPO WH':        r?.vpoDetails?.wh || '-',
      'Action':        r.action
    }));

    const colTotalsMap = {
      'Ordered': totals.ordered,
      'Balance': totals.balance,
      'Value': totals.value,
      'Stock Qty': totals.stockQty,
      'Allocated Qty': totals.allocatedQty,
      'PO Ordered': totals.poOrdered,
      'PO Balance': totals.poBalance,
      'PO Allocated': totals.poAllocated,
      'VPO Ordered': totals.vpoOrdered,
      'VPO Open': totals.vpoOpen,
      'VPO Allocated': totals.vpoAllocated
    };

    const StatCard = ({ label, value, subValue, color }) => (
      <div style={{ 
        background: '#fff', padding: '10px 15px', borderRadius: '8px', 
        border: `1px solid ${color}22`, borderLeft: `4px solid ${color}`,
        boxShadow: '0 1px 3px rgba(0,0,0,0.05)', display: 'flex', flexDirection: 'column', gap: '2px', minWidth: '130px'
      }}>
        <span style={{ fontSize: '9px', color: '#64748b', fontWeight: '700', textTransform: 'uppercase', letterSpacing: '0.5px' }}>{label}</span>
        <span style={{ fontSize: '15px', color: '#0f172a', fontWeight: '800' }}>{value}</span>
        {subValue && <span style={{ fontSize: '10px', color: color, fontWeight: '600' }}>{subValue}</span>}
      </div>
    );

    return (
      <div className="card animate-in">
        <div style={{ display: 'flex', gap: '12px', flexWrap: 'wrap', marginBottom: '20px', padding: '4px' }}>
          <StatCard label="Total Rows" value={fmtNum(rows.length)} color="#64748b" />
          <StatCard label="Due Orders" value={fmtNum(totals.due)} color="#e11d48" />
          <StatCard label="Scheduled" value={fmtNum(totals.scheduled)} color="#10b981" />
          <StatCard label="Total Ordered" value={fmtNum(totals.ordered)} subValue={`Value: ${fmtNum(totals.value, 2)}`} color="#2563eb" />
          <StatCard label="Total Balance" value={fmtNum(totals.balance)} color="#f59e0b" />
          <StatCard label="Stock Covered" value={fmtNum(totals.allocatedQty)} color="#0ea5e9" />
          <StatCard label="VPO Covered" value={fmtNum(totals.vpoAllocated)} color="#ec4899" />
        </div>

        <DataTable 
          list={display} 
          columnTotals={colTotalsMap}
          onExport={exportToExcel}
          columnGroups={[
            { label: 'Sales Order Details', span: 10, bg: '#1e40af' },
            { label: 'Stock Status', span: 3, bg: '#0369a1' },
            { label: 'Pending PO Details', span: 7, bg: '#b45309' },
            { label: 'Vendor PO Details', span: 11, bg: '#7e22ce' },
            { label: 'Action Required', span: 1, bg: '#111827' }
          ]}
          rowStyle={(r) => {
            if (r['Category'] === 'Due' || r['Category'] === 'Due Order') return { background: '#fff1f2' };
            return {};
          }}
        />
      </div>
    );
  };

  // ── Render ───────────────────────────────────────────────────────────────────
  const renderSOTab = () => {
    if (!data.so.length) return <div className="empty">No data uploaded yet.</div>;
    const TODAY = new Date(); TODAY.setHours(0, 0, 0, 0);
    const allCols = Object.keys(data.so[0]);
    const dueDateCol = allCols.find(k => k.toLowerCase().includes('due')) ||
                       allCols.find(k => k.toLowerCase().includes('delivery')) || null;
    const overdue = dueDateCol
      ? data.so.filter(r => { const d = toDate(r[dueDateCol]); return d && d <= TODAY; })
      : data.so;
    return (
      <div>
        <div className="tbl-info">Showing <strong>{overdue.length}</strong> overdue rows out of {data.so.length} total</div>
        {renderTable(overdue, dueDateCol)}
      </div>
    );
  };

  const renderPOTab = () => {
    if (!data.po.length) return <div className="empty">No data uploaded yet.</div>;
    const TODAY = new Date(); TODAY.setHours(0, 0, 0, 0);
    const allCols = Object.keys(data.po[0]);
    const dueDateCol = allCols.find(k => k.toLowerCase().includes('due')) ||
                       allCols.find(k => k.toLowerCase().includes('delivery')) || null;
    const overdue = dueDateCol
      ? data.po.filter(r => { const d = toDate(r[dueDateCol]); return d && d <= TODAY; })
      : data.po;
    return (
      <div>
        <div className="tbl-info">Showing <strong>{overdue.length}</strong> overdue rows out of {data.po.length} total</div>
        {renderTable(overdue)}
      </div>
    );
  };

  return (
    <div className="app">
      <header className="header">
        <div style={{ display:'flex', alignItems:'center', gap:'10px' }}>
          <div className="header-logo">SKC</div>
          <div className="header-sub">Material Allocation Viewer</div>
        </div>
        <button onClick={clearAllData} style={{ marginLeft:'auto', background:'#fee2e2', color:'#b91c1c', border:'1px solid #fecaca', padding:'4px 10px', borderRadius:'4px', cursor:'pointer', fontSize:'9px', fontWeight:'bold', textTransform:'uppercase' }}>
          🗑️ Clear All Data
        </button>
      </header>

      <div className="upload-section">
        {[
          { label: '📦 Stock File',         type: 'stock' },
          { label: '📋 Pending SO File',    type: 'so'    },
          { label: '🛒 Pending PO File',    type: 'po'    },
          { label: '📂 Open Vendor Orders', type: 'oo'    },
        ].map(({ label, type }) => (
          <div key={type} className="upload-card">
            <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:'4px' }}>
              <label style={{ margin:0 }}>{label}</label>
              <button onClick={() => downloadTemplate(type)} style={{ background:'none', border:'none', color:'#2563eb', fontSize:'9px', cursor:'pointer', fontWeight:'bold', textDecoration:'underline', padding:0 }}>
                📥 Download Template
              </button>
            </div>
            <input type="file" accept=".xlsx,.xls" onChange={e => handleUpload(e, type)} />
            {data[type].length > 0 && <span className="badge">{data[type].length} rows</span>}
          </div>
        ))}
      </div>

      <div className="tabs">
        {[
          { id: 'master', label: '📊 Master Report', count: null },
          { id: 'stock',  label: 'Stock',            count: data.stock.length },
          { id: 'so',     label: 'Pending SO',       count: data.so.length },
          { id: 'po',     label: 'Pending PO',       count: data.po.length },
          { id: 'oo',     label: 'Open Vendor Orders', count: data.oo.length },
        ].map(({ id, label, count }) => (
          <button key={id} className={activeTab === id ? 'active' : ''} onClick={() => setActiveTab(id)}>
            {label}{count ? ` (${count})` : ''}
          </button>
        ))}
      </div>

      <main className="content">
        {activeTab === 'master' && renderMaster()}
        {activeTab === 'stock'  && (() => {
          const totalQty = data.stock.reduce((s, r) => s + (Number(r['Quantity'] || r['Qty'] || r['Closing Stock'] || 0)), 0);
          const totalVal = data.stock.reduce((s, r) => s + (Number(r['Value'] || r['Amount'] || 0)), 0);
          return (
            <div>
              {data.stock.length > 0 && (
                <div className="stock-totals">
                  <span>Total Items: <strong>{data.stock.length}</strong></span>
                  <span>Total Qty: <strong>{fmtNum(totalQty)}</strong></span>
                  <span>Total Value: <strong>{fmtNum(totalVal, 2)}</strong></span>
                </div>
              )}
              {renderTable(data.stock)}
            </div>
          );
        })()}
        {activeTab === 'so'     && renderSOTab()}
        {activeTab === 'po'     && renderPOTab()}
        {activeTab === 'oo'     && renderTable(data.oo)}
      </main>

      <style>{css}</style>
    </div>
  );
};

const css = `
  :root {
    --bg: #f8fafc; --surface: #fff; --border: #cbd5e1;
    --accent: #1e40af; --text: #0f172a; --muted: #475569;
    --font: 'Cambria', 'Georgia', serif;
  }
  * { box-sizing: border-box; }
  body { margin: 0; font-family: var(--font); background: var(--bg); color: var(--text); font-size: 10px; }

  .header { background: #fff; padding: 8px 16px; border-bottom: 1px solid var(--border); display: flex; align-items: center; gap: 10px; }
  .header-logo { font-size: 20px; font-weight: bold; color: var(--accent); font-family: var(--font); }
  .header-sub { font-size: 9px; color: var(--muted); text-transform: uppercase; font-weight: bold; }

  .upload-section { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; padding: 10px 16px; }
  .upload-card { background: #fff; padding: 8px 10px; border: 1px solid var(--border); border-radius: 4px; }
  .upload-card label { display: block; margin-bottom: 4px; font-weight: bold; font-size: 10px; }
  .upload-card input { font-size: 10px; font-family: var(--font); width: 100%; }
  .badge { display: inline-block; margin-top: 3px; background: #dbeafe; color: #1e40af; border-radius: 3px; padding: 1px 5px; font-size: 9px; }

  .tabs { display: flex; background: #fff; border-bottom: 1px solid var(--border); padding: 0 16px; }
  .tabs button { padding: 6px 14px; border: none; background: none; cursor: pointer; font-weight: bold; color: var(--muted); border-bottom: 2px solid transparent; font-family: var(--font); font-size: 10px; }
  .tabs button.active { color: var(--accent); border-bottom-color: var(--accent); }

  .content { padding: 8px 16px; }
  .empty { padding: 20px; color: var(--muted); font-size: 10px; }
  .tbl-info { padding: 3px 0 5px; font-size: 10px; color: var(--muted); }
  .tbl-toolbar { display: flex; align-items: center; gap: 8px; padding: 5px 0; }
  .tbl-search { border: 1px solid var(--border); border-radius: 3px; padding: 3px 8px; font-size: 10px; font-family: var(--font); width: 220px; outline: none; }
  .tbl-search:focus { border-color: var(--accent); }
  .tbl-clear { border: 1px solid var(--border); background: #fff; border-radius: 3px; padding: 3px 8px; font-size: 10px; cursor: pointer; color: var(--muted); font-family: var(--font); }
  .tbl-clear:hover { background: #fee2e2; color: #b91c1c; border-color: #b91c1c; }
  .tbl-count { font-size: 9px; color: var(--muted); }
  th:hover { background: #e2e8f0; }

  .table-container { background: #fff; border: 1px solid var(--border); border-radius: 4px; overflow: auto; max-height: calc(100vh - 200px); }
  table { width: 100%; border-collapse: collapse; font-family: var(--font); font-size: 10px; }
  th { background: #f1f5f9; padding: 4px 6px; border-bottom: 1px solid var(--border); border-right: 1px solid var(--border); position: sticky; top: 0; white-space: nowrap; font-size: 10px; z-index: 2; }
  td { padding: 3px 6px; border-bottom: 1px solid #f1f5f9; border-right: 1px solid #e2e8f0; white-space: nowrap; font-size: 10px; }
  tr:hover td { background: #f8fafc; }

  .master-summary { display: flex; gap: 20px; padding: 5px 2px 7px; font-size: 10px; color: var(--muted); flex-wrap: wrap; }
  .master-summary span strong { color: var(--text); }
  .stock-totals { display: flex; gap: 24px; padding: 5px 2px 7px; font-size: 10px; color: var(--muted); background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 4px; padding: 5px 12px; margin-bottom: 6px; }
  .stock-totals strong { color: #15803d; }

  .badge-due { background: #fee2e2; color: #b91c1c; border-radius: 3px; padding: 1px 5px; font-size: 9px; white-space: nowrap; }
  .badge-sch { background: #dbeafe; color: #1e40af; border-radius: 3px; padding: 1px 5px; font-size: 9px; white-space: nowrap; }
`;

export default function WrappedApp() {
  return (
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  );
}
