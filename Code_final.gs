// ============================================================
// Recykal Import Lead Tracker — Google Apps Script Backend
// ============================================================

const SPREADSHEET_ID = '1rKJWsvOpSGkjO6uapsorNFv-WShe37-krDxIKJ6Cmzc';
const ALLOWED_DOMAIN = 'recykal.com';

function verifyToken(token) {
  if (!token) return null;
  try {
    const parts = token.split('.');
    if (parts.length !== 3) return null;
    const payload = JSON.parse(
      Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[1] + '==')).getDataAsString()
    );
    if (!payload.email) return null;
    if (!payload.email.endsWith('@' + ALLOWED_DOMAIN)) return null;
    if (payload.exp < Math.floor(Date.now() / 1000)) return null;
    return payload.email;
  } catch(e) {
    return null;
  }
}

function getSheet(name) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i] === '' ? null : row[i]; });
      return obj;
    })
    .filter(r => r.id);
}

function getAll(data) {
  const buyers = getLeads({type: 'buyer'}).leads || [];
  const suppliers = getLeads({type: 'supplier'}).leads || [];
  const rfqs = getAllRFQs({}).rfqs || [];
  const orders = getOrders({}).orders || [];
  return { buyers, suppliers, rfqs, orders };
}

function doGet(e) {
  let result;
  try {
    const email = verifyToken(e.parameter.token);
    if (!email) {
      result = { error: 'Unauthorized' };
    } else {
      const action = e.parameter.action;
      const raw = e.parameter.data;
      const data = raw ? JSON.parse(raw) : {};
      switch (action) {
        case 'getLeads':                result = getLeads(data);                break;
        case 'addLead':                 result = addLead(data);                 break;
        case 'updateLead':              result = updateLead(data);              break;
        case 'getRFQs':                 result = getRFQs(data);                 break;
        case 'getAllRFQs':              result = getAllRFQs(data);              break;
        case 'addRFQ':                  result = addRFQ(data);                  break;
        case 'updateRFQ':               result = updateRFQ(data);               break;
        case 'getLeadRemarks':          result = getLeadRemarks(data);          break;
        case 'addLeadRemark':           result = addLeadRemark(data);           break;
        case 'getRFQRemarks':           result = getRFQRemarks(data);           break;
        case 'addRFQRemark':            result = addRFQRemark(data);            break;
        case 'getAll':                  result = getAll(data);                  break;
        case 'getOrders':               result = getOrders(data);               break;
        case 'addOrder':                result = addOrder(data);                break;
        case 'updateOrder':             result = updateOrder(data);             break;
        case 'getOrderPayments':        result = getOrderPayments(data);        break;
        case 'addOrderPayment':         result = addOrderPayment(data);         break;
        case 'getOrderRemarks':         result = getOrderRemarks(data);         break;
        case 'addOrderRemark':          result = addOrderRemark(data);          break;
        case 'getImportTransactions':   result = getImportTransactions(data);   break;
        case 'addImportTransaction':    result = addImportTransaction(data);    break;
        case 'updateImportTransaction': result = updateImportTransaction(data); break;
        case 'getImportDocs':           result = getImportDocs(data);           break;
        case 'updateImportDocs':        result = updateImportDocs(data);        break;
        case 'getImportPayments':       result = getImportPayments(data);       break;
        case 'addImportPayment':        result = addImportPayment(data);        break;
        case 'updateImportPayment':     result = updateImportPayment(data);     break;
        case 'getImportLineItems':      result = getImportLineItems(data);      break;
        case 'addImportLineItem':       result = addImportLineItem(data);       break;
        case 'deleteImportLineItem':    result = deleteImportLineItem(data);    break;
        case 'getImportCosts':          result = getImportCosts(data);          break;
        case 'addImportCost':           result = addImportCost(data);           break;
        case 'deleteImportCost':        result = deleteImportCost(data);        break;
        case 'getImportTimeline':       result = getImportTimeline(data);       break;
        case 'updateImportTimeline':    result = updateImportTimeline(data);    break;
        case 'addImportRemark':         result = addImportRemark(data);         break;
        case 'getImportRemarks':        result = getImportRemarks(data);        break;
        case 'getImportDocReviews':     result = getImportDocReviews(data);     break;
        case 'addImportDocReview':      result = addImportDocReview(data);      break;
        default: result = { error: 'Unknown action: ' + action };
      }
    }
  } catch (err) {
    result = { error: err.toString() };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- LEADS ----

function getLeads(params) {
  const leads = sheetToObjects(getSheet('Leads'));
  const filtered = params.type ? leads.filter(l => l.type === params.type) : leads;
  return { leads: filtered };
}

function addLead(data) {
  const sheet = getSheet('Leads');
  const all = sheetToObjects(sheet);
  const dup = all.find(l =>
    (l.company_name || '').toLowerCase() === (data.company_name || '').toLowerCase() &&
    (l.country || '').toLowerCase() === (data.country || '').toLowerCase() &&
    l.type === data.type
  );
  if (dup) return { error: 'DUPLICATE', existing: dup };

  const id = Utilities.getUuid();
  const now = new Date().toISOString();
  sheet.appendRow([
    id, data.type, data.company_name, data.country, data.location || '',
    data.stage || 'Prospect', data.assigned_to || '', data.poc_name || '',
    data.phone || '', data.email || '', data.materials || '', data.notes || '',
    now, data.by, now, data.by
  ]);
  return { success: true, id };
}

function updateLead(data) {
  const sheet = getSheet('Leads');
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('id');

  for (let i = 1; i < values.length; i++) {
    if (values[i][idCol] === data.id) {
      const map = {
        company_name: data.company_name, country: data.country,
        location: data.location, stage: data.stage, assigned_to: data.assigned_to,
        poc_name: data.poc_name, phone: data.phone, email: data.email,
        materials: data.materials, notes: data.notes,
        updated_at: new Date().toISOString(), updated_by: data.by,
      };
      headers.forEach((h, col) => {
        if (h in map && map[h] !== undefined)
          sheet.getRange(i + 1, col + 1).setValue(map[h] ?? '');
      });
      return { success: true };
    }
  }
  return { error: 'Lead not found' };
}

// ---- RFQs ----

function getAllRFQs(params) {
  params = params || {};
  const rfqs = sheetToObjects(getSheet('RFQs'));
  const leads = sheetToObjects(getSheet('Leads'));
  const leadMap = {};
  leads.forEach(function(l) { leadMap[l.id] = l; });
  let filtered = rfqs;
  if (params.status) filtered = filtered.filter(function(r) { return r.status === params.status; });
  if (params.material) filtered = filtered.filter(function(r) { return (r.material||'').toLowerCase().indexOf(params.material.toLowerCase()) !== -1; });
  if (params.assigned_to) filtered = filtered.filter(function(r) { return r.assigned_to === params.assigned_to; });
  if (params.lead_type) filtered = filtered.filter(function(r) { return (leadMap[r.lead_id]||{}).type === params.lead_type; });
  filtered = filtered.map(function(r) {
    var lead = leadMap[r.lead_id] || {};
    return Object.assign({}, r, { company_name: lead.company_name||'', country: lead.country||'', lead_type: lead.type||'' });
  });
  return { rfqs: filtered };
}

function getRFQs(params) {
  const rfqs = sheetToObjects(getSheet('RFQs'));
  return { rfqs: rfqs.filter(r => r.lead_id === params.lead_id) };
}

function addRFQ(data) {
  const sheet = getSheet('RFQs');
  const id = Utilities.getUuid();
  const now = new Date().toISOString();
  sheet.appendRow([
    id, data.lead_id, data.material || '', data.qty || '', data.unit || 'MT',
    data.attachment_pct || '', data.expected_price || '', data.deal_closure_price || '',
    data.currency || 'USD', data.status || 'RFQ', data.not_converted_reason || '',
    data.notes || '', data.assigned_to || '', now, data.by, now
  ]);
  return { success: true, id };
}

function updateRFQ(data) {
  const sheet = getSheet('RFQs');
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('id');

  for (let i = 1; i < values.length; i++) {
    if (values[i][idCol] === data.id) {
      const map = {
        material: data.material, qty: data.qty, unit: data.unit,
        attachment_pct: data.attachment_pct, expected_price: data.expected_price,
        deal_closure_price: data.deal_closure_price, currency: data.currency,
        status: data.status, not_converted_reason: data.not_converted_reason,
        notes: data.notes, assigned_to: data.assigned_to,
        buyer_lead_id: data.buyer_lead_id, buyer_name: data.buyer_name,
        supplier_lead_id: data.supplier_lead_id, supplier_name: data.supplier_name,
        incoterms: data.incoterms, payment_terms: data.payment_terms,
        buyer_price: data.buyer_price, supplier_price: data.supplier_price,
        costs: data.costs,
        updated_at: new Date().toISOString(),
      };
      headers.forEach((h, col) => {
        if (h in map && map[h] !== undefined)
          sheet.getRange(i + 1, col + 1).setValue(map[h] ?? '');
      });
      return { success: true };
    }
  }
  return { error: 'RFQ not found' };
}

// ---- REMARKS ----

function getLeadRemarks(params) {
  const all = sheetToObjects(getSheet('Lead_Remarks'));
  const remarks = all.filter(r => r.lead_id === params.lead_id).reverse();
  return { remarks };
}

function addLeadRemark(data) {
  const sheet = getSheet('Lead_Remarks');
  const id = Utilities.getUuid();
  const now = new Date().toISOString();
  sheet.appendRow([id, data.lead_id, data.remark, data.by, now]);
  return { success: true, id, created_at: now };
}

function getRFQRemarks(params) {
  const all = sheetToObjects(getSheet('RFQ_Remarks'));
  const remarks = all.filter(r => r.rfq_id === params.rfq_id).reverse();
  return { remarks };
}

function addRFQRemark(data) {
  const sheet = getSheet('RFQ_Remarks');
  const id = Utilities.getUuid();
  const now = new Date().toISOString();
  sheet.appendRow([id, data.rfq_id, data.remark, data.by, now]);
  return { success: true, id, created_at: now };
}

// ---- ORDERS ----

function getOrders(params) {
  params = params || {};
  const orders = sheetToObjects(getSheet('Orders'));
  const leads = sheetToObjects(getSheet('Leads'));
  const leadMap = {};
  leads.forEach(function(l) { leadMap[l.id] = l; });
  let filtered = orders;
  if (params.status) filtered = filtered.filter(function(o) { return o.status === params.status; });
  filtered = filtered.map(function(o) {
    const buyer = leadMap[o.buyer_lead_id] || {};
    const supplier = leadMap[o.supplier_lead_id] || {};
    return Object.assign({}, o, {
      buyer_company: buyer.company_name || o.buyer_name || '',
      supplier_company: supplier.company_name || o.supplier_name || '',
      buyer_country: buyer.country || '',
      supplier_country: supplier.country || '',
    });
  });
  return { orders: filtered };
}

function addOrder(data) {
  const sheet = getSheet('Orders');
  const id = Utilities.getUuid();
  const now = new Date().toISOString();
  const orderRef = data.order_ref || ('ORD-' + now.slice(0,10).replace(/-/g,'') + '-' + id.slice(0,4).toUpperCase());
  sheet.appendRow([
    id, orderRef,
    data.buyer_lead_id || '', data.buyer_name || '',
    data.supplier_lead_id || '', data.supplier_name || '',
    data.material || '', data.qty || '', data.unit || 'MT',
    data.currency || 'USD',
    data.buyer_price || '', data.supplier_price || '',
    data.incoterms || '', data.payment_terms || '',
    data.status || 'Confirmed',
    data.container_number || '', data.bl_number || '',
    data.vessel_name || '', data.loading_port || '', data.discharge_port || '',
    data.etd || '', data.eta || '',
    data.notes || '',
    data.line_items || '',
    now, data.by, now, data.by
  ]);
  return { success: true, id, order_ref: orderRef };
}

function updateOrder(data) {
  const sheet = getSheet('Orders');
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('id');

  for (let i = 1; i < values.length; i++) {
    if (values[i][idCol] === data.id) {
      const map = {
        order_ref: data.order_ref,
        buyer_lead_id: data.buyer_lead_id, buyer_name: data.buyer_name,
        supplier_lead_id: data.supplier_lead_id, supplier_name: data.supplier_name,
        material: data.material, qty: data.qty, unit: data.unit,
        currency: data.currency,
        buyer_price: data.buyer_price, supplier_price: data.supplier_price,
        incoterms: data.incoterms, payment_terms: data.payment_terms,
        costs: data.costs,
        line_items: data.line_items,
        status: data.status,
        container_number: data.container_number, bl_number: data.bl_number,
        vessel_name: data.vessel_name, loading_port: data.loading_port,
        discharge_port: data.discharge_port,
        etd: data.etd, eta: data.eta,
        notes: data.notes,
        updated_at: new Date().toISOString(), updated_by: data.by,
      };
      headers.forEach((h, col) => {
        if (h in map && map[h] !== undefined)
          sheet.getRange(i + 1, col + 1).setValue(map[h] ?? '');
      });
      return { success: true };
    }
  }
  return { error: 'Order not found' };
}

function getOrderPayments(params) {
  const all = sheetToObjects(getSheet('Order_Payments'));
  const payments = all.filter(function(p) { return p.order_id === params.order_id; });
  return { payments: payments };
}

function addOrderPayment(data) {
  const sheet = getSheet('Order_Payments');
  const id = Utilities.getUuid();
  const now = new Date().toISOString();
  sheet.appendRow([
    id, data.order_id, data.type || '', data.amount || 0,
    data.currency || 'USD', data.payment_date || now.slice(0,10),
    data.reference || '', data.notes || '', data.by, now
  ]);
  return { success: true, id, created_at: now };
}

function getOrderRemarks(params) {
  const all = sheetToObjects(getSheet('Order_Remarks'));
  const remarks = all.filter(function(r) { return r.order_id === params.order_id; }).reverse();
  return { remarks: remarks };
}

function addOrderRemark(data) {
  const sheet = getSheet('Order_Remarks');
  const id = Utilities.getUuid();
  const now = new Date().toISOString();
  sheet.appendRow([id, data.order_id, data.remark, data.by, now]);
  return { success: true, id, created_at: now };
}

// ---- SETUP HELPER ----
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  function ensureSheet(name, headers) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getLastRow() === 0) sheet.appendRow(headers);
    return sheet;
  }

  ensureSheet('Leads', [
    'id','type','company_name','country','location','stage','assigned_to',
    'poc_name','phone','email','materials','notes',
    'created_at','created_by','updated_at','updated_by'
  ]);
  ensureSheet('RFQs', [
    'id','lead_id','material','qty','unit','attachment_pct',
    'expected_price','deal_closure_price','currency','status',
    'not_converted_reason','notes','assigned_to',
    'buyer_lead_id','buyer_name','supplier_lead_id','supplier_name',
    'buyer_price','supplier_price','incoterms','payment_terms','costs',
    'created_at','created_by','updated_at'
  ]);
  ensureSheet('Lead_Remarks', ['id','lead_id','remark','added_by','created_at']);
  ensureSheet('RFQ_Remarks',  ['id','rfq_id','remark','added_by','created_at']);
  ensureSheet('Orders', [
    'id','order_ref',
    'buyer_lead_id','buyer_name',
    'supplier_lead_id','supplier_name',
    'material','qty','unit','currency',
    'buyer_price','supplier_price',
    'incoterms','payment_terms','costs','status',
    'container_number','bl_number','vessel_name',
    'loading_port','discharge_port','etd','eta',
    'notes','line_items','created_at','created_by','updated_at','updated_by'
  ]);
  ensureSheet('Order_Payments', [
    'id','order_id','type','amount','currency',
    'payment_date','reference','notes','recorded_by','created_at'
  ]);
  ensureSheet('Order_Remarks', ['id','order_id','remark','added_by','created_at']);

  Logger.log('Setup complete.');
}

function addRFQColumns() {
  const sheet = getSheet('RFQs');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const needed = ['buyer_lead_id','buyer_name','supplier_lead_id','supplier_name','buyer_price','supplier_price','incoterms','payment_terms','costs'];
  needed.forEach(function(col) {
    if (!headers.includes(col)) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(col);
      Logger.log('Added column: ' + col);
    }
  });
  Logger.log('Done.');
}

// Run this once to add line_items column to existing Orders sheet
function addOrderLineItemsColumn() {
  const sheet = getSheet('Orders');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const needed = ['line_items'];
  needed.forEach(function(col) {
    if (!headers.includes(col)) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(col);
      Logger.log('Added column: ' + col);
    }
  });
  Logger.log('Done.');
}

// ============================================================
// IMPORT TRANSACTION TRACKER
// ============================================================

function setupImportSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = {
    ImportTransactions: [
      'id','txn_ref','order_ref','buy_direct','supplier','supplier_country',
      'intermediary','intermediary_price','buy_price','buy_currency','buy_payment_terms',
      'buyer','buyer_country','sell_price','sell_currency','sell_payment_terms',
      'material','qty','unit','incoterms','vessel','container','bl',
      'port_loading','port_discharge','etd','eta','stage','rag','assigned_to',
      'notes','docs_total','docs_approved','created_by','created_at','updated_at'
    ],
    ImportDocs: [
      'id','txn_id','name','status','reviewer','party',
      'expected_date','received_date','notes','updated_at'
    ],
    ImportPayments: [
      'id','txn_id','leg','type','amount','currency','conversion_rate','party',
      'due_date','paid_date','reference','notes','created_by','created_at'
    ],
    ImportCosts: [
      'id','txn_id','type','amount','currency','notes','created_by','created_at'
    ],
    ImportTimeline: [
      'id','txn_id','stage_data','updated_by','updated_at'
    ],
    ImportRemarks: [
      'id','txn_id','remark','added_by','created_at'
    ],
    ImportDocReviews: [
      'id','txn_id','doc_name','action','round','notes','actor','created_at'
    ],
    ImportLineItems: [
      'id','txn_id','material','hsn','qty','unit','price','currency','amount','created_by','created_at'
    ]
  };
  Object.keys(sheets).forEach(function(name) {
    var headers = sheets[name];
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold');
    }
  });
  return { success: true, message: 'Import sheets created' };
}

function getImportSheet(name) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
}

function impHeaders(sheetName) {
  var sh = getImportSheet(sheetName);
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
}

function impRows(sheetName) {
  var sh = getImportSheet(sheetName);
  if (sh.getLastRow() < 2) return [];
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  return sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues().map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i] !== undefined && row[i] !== null ? String(row[i]) : ''; });
    return obj;
  });
}

function impUid() {
  return 'IMP' + Date.now() + Math.random().toString(36).slice(2,6).toUpperCase();
}

function impNow() {
  return new Date().toISOString();
}

function impAppend(sheetName, headers, obj) {
  var sh = getImportSheet(sheetName);
  sh.appendRow(headers.map(function(h) { return obj[h] !== undefined ? obj[h] : ''; }));
}

function impUpdate(sheetName, id, updates) {
  var sh = getImportSheet(sheetName);
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (sh.getLastRow() < 2) return false;
  var data = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
  var idIdx = headers.indexOf('id');
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(id)) {
      headers.forEach(function(h, j) {
        if (updates[h] !== undefined) sh.getRange(i+2, j+1).setValue(updates[h]);
      });
      return true;
    }
  }
  return false;
}

function getImportTransactions(data) {
  return { transactions: impRows('ImportTransactions') };
}

function addImportTransaction(data) {
  var id = impUid();
  var now = impNow();
  var txn_ref = data.txn_ref || ('IMP-' + new Date().getFullYear() + '-' + String(Math.floor(Math.random()*9000)+1000));
  var headers = impHeaders('ImportTransactions');
  var row = {};
  headers.forEach(function(h) { row[h] = data[h] || ''; });
  row.id = id;
  row.txn_ref = txn_ref;
  row.docs_total = 0;
  row.docs_approved = 0;
  row.created_by = data.by || '';
  row.created_at = now;
  row.updated_at = now;
  impAppend('ImportTransactions', headers, row);
  initImportDocs(id, now);
  initImportTimeline(id, now, data.by || '');
  return { success: true, id: id, txn_ref: txn_ref };
}

function initImportDocs(txn_id, now) {
  var DOC_TYPES = ['Sales Contract','Invoice','OBL','Draft BL','PSIC','Packing List','COO (Certificate of Origin)','Form 9','Form 6','Freight Certificate','Insurance','LC','Shipping Instruction','Payment(s)','Loading Images'];
  var headers = impHeaders('ImportDocs');
  DOC_TYPES.forEach(function(name) {
    impAppend('ImportDocs', headers, {
      id: impUid(), txn_id: txn_id, name: name, status: 'Pending',
      reviewer: '', party: '', expected_date: '', received_date: '', notes: '', updated_at: now
    });
  });
}

function initImportTimeline(txn_id, now, by) {
  var STAGES = ['Confirmation','Loading','Shipping','Reaches India','Clearance'];
  var timeline = STAGES.map(function(stage) {
    return { stage: stage, expected_date: '', actual_date: '', responsible: '', delay_reason: '', notes: '' };
  });
  var headers = impHeaders('ImportTimeline');
  impAppend('ImportTimeline', headers, {
    id: impUid(), txn_id: txn_id,
    stage_data: JSON.stringify(timeline),
    updated_by: by, updated_at: now
  });
}

function updateImportTransaction(data) {
  var now = impNow();
  var updates = {};
  Object.keys(data).forEach(function(k) { if (k !== 'by') updates[k] = data[k]; });
  updates.updated_at = now;
  var ok = impUpdate('ImportTransactions', data.id, updates);
  return ok ? { success: true } : { error: 'Transaction not found' };
}

function getImportDocs(data) {
  return { docs: impRows('ImportDocs').filter(function(r) { return r.txn_id === data.txn_id; }) };
}

function updateImportDocs(data) {
  var now = impNow();
  var docs;
  try { docs = JSON.parse(data.docs); } catch(e) { return { error: 'Invalid docs JSON' }; }
  var sh = getImportSheet('ImportDocs');
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (sh.getLastRow() < 2) return { error: 'No doc rows found' };
  var rows = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
  var txnIdIdx = headers.indexOf('txn_id');
  var nameIdx  = headers.indexOf('name');
  var updIdx   = headers.indexOf('updated_at');
  var nameToRowIdx = {};
  rows.forEach(function(row, i) {
    if (String(row[txnIdIdx]) === data.txn_id) {
      nameToRowIdx[String(row[nameIdx])] = i;
    }
  });
  docs.forEach(function(d) {
    var ri = nameToRowIdx[d.name];
    if (ri !== undefined) {
      headers.forEach(function(h, j) {
        if (d[h] !== undefined) sh.getRange(ri+2, j+1).setValue(d[h]);
      });
      if (updIdx >= 0) sh.getRange(ri+2, updIdx+1).setValue(now);
    }
  });
  var approved = docs.filter(function(d) { return d.status === 'Approved'; }).length;
  impUpdate('ImportTransactions', data.txn_id, { docs_total: docs.length, docs_approved: approved, updated_at: now });
  return { success: true };
}

function getImportPayments(data) {
  return { payments: impRows('ImportPayments').filter(function(r) { return r.txn_id === data.txn_id; }) };
}

function addImportPayment(data) {
  var id = impUid();
  var now = impNow();
  var headers = impHeaders('ImportPayments');
  var row = {};
  headers.forEach(function(h) { row[h] = data[h] || ''; });
  row.id = id;
  row.created_by = data.by || '';
  row.created_at = now;
  impAppend('ImportPayments', headers, row);
  impUpdate('ImportTransactions', data.txn_id, { updated_at: now });
  return { success: true, id: id, created_at: now };
}

function updateImportPayment(data) {
  var updates = {};
  ['type','amount','currency','conversion_rate','party','due_date','paid_date','reference','notes'].forEach(function(f) {
    if (data[f] !== undefined) updates[f] = data[f];
  });
  var ok = impUpdate('ImportPayments', data.id, updates);
  if (ok) impUpdate('ImportTransactions', data.txn_id, { updated_at: impNow() });
  return ok ? { success: true } : { success: false, error: 'Not found' };
}

function getImportLineItems(data) {
  return { items: impRows('ImportLineItems').filter(function(r) { return r.txn_id === data.txn_id; }) };
}

function addImportLineItem(data) {
  var id = impUid();
  var now = impNow();
  var headers = impHeaders('ImportLineItems');
  var row = {};
  headers.forEach(function(h) { row[h] = data[h] || ''; });
  row.id = id;
  row.created_by = data.by || '';
  row.created_at = now;
  // Auto-calc amount if not provided
  if (!row.amount && row.qty && row.price) {
    row.amount = (parseFloat(row.qty) * parseFloat(row.price)).toFixed(2);
  }
  impAppend('ImportLineItems', headers, row);
  impUpdate('ImportTransactions', data.txn_id, { updated_at: now });
  return { success: true, id: id, amount: row.amount };
}

function deleteImportLineItem(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('ImportLineItems');
  if (!sheet) return { success: false };
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var idCol = headers.indexOf('id');
  for (var i = rows.length - 1; i >= 1; i--) {
    if (rows[i][idCol] === data.id) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false, error: 'Not found' };
}

function getImportCosts(data) {
  return { costs: impRows('ImportCosts').filter(function(r) { return r.txn_id === data.txn_id; }) };
}

function addImportCost(data) {
  var id = impUid();
  var now = impNow();
  var headers = impHeaders('ImportCosts');
  var row = {};
  headers.forEach(function(h) { row[h] = data[h] || ''; });
  row.id = id;
  row.created_by = data.by || '';
  row.created_at = now;
  impAppend('ImportCosts', headers, row);
  impUpdate('ImportTransactions', data.txn_id, { updated_at: now });
  return { success: true, id: id };
}

function deleteImportCost(data) {
  var sh = getImportSheet('ImportCosts');
  if (sh.getLastRow() < 2) return { error: 'Not found' };
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var idIdx = headers.indexOf('id');
  var rows = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][idIdx]) === String(data.cost_id)) {
      sh.deleteRow(i+2);
      return { success: true };
    }
  }
  return { error: 'Cost not found' };
}

function getImportTimeline(data) {
  var rows = impRows('ImportTimeline').filter(function(r) { return r.txn_id === data.txn_id; });
  if (!rows.length) return { timeline: [] };
  try {
    return { timeline: JSON.parse(rows[rows.length-1].stage_data) };
  } catch(e) { return { timeline: [] }; }
}

function updateImportTimeline(data) {
  var now = impNow();
  var sh = getImportSheet('ImportTimeline');
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var txnIdx = headers.indexOf('txn_id');
  var sdIdx  = headers.indexOf('stage_data');
  var ubIdx  = headers.indexOf('updated_by');
  var uaIdx  = headers.indexOf('updated_at');
  if (sh.getLastRow() >= 2) {
    var rows = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][txnIdx]) === data.txn_id) {
        if (sdIdx >= 0) sh.getRange(i+2, sdIdx+1).setValue(data.timeline);
        if (ubIdx >= 0) sh.getRange(i+2, ubIdx+1).setValue(data.by || '');
        if (uaIdx >= 0) sh.getRange(i+2, uaIdx+1).setValue(now);
        impUpdate('ImportTransactions', data.txn_id, { updated_at: now });
        return { success: true };
      }
    }
  }
  impAppend('ImportTimeline', headers, {
    id: impUid(), txn_id: data.txn_id, stage_data: data.timeline,
    updated_by: data.by || '', updated_at: now
  });
  return { success: true };
}

function getImportRemarks(data) {
  var remarks = impRows('ImportRemarks').filter(function(r) { return r.txn_id === data.txn_id; });
  return { remarks: remarks.reverse() };
}

function addImportRemark(data) {
  var id = impUid();
  var now = impNow();
  var headers = impHeaders('ImportRemarks');
  impAppend('ImportRemarks', headers, {
    id: id, txn_id: data.txn_id, remark: data.remark, added_by: data.by || '', created_at: now
  });
  return { success: true, id: id, created_at: now };
}

function getImportDocReviews(data) {
  return { reviews: impRows('ImportDocReviews').filter(function(r) { return r.txn_id === data.txn_id; }) };
}

function addImportDocReview(data) {
  var id = impUid();
  var now = impNow();
  var headers = impHeaders('ImportDocReviews');
  impAppend('ImportDocReviews', headers, {
    id: id,
    txn_id:   data.txn_id,
    doc_name: data.doc_name,
    action:   data.action,
    round:    data.round || 1,
    notes:    data.notes || '',
    actor:    data.actor || '',
    created_at: now
  });
  impUpdate('ImportTransactions', data.txn_id, { updated_at: now });
  return { success: true, id: id, created_at: now };
}
