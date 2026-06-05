'use client';

import React, { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, LineChart, Line, CartesianGrid, PieChart, Pie, Cell } from 'recharts';
import { Upload, Users, BriefcaseBusiness, MapPin, CalendarCheck, Lightbulb, Package, Euro, AlertTriangle, Globe2, Database, Download, ClipboardList, Cloud, CloudOff } from 'lucide-react';
import { supabase, isSupabaseConfigured } from '../lib/supabase';

type Customer = {
  code: string; name: string; city: string; province: string; region: string; country: string; agent: string; email: string; phone: string;
};
type InvoiceRow = { code: string; description: string; qty: number; total: number; price: number };
type Invoice = { id: string; date: string; number: string; customerCode: string; customerName: string; city: string; province: string; agent: string; total: number; rows: InvoiceRow[] };
type Visit = { id: string; date: string; customerCode: string; customerName: string; agent: string; outcome: string; nextDate: string; notes: string };

const demoCustomers: Customer[] = [
  { code: '3273', name: 'Cliente Demo Napoli', city: 'Napoli', province: 'NA', region: 'Campania', country: 'Italia', agent: 'Terrana Mattia', email: '', phone: '' },
  { code: '2548', name: '68 mark avenue', city: 'Abidjan', province: 'EE', region: '', country: 'Costa d’Avorio', agent: '', email: '68markavenueabidjan@gmail.com', phone: '00221773132293' },
  { code: '3569', name: 'U-TEX', city: 'S. Antonino', province: 'EE', region: '', country: 'Svizzera', agent: 'Mauro', email: 'info@utex.ch', phone: '0793757763' }
];

const demoInvoices: Invoice[] = [
  { id: '155', date: '2025-06-09', number: '155', customerCode: '3569', customerName: 'U-TEX', city: 'S. Antonino', province: 'EE', agent: 'Mauro', total: 331.72, rows: [{ code: '1324', description: 'ZAMELIA collana', qty: 1, price: 12, total: 3.6 }] },
  { id: '1024', date: '2026-05-24', number: '1024', customerCode: '3273', customerName: 'Cliente Demo Napoli', city: 'Napoli', province: 'NA', agent: 'Terrana Mattia', total: 1280, rows: [{ code: 'B001', description: 'Bracciale pietra naturale', qty: 90, price: 14.2, total: 1280 }] }
];

const provinceCoords: Record<string, { x: number; y: number }> = { NA:{x:55,y:68}, SA:{x:57,y:72}, RM:{x:48,y:58}, MI:{x:38,y:25}, FI:{x:43,y:45}, VE:{x:55,y:31}, BA:{x:67,y:70}, EE:{x:82,y:38} };
const inactiveDays = 120;
const targetVisits = 4;

function money(v: number) { return new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 }).format(v || 0); }
function num(v: any) { return Number(String(v ?? '0').replace(',', '.')) || 0; }
function daysSince(date: string) { if (!date) return 9999; return Math.floor((Date.now() - new Date(date).getTime()) / 86400000); }
function isItaly(country: string) { const c = String(country || '').trim().toLowerCase(); return !c || c === 'italia' || c === 'italy' || c === 'it'; }
function cleanProvince(p: string, country: string) { const v = String(p || '').trim().toUpperCase(); if (!isItaly(country)) return 'ESTERO'; return v || 'ND'; }
function customerSegment(c: any) { if (!isItaly(c.country)) return 'Estero'; if (!c.province || c.province === 'ND' || !c.city) return 'Dati incompleti'; return 'Italia completa'; }
function getXmlValue(chunk: string, tag: string) { const m = chunk.match(new RegExp(`<${tag}>([\\s\\S]*?)</${tag}>`, 'i')); return m ? decode(m[1]) : ''; }
function decode(s: string) { return s.replace(/<!\[CDATA\[|\]\]>/g, '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').trim(); }

function parseCustomersWorkbook(buffer: ArrayBuffer): Customer[] {
  const wb = XLSX.read(buffer, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows: any[] = XLSX.utils.sheet_to_json(ws, { defval: '' });
  return rows.map(r => ({
    code: String(r['Cod.'] || r['Cod'] || '').trim(),
    name: String(r['Denominazione'] || '').trim(),
    city: String(r['Città'] || r['Citta'] || '').trim(),
    province: cleanProvince(String(r['Prov.'] || r['Prov'] || '').trim(), String(r['Nazione'] || '').trim()),
    region: String(r['Regione'] || '').trim(),
    country: String(r['Nazione'] || '').trim(),
    agent: String(r['Agente'] || '').trim(),
    email: String(r['e-mail'] || r['Email'] || '').trim(),
    phone: String(r['Tel.'] || r['Cell'] || '').trim()
  })).filter(c => c.code || c.name);
}

function parseInvoicesXml(xml: string): Invoice[] {
  const invoices: Invoice[] = [];
  const docs = xml.match(/<Document>[\s\S]*?<\/Document>/g) || [];
  for (const doc of docs) {
    const rowsChunk = getXmlValue(doc, 'Rows');
    const rowChunks = rowsChunk.match(/<Row>[\s\S]*?<\/Row>/g) || [];
    const rows = rowChunks.map(row => ({
      code: getXmlValue(row, 'Code'),
      description: getXmlValue(row, 'Description'),
      qty: num(getXmlValue(row, 'Qty')),
      price: num(getXmlValue(row, 'Price')),
      total: num(getXmlValue(row, 'Total'))
    })).filter(r => r.code || r.description || r.qty || r.total);
    invoices.push({
      id: `${getXmlValue(doc, 'Date')}-${getXmlValue(doc, 'Number')}`,
      date: getXmlValue(doc, 'Date'),
      number: getXmlValue(doc, 'Number'),
      customerCode: getXmlValue(doc, 'CustomerCode'),
      customerName: getXmlValue(doc, 'CustomerName'),
      city: getXmlValue(doc, 'CustomerCity'),
      province: getXmlValue(doc, 'CustomerProvince') || 'ND',
      agent: getXmlValue(doc, 'SalesAgent'),
      total: num(getXmlValue(doc, 'Total')),
      rows
    });
  }
  return invoices.filter(i => i.customerName || i.customerCode);
}

export default function Page() {
  const [customers, setCustomers] = useState<Customer[]>(demoCustomers);
  const [invoices, setInvoices] = useState<Invoice[]>(demoInvoices);
  const [visits, setVisits] = useState<Visit[]>([]);
  const [tab, setTab] = useState('dashboard');
  const [q, setQ] = useState('');
  const [message, setMessage] = useState(isSupabaseConfigured() ? 'Supabase configurato. Puoi caricare i file e salvare nel database.' : 'Carica clienti danea.xlsx e fatture.DefXml. Per salvare online configura Supabase nel file .env.local.');
  const [saving, setSaving] = useState(false);
  const [visitForm, setVisitForm] = useState({ customerCode: '', date: new Date().toISOString().slice(0,10), outcome: 'Visita effettuata', nextDate: '', notes: '' });

  const [selectedCustomer, setSelectedCustomer] = useState<any | null>(null);

  useEffect(() => {
  if (isSupabaseConfigured()) {
    loadFromSupabase();
  }
}, []);

  const mergedCustomers = useMemo(() => {
    const m = new Map<string, any>();
    customers.forEach(c => m.set(c.code || c.name, { ...c, amount: 0, orders: 0, qty: 0, lastOrder: '', topItems: {}, visits: 0, lastVisit: '', segment: customerSegment(c) }));
    invoices.forEach(inv => {
      const key = inv.customerCode || inv.customerName;
      const c = m.get(key) || { code: inv.customerCode, name: inv.customerName, city: inv.city, province: inv.province || 'ND', region: '', country: 'Italia', agent: inv.agent, amount: 0, orders: 0, qty: 0, lastOrder: '', topItems: {}, visits: 0, lastVisit: '', segment: 'Dati incompleti' };
      c.amount += inv.total; c.orders += 1; c.agent = c.agent || inv.agent; c.city = c.city || inv.city; c.province = c.province || inv.province;
      if (!c.lastOrder || inv.date > c.lastOrder) c.lastOrder = inv.date;
      inv.rows.forEach(r => { c.qty += r.qty; if (r.description) c.topItems[r.description] = (c.topItems[r.description] || 0) + r.qty; });
      m.set(key, c);
    });
    visits.forEach(v => {
      const key = v.customerCode || v.customerName; const c = m.get(key); if (!c) return;
      c.visits += 1; if (!c.lastVisit || v.date > c.lastVisit) c.lastVisit = v.date;
    });
    return Array.from(m.values()).map(c => {
      const top = Object.entries(c.topItems || {}).sort((a:any,b:any)=>b[1]-a[1])[0];
      return { ...c, province: c.province || 'ND', country: c.country || 'Italia', segment: customerSegment(c), agent: c.agent || 'Senza agente', status: daysSince(c.lastOrder) <= inactiveDays ? 'Attivo' : 'Inattivo', topItem: top ? top[0] : '-', avgOrder: c.orders ? c.amount / c.orders : 0, nextAction: !c.lastOrder ? 'Verifica anagrafica' : daysSince(c.lastOrder) > 180 ? 'Recupero urgente' : daysSince(c.lastOrder) > inactiveDays ? 'Richiamare' : c.visits < 1 ? 'Prima visita' : 'Monitorare' }; 
    });
  }, [customers, invoices, visits]);

  const agents = useMemo(() => {
    const m = new Map<string, any>();
    mergedCustomers.forEach(c => {
      const a = m.get(c.agent) || { agent: c.agent, customers: 0, active: 0, inactive: 0, amount: 0, orders: 0, qty: 0, visits: 0, zones: new Set(), lowVisit: 0 };
      a.customers++; a.active += c.status === 'Attivo' ? 1 : 0; a.inactive += c.status === 'Inattivo' ? 1 : 0; a.amount += c.amount; a.orders += c.orders; a.qty += c.qty; a.visits += c.visits; a.zones.add(c.province || 'ND'); if (c.visits < targetVisits) a.lowVisit++;
      m.set(c.agent, a);
    });
    return Array.from(m.values()).map(a => ({ ...a, avgOrder: a.orders ? a.amount / a.orders : 0, avgVisits: a.customers ? a.visits / a.customers : 0, zonesText: Array.from(a.zones).join(', ') })).sort((a,b)=>b.amount-a.amount);
  }, [mergedCustomers]);

  const zones = useMemo(() => {
    const m = new Map<string, any>();
    mergedCustomers.forEach(c => {
      const z = m.get(c.province || 'ND') || { province: c.province || 'ND', amount: 0, customers: 0, active: 0, inactive: 0, uncovered: 0, agents: new Set() };
      z.amount += c.amount; z.customers++; z.active += c.status === 'Attivo' ? 1 : 0; z.inactive += c.status === 'Inattivo' ? 1 : 0; if (c.agent === 'Senza agente') z.uncovered++; else z.agents.add(c.agent);
      m.set(z.province, z);
    });
    return Array.from(m.values()).map(z => ({...z, agentsText: Array.from(z.agents).join(', ') || 'Nessun agente'}));
  }, [mergedCustomers]);

  const segments = useMemo(() => {
    const m = new Map<string, any>();
    mergedCustomers.forEach(c => {
      const key = c.segment || 'Dati incompleti';
      const x = m.get(key) || { segment: key, customers: 0, amount: 0, inactive: 0, unassigned: 0 };
      x.customers++; x.amount += c.amount; if (c.status === 'Inattivo') x.inactive++; if (c.agent === 'Senza agente') x.unassigned++;
      m.set(key, x);
    });
    return Array.from(m.values()).sort((a,b)=>b.customers-a.customers);
  }, [mergedCustomers]);

  const countries = useMemo(() => {
    const m = new Map<string, any>();
    mergedCustomers.forEach(c => {
      const key = isItaly(c.country) ? 'Italia' : (c.country || 'Estero non specificato');
      const x = m.get(key) || { country: key, customers: 0, amount: 0, active: 0, inactive: 0, agents: new Set() };
      x.customers++; x.amount += c.amount; x.active += c.status === 'Attivo' ? 1 : 0; x.inactive += c.status === 'Inattivo' ? 1 : 0; if (c.agent !== 'Senza agente') x.agents.add(c.agent);
      m.set(key, x);
    });
    return Array.from(m.values()).map(x=>({...x, agentsText: Array.from(x.agents).join(', ') || 'Nessun agente'})).sort((a,b)=>b.amount-a.amount);
  }, [mergedCustomers]);

  const cleanupList = useMemo(() => mergedCustomers.filter(c => c.segment === 'Dati incompleti' || c.agent === 'Senza agente' || !c.email).sort((a,b)=>b.amount-a.amount).slice(0,300), [mergedCustomers]);

  const items = useMemo(() => {
    const m = new Map<string, any>();
    invoices.forEach(inv => inv.rows.forEach(r => { const k = r.description || r.code; if(!k) return; const x = m.get(k) || { item: k, qty: 0, amount: 0 }; x.qty += r.qty; x.amount += r.total; m.set(k, x); }));
    return Array.from(m.values()).sort((a,b)=>b.qty-a.qty).slice(0,20);
  }, [invoices]);

  const monthly = useMemo(() => {
    const m = new Map<string, number>(); invoices.forEach(i => { const month = (i.date || '').slice(0,7) || 'ND'; m.set(month, (m.get(month)||0)+i.total); });
    return Array.from(m.entries()).map(([month, amount]) => ({ month, amount })).sort((a,b)=>a.month.localeCompare(b.month));
  }, [invoices]);

  const filteredCustomers = mergedCustomers.filter(c => `${c.name} ${c.code} ${c.agent} ${c.city} ${c.province}`.toLowerCase().includes(q.toLowerCase()));
  const totalAmount = invoices.reduce((s,i)=>s+i.total,0); const totalQty = invoices.reduce((s,i)=>s+i.rows.reduce((a,r)=>a+r.qty,0),0);
  const inactive = mergedCustomers.filter(c=>c.status==='Inattivo').length; const unassigned = mergedCustomers.filter(c=>c.agent==='Senza agente').length;

  async function handleCustomersFile(file: File) { const buf = await file.arrayBuffer(); const parsed = parseCustomersWorkbook(buf); setCustomers(parsed); setMessage(`Importati ${parsed.length} clienti dal file Excel.`); }
  async function handleInvoicesFile(file: File) { const text = await file.text(); const parsed = parseInvoicesXml(text); setInvoices(parsed); setMessage(`Importate ${parsed.length} fatture/documenti dal file XML.`); }
 async function addVisit() {
  const c = mergedCustomers.find(
    x => x.code === visitForm.customerCode
  );

  if (!c) {
    alert('Seleziona un cliente.');
    return;
  }

  const nuovaVisita = {
    id: crypto.randomUUID(),
    customerCode: c.code,
    customerName: c.name,
    agent: c.agent,
    date: visitForm.date,
    outcome: visitForm.outcome,
    nextDate: visitForm.nextDate,
    notes: visitForm.notes
  };

  setVisits(v => [nuovaVisita, ...v]);

  if (supabase) {
    const { error } = await supabase
      .from('visite')
      .insert({
        id: nuovaVisita.id,
        date: nuovaVisita.date,
        customer_code: nuovaVisita.customerCode,
        customer_name: nuovaVisita.customerName,
        agent: nuovaVisita.agent,
        outcome: nuovaVisita.outcome,
        next_date: nuovaVisita.nextDate || null,
        notes: nuovaVisita.notes
      });

    if (error) {
      alert(error.message);
      return;
    }
  }

  setVisitForm({
    customerCode: '',
    date: new Date().toISOString().slice(0,10),
    outcome: 'Visita effettuata',
    nextDate: '',
    notes: ''
  });
}
  async function saveToSupabase() {
    if (!supabase) return alert('Supabase non è configurato. Crea .env.local con URL e anon key.');
    setSaving(true);
    try {
      const customersRows = customers.map(c => ({ code: c.code || null, name: c.name || 'Senza nome', city: c.city, province: c.province, region: c.region, country: c.country, agent: c.agent, email: c.email, phone: c.phone }));
      if (customersRows.length) {
        const { error } = await supabase.from('clienti').upsert(customersRows, { onConflict: 'code' });
        if (error) throw error;
      }

      const invoiceRows = invoices.map(i => ({ id: i.id, date: i.date || null, number: i.number, customer_code: i.customerCode, customer_name: i.customerName, city: i.city, province: i.province, agent: i.agent, total: i.total }));
      if (invoiceRows.length) {
        const { error } = await supabase.from('fatture').upsert(invoiceRows, { onConflict: 'id' });
        if (error) throw error;
        await supabase.from('righe_fattura').delete().in('invoice_id', invoices.map(i => i.id));
        const lines = invoices.flatMap(i => i.rows.map(r => ({ invoice_id: i.id, code: r.code, description: r.description, qty: r.qty, price: r.price, total: r.total })));
        if (lines.length) {
          const { error: rowsError } = await supabase.from('righe_fattura').insert(lines);
          if (rowsError) throw rowsError;
        }
      }

      const visitRows = visits.map(v => ({ id: v.id, date: v.date, customer_code: v.customerCode, customer_name: v.customerName, agent: v.agent, outcome: v.outcome, next_date: v.nextDate || null, notes: v.notes }));
      if (visitRows.length) {
        const { error } = await supabase.from('visite').upsert(visitRows, { onConflict: 'id' });
        if (error) throw error;
      }
      setMessage(`Salvati su Supabase: ${customersRows.length} clienti, ${invoiceRows.length} fatture, ${visits.length} visite.`);
    } catch (e: any) {
      setMessage('Errore Supabase: ' + (e.message || e));
    } finally {
      setSaving(false);
    }
  }

  async function loadFromSupabase() {
    if (!supabase) return alert('Supabase non è configurato. Crea .env.local con URL e anon key.');
    setSaving(true);
    try {
      const [{ data: c, error: ce }, { data: f, error: fe }, { data: r, error: re }, { data: v, error: ve }] = await Promise.all([
        supabase.from('clienti').select('*').limit(5000),
        supabase.from('fatture').select('*').limit(10000),
        supabase.from('righe_fattura').select('*').limit(50000),
        supabase.from('visite').select('*').order('date', { ascending: false }).limit(5000)
      ]);
      if (ce || fe || re || ve) throw ce || fe || re || ve;
      setCustomers((c || []).map((x: any) => ({ code: x.code || '', name: x.name || '', city: x.city || '', province: x.province || '', region: x.region || '', country: x.country || '', agent: x.agent || '', email: x.email || '', phone: x.phone || '' })));
      const rowsByInvoice = new Map<string, any[]>();
      (r || []).forEach((row: any) => { const arr = rowsByInvoice.get(row.invoice_id) || []; arr.push({ code: row.code || '', description: row.description || '', qty: Number(row.qty || 0), price: Number(row.price || 0), total: Number(row.total || 0) }); rowsByInvoice.set(row.invoice_id, arr); });
      setInvoices((f || []).map((x: any) => ({ id: x.id, date: x.date || '', number: x.number || '', customerCode: x.customer_code || '', customerName: x.customer_name || '', city: x.city || '', province: x.province || '', agent: x.agent || '', total: Number(x.total || 0), rows: rowsByInvoice.get(x.id) || [] })));
      setVisits((v || []).map((x: any) => ({ id: x.id, date: x.date || '', customerCode: x.customer_code || '', customerName: x.customer_name || '', agent: x.agent || '', outcome: x.outcome || '', nextDate: x.next_date || '', notes: x.notes || '' })));
      setMessage(`Caricati da Supabase: ${(c || []).length} clienti, ${(f || []).length} fatture, ${(v || []).length} visite.`);
    } catch (e: any) {
      setMessage('Errore caricamento Supabase: ' + (e.message || e));
    } finally {
      setSaving(false);
    }
  }

  function exportVisits() { const ws = XLSX.utils.json_to_sheet(visits); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Visite'); XLSX.writeFile(wb, 'visite_agenti.xlsx'); }
  function exportCustomersToFix() { const rows = cleanupList.map(c => ({ Codice: c.code, Cliente: c.name, Citta: c.city, Provincia: c.province, Nazione: c.country, Agente: c.agent, Vendite: c.amount, Stato: c.status, Azione: c.nextAction, Email: c.email })); const ws = XLSX.utils.json_to_sheet(rows); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Da sistemare'); XLSX.writeFile(wb, 'clienti_da_sistemare.xlsx'); }
  function exportAgentWorklist() { const rows = mergedCustomers.filter(c => c.nextAction !== 'Monitorare').map(c => ({ Agente: c.agent, Codice: c.code, Cliente: c.name, Provincia: c.province, Nazione: c.country, Vendite: c.amount, Ultimo_ordine: c.lastOrder, Visite: c.visits, Azione: c.nextAction, Email: c.email, Telefono: c.phone })); const ws = XLSX.utils.json_to_sheet(rows); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Lista lavoro agenti'); XLSX.writeFile(wb, 'lista_lavoro_agenti.xlsx'); }

  return <main style={{maxWidth: 1280, margin: '0 auto', padding: 24}}>
    <header style={{display:'flex', justifyContent:'space-between', gap:16, flexWrap:'wrap', alignItems:'center', marginBottom:24}}>
      <div><div className="small">CRM commerciale per export Danea Easyfatt</div><h1 style={{fontSize:42, margin:'6px 0'}}>Danea CRM Agenti</h1><p className="small">Clienti, fatture, agenti, zone e visite in un’unica dashboard.</p></div>
      <div style={{display:'flex', gap:10, flexWrap:'wrap'}}>
        <label className="btn"><Upload size={16}/> Clienti Excel <input hidden type="file" accept=".xlsx,.xls" onChange={e=>e.target.files?.[0] && handleCustomersFile(e.target.files[0])}/></label>
        <label className="btn"><Upload size={16}/> Fatture XML <input hidden type="file" accept=".xml,.DefXml,.txt" onChange={e=>e.target.files?.[0] && handleInvoicesFile(e.target.files[0])}/></label>
        <button className="btn secondary" onClick={saveToSupabase} disabled={saving}>{isSupabaseConfigured() ? <Cloud size={16}/> : <CloudOff size={16}/>} Salva DB</button>
        <button className="btn secondary" onClick={loadFromSupabase} disabled={saving}><Database size={16}/> Carica DB</button>
      </div>
    </header>
    <div className="card" style={{padding:14, marginBottom:16}}>{message}</div>
    <section className="grid grid-4" style={{marginBottom:18}}>
      <Kpi icon={<Euro/>} label="Fatturato" value={money(totalAmount)} />
      <Kpi icon={<Package/>} label="Pezzi venduti" value={Math.round(totalQty)} />
      <Kpi icon={<Users/>} label="Clienti" value={mergedCustomers.length} />
      <Kpi icon={<AlertTriangle/>} label="Inattivi / senza agente" value={`${inactive} / ${unassigned}`} />
      <Kpi icon={<Globe2/>} label="Clienti estero" value={mergedCustomers.filter(c=>c.segment==='Estero').length} />
      <Kpi icon={<Database/>} label="Dati incompleti" value={mergedCustomers.filter(c=>c.segment==='Dati incompleti').length} />
    </section>
    <nav className="nav" style={{marginBottom:18}}>{['dashboard','clienti','agenti','visite','zone','estero','pulizia','articoli','consigli'].map(t=><button key={t} className={tab===t?'active':''} onClick={()=>setTab(t)}>{t[0].toUpperCase()+t.slice(1)}</button>)}</nav>

    {tab==='dashboard' && <section className="grid grid-2"><Panel title="Fatturato mensile"><ResponsiveContainer width="100%" height={320}><LineChart data={monthly}><CartesianGrid strokeDasharray="3 3"/><XAxis dataKey="month"/><YAxis/><Tooltip formatter={(v:any)=>money(v)}/><Line dataKey="amount" strokeWidth={3}/></LineChart></ResponsiveContainer></Panel><Panel title="Stato clienti"><ResponsiveContainer width="100%" height={320}><PieChart><Pie data={[{name:'Attivi',value:mergedCustomers.length-inactive},{name:'Inattivi',value:inactive}]} dataKey="value" label outerRadius={110}>{[0,1].map(i=><Cell key={i}/>)}</Pie><Tooltip/></PieChart></ResponsiveContainer></Panel></section>}

    {tab==='clienti' && <section><input className="input" placeholder="Cerca cliente, agente, città, provincia..." value={q} onChange={e=>setQ(e.target.value)} style={{marginBottom:12}}
    />
    {selectedCustomer && (
  <div className="card" style={{ marginBottom: 12 }}>
    <h2>Scheda cliente</h2>

    <p><b>Cliente:</b> {selectedCustomer.name}</p>
    <p><b>Codice:</b> {selectedCustomer.code}</p>
    <p><b>Agente:</b> {selectedCustomer.agent}</p>
    <p><b>Città:</b> {selectedCustomer.city}</p>
    <p><b>Provincia:</b> {selectedCustomer.province}</p>
    <p><b>Vendite:</b> {money(selectedCustomer.amount)}</p>
    <p><b>Ordine medio:</b> {money(selectedCustomer.avgOrder)}</p>
    <p><b>Ultimo ordine:</b> {selectedCustomer.lastOrder || '-'}</p>

    <h3>Storico visite</h3>

<table>
  <thead>
    <tr>
      <th>Data</th>
      <th>Esito</th>
      <th>Prossima</th>
      <th>Note</th>
    </tr>
  </thead>
  <tbody>
    {visits
      .filter(v => v.customerCode === selectedCustomer.code)
      .map(v => (
        <tr key={v.id}>
          <td>{v.date}</td>
          <td>{v.outcome}</td>
          <td>{v.nextDate || '-'}</td>
          <td>{v.notes || '-'}</td>
        </tr>
      ))}
  </tbody>
</table>

<h3>Ultime fatture</h3>

<table>
  <thead>
    <tr>
      <th>Data</th>
      <th>Numero</th>
      <th>Totale</th>
    </tr>
  </thead>
  <tbody>
    {invoices
      .filter(i => i.customerCode === selectedCustomer.code)
      .sort((a, b) => b.date.localeCompare(a.date))
      .slice(0, 10)
      .map(i => (
        <tr key={i.id}>
          <td>{i.date}</td>
          <td>{i.number}</td>
          <td>{money(i.total)}</td>
        </tr>
      ))}
  </tbody>
</table>

<h3>Articoli acquistati</h3>

<table>
  <thead>
    <tr>
      <th>Articolo</th>
      <th>Quantità</th>
      <th>Totale</th>
    </tr>
  </thead>
  <tbody>
    {Object.entries(
      invoices
        .filter(i => i.customerCode === selectedCustomer.code)
        .flatMap(i => i.rows)
        .reduce((acc: any, r: any) => {
          const key = r.description || r.code || 'Articolo senza nome';

          if (!acc[key]) {
            acc[key] = {
              description: key,
              qty: 0,
              total: 0
            };
          }

          acc[key].qty += Number(r.qty || 0);
          acc[key].total += Number(r.total || 0);

          return acc;
        }, {})
    )
      .map(([, item]: any) => item)
      .sort((a: any, b: any) => b.qty - a.qty)
      .slice(0, 10)
      .map((item: any) => (
        <tr key={item.description}>
          <td>{item.description}</td>
          <td>{item.qty}</td>
          <td>{money(item.total)}</td>
        </tr>
      ))}
  </tbody>
  
</table>
<button
  className="btn"
  onClick={() => {
    setVisitForm(v => ({ ...v, customerCode: selectedCustomer.code }));
    setTab('visite');
  }}
>
  Nuova visita
</button>

    <button onClick={() => setSelectedCustomer(null)}>
      Chiudi
    </button>
  </div>
)}<div className="card table-wrap"><table><thead><tr><th>Cliente</th><th>Agente</th><th>Zona</th><th>Vendite</th><th>Ordine medio</th><th>Ultimo ordine</th><th>Visite</th><th>Top articolo</th><th>Segmento</th><th>Azione</th><th>Stato</th></tr></thead><tbody>{filteredCustomers.map(c=><tr
  key={c.code || c.name}
  onClick={() => setSelectedCustomer(c)}
  style={{ cursor: 'pointer' }}
><td><b>{c.name}</b><br/><span className="small">{c.code} · {c.email}</span></td><td>{c.agent}</td><td>{c.city} {c.province}</td><td>{money(c.amount)}</td><td>{money(c.avgOrder)}</td><td>{c.lastOrder || '-'}</td><td>{c.visits}</td><td>{c.topItem}</td><td><span className="badge">{c.segment}</span></td><td>{c.nextAction}</td><td><span className={`badge ${c.status==='Attivo'?'ok':'bad'}`}>{c.status}</span></td></tr>)}</tbody></table></div></section>}

    {tab==='agenti' && <section className="grid"><div className="grid grid-2"><Panel title="Vendite per agente"><ResponsiveContainer width="100%" height={320}><BarChart data={agents}><XAxis dataKey="agent"/><YAxis/><Tooltip formatter={(v:any)=>money(v)}/><Bar dataKey="amount"/></BarChart></ResponsiveContainer></Panel><Panel title="Clienti attivi/inattivi"><ResponsiveContainer width="100%" height={320}><BarChart data={agents}><XAxis dataKey="agent"/><YAxis/><Tooltip/><Bar dataKey="active" stackId="a"/><Bar dataKey="inactive" stackId="a"/></BarChart></ResponsiveContainer></Panel></div><div className="card table-wrap"><table><thead><tr><th>Agente</th><th>Clienti</th><th>Vendite</th><th>Ordine medio</th><th>Visite medie</th><th>Clienti sotto visite</th><th>Zone</th></tr></thead><tbody>{agents.map(a=><tr key={a.agent}><td><b>{a.agent}</b></td><td>{a.customers}</td><td>{money(a.amount)}</td><td>{money(a.avgOrder)}</td><td>{a.avgVisits.toFixed(1)}</td><td><span className={`badge ${a.lowVisit?'warn':'ok'}`}>{a.lowVisit}</span></td><td>{a.zonesText}</td></tr>)}</tbody></table></div></section>}

    {tab==='visite' && <section className="grid grid-2"><div className="card" style={{padding:18}}><h2>Registra visita</h2><select className="input" value={visitForm.customerCode} onChange={e=>setVisitForm({...visitForm, customerCode:e.target.value})}><option value="">Seleziona cliente</option>{mergedCustomers.map(c=><option key={c.code || c.name} value={c.code}>{c.name} — {c.agent}</option>)}</select><br/><br/><input className="input" type="date" value={visitForm.date} onChange={e=>setVisitForm({...visitForm,date:e.target.value})}/><br/><br/><select className="input" value={visitForm.outcome} onChange={e=>setVisitForm({...visitForm,outcome:e.target.value})}><option>Visita effettuata</option><option>Cliente interessato</option><option>Ordine previsto</option><option>Nessun interesse</option><option>Cliente chiuso / da recuperare</option></select><br/><br/><input className="input" type="date" value={visitForm.nextDate} onChange={e=>setVisitForm({...visitForm,nextDate:e.target.value})}/><br/><br/><textarea className="input" placeholder="Note visita" value={visitForm.notes} onChange={e=>setVisitForm({...visitForm,notes:e.target.value})}/><br/><br/><button className="btn" onClick={addVisit}>Salva visita</button> <button className="btn secondary" onClick={exportVisits}>Esporta visite</button></div><div className="card table-wrap"><table><thead><tr><th>Data</th><th>Cliente</th><th>Agente</th><th>Esito</th><th>Prossima</th><th>Note</th></tr></thead><tbody>{visits.map(v=><tr key={v.id}><td>{v.date}</td><td>{v.customerName}</td><td>{v.agent}</td><td>{v.outcome}</td><td>{v.nextDate}</td><td>{v.notes}</td></tr>)}</tbody></table></div></section>}

    {tab==='zone' && <section className="grid grid-2"><div className="card" style={{padding:18}}><h2><MapPin size={20}/> Mappa zone</h2><div className="map">{zones.map(z=>{const p=provinceCoords[z.province]||{x:50,y:50}; return <div key={z.province} className={`pin ${z.uncovered?'bad':'ok'}`} style={{left:`${p.x}%`,top:`${p.y}%`}} title={`${z.province}: ${z.customers} clienti`}>{z.province}</div>})}</div></div><div className="card table-wrap"><table><thead><tr><th>Provincia</th><th>Clienti</th><th>Vendite</th><th>Agenti</th><th>Scoperti</th></tr></thead><tbody>{zones.map(z=><tr key={z.province}><td><b>{z.province}</b></td><td>{z.customers}</td><td>{money(z.amount)}</td><td>{z.agentsText}</td><td><span className={`badge ${z.uncovered?'bad':'ok'}`}>{z.uncovered}</span></td></tr>)}</tbody></table></div></section>}

    {tab==='estero' && <section className="grid grid-2"><Panel title="Clienti per segmento"><ResponsiveContainer width="100%" height={320}><BarChart data={segments}><XAxis dataKey="segment"/><YAxis/><Tooltip formatter={(v:any,n:any)=>n==='amount'?money(v):v}/><Bar dataKey="customers" name="Clienti"/><Bar dataKey="amount" name="Vendite"/></BarChart></ResponsiveContainer></Panel><div className="card table-wrap"><table><thead><tr><th>Nazione</th><th>Clienti</th><th>Vendite</th><th>Attivi</th><th>Inattivi</th><th>Agenti</th></tr></thead><tbody>{countries.map(c=><tr key={c.country}><td><b>{c.country}</b></td><td>{c.customers}</td><td>{money(c.amount)}</td><td>{c.active}</td><td>{c.inactive}</td><td>{c.agentsText}</td></tr>)}</tbody></table></div></section>}

    {tab==='pulizia' && <section className="grid"><div className="card" style={{padding:18, display:'flex', gap:12, flexWrap:'wrap', alignItems:'center', justifyContent:'space-between'}}><div><h2 style={{margin:'0 0 6px'}}>Pulizia dati e liste operative</h2><p className="small" style={{margin:0}}>Qui trovi clienti senza agente, senza provincia/città o senza email. Sono i dati che conviene sistemare in Danea o assegnare agli agenti.</p></div><div style={{display:'flex', gap:10, flexWrap:'wrap'}}><button className="btn" onClick={exportCustomersToFix}><Download size={16}/> Esporta da sistemare</button><button className="btn secondary" onClick={exportAgentWorklist}><ClipboardList size={16}/> Lista lavoro agenti</button></div></div><div className="card table-wrap"><table><thead><tr><th>Cliente</th><th>Problema</th><th>Agente</th><th>Zona</th><th>Nazione</th><th>Vendite</th><th>Azione consigliata</th><th>Email</th></tr></thead><tbody>{cleanupList.map(c=>{ const problems = [c.agent==='Senza agente'?'Senza agente':'', c.segment==='Dati incompleti'?'Indirizzo incompleto':'', !c.email?'Senza email':''].filter(Boolean).join(', '); return <tr key={c.code || c.name}><td><b>{c.name}</b><br/><span className="small">{c.code}</span></td><td>{problems}</td><td>{c.agent}</td><td>{c.city} {c.province}</td><td>{c.country || 'Italia'}</td><td>{money(c.amount)}</td><td>{c.nextAction}</td><td>{c.email || '-'}</td></tr>})}</tbody></table></div></section>}

    {tab==='articoli' && <section><Panel title="Articoli più ordinati"><ResponsiveContainer width="100%" height={420}><BarChart data={items}><XAxis dataKey="item" tick={{fontSize:11}}/><YAxis/><Tooltip/><Bar dataKey="qty"/></BarChart></ResponsiveContainer></Panel></section>}

    {tab==='consigli' && <section className="grid grid-3">
      <Advice text={`${inactive} clienti sono inattivi da oltre ${inactiveDays} giorni: prepara lista richiamo per agente.`}/>
      <Advice text={`${unassigned} clienti sono senza agente: assegnali o crea una zona scoperta da sviluppare.`}/>
      <Advice text={`${mergedCustomers.filter(c=>c.segment==='Dati incompleti').length} clienti hanno dati incompleti: sistemare provincia/città migliorerà mappe e zone.`}/>
      <Advice text={`${mergedCustomers.filter(c=>c.segment==='Estero').length} clienti sono esteri: conviene analizzarli separatamente per nazione e agente.`}/>
      <Advice text={`Obiettivo visite: ${targetVisits} visite/anno per cliente. Gli agenti con molti clienti sotto soglia vanno seguiti.`}/>
      <Advice text={items[0] ? `Prodotto più ordinato: ${items[0].item}. Valuta riassortimento e proposta abbinata.` : 'Importa fatture per vedere i prodotti più venduti.'}/>
      <Advice text={`Ordine medio globale: ${money(invoices.length ? totalAmount / invoices.length : 0)}. Usa questo dato per confrontare agenti e clienti.`}/>
      <Advice text="Prossimo sviluppo consigliato: login agenti e database Supabase per usare l’app da più dispositivi."/>

     
    </section>}
  </main>;
}

function Kpi({icon,label,value}: any) { return <div className="card" style={{padding:18}}><div className="small" style={{display:'flex',justifyContent:'space-between'}}>{label}{React.cloneElement(icon,{size:18})}</div><div style={{fontSize:26,fontWeight:800,marginTop:8}}>{value}</div></div> }
function Panel({title, children}: any) { return <div className="card" style={{padding:18}}><h2 style={{marginTop:0}}>{title}</h2>{children}</div> }
function Advice({text}: {text:string}) { return <div className="card" style={{padding:18, display:'flex', gap:12}}><Lightbulb size={22}/><p style={{margin:0}}>{text}</p></div> }
