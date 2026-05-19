import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ── Storage ───────────────────────────────────────────────────────────────────
const store = {
  get: async (k) => { try { const r = await window.storage.get(k); return r ? JSON.parse(r.value) : null; } catch { return null; } },
  set: async (k, v) => { try { await window.storage.set(k, JSON.stringify(v)); } catch {} },
};

// ── Helpers ───────────────────────────────────────────────────────────────────
const today = () => new Date().toISOString().split("T")[0];
const fmt = (v) => new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" }).format(v || 0);
const fmtDate = (d) => { if (!d) return ""; const [y, m, day] = d.split("-"); return `${day}/${m}/${y}`; };
const uid = () => `${Date.now()}_${Math.random().toString(36).slice(2,7)}`;

function excelDateToISO(s) {
  if (!s) return today();
  if (typeof s === "string" && s.includes("/")) {
    const p = s.split("/");
    if (p.length === 3) {
      const [a,b,c] = p;
      if (c.length===4) return `${c}-${b.padStart(2,"0")}-${a.padStart(2,"0")}`;
      if (a.length===4) return `${a}-${b.padStart(2,"0")}-${c.padStart(2,"0")}`;
    }
    return today();
  }
  if (typeof s === "string" && s.includes("-")) return s.slice(0,10);
  if (typeof s === "number") return new Date(Math.round((s-25569)*86400*1000)).toISOString().split("T")[0];
  return today();
}
function guessPayment(v) {
  if (!v) return "pix";
  const s = String(v).toLowerCase();
  if (s.includes("pix")) return "pix";
  if (s.includes("dinheiro")||s.includes("especie")||s.includes("cash")) return "dinheiro";
  if (s.includes("débit")||s.includes("debit")) return "debito";
  if (s.includes("crédit")||s.includes("credit")) return "credito";
  return "pix";
}
function guessTipo(v) {
  if (!v) return "despesa";
  const s = String(v).toLowerCase();
  if (s.includes("receita")||s.includes("entrada")||s.includes("recebido")||s.includes("credit")) return "receita";
  return "despesa";
}

// ── Constants ─────────────────────────────────────────────────────────────────
const PAYMENT_METHODS = [
  { id:"pix",     label:"Pix",      icon:"⚡", color:"#00BDAE" },
  { id:"dinheiro",label:"Dinheiro", icon:"💵", color:"#4CAF50" },
  { id:"debito",  label:"Débito",   icon:"💳", color:"#5B8BFF" },
  { id:"credito", label:"Crédito",  icon:"🔷", color:"#9C6FFF" },
];
const CAT_REC  = ["Receita Móveis","Receita Serviço","Receita MO","Cashback","Outros"];
const CAT_DESP = ["Matéria Prima","Gasolina","Alimentação","Fornecedor","Pró-Labore","Impostos","Marketing","Outros"];

// ── Icons ─────────────────────────────────────────────────────────────────────
const Ico = {
  Plus:   ()=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" width="20" height="20"><path d="M12 5v14M5 12h14"/></svg>,
  Dash:   ()=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" width="20" height="20"><rect x="3" y="3" width="18" height="18" rx="3"/><path d="M8 12h8M12 8v8"/></svg>,
  List:   ()=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="20" height="20"><path d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2"/></svg>,
  Upload: ()=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="20" height="20"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M17 8l-5-5-5 5M12 3v12"/></svg>,
  Trash:  ()=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="15" height="15"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6M10 11v6M14 11v6M9 6V4h6v2"/></svg>,
  Person: ()=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="16" height="16"><circle cx="12" cy="8" r="4"/><path d="M4 20c0-4 3.6-7 8-7s8 3 8 7"/></svg>,
  Edit:   ()=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="15" height="15"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>,
  Bank:   ()=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="20" height="20"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>,
  Users:  ()=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="20" height="20"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87M16 3.13a4 4 0 010 7.75"/></svg>,
  Phone:  ()=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="14" height="14"><path d="M22 16.92v3a2 2 0 01-2.18 2 19.79 19.79 0 01-8.63-3.07A19.5 19.5 0 013.07 9.8a19.79 19.79 0 01-3.07-8.67A2 2 0 012 .99h3a2 2 0 012 1.72c.127.96.361 1.903.7 2.81a2 2 0 01-.45 2.11L6.09 8.91a16 16 0 006 6l1.27-1.27a2 2 0 012.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0122 16.92z"/></svg>,
  Map:    ()=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="14" height="14"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0118 0z"/><circle cx="12" cy="10" r="3"/></svg>,
  X:      ()=><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" width="16" height="16"><path d="M18 6L6 18M6 6l12 12"/></svg>,
};

// ── App Root ──────────────────────────────────────────────────────────────────
export default function App() {
  const [tab, setTab] = useState("novo");
  const [transactions, setTransactions] = useState([]);
  const [clientNames, setClientNames] = useState(["Geral"]);
  const [clientDB, setClientDB] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    (async () => {
      const [t, c, db] = await Promise.all([
        store.get("cf_transactions"),
        store.get("cf_clients"),
        store.get("cf_clientdb"),
      ]);
      setTransactions(t || []);
      setClientNames(c || ["Geral"]);
      setClientDB(db || []);
      setLoading(false);
    })();
  }, []);

  const addTx = useCallback(async (tx) => {
    const updated = [tx, ...transactions];
    setTransactions(updated); await store.set("cf_transactions", updated);
    if (tx.cliente && !clientNames.includes(tx.cliente)) {
      const nc = [...clientNames, tx.cliente]; setClientNames(nc); await store.set("cf_clients", nc);
    }
  }, [transactions, clientNames]);

  const addManyTx = useCallback(async (list) => {
    const updated = [...list, ...transactions];
    setTransactions(updated); await store.set("cf_transactions", updated);
    const nc = [...clientNames];
    list.forEach(tx => { if (tx.cliente && !nc.includes(tx.cliente)) nc.push(tx.cliente); });
    if (nc.length !== clientNames.length) { setClientNames(nc); await store.set("cf_clients", nc); }
  }, [transactions, clientNames]);

  const deleteTx = useCallback(async (id) => {
    const updated = transactions.filter(t => t.id !== id);
    setTransactions(updated); await store.set("cf_transactions", updated);
  }, [transactions]);

  const saveClientDB = useCallback(async (list) => {
    setClientDB(list); await store.set("cf_clientdb", list);
    const names = ["Geral", ...list.map(c => c.nome)];
    setClientNames(names); await store.set("cf_clients", names);
  }, []);

  if (loading) return (
    <div style={S.loader}><div style={S.spin}/><span style={{color:"#94a3b8",fontFamily:"'DM Sans',sans-serif"}}>Carregando...</span></div>
  );

  const NAV = [
    { id:"novo",      label:"Lançar",    icon:<Ico.Plus/> },
    { id:"clientes",  label:"Clientes",  icon:<Ico.Users/> },
    { id:"extratos",  label:"Extratos",  icon:<Ico.Bank/> },
    { id:"importar",  label:"Importar",  icon:<Ico.Upload/> },
    { id:"dashboard", label:"Dashboard", icon:<Ico.Dash/> },
    { id:"historico", label:"Histórico", icon:<Ico.List/> },
  ];

  return (
    <div style={S.root}>
      <style>{css}</style>
      <header style={S.header}>
        <div style={S.headerInner}>
          <div style={S.logo}><span style={S.logoIcon}>◈</span><span style={S.logoText}>Caixa</span></div>
          <nav style={S.nav}>
            {NAV.map(n => (
              <button key={n.id} style={{...S.navBtn,...(tab===n.id?S.navActive:{})}} onClick={()=>setTab(n.id)}>
                {n.icon}<span style={S.navLabel}>{n.label}</span>
              </button>
            ))}
          </nav>
        </div>
      </header>
      <main style={S.main}>
        {tab==="novo"      && <NovoLancamento onAdd={addTx} clientNames={clientNames} clientDB={clientDB}/>}
        {tab==="clientes"  && <Clientes clientDB={clientDB} onSave={saveClientDB}/>}
        {tab==="extratos"  && <Extratos onAddMany={addManyTx} onDone={()=>setTab("historico")}/>}
        {tab==="importar"  && <Importar onAddMany={addManyTx} onDone={()=>setTab("historico")}/>}
        {tab==="dashboard" && <Dashboard transactions={transactions} clientNames={clientNames}/>}
        {tab==="historico" && <Historico transactions={transactions} onDelete={deleteTx} clientNames={clientNames}/>}
      </main>
    </div>
  );
}

// ── Clientes ──────────────────────────────────────────────────────────────────
function Clientes({ clientDB, onSave }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [search, setSearch] = useState("");
  const [form, setForm] = useState({ nome:"", telefone:"", endereco:"", obs:"" });
  const setF = (k,v) => setForm(f=>({...f,[k]:v}));

  const openNew  = () => { setForm({nome:"",telefone:"",endereco:"",obs:""}); setEditing(null); setShowForm(true); };
  const openEdit = (c) => { setForm({...c}); setEditing(c.id); setShowForm(true); };

  const handleSave = async () => {
    if (!form.nome.trim()) return;
    let updated;
    if (editing) {
      updated = clientDB.map(c => c.id===editing ? {...form, id:editing} : c);
    } else {
      updated = [...clientDB, {...form, id:uid()}];
    }
    await onSave(updated);
    setShowForm(false);
  };

  const handleDelete = async (id) => {
    await onSave(clientDB.filter(c=>c.id!==id));
  };

  const filtered = clientDB.filter(c => c.nome.toLowerCase().includes(search.toLowerCase()));

  if (showForm) return (
    <div style={S.card} className="fade-in">
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"20px"}}>
        <h2 style={{...S.cardTitle,marginBottom:0}}>{editing?"Editar":"Novo"} Cliente</h2>
        <button style={S.iconBtn} onClick={()=>setShowForm(false)}><Ico.X/></button>
      </div>
      {[
        { key:"nome",      label:"Nome *",      placeholder:"Nome completo",        type:"text" },
        { key:"telefone",  label:"Telefone",     placeholder:"(00) 00000-0000",      type:"tel"  },
        { key:"endereco",  label:"Endereço",     placeholder:"Rua, número, bairro",  type:"text" },
        { key:"obs",       label:"Observações",  placeholder:"Anotações livres",     type:"text" },
      ].map(f=>(
        <div key={f.key} style={S.field}>
          <label style={S.label}>{f.label}</label>
          <input style={S.input} type={f.type} placeholder={f.placeholder} value={form[f.key]} onChange={e=>setF(f.key,e.target.value)}/>
        </div>
      ))}
      <div style={{display:"flex",gap:"10px",marginTop:"8px"}}>
        <button style={S.btnSec} onClick={()=>setShowForm(false)}>Cancelar</button>
        <button style={{...S.btnPrimary,flex:1}} onClick={handleSave}>✓ Salvar Cliente</button>
      </div>
    </div>
  );

  return (
    <div className="fade-in">
      <div style={{...S.card,marginBottom:"12px"}}>
        <div style={{display:"flex",gap:"10px",alignItems:"center"}}>
          <input style={{...S.input,flex:1,marginBottom:0}} placeholder="🔍 Buscar cliente..." value={search} onChange={e=>setSearch(e.target.value)}/>
          <button style={{...S.btnPrimary,whiteSpace:"nowrap",marginTop:0}} onClick={openNew}>+ Novo</button>
        </div>
      </div>

      {filtered.length===0 ? (
        <div style={S.empty}>{search ? "Nenhum cliente encontrado" : "Nenhum cliente cadastrado ainda"}</div>
      ) : (
        <div style={{display:"flex",flexDirection:"column",gap:"8px"}}>
          {filtered.map(c=>(
            <div key={c.id} style={S.clientCard}>
              <div style={S.clientAvatar}>{c.nome[0].toUpperCase()}</div>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontWeight:"700",color:"#f1f5f9",marginBottom:"4px"}}>{c.nome}</div>
                {c.telefone && <div style={S.clientMeta}><Ico.Phone/> {c.telefone}</div>}
                {c.endereco && <div style={S.clientMeta}><Ico.Map/> {c.endereco}</div>}
                {c.obs && <div style={{fontSize:"0.72rem",color:"#64748b",marginTop:"4px",fontStyle:"italic"}}>{c.obs}</div>}
              </div>
              <div style={{display:"flex",gap:"6px",flexShrink:0}}>
                <button style={S.iconBtn} onClick={()=>openEdit(c)}><Ico.Edit/></button>
                <button style={{...S.iconBtn,color:"#ef4444"}} onClick={()=>handleDelete(c.id)}><Ico.Trash/></button>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ── Extratos ──────────────────────────────────────────────────────────────────
function Extratos({ onAddMany, onDone }) {
  const [banco, setBanco] = useState(null); // "infinity" | "pagbank"
  const [step, setStep]   = useState("pick"); // pick | upload | map | preview | done
  const [rows, setRows]   = useState([]);
  const [headers, setHeaders] = useState([]);
  const [mapping, setMapping] = useState({ data:"", descricao:"", valor:"", tipo:"" });
  const [preview, setPreview] = useState([]);
  const [imported, setImported] = useState(0);
  const [dragOver, setDragOver] = useState(false);
  const fileRef = useRef();

  const BANKS = [
    {
      id: "infinity", name: "Infinity",
      color: "#6366f1",
      logo: <div style={{width:"44px",height:"44px",borderRadius:"12px",background:"#1e1b4b",display:"flex",alignItems:"center",justifyContent:"center"}}>
        <svg viewBox="0 0 44 22" width="32" height="16" fill="none">
          <path d="M11 11C11 7 14 4 18 4C22 4 24 8 26.5 11C29 14 31 18 35 18C39 18 42 15 42 11C42 7 39 4 35 4C31 4 29 8 26.5 11C24 14 22 18 18 18C14 18 11 15 11 11Z" stroke="#818cf8" strokeWidth="2.5" fill="none"/>
          <path d="M11 11C11 15 8 18 4 18" stroke="#818cf8" strokeWidth="2.5" strokeLinecap="round"/>
          <circle cx="4" cy="11" r="3" fill="#818cf8"/>
        </svg>
      </div>
    },
    {
      id: "pagbank", name: "PagBank",
      color: "#f59e0b",
      logo: <div style={{width:"44px",height:"44px",borderRadius:"12px",background:"#fef3c7",display:"flex",alignItems:"center",justifyContent:"center"}}>
        <svg viewBox="0 0 44 44" width="38" height="38" fill="none">
          <circle cx="22" cy="22" r="20" fill="#f59e0b"/>
          <text x="22" y="29" textAnchor="middle" fontSize="18" fontWeight="900" fontFamily="Arial" fill="white">P</text>
        </svg>
      </div>
    },
  ];

  const parseFile = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type:"array" });
        let best = wb.SheetNames[0], bestN = 0;
        wb.SheetNames.forEach(n => { const j = XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,defval:""}); if(j.length>bestN){bestN=j.length;best=n;} });
        const json = XLSX.utils.sheet_to_json(wb.Sheets[best], { header:1, defval:"" });
        let hi = 0;
        for (let i=0;i<Math.min(10,json.length);i++) { if(json[i].filter(c=>c!=="").length>=2){hi=i;break;} }
        const hdrs = json[hi].map((h,i)=>h!==""?String(h):`Col ${i+1}`);
        const dataRows = json.slice(hi+1).filter(r=>r.some(c=>c!==""));
        setHeaders(hdrs); setRows(dataRows);
        // auto-map
        const am = {data:"",descricao:"",valor:"",tipo:""};
        hdrs.forEach((h,i)=>{
          const hl=String(h).toLowerCase();
          if(!am.data&&(hl.includes("data")||hl.includes("date"))) am.data=String(i);
          if(!am.descricao&&(hl.includes("desc")||hl.includes("hist")||hl.includes("memo")||hl.includes("lançamento"))) am.descricao=String(i);
          if(!am.valor&&(hl.includes("valor")||hl.includes("value")||hl.includes("amount")||hl.includes("montante"))) am.valor=String(i);
          if(!am.tipo&&(hl.includes("tipo")||hl.includes("natureza")||hl.includes("débito")||hl.includes("crédito")||hl.includes("d/c"))) am.tipo=String(i);
        });
        setMapping(am); setStep("map");
      } catch { alert("Erro ao ler arquivo. Use .xlsx, .xls ou .csv."); }
    };
    reader.readAsArrayBuffer(file);
  };

  const buildPreview = () => {
    const get = (row,key) => mapping[key]!==""?row[parseInt(mapping[key])]:"";
    const built = rows.slice(0,500).map((row,i)=>{
      const rawVal = get(row,"valor");
      let valor = parseFloat(String(rawVal).replace(/[R$\s.]/g,"").replace(",","."));
      if (isNaN(valor)) valor=0;
      const tipoRaw = get(row,"tipo");
      const tipo = Math.abs(valor)>0 && valor>0 ? "receita" : (valor<0?"despesa":guessTipo(tipoRaw));
      return {
        id: uid(),
        data: excelDateToISO(get(row,"data")),
        tipo,
        valor: Math.abs(valor),
        categoria: "Outros",
        pagamento: "pix",
        cliente: "Geral",
        descricao: String(get(row,"descricao")||"").trim(),
        conta: banco,
      };
    }).filter(r=>r.valor>0);
    setPreview(built); setStep("preview");
  };

  const handleImport = async () => {
    await onAddMany(preview);
    setImported(preview.length); setStep("done");
  };

  const reset = () => { setBanco(null); setStep("pick"); setRows([]); setHeaders([]); setPreview([]); };

  // Step: pick bank
  if (step==="pick") return (
    <div className="fade-in">
      <div style={{...S.card,marginBottom:"12px"}}>
        <h2 style={S.cardTitle}>Upload de Extrato</h2>
        <p style={{color:"#94a3b8",fontSize:"0.85rem",marginBottom:"20px"}}>Selecione o banco para fazer o upload do extrato:</p>
        <div style={{display:"flex",gap:"12px"}}>
          {BANKS.map(b=>(
            <button key={b.id} style={{...S.bankPickBtn,...(banco===b.id?{borderColor:b.color,background:b.color+"18"}:{})}}
              onClick={()=>{ setBanco(b.id); setStep("upload"); }}>
              {b.logo}
              <span style={{fontWeight:"700",color:"#f1f5f9",marginTop:"8px"}}>{b.name}</span>
            </button>
          ))}
        </div>
      </div>
      <div style={S.tipBox}>
        <div style={{color:"#60a5fa",fontWeight:"700",fontSize:"0.8rem",marginBottom:"6px"}}>💡 Formatos aceitos</div>
        <div style={{color:"#94a3b8",fontSize:"0.75rem",lineHeight:"1.8"}}>.xlsx · .xls · .csv exportado do internet banking</div>
      </div>
    </div>
  );

  const selectedBank = BANKS.find(b=>b.id===banco);

  if (step==="upload") return (
    <div style={S.card} className="fade-in">
      <div style={{display:"flex",alignItems:"center",gap:"12px",marginBottom:"20px"}}>
        <button style={S.iconBtn} onClick={reset}><Ico.X/></button>
        {selectedBank.logo}
        <h2 style={{...S.cardTitle,marginBottom:0}}>Extrato {selectedBank.name}</h2>
      </div>
      <div style={{...S.dropZone,...(dragOver?S.dropZoneActive:{})}}
        onDragOver={e=>{e.preventDefault();setDragOver(true);}}
        onDragLeave={()=>setDragOver(false)}
        onDrop={e=>{e.preventDefault();setDragOver(false);if(e.dataTransfer.files[0])parseFile(e.dataTransfer.files[0]);}}
        onClick={()=>fileRef.current.click()}>
        <div style={{fontSize:"2.5rem",marginBottom:"10px"}}>📄</div>
        <div style={{fontWeight:"700",color:"#e2e8f0",marginBottom:"4px"}}>Arraste o extrato aqui</div>
        <div style={{color:"#64748b",fontSize:"0.8rem"}}>ou clique para selecionar</div>
        <div style={{color:"#475569",fontSize:"0.7rem",marginTop:"6px"}}>.xlsx · .xls · .csv</div>
        <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>{if(e.target.files[0])parseFile(e.target.files[0]);}}/>
      </div>
    </div>
  );

  if (step==="map") return (
    <div style={S.card} className="fade-in">
      <h2 style={S.cardTitle}>Mapear Colunas — {selectedBank.name}</h2>
      <p style={{color:"#94a3b8",fontSize:"0.82rem",marginBottom:"18px"}}>{rows.length} linhas encontradas. Confirme as colunas:</p>
      <div style={{display:"flex",flexDirection:"column",gap:"12px",marginBottom:"20px"}}>
        {[
          {key:"data",      label:"Data *"},
          {key:"descricao", label:"Descrição / Histórico"},
          {key:"valor",     label:"Valor *"},
          {key:"tipo",      label:"Tipo (Déb/Créd)"},
        ].map(f=>(
          <div key={f.key} style={S.mapRow}>
            <div style={S.mapLabel}>{f.label}</div>
            <select style={{...S.input,flex:1}} value={mapping[f.key]} onChange={e=>setMapping(m=>({...m,[f.key]:e.target.value}))}>
              <option value="">— Ignorar —</option>
              {headers.map((h,i)=><option key={i} value={String(i)}>{h}</option>)}
            </select>
          </div>
        ))}
      </div>
      <div style={{display:"flex",gap:"10px"}}>
        <button style={S.btnSec} onClick={()=>setStep("upload")}>← Voltar</button>
        <button style={{...S.btnPrimary,flex:1}} onClick={buildPreview} disabled={!mapping.valor||!mapping.data}>Pré-visualizar →</button>
      </div>
    </div>
  );

  if (step==="preview") return (
    <div className="fade-in">
      <div style={{...S.card,marginBottom:"12px"}}>
        <h2 style={S.cardTitle}>Pré-visualização — {selectedBank.name}</h2>
        <div style={{display:"flex",gap:"8px",marginBottom:"16px",flexWrap:"wrap"}}>
          <div style={S.prevStat}><span style={{color:"#22c55e",fontWeight:"700"}}>{preview.filter(r=>r.tipo==="receita").length}</span><span style={{color:"#64748b",fontSize:"0.7rem"}}>Créditos</span></div>
          <div style={S.prevStat}><span style={{color:"#f87171",fontWeight:"700"}}>{preview.filter(r=>r.tipo==="despesa").length}</span><span style={{color:"#64748b",fontSize:"0.7rem"}}>Débitos</span></div>
          <div style={S.prevStat}><span style={{color:"#60a5fa",fontWeight:"700",fontSize:"0.9rem"}}>{fmt(preview.filter(r=>r.tipo==="receita").reduce((s,r)=>s+r.valor,0)-preview.filter(r=>r.tipo==="despesa").reduce((s,r)=>s+r.valor,0))}</span><span style={{color:"#64748b",fontSize:"0.7rem"}}>Saldo</span></div>
        </div>
        <div style={{display:"flex",gap:"10px"}}>
          <button style={S.btnSec} onClick={()=>setStep("map")}>← Ajustar</button>
          <button style={{...S.btnPrimary,flex:1}} onClick={handleImport}>✓ Importar {preview.length} lançamentos</button>
        </div>
      </div>
      <div style={{display:"flex",flexDirection:"column",gap:"6px",maxHeight:"380px",overflowY:"auto"}}>
        {preview.slice(0,60).map((r,i)=>(
          <div key={i} style={S.txRow}>
            <div style={{...S.txBar,background:r.tipo==="receita"?"#22c55e":"#f87171"}}/>
            <div style={S.txBody}>
              <div style={S.txTop}>
                <span style={{...S.txCat,fontSize:"0.8rem"}}>{r.descricao||"—"}</span>
                <span style={{...S.txVal,color:r.tipo==="receita"?"#22c55e":"#f87171"}}>{r.tipo==="receita"?"+":"-"}{fmt(r.valor)}</span>
              </div>
              <span style={{color:"#64748b",fontSize:"0.7rem"}}>{fmtDate(r.data)}</span>
            </div>
          </div>
        ))}
        {preview.length>60&&<div style={S.empty}>… e mais {preview.length-60} lançamentos</div>}
      </div>
    </div>
  );

  return (
    <div style={{...S.card,textAlign:"center",padding:"48px 24px"}} className="fade-in">
      <div style={{fontSize:"4rem",marginBottom:"16px"}}>🎉</div>
      <h2 style={{fontSize:"1.4rem",fontWeight:"800",color:"#f1f5f9",marginBottom:"8px"}}>Extrato importado!</h2>
      <p style={{color:"#94a3b8",marginBottom:"24px"}}>{imported} lançamentos do {selectedBank.name} adicionados.</p>
      <div style={{display:"flex",gap:"10px"}}>
        <button style={{...S.btnSec,flex:1}} onClick={reset}>Importar outro</button>
        <button style={{...S.btnPrimary,flex:1}} onClick={onDone}>Ver Histórico →</button>
      </div>
    </div>
  );
}

// ── Importar Planilha ─────────────────────────────────────────────────────────
function Importar({ onAddMany, onDone }) {
  const [step, setStep] = useState("upload");
  const [rows, setRows] = useState([]);
  const [headers, setHeaders] = useState([]);
  const [mapping, setMapping] = useState({ data:"", tipo:"", valor:"", categoria:"", pagamento:"", cliente:"", descricao:"" });
  const [preview, setPreview] = useState([]);
  const [imported, setImported] = useState(0);
  const [dragOver, setDragOver] = useState(false);
  const fileRef = useRef();

  const FIELDS = [
    {key:"data",      label:"Data",            required:true},
    {key:"tipo",      label:"Entrada / Saída",  required:true},
    {key:"valor",     label:"Valor (R$)",        required:true},
    {key:"categoria", label:"Categoria",         required:false},
    {key:"pagamento", label:"Forma de Pagamento",required:false},
    {key:"cliente",   label:"Cliente",           required:false},
    {key:"descricao", label:"Descrição",         required:false},
  ];

  const parseFile = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result),{type:"array"});
        let best=wb.SheetNames[0],bestN=0;
        wb.SheetNames.forEach(n=>{const j=XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,defval:""});if(j.length>bestN){bestN=j.length;best=n;}});
        const json = XLSX.utils.sheet_to_json(wb.Sheets[best],{header:1,defval:""});
        let hi=0;
        for(let i=0;i<Math.min(10,json.length);i++){if(json[i].filter(c=>c!=="").length>=3){hi=i;break;}}
        const hdrs=json[hi].map((h,i)=>h!==""?String(h):`Col ${i+1}`);
        const dataRows=json.slice(hi+1).filter(r=>r.some(c=>c!==""));
        setHeaders(hdrs); setRows(dataRows);
        const am={data:"",tipo:"",valor:"",categoria:"",pagamento:"",cliente:"",descricao:""};
        hdrs.forEach((h,i)=>{
          const hl=String(h).toLowerCase();
          if(!am.data&&(hl.includes("data")||hl.includes("date"))) am.data=String(i);
          if(!am.tipo&&(hl.includes("tipo")||hl.includes("receita")||hl.includes("despesa")||hl.includes("entrada"))) am.tipo=String(i);
          if(!am.valor&&(hl.includes("valor")||hl.includes("value")||hl.includes("amount"))) am.valor=String(i);
          if(!am.categoria&&(hl.includes("categ")||hl.includes("classif"))) am.categoria=String(i);
          if(!am.pagamento&&(hl.includes("pgto")||hl.includes("pagamento")||hl.includes("forma"))) am.pagamento=String(i);
          if(!am.cliente&&(hl.includes("cliente")||hl.includes("client")||hl.includes("item"))) am.cliente=String(i);
          if(!am.descricao&&(hl.includes("desc")||hl.includes("obs"))) am.descricao=String(i);
        });
        setMapping(am); setStep("map");
      } catch { alert("Erro ao ler arquivo."); }
    };
    reader.readAsArrayBuffer(file);
  };

  const buildPreview = () => {
    const get=(row,key)=>mapping[key]!==""?row[parseInt(mapping[key])]:"";
    const built=rows.slice(0,300).map((row,i)=>{
      const rawVal=get(row,"valor");
      let valor=parseFloat(String(rawVal).replace(/[R$\s.]/g,"").replace(",","."));
      if(isNaN(valor)) valor=0;
      return {id:uid(),data:excelDateToISO(get(row,"data")),tipo:guessTipo(get(row,"tipo")),valor:Math.abs(valor),categoria:String(get(row,"categoria")||"Outros").trim(),pagamento:guessPayment(get(row,"pagamento")),cliente:String(get(row,"cliente")||"Geral").trim()||"Geral",descricao:String(get(row,"descricao")||"").trim()};
    }).filter(r=>r.valor>0);
    setPreview(built); setStep("preview");
  };

  const handleImport = async () => { await onAddMany(preview); setImported(preview.length); setStep("done"); };

  if (step==="upload") return (
    <div style={S.card} className="fade-in">
      <h2 style={S.cardTitle}>Importar Planilha</h2>
      <p style={{color:"#94a3b8",fontSize:"0.85rem",marginBottom:"20px"}}>Importe seu Excel ou CSV com lançamentos anteriores.</p>
      <div style={{...S.dropZone,...(dragOver?S.dropZoneActive:{})}}
        onDragOver={e=>{e.preventDefault();setDragOver(true);}} onDragLeave={()=>setDragOver(false)}
        onDrop={e=>{e.preventDefault();setDragOver(false);if(e.dataTransfer.files[0])parseFile(e.dataTransfer.files[0]);}}
        onClick={()=>fileRef.current.click()}>
        <div style={{fontSize:"2.5rem",marginBottom:"10px"}}>📂</div>
        <div style={{fontWeight:"700",color:"#e2e8f0",marginBottom:"4px"}}>Arraste o arquivo aqui</div>
        <div style={{color:"#64748b",fontSize:"0.8rem"}}>ou clique para selecionar</div>
        <div style={{color:"#475569",fontSize:"0.7rem",marginTop:"6px"}}>.xlsx · .xls · .csv</div>
        <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>{if(e.target.files[0])parseFile(e.target.files[0]);}}/>
      </div>
    </div>
  );

  if (step==="map") return (
    <div style={S.card} className="fade-in">
      <h2 style={S.cardTitle}>Mapear Colunas</h2>
      <p style={{color:"#94a3b8",fontSize:"0.82rem",marginBottom:"18px"}}>{rows.length} linhas encontradas.</p>
      <div style={{display:"flex",flexDirection:"column",gap:"12px",marginBottom:"20px"}}>
        {FIELDS.map(f=>(
          <div key={f.key} style={S.mapRow}>
            <div style={S.mapLabel}>{f.label}{f.required&&<span style={{color:"#f87171"}}> *</span>}</div>
            <select style={{...S.input,flex:1}} value={mapping[f.key]} onChange={e=>setMapping(m=>({...m,[f.key]:e.target.value}))}>
              <option value="">— Ignorar —</option>
              {headers.map((h,i)=><option key={i} value={String(i)}>{h}</option>)}
            </select>
          </div>
        ))}
      </div>
      <div style={{display:"flex",gap:"10px"}}>
        <button style={S.btnSec} onClick={()=>setStep("upload")}>← Voltar</button>
        <button style={{...S.btnPrimary,flex:1}} onClick={buildPreview} disabled={!mapping.valor||!mapping.data}>Pré-visualizar →</button>
      </div>
    </div>
  );

  if (step==="preview") return (
    <div className="fade-in">
      <div style={{...S.card,marginBottom:"12px"}}>
        <h2 style={S.cardTitle}>Pré-visualização</h2>
        <div style={{display:"flex",gap:"8px",marginBottom:"16px",flexWrap:"wrap"}}>
          <div style={S.prevStat}><span style={{color:"#22c55e",fontWeight:"700"}}>{preview.filter(r=>r.tipo==="receita").length}</span><span style={{color:"#64748b",fontSize:"0.7rem"}}>Entradas</span></div>
          <div style={S.prevStat}><span style={{color:"#f87171",fontWeight:"700"}}>{preview.filter(r=>r.tipo==="despesa").length}</span><span style={{color:"#64748b",fontSize:"0.7rem"}}>Saídas</span></div>
          <div style={S.prevStat}><span style={{color:"#60a5fa",fontWeight:"700",fontSize:"0.9rem"}}>{fmt(preview.filter(r=>r.tipo==="receita").reduce((s,r)=>s+r.valor,0)-preview.filter(r=>r.tipo==="despesa").reduce((s,r)=>s+r.valor,0))}</span><span style={{color:"#64748b",fontSize:"0.7rem"}}>Saldo</span></div>
        </div>
        <div style={{display:"flex",gap:"10px"}}>
          <button style={S.btnSec} onClick={()=>setStep("map")}>← Ajustar</button>
          <button style={{...S.btnPrimary,flex:1}} onClick={handleImport}>✓ Importar {preview.length} lançamentos</button>
        </div>
      </div>
      <div style={{display:"flex",flexDirection:"column",gap:"6px",maxHeight:"380px",overflowY:"auto"}}>
        {preview.slice(0,50).map((r,i)=>{
          const pm=PAYMENT_METHODS.find(p=>p.id===r.pagamento);
          return (
            <div key={i} style={S.txRow}>
              <div style={{...S.txBar,background:r.tipo==="receita"?"#22c55e":"#f87171"}}/>
              <div style={S.txBody}>
                <div style={S.txTop}><span style={S.txCat}>{r.categoria}</span><span style={{...S.txVal,color:r.tipo==="receita"?"#22c55e":"#f87171"}}>{r.tipo==="receita"?"+":"-"}{fmt(r.valor)}</span></div>
                <div style={S.txMeta}>
                  <span style={{background:pm?.color+"22",color:pm?.color,padding:"1px 6px",borderRadius:"4px",fontSize:"0.65rem"}}>{pm?.icon} {pm?.label}</span>
                  <span style={{color:"#64748b",fontSize:"0.7rem"}}>{fmtDate(r.data)}</span>
                </div>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );

  return (
    <div style={{...S.card,textAlign:"center",padding:"48px 24px"}} className="fade-in">
      <div style={{fontSize:"4rem",marginBottom:"16px"}}>🎉</div>
      <h2 style={{fontSize:"1.4rem",fontWeight:"800",color:"#f1f5f9",marginBottom:"8px"}}>Importação concluída!</h2>
      <p style={{color:"#94a3b8",marginBottom:"24px"}}>{imported} lançamentos importados com sucesso.</p>
      <button style={S.btnPrimary} onClick={onDone}>Ver Histórico →</button>
    </div>
  );
}

// ── Novo Lançamento ───────────────────────────────────────────────────────────
function NovoLancamento({ onAdd, clientNames, clientDB }) {
  const [tipo, setTipo] = useState("receita");
  const [form, setForm] = useState({ data:today(), valor:"", categoria:"", pagamento:"pix", cliente:"", descricao:"" });
  const [success, setSuccess] = useState(false);
  const [clientSearch, setClientSearch] = useState("");
  const [showClientList, setShowClientList] = useState(false);
  const set = (k,v) => setForm(f=>({...f,[k]:v}));

  const handleSubmit = async () => {
    if (!form.valor||!form.categoria) return;
    const val = parseFloat(form.valor.replace(",","."));
    if (isNaN(val)||val<=0) return;
    await onAdd({id:uid(),tipo,...form,valor:val,cliente:form.cliente||"Geral"});
    setForm({data:today(),valor:"",categoria:"",pagamento:"pix",cliente:"",descricao:""});
    setClientSearch(""); setSuccess(true); setTimeout(()=>setSuccess(false),2000);
  };

  const cats = tipo==="receita"?CAT_REC:CAT_DESP;
  const filteredClients = clientDB.filter(c=>c.nome.toLowerCase().includes(clientSearch.toLowerCase()));

  return (
    <div style={S.card} className="fade-in">
      <h2 style={S.cardTitle}>Novo Lançamento</h2>
      <div style={S.tipoRow}>
        <button style={{...S.tipoBtn,...(tipo==="receita"?S.tipoBtnRec:S.tipoBtnOff)}} onClick={()=>{setTipo("receita");set("categoria","");}}>↑ Entrada</button>
        <button style={{...S.tipoBtn,...(tipo==="despesa"?S.tipoBtnDesp:S.tipoBtnOff)}} onClick={()=>{setTipo("despesa");set("categoria","");}}>↓ Saída</button>
      </div>
      <div style={S.field}>
        <label style={S.label}>Valor (R$)</label>
        <div style={S.valueWrap}>
          <span style={S.valuePrefix}>R$</span>
          <input style={{...S.input,paddingLeft:"3rem",fontSize:"1.5rem",fontWeight:"700",color:tipo==="receita"?"#22c55e":"#f87171"}}
            type="number" inputMode="decimal" placeholder="0,00" value={form.valor} onChange={e=>set("valor",e.target.value)}/>
        </div>
      </div>
      <div style={S.field}>
        <label style={S.label}>Data</label>
        <input style={S.input} type="date" value={form.data} onChange={e=>set("data",e.target.value)}/>
      </div>
      <div style={S.field}>
        <label style={S.label}>Forma de Pagamento</label>
        <div style={S.payRow}>
          {PAYMENT_METHODS.map(pm=>(
            <button key={pm.id} style={{...S.payBtn,...(form.pagamento===pm.id?{...S.payBtnActive,borderColor:pm.color,color:pm.color,background:pm.color+"18"}:{})}} onClick={()=>set("pagamento",pm.id)}>
              <span>{pm.icon}</span><span style={{fontSize:"0.7rem",marginTop:"2px"}}>{pm.label}</span>
            </button>
          ))}
        </div>
      </div>
      <div style={S.field}>
        <label style={S.label}>Categoria</label>
        <select style={S.input} value={form.categoria} onChange={e=>set("categoria",e.target.value)}>
          <option value="">Selecione...</option>
          {cats.map(c=><option key={c} value={c}>{c}</option>)}
        </select>
      </div>
      <div style={S.field}>
        <label style={S.label}>Cliente</label>
        <div style={{position:"relative"}}>
          <input style={S.input} placeholder="Buscar ou digitar cliente..." value={clientSearch||form.cliente}
            onChange={e=>{ setClientSearch(e.target.value); set("cliente",e.target.value); setShowClientList(true); }}
            onFocus={()=>setShowClientList(true)}
            onBlur={()=>setTimeout(()=>setShowClientList(false),150)}/>
          {showClientList && clientDB.length>0 && (
            <div style={S.dropdown}>
              {filteredClients.slice(0,6).map(c=>(
                <div key={c.id} style={S.dropItem} onMouseDown={()=>{ set("cliente",c.nome); setClientSearch(c.nome); setShowClientList(false); }}>
                  <div style={{fontWeight:"600",color:"#f1f5f9",fontSize:"0.85rem"}}>{c.nome}</div>
                  {c.telefone&&<div style={{color:"#64748b",fontSize:"0.72rem"}}>{c.telefone}</div>}
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
      <div style={S.field}>
        <label style={S.label}>Descrição (opcional)</label>
        <input style={S.input} placeholder="Ex: Painel TV + instalação" value={form.descricao} onChange={e=>set("descricao",e.target.value)}/>
      </div>
      <button style={{...S.btnPrimary,...(success?S.btnSuccess:{})}} onClick={handleSubmit} className="press">
        {success?"✓ Lançado!":`Registrar ${tipo==="receita"?"Entrada":"Saída"}`}
      </button>
    </div>
  );
}

// ── Wallet Panel ──────────────────────────────────────────────────────────────
function WalletPanel() {
  const [wallets, setWallets] = useState({infinity:0,pagbank:0,dinheiro:0});
  const [editing, setEditing] = useState(null);
  const [tempVal, setTempVal] = useState("");
  const [loaded, setLoaded] = useState(false);
  useEffect(()=>{ store.get("cf_wallets").then(w=>{ setWallets(w||{infinity:0,pagbank:0,dinheiro:0}); setLoaded(true); }); },[]);
  const startEdit=(k)=>{ setEditing(k); setTempVal(String(wallets[k]||"")); };
  const confirmEdit=async(k)=>{ const v=parseFloat(tempVal.replace(",",".")); const u={...wallets,[k]:isNaN(v)?0:v}; setWallets(u); await store.set("cf_wallets",u); setEditing(null); };
  const total=wallets.infinity+wallets.pagbank+wallets.dinheiro;
  const WALLETS=[
    {key:"infinity",name:"Infinity",color:"#6366f1",logo:<div style={{width:"36px",height:"36px",borderRadius:"10px",background:"#1e1b4b",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}><svg viewBox="0 0 40 20" width="26" height="13" fill="none"><path d="M10 10C10 6 13 3 17 3C21 3 23 7 25 10C27 13 29 17 33 17C37 17 40 14 40 10C40 6 37 3 33 3C29 3 27 7 25 10C23 13 21 17 17 17C13 17 10 14 10 10Z" stroke="#818cf8" strokeWidth="2.5" fill="none"/><path d="M10 10C10 14 7 17 3 17" stroke="#818cf8" strokeWidth="2.5" strokeLinecap="round" fill="none"/><circle cx="3" cy="10" r="2.5" fill="#818cf8"/></svg></div>},
    {key:"pagbank",name:"PagBank",color:"#f59e0b",logo:<div style={{width:"36px",height:"36px",borderRadius:"10px",background:"#fef3c7",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}><svg viewBox="0 0 36 36" width="30" height="30" fill="none"><circle cx="18" cy="18" r="16" fill="#f59e0b"/><text x="18" y="24" textAnchor="middle" fontSize="15" fontWeight="900" fontFamily="Arial" fill="white">P</text></svg></div>},
    {key:"dinheiro",name:"Dinheiro",color:"#22c55e",logo:<div style={{width:"36px",height:"36px",borderRadius:"10px",background:"#052e16",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}><svg viewBox="0 0 36 36" width="26" height="26" fill="none"><rect x="3" y="9" width="30" height="18" rx="3" stroke="#22c55e" strokeWidth="2"/><circle cx="18" cy="18" r="5" stroke="#22c55e" strokeWidth="2"/><circle cx="18" cy="18" r="2" fill="#22c55e"/><circle cx="7" cy="18" r="1.5" fill="#22c55e"/><circle cx="29" cy="18" r="1.5" fill="#22c55e"/></svg></div>},
  ];
  if(!loaded) return null;
  return (
    <div style={S.card}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"14px"}}>
        <h3 style={{...S.sectionTitle,marginBottom:0}}>💰 Saldo por Conta</h3>
        <span style={{fontSize:"0.68rem",color:"#64748b"}}>Toque no valor para editar</span>
      </div>
      <div style={{display:"flex",flexDirection:"column",gap:"10px",marginBottom:"12px"}}>
        {WALLETS.map(w=>(
          <div key={w.key} style={{display:"flex",alignItems:"center",gap:"12px",background:"#1a2235",borderRadius:"12px",padding:"12px 14px"}}>
            {w.logo}
            <div style={{flex:1}}>
              <div style={{fontSize:"0.72rem",color:"#94a3b8",marginBottom:"2px"}}>{w.name}</div>
              {editing===w.key?(
                <div style={{display:"flex",gap:"6px",alignItems:"center"}}>
                  <span style={{color:"#64748b",fontSize:"0.85rem"}}>R$</span>
                  <input autoFocus style={{...S.input,padding:"4px 8px",fontSize:"0.95rem",fontWeight:"700",width:"120px"}} value={tempVal} onChange={e=>setTempVal(e.target.value)} onKeyDown={e=>{if(e.key==="Enter")confirmEdit(w.key);if(e.key==="Escape")setEditing(null);}}/>
                  <button style={{...S.addClientBtn,padding:"4px 10px",fontSize:"0.8rem"}} onClick={()=>confirmEdit(w.key)}>✓</button>
                </div>
              ):(
                <div style={{fontWeight:"800",fontSize:"1.05rem",color:w.color,cursor:"pointer"}} onClick={()=>startEdit(w.key)}>{fmt(wallets[w.key])}</div>
              )}
            </div>
          </div>
        ))}
      </div>
      <div style={{borderTop:"1px solid #1e293b",paddingTop:"12px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <span style={{color:"#94a3b8",fontSize:"0.8rem",fontWeight:"600"}}>TOTAL EM CAIXA</span>
        <span style={{fontWeight:"800",fontSize:"1.15rem",color:"#f1f5f9"}}>{fmt(total)}</span>
      </div>
    </div>
  );
}

// ── Dashboard ─────────────────────────────────────────────────────────────────
function Dashboard({ transactions, clientNames }) {
  const [filterMonth, setFilterMonth] = useState("todos");
  const months=["todos","01","02","03","04","05","06","07","08","09","10","11","12"];
  const mNames=["Todos","Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];
  const filtered=filterMonth==="todos"?transactions:transactions.filter(t=>t.data&&t.data.slice(5,7)===filterMonth);
  const totalRec=filtered.filter(t=>t.tipo==="receita").reduce((s,t)=>s+t.valor,0);
  const totalDesp=filtered.filter(t=>t.tipo==="despesa").reduce((s,t)=>s+t.valor,0);
  const saldo=totalRec-totalDesp;
  const byPay=PAYMENT_METHODS.map(pm=>({...pm,total:filtered.filter(t=>t.pagamento===pm.id).reduce((s,t)=>s+t.valor,0),count:filtered.filter(t=>t.pagamento===pm.id).length})).filter(p=>p.count>0);
  const clientStats=clientNames.map(c=>{
    const txs=filtered.filter(t=>(t.cliente||"Geral")===c);
    const rec=txs.filter(t=>t.tipo==="receita").reduce((s,t)=>s+t.valor,0);
    const desp=txs.filter(t=>t.tipo==="despesa").reduce((s,t)=>s+t.valor,0);
    return {name:c,receitas:rec,despesas:desp,saldo:rec-desp,total:rec+desp};
  }).filter(c=>c.total>0).sort((a,b)=>b.receitas-a.receitas);
  const catDesp={};
  filtered.filter(t=>t.tipo==="despesa").forEach(t=>{catDesp[t.categoria]=(catDesp[t.categoria]||0)+t.valor;});
  const topCats=Object.entries(catDesp).sort((a,b)=>b[1]-a[1]).slice(0,5);
  const maxCat=topCats[0]?.[1]||1;

  return (
    <div style={{display:"flex",flexDirection:"column",gap:"16px"}} className="fade-in">
      <WalletPanel/>
      <div style={S.monthScroll}>
        {months.map((m,i)=>(
          <button key={m} style={{...S.monthBtn,...(filterMonth===m?S.monthBtnActive:{})}} onClick={()=>setFilterMonth(m)}>{mNames[i]}</button>
        ))}
      </div>
      <div style={S.kpiRow}>
        <div style={{...S.kpi,borderTop:"3px solid #22c55e"}}><span style={S.kpiLabel}>Entradas</span><span style={{...S.kpiValue,color:"#22c55e"}}>{fmt(totalRec)}</span></div>
        <div style={{...S.kpi,borderTop:"3px solid #f87171"}}><span style={S.kpiLabel}>Saídas</span><span style={{...S.kpiValue,color:"#f87171"}}>{fmt(totalDesp)}</span></div>
        <div style={{...S.kpi,borderTop:`3px solid ${saldo>=0?"#60a5fa":"#fb923c"}`}}><span style={S.kpiLabel}>Saldo</span><span style={{...S.kpiValue,color:saldo>=0?"#60a5fa":"#fb923c"}}>{fmt(saldo)}</span></div>
      </div>
      {byPay.length>0&&(
        <div style={S.card}>
          <h3 style={S.sectionTitle}>Formas de Pagamento</h3>
          <div style={S.payStatGrid}>
            {byPay.map(pm=>(
              <div key={pm.id} style={{...S.payStat,borderLeft:`3px solid ${pm.color}`}}>
                <span style={{fontSize:"1.3rem"}}>{pm.icon}</span>
                <div><div style={{fontSize:"0.7rem",color:"#94a3b8",textTransform:"uppercase"}}>{pm.label}</div><div style={{fontWeight:"700",color:"#e2e8f0",fontSize:"0.9rem"}}>{fmt(pm.total)}</div><div style={{fontSize:"0.65rem",color:"#64748b"}}>{pm.count} lançamentos</div></div>
              </div>
            ))}
          </div>
        </div>
      )}
      {clientStats.length>0&&(
        <div style={S.card}>
          <h3 style={S.sectionTitle}><Ico.Person/> Por Cliente</h3>
          <div style={{display:"flex",flexDirection:"column",gap:"10px"}}>
            {clientStats.map(c=>(
              <div key={c.name} style={S.clientRow}>
                <div style={S.clientName}><div style={S.clientAvatar}>{c.name[0].toUpperCase()}</div><span>{c.name}</span></div>
                <div style={S.clientStats}>
                  <span style={{color:"#22c55e",fontSize:"0.8rem"}}>↑ {fmt(c.receitas)}</span>
                  <span style={{color:"#f87171",fontSize:"0.8rem"}}>↓ {fmt(c.despesas)}</span>
                  <span style={{color:c.saldo>=0?"#60a5fa":"#fb923c",fontWeight:"700",fontSize:"0.9rem"}}>{fmt(c.saldo)}</span>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}
      {topCats.length>0&&(
        <div style={S.card}>
          <h3 style={S.sectionTitle}>Top Categorias de Saída</h3>
          <div style={{display:"flex",flexDirection:"column",gap:"10px"}}>
            {topCats.map(([cat,val])=>(
              <div key={cat}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:"4px"}}>
                  <span style={{fontSize:"0.8rem",color:"#cbd5e1"}}>{cat}</span>
                  <span style={{fontSize:"0.8rem",color:"#f87171",fontWeight:"600"}}>{fmt(val)}</span>
                </div>
                <div style={S.barBg}><div style={{...S.barFill,width:`${(val/maxCat)*100}%`}}/></div>
              </div>
            ))}
          </div>
        </div>
      )}
      {filtered.length===0&&<div style={S.empty}>Nenhum lançamento {filterMonth!=="todos"?"neste mês":""}</div>}
    </div>
  );
}

// ── Histórico ─────────────────────────────────────────────────────────────────
function Historico({ transactions, onDelete, clientNames }) {
  const [filter, setFilter] = useState({tipo:"todos",cliente:"todos",pagamento:"todos"});
  const set=(k,v)=>setFilter(f=>({...f,[k]:v}));
  const filtered=transactions.filter(t=>{
    if(filter.tipo!=="todos"&&t.tipo!==filter.tipo) return false;
    if(filter.cliente!=="todos"&&(t.cliente||"Geral")!==filter.cliente) return false;
    if(filter.pagamento!=="todos"&&t.pagamento!==filter.pagamento) return false;
    return true;
  });
  return (
    <div style={{display:"flex",flexDirection:"column",gap:"12px"}} className="fade-in">
      <div style={S.card}>
        <div style={S.filterRow}>
          <select style={S.filterSel} value={filter.tipo} onChange={e=>set("tipo",e.target.value)}>
            <option value="todos">Todos</option><option value="receita">Entradas</option><option value="despesa">Saídas</option>
          </select>
          <select style={S.filterSel} value={filter.cliente} onChange={e=>set("cliente",e.target.value)}>
            <option value="todos">Todos clientes</option>
            {clientNames.map(c=><option key={c} value={c}>{c}</option>)}
          </select>
          <select style={S.filterSel} value={filter.pagamento} onChange={e=>set("pagamento",e.target.value)}>
            <option value="todos">Todos pgtos</option>
            {PAYMENT_METHODS.map(pm=><option key={pm.id} value={pm.id}>{pm.label}</option>)}
          </select>
        </div>
      </div>
      {filtered.length===0?<div style={S.empty}>Nenhum lançamento encontrado</div>:(
        <div style={{display:"flex",flexDirection:"column",gap:"8px"}}>
          {filtered.map(t=>{
            const pm=PAYMENT_METHODS.find(p=>p.id===t.pagamento);
            return (
              <div key={t.id} style={S.txRow} className="tx-row">
                <div style={{...S.txBar,background:t.tipo==="receita"?"#22c55e":"#f87171"}}/>
                <div style={S.txBody}>
                  <div style={S.txTop}>
                    <span style={S.txCat}>{t.categoria||"—"}</span>
                    <span style={{...S.txVal,color:t.tipo==="receita"?"#22c55e":"#f87171"}}>{t.tipo==="receita"?"+":"-"}{fmt(t.valor)}</span>
                  </div>
                  <div style={S.txMeta}>
                    <span style={{background:pm?.color+"22",color:pm?.color,padding:"1px 6px",borderRadius:"4px",fontSize:"0.65rem"}}>{pm?.icon} {pm?.label}</span>
                    {t.cliente&&t.cliente!=="Geral"&&<span style={S.txClient}><Ico.Person/> {t.cliente}</span>}
                    {t.conta&&<span style={{fontSize:"0.65rem",color:"#64748b",background:"#1e293b",padding:"1px 6px",borderRadius:"4px"}}>{t.conta==="infinity"?"Infinity":"PagBank"}</span>}
                    <span style={{color:"#64748b",fontSize:"0.7rem"}}>{fmtDate(t.data)}</span>
                  </div>
                  {t.descricao&&<div style={S.txDesc}>{t.descricao}</div>}
                </div>
                <button style={S.deleteBtn} className="press" onClick={()=>onDelete(t.id)}><Ico.Trash/></button>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ── Styles ────────────────────────────────────────────────────────────────────
const S = {
  root:{minHeight:"100vh",background:"#0a0f1e",fontFamily:"'DM Sans',sans-serif",color:"#e2e8f0"},
  loader:{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",height:"100vh",gap:"12px"},
  spin:{width:"32px",height:"32px",border:"3px solid #1e293b",borderTop:"3px solid #60a5fa",borderRadius:"50%",animation:"spin 0.8s linear infinite"},
  header:{background:"#0d1527",borderBottom:"1px solid #1e293b",position:"sticky",top:0,zIndex:100},
  headerInner:{maxWidth:"900px",margin:"0 auto",padding:"0 12px",display:"flex",alignItems:"center",justifyContent:"space-between",height:"52px"},
  logo:{display:"flex",alignItems:"center",gap:"8px"},
  logoIcon:{color:"#60a5fa",fontSize:"1.3rem"},
  logoText:{fontSize:"1.1rem",fontWeight:"800",background:"linear-gradient(135deg,#60a5fa,#a78bfa)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent"},
  nav:{display:"flex",gap:"1px"},
  navBtn:{display:"flex",flexDirection:"column",alignItems:"center",gap:"1px",padding:"5px 7px",borderRadius:"8px",border:"none",background:"transparent",color:"#64748b",cursor:"pointer",transition:"all 0.2s"},
  navActive:{background:"#1e3a5f",color:"#60a5fa"},
  navLabel:{fontSize:"0.58rem"},
  main:{maxWidth:"900px",margin:"0 auto",padding:"14px"},
  card:{background:"#111827",border:"1px solid #1e293b",borderRadius:"16px",padding:"20px"},
  cardTitle:{fontSize:"1.15rem",fontWeight:"700",marginBottom:"18px",color:"#f1f5f9"},
  tipoRow:{display:"flex",gap:"10px",marginBottom:"16px"},
  tipoBtn:{flex:1,padding:"12px",borderRadius:"12px",border:"2px solid transparent",fontWeight:"700",fontSize:"1rem",cursor:"pointer",transition:"all 0.2s",fontFamily:"inherit"},
  tipoBtnRec:{background:"#14532d",borderColor:"#22c55e",color:"#22c55e"},
  tipoBtnDesp:{background:"#450a0a",borderColor:"#f87171",color:"#f87171"},
  tipoBtnOff:{background:"#1e293b",color:"#475569"},
  field:{marginBottom:"14px"},
  label:{display:"block",fontSize:"0.72rem",color:"#94a3b8",marginBottom:"6px",textTransform:"uppercase",letterSpacing:"0.05em"},
  input:{width:"100%",background:"#1e293b",border:"1px solid #334155",borderRadius:"10px",padding:"10px 14px",color:"#f1f5f9",fontSize:"0.95rem",outline:"none",boxSizing:"border-box",fontFamily:"inherit"},
  valueWrap:{position:"relative"},
  valuePrefix:{position:"absolute",left:"14px",top:"50%",transform:"translateY(-50%)",color:"#64748b",fontSize:"1rem",pointerEvents:"none"},
  payRow:{display:"flex",gap:"8px"},
  payBtn:{flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:"4px",padding:"10px 4px",borderRadius:"10px",border:"2px solid #1e293b",background:"#1a2235",color:"#64748b",cursor:"pointer",fontSize:"0.75rem",transition:"all 0.2s"},
  payBtnActive:{},
  addClientBtn:{padding:"10px 14px",borderRadius:"10px",border:"1px solid #334155",background:"#1e293b",color:"#94a3b8",cursor:"pointer",whiteSpace:"nowrap",fontFamily:"inherit",fontSize:"0.85rem"},
  btnPrimary:{width:"100%",padding:"13px",borderRadius:"12px",border:"none",background:"linear-gradient(135deg,#3b82f6,#8b5cf6)",color:"#fff",fontWeight:"700",fontSize:"0.95rem",cursor:"pointer",marginTop:"6px",transition:"all 0.2s",fontFamily:"inherit"},
  btnSuccess:{background:"linear-gradient(135deg,#16a34a,#15803d)"},
  btnSec:{padding:"13px 18px",borderRadius:"12px",border:"1px solid #334155",background:"#1e293b",color:"#94a3b8",fontWeight:"600",cursor:"pointer",fontFamily:"inherit",fontSize:"0.9rem"},
  iconBtn:{padding:"7px",borderRadius:"8px",border:"1px solid #334155",background:"#1e293b",color:"#94a3b8",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"},
  dropZone:{border:"2px dashed #334155",borderRadius:"14px",padding:"36px 20px",textAlign:"center",cursor:"pointer",transition:"all 0.2s",marginBottom:"14px"},
  dropZoneActive:{borderColor:"#60a5fa",background:"rgba(30,58,95,0.2)"},
  tipBox:{background:"#0d1527",border:"1px solid #1e293b",borderRadius:"10px",padding:"14px"},
  mapRow:{display:"flex",alignItems:"center",gap:"12px",flexWrap:"wrap"},
  mapLabel:{fontSize:"0.8rem",color:"#cbd5e1",fontWeight:"600",minWidth:"150px"},
  prevStat:{display:"flex",flexDirection:"column",alignItems:"center",gap:"2px",background:"#1a2235",borderRadius:"10px",padding:"10px 14px",flex:1},
  bankPickBtn:{flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:"8px",padding:"20px 12px",borderRadius:"14px",border:"2px solid #1e293b",background:"#1a2235",cursor:"pointer",transition:"all 0.2s",fontFamily:"inherit"},
  dropdown:{position:"absolute",top:"100%",left:0,right:0,background:"#1e293b",border:"1px solid #334155",borderRadius:"10px",zIndex:50,maxHeight:"200px",overflowY:"auto",marginTop:"4px"},
  dropItem:{padding:"10px 14px",cursor:"pointer",borderBottom:"1px solid #0f172a"},
  clientCard:{display:"flex",alignItems:"flex-start",gap:"12px",background:"#1a2235",borderRadius:"12px",padding:"14px"},
  clientMeta:{display:"flex",alignItems:"center",gap:"5px",color:"#94a3b8",fontSize:"0.75rem",marginTop:"3px"},
  monthScroll:{display:"flex",gap:"6px",overflowX:"auto",paddingBottom:"4px"},
  monthBtn:{flexShrink:0,padding:"6px 13px",borderRadius:"20px",border:"1px solid #1e293b",background:"#111827",color:"#64748b",cursor:"pointer",fontSize:"0.78rem",transition:"all 0.2s",fontFamily:"inherit"},
  monthBtnActive:{background:"#1e3a5f",borderColor:"#60a5fa",color:"#60a5fa"},
  kpiRow:{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"10px"},
  kpi:{background:"#111827",border:"1px solid #1e293b",borderRadius:"12px",padding:"14px 12px",display:"flex",flexDirection:"column",gap:"4px"},
  kpiLabel:{fontSize:"0.62rem",color:"#64748b",textTransform:"uppercase",letterSpacing:"0.05em"},
  kpiValue:{fontSize:"0.9rem",fontWeight:"800"},
  sectionTitle:{fontSize:"0.82rem",color:"#94a3b8",textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:"14px",display:"flex",alignItems:"center",gap:"6px"},
  payStatGrid:{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"10px"},
  payStat:{display:"flex",alignItems:"center",gap:"10px",background:"#1a2235",borderRadius:"10px",padding:"12px"},
  clientRow:{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 12px",background:"#1a2235",borderRadius:"10px",flexWrap:"wrap",gap:"8px"},
  clientName:{display:"flex",alignItems:"center",gap:"10px",fontWeight:"600",fontSize:"0.88rem"},
  clientAvatar:{width:"32px",height:"32px",borderRadius:"50%",background:"linear-gradient(135deg,#3b82f6,#8b5cf6)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"0.8rem",fontWeight:"700",flexShrink:0},
  clientStats:{display:"flex",gap:"10px",alignItems:"center",flexWrap:"wrap"},
  barBg:{height:"6px",background:"#1e293b",borderRadius:"3px",overflow:"hidden"},
  barFill:{height:"100%",background:"linear-gradient(90deg,#f87171,#fb923c)",borderRadius:"3px",transition:"width 0.5s ease"},
  filterRow:{display:"flex",gap:"8px",flexWrap:"wrap"},
  filterSel:{flex:1,minWidth:"90px",background:"#1e293b",border:"1px solid #334155",borderRadius:"8px",padding:"8px 10px",color:"#f1f5f9",fontSize:"0.78rem",outline:"none",fontFamily:"inherit"},
  txRow:{display:"flex",background:"#111827",border:"1px solid #1e293b",borderRadius:"12px",overflow:"hidden",alignItems:"stretch"},
  txBar:{width:"4px",flexShrink:0},
  txBody:{flex:1,padding:"12px 14px"},
  txTop:{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"6px"},
  txCat:{fontWeight:"600",fontSize:"0.88rem",color:"#f1f5f9"},
  txVal:{fontWeight:"700",fontSize:"0.95rem"},
  txMeta:{display:"flex",alignItems:"center",gap:"7px",flexWrap:"wrap"},
  txClient:{display:"flex",alignItems:"center",gap:"4px",color:"#94a3b8",fontSize:"0.7rem"},
  txDesc:{marginTop:"5px",fontSize:"0.72rem",color:"#64748b",fontStyle:"italic"},
  deleteBtn:{padding:"0 12px",border:"none",background:"transparent",color:"#ef4444",cursor:"pointer",opacity:0.3,transition:"opacity 0.2s"},
  empty:{textAlign:"center",color:"#475569",padding:"40px",fontStyle:"italic"},
};

const css = `
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;600;700;800&display=swap');
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: #0a0f1e; }
  input[type=date]::-webkit-calendar-picker-indicator { filter: invert(0.5); }
  select option { background: #1e293b; }
  .fade-in { animation: fadeIn 0.3s ease; }
  @keyframes fadeIn { from{opacity:0;transform:translateY(8px)} to{opacity:1;transform:translateY(0)} }
  @keyframes spin { to { transform: rotate(360deg); } }
  .press:active { transform: scale(0.96); }
  .tx-row:hover button { opacity: 1 !important; }
  ::-webkit-scrollbar { height: 4px; width: 4px; }
  ::-webkit-scrollbar-track { background: transparent; }
  ::-webkit-scrollbar-thumb { background: #334155; border-radius: 2px; }
`;
