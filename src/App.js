import { useState, useEffect, useMemo, useCallback } from "react";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from "recharts";
import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient";
import Auth from "./Auth";

// ─── Constants ────────────────────────────────────────────────────────────────

const CATEGORIES = {
  receita: [
    { id: "salario",       label: "Salário",      icon: "💼" },
    { id: "imoveis",       label: "Imóveis",      icon: "🏢" },
    { id: "investimentos", label: "Investimentos",icon: "📈" },
    { id: "outros_r",      label: "Outros",       icon: "➕" },
  ],
  despesa: [
    { id: "moradia",     label: "Moradia",     icon: "🏠" },
    { id: "alimentacao", label: "Alimentação", icon: "🍽️" },
    { id: "transporte",  label: "Transporte",  icon: "🚗" },
    { id: "saude",       label: "Saúde",       icon: "💊" },
    { id: "educacao",    label: "Educação",    icon: "📚" },
    { id: "compras",     label: "Compras",     icon: "🛍️" },
    { id: "nina",        label: "Nina",        icon: "🐾" },
    { id: "viagens",     label: "Viagens",     icon: "✈️" },
    { id: "empregada",   label: "Empregada",   icon: "🧹" },
    { id: "corpo",       label: "Corpo",       icon: "💆" },
    { id: "outros_d",    label: "Outros",      icon: "➖" },
  ],
};

const DEFAULT_BUDGETS = {
  salario:       17000,
  imoveis:       4567.38,
  investimentos: 7195.28,
  outros_r:      5000,
  moradia:       5660,
  alimentacao:   1400,
  transporte:    650,
  saude:         153.79,
  educacao:      0,
  compras:       1200,
  nina:          3822,
  viagens:       2000,
  empregada:     3782.44,
  corpo:         385,
  outros_d:      287.5,
};

const INV_BANKS = [
  { id: "itau",     label: "Itaú",     icon: "🟠", color: "#f97316" },
  { id: "xp",       label: "XP",       icon: "⚫", color: "#a3a3a3" },
  { id: "bradesco", label: "Bradesco", icon: "🔴", color: "#ef4444" },
];

const MONTHS      = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];
const MONTHS_FULL = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];

const ALL_CATS = [...CATEGORIES.receita, ...CATEGORIES.despesa];
const fmt      = (v) => v.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
const fmtNum   = (v) => Number(v.toFixed(2));
const getCat   = (id) => ALL_CATS.find(c => c.id === id);
const getCatLabel = (id) => { const c = getCat(id); return c ? `${c.icon} ${c.label}` : id; };

// ─── Excel Export ─────────────────────────────────────────────────────────────

function exportToExcel(transactions, investments, budgets, year) {
  const wb = XLSX.utils.book_new();

  const txRows = transactions
    .filter(t => new Date(t.date+"T12:00:00").getFullYear()===year)
    .sort((a,b) => a.date.localeCompare(b.date))
    .map(t => ({ Data:t.date, Tipo:t.type==="receita"?"Receita":"Despesa", Categoria:getCatLabel(t.category), "Descrição":t.description||"", Valor:fmtNum(t.amount) }));
  const ws1 = XLSX.utils.json_to_sheet(txRows);
  ws1["!cols"] = [{wch:12},{wch:10},{wch:20},{wch:28},{wch:14}];
  XLSX.utils.book_append_sheet(wb, ws1, "Lançamentos");

  const summaryRows = MONTHS.map((m,i) => {
    const txs = transactions.filter(t => { const d=new Date(t.date+"T12:00:00"); return d.getMonth()===i && d.getFullYear()===year; });
    const r = txs.filter(t=>t.type==="receita").reduce((s,t)=>s+t.amount,0);
    const d = txs.filter(t=>t.type==="despesa").reduce((s,t)=>s+t.amount,0);
    return { "Mês":MONTHS_FULL[i], Receitas:fmtNum(r), Despesas:fmtNum(d), Saldo:fmtNum(r-d) };
  });
  const tR=summaryRows.reduce((s,r)=>s+r.Receitas,0), tD=summaryRows.reduce((s,r)=>s+r.Despesas,0);
  summaryRows.push({ "Mês":"TOTAL", Receitas:fmtNum(tR), Despesas:fmtNum(tD), Saldo:fmtNum(tR-tD) });
  const ws2 = XLSX.utils.json_to_sheet(summaryRows);
  ws2["!cols"] = [{wch:14},{wch:16},{wch:16},{wch:16}];
  XLSX.utils.book_append_sheet(wb, ws2, "Resumo Mensal");

  const invRows = MONTHS.map((m,i) => {
    const snap = investments.find(inv=>inv.year===year&&inv.month===i)||{};
    const it=snap.itau||0,xp=snap.xp||0,br=snap.bradesco||0;
    return { "Mês":MONTHS_FULL[i],"Itaú":fmtNum(it),XP:fmtNum(xp),Bradesco:fmtNum(br),Total:fmtNum(it+xp+br) };
  }).filter(r=>r.Total>0);
  if (invRows.length>0) {
    const ws3 = XLSX.utils.json_to_sheet(invRows);
    ws3["!cols"] = [{wch:14},{wch:14},{wch:14},{wch:14},{wch:14}];
    XLSX.utils.book_append_sheet(wb, ws3, "Investimentos");
  }

  XLSX.writeFile(wb, `Balance_${year}.xlsx`);
}

// ─── Budget helpers ───────────────────────────────────────────────────────────

function getBudgetAmount(budgets, categoryId, year, month) {
  // Priority: month-specific > base (month=null) > default
  const monthly = budgets.find(b => b.category===categoryId && b.year===year && b.month===month);
  if (monthly) return monthly.amount;
  const base = budgets.find(b => b.category===categoryId && b.year===year && b.month===null);
  if (base) return base.amount;
  return DEFAULT_BUDGETS[categoryId] ?? 0;
}

// ─── Main App ─────────────────────────────────────────────────────────────────

export default function App() {
  const [session,      setSession]      = useState(undefined);
  const [transactions, setTransactions] = useState([]);
  const [investments,  setInvestments]  = useState([]);
  const [budgets,      setBudgets]      = useState([]);
  const [dbLoading,    setDbLoading]    = useState(false);

  const [view,          setView]          = useState("dashboard");
  const [type,          setType]          = useState("despesa");
  const [category,      setCategory]      = useState("");
  const [amount,        setAmount]        = useState("");
  const [desc,          setDesc]          = useState("");
  const [date,          setDate]          = useState(() => new Date().toISOString().slice(0,10));
  const [selectedMonth, setSelectedMonth] = useState(new Date().getMonth());
  const [selectedYear]                    = useState(new Date().getFullYear());
  const [toast,         setToast]         = useState(null);
  const [invInputs,     setInvInputs]     = useState({itau:"",xp:"",bradesco:""});
  const [hidden,        setHidden]        = useState(true);
  const [isDesktop,     setIsDesktop]     = useState(window.innerWidth>=768);
  const [syncing,       setSyncing]       = useState(false);

  // Budget editing state
  const [show2FA,      setShow2FA]      = useState(false);
  const [totp2FACode,  setTotp2FACode]  = useState("");
  const [totp2FAQr,    setTotp2FAQr]    = useState("");
  const [totp2FAId,    setTotp2FAId]    = useState("");
  const [totp2FAStep,  setTotp2FAStep]  = useState("qr"); // qr | done
  const [budgetScope,   setBudgetScope]   = useState("month");
  const [editingTx,     setEditingTx]     = useState(null); // {id, amount, description}
  const [editAmount,    setEditAmount]    = useState("");
  const [editDesc,      setEditDesc]      = useState("");

  useEffect(() => {
    supabase.auth.getSession().then(({data:{session}}) => setSession(session));
    const {data:{subscription}} = supabase.auth.onAuthStateChange((_e,s) => setSession(s));
    return () => subscription.unsubscribe();
  }, []);

  useEffect(() => {
    const fn = () => setIsDesktop(window.innerWidth>=768);
    window.addEventListener("resize", fn);
    return () => window.removeEventListener("resize", fn);
  }, []);

  const loadData = useCallback(async () => {
    if (!session) return;
    setDbLoading(true);
    const [{ data:tx },{ data:inv },{ data:bud }] = await Promise.all([
      supabase.from("transactions").select("*").order("date",{ascending:false}),
      supabase.from("investments").select("*"),
      supabase.from("budgets").select("*"),
    ]);
    setTransactions(tx||[]);
    setInvestments(inv||[]);
    setBudgets(bud||[]);
    setDbLoading(false);
  }, [session]);

  useEffect(() => { loadData(); }, [loadData]);

  const showToast = (msg,ok=true) => { setToast({msg,ok}); setTimeout(()=>setToast(null),2500); };
  const mask = (val) => hidden ? "••••••" : val;

  // ─── CRUD ──────────────────────────────────────────────────────────────────

  const addTransaction = async () => {
    if (!amount||!category) return showToast("Preencha todos os campos",false);
    const val = parseFloat(amount.replace(",","."));
    if (isNaN(val)||val<=0) return showToast("Valor inválido",false);
    setSyncing(true);
    const {data,error} = await supabase.from("transactions").insert({
      type,category,amount:val,description:desc,date,user_id:session.user.id,
    }).select().single();
    setSyncing(false);
    if (error) return showToast("Erro ao salvar",false);
    setTransactions(prev=>[data,...prev]);
    setAmount(""); setDesc(""); setCategory("");
    showToast("Lançamento salvo ☁️");
    setView("dashboard");
  };

  const deleteTransaction = async (id) => {
    setSyncing(true);
    const {error} = await supabase.from("transactions").delete().eq("id",id);
    setSyncing(false);
    if (error) return showToast("Erro ao remover",false);
    setTransactions(prev=>prev.filter(t=>t.id!==id));
    showToast("Removido");
  };

  const openEditTx = (t) => {
    setEditingTx(t);
    setEditAmount(String(t.amount));
    setEditDesc(t.description||"");
  };

  const saveEditTx = async () => {
    const val = parseFloat(editAmount.replace(",","."));
    if (isNaN(val)||val<=0) return showToast("Valor inválido",false);
    setSyncing(true);
    const {data,error} = await supabase.from("transactions")
      .update({amount:val, description:editDesc})
      .eq("id",editingTx.id)
      .select().single();
    setSyncing(false);
    if (error) return showToast("Erro ao editar",false);
    setTransactions(prev=>prev.map(t=>t.id===editingTx.id?data:t));
    setEditingTx(null);
    showToast("Lançamento atualizado ☁️");
  };

  const saveInvestMonth = async () => {
    const itau     = parseFloat(invInputs.itau.replace(",","."))||0;
    const xp       = parseFloat(invInputs.xp.replace(",","."))||0;
    const bradesco = parseFloat(invInputs.bradesco.replace(",","."))||0;
    setSyncing(true);
    const {data,error} = await supabase.from("investments").upsert({
      user_id:session.user.id, year:selectedYear, month:selectedMonth, itau, xp, bradesco,
    },{onConflict:"user_id,year,month"}).select().single();
    setSyncing(false);
    if (error) return showToast("Erro ao salvar",false);
    setInvestments(prev=>[...prev.filter(i=>!(i.year===selectedYear&&i.month===selectedMonth)),data]);
    showToast("Investimentos salvos ☁️");
  };

  const saveBudgets = async () => {
    const rows = Object.entries(budgetEdits)
      .filter(([,v]) => v !== "")
      .map(([catId, v]) => ({
        user_id: session.user.id,
        category: catId,
        year: selectedYear,
        month: budgetScope === "month" ? selectedMonth : null,
        amount: parseFloat(String(v).replace(",",".")) || 0,
      }));
    if (rows.length===0) return showToast("Nenhuma alteração",false);
    setSyncing(true);
    const {data,error} = await supabase.from("budgets")
      .upsert(rows,{onConflict:"user_id,category,year,month"})
      .select();
    setSyncing(false);
    if (error) return showToast("Erro ao salvar orçamento",false);
    setBudgets(prev => {
      const updated = [...prev];
      data.forEach(newB => {
        const idx = updated.findIndex(b=>b.category===newB.category&&b.year===newB.year&&b.month===newB.month);
        if (idx>=0) updated[idx]=newB; else updated.push(newB);
      });
      return updated;
    });
    setBudgetEdits({});
    showToast("Orçamento salvo ☁️");
  };

  const openBudget = () => {
    // Pre-fill inputs with current values
    const edits = {};
    ALL_CATS.forEach(c => {
      const val = getBudgetAmount(budgets, c.id, selectedYear, selectedMonth);
      edits[c.id] = val > 0 ? String(val) : "";
    });
    setBudgetEdits(edits);
    setBudgetScope("month");
    setView("budget");
  };

  const handle2FAEnroll = async () => {
    setSyncing(true);
    const { data, error } = await supabase.auth.mfa.enroll({ factorType: "totp", issuer: "Balance App" });
    setSyncing(false);
    if (error) return showToast("Erro ao iniciar 2FA: " + error.message, false);
    setTotp2FAQr(data.totp.qr_code);
    setTotp2FAId(data.id);
    setTotp2FAStep("qr");
    setTotp2FACode("");
    setShow2FA(true);
  };

  const handle2FAConfirm = async () => {
    if (!totp2FACode || totp2FACode.length !== 6) return showToast("Digite o código de 6 dígitos", false);
    setSyncing(true);
    const { error } = await supabase.auth.mfa.challengeAndVerify({ factorId: totp2FAId, code: totp2FACode });
    setSyncing(false);
    if (error) return showToast("Código inválido", false);
    setShow2FA(false);
    setTotp2FAStep("done");
    showToast("2FA ativado com sucesso! 🔐");
  };
    await supabase.auth.signOut();
    setTransactions([]); setInvestments([]); setBudgets([]);
  };

  const handleExport = () => { exportToExcel(transactions,investments,budgets,selectedYear); showToast(`Excel ${selectedYear} exportado!`); };

  // ─── Derived data ──────────────────────────────────────────────────────────

  const getInvestMonth = (month,year) => investments.find(i=>i.year===year&&i.month===month)||{itau:0,xp:0,bradesco:0};

  const openInvest = () => {
    const snap = getInvestMonth(selectedMonth,selectedYear);
    setInvInputs({itau:snap.itau>0?String(snap.itau):"",xp:snap.xp>0?String(snap.xp):"",bradesco:snap.bradesco>0?String(snap.bradesco):""});
    setView("invest");
  };

  const currentInvest = getInvestMonth(selectedMonth,selectedYear);
  const totalInvest   = INV_BANKS.reduce((s,b)=>s+(currentInvest[b.id]||0),0);

  const monthlyTx = useMemo(()=>
    transactions.filter(t=>{ const d=new Date(t.date+"T12:00:00"); return d.getMonth()===selectedMonth&&d.getFullYear()===selectedYear; }),
    [transactions,selectedMonth,selectedYear]);

  const totalReceita = monthlyTx.filter(t=>t.type==="receita").reduce((s,t)=>s+t.amount,0);
  const totalDespesa = monthlyTx.filter(t=>t.type==="despesa").reduce((s,t)=>s+t.amount,0);
  const saldo        = totalReceita-totalDespesa;

  const byCategory = useMemo(()=>{
    const map={};
    monthlyTx.forEach(t=>{ if(!map[t.category]) map[t.category]={receita:0,despesa:0}; map[t.category][t.type]+=t.amount; });
    return map;
  },[monthlyTx]);

  const yearlyData = useMemo(()=>
    MONTHS.map((m,i)=>{
      const txs=transactions.filter(t=>{ const d=new Date(t.date+"T12:00:00"); return d.getMonth()===i&&d.getFullYear()===selectedYear; });
      return { name:m, Receita:txs.filter(t=>t.type==="receita").reduce((s,t)=>s+t.amount,0), Despesa:txs.filter(t=>t.type==="despesa").reduce((s,t)=>s+t.amount,0) };
    }),[transactions,selectedYear]);

  const yearlyInvestData = useMemo(()=>
    MONTHS.map((m,i)=>{ const snap=getInvestMonth(i,selectedYear); return {name:m,Total:INV_BANKS.reduce((s,b)=>s+(snap[b.id]||0),0)}; }),
    [investments,selectedYear]);

  // Budget rows for current month — all categories that have budget or spending
  const budgetRows = useMemo(()=>{
    const rows=[];
    ALL_CATS.forEach(cat=>{
      const budgeted = getBudgetAmount(budgets,cat.id,selectedYear,selectedMonth);
      const spent    = byCategory[cat.id]?.[cat.type==="receita"?"receita":"despesa"] ||
                       (byCategory[cat.id]?.receita||0)+(byCategory[cat.id]?.despesa||0);
      // figure out real spent for this category
      const realSpent = monthlyTx.filter(t=>t.category===cat.id).reduce((s,t)=>s+t.amount,0);
      if (budgeted===0&&realSpent===0) return;
      const isReceita = CATEGORIES.receita.find(c=>c.id===cat.id);
      const pct = budgeted>0 ? Math.min((realSpent/budgeted)*100,100) : 100;
      const over = budgeted>0 && realSpent>budgeted;
      const diff = budgeted - realSpent; // positive = remaining, negative = over
      rows.push({ cat, budgeted, realSpent, pct, over, diff, isReceita });
    });
    return rows;
  },[budgets,monthlyTx,selectedMonth,selectedYear]);

  if (session===undefined) return (
    <div style={{minHeight:"100vh",background:"#020817",display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{color:"#475569",fontSize:15}}>Carregando...</div>
    </div>
  );
  if (!session) return <Auth />;

  // ─── Layout helpers ────────────────────────────────────────────────────────

  const S = isDesktop ? desktopStyles : mobileStyles;

  const NavButton = ({id,icon,label,onClick}) => (
    <button style={{...S.navBtn,...(view===id?S.navBtnActive:{})}} onClick={onClick||(()=>setView(id))}>
      <span style={S.navIcon}>{icon}</span>
      <span style={{...S.navLabel,color:view===id?"#60a5fa":"#475569"}}>{label}</span>
    </button>
  );

  // ─── Section components ────────────────────────────────────────────────────

  const budgetBarJSX = (row) => {
    const {cat,budgeted,realSpent,pct,over,diff,isReceita} = row;
    const barColor = isReceita ? (realSpent>=budgeted?"#4ade80":"#3b82f6") : (over?"#f87171":"#4ade80");
    const diffColor = isReceita ? (diff<=0?"#4ade80":"#f87171") : (over?"#f87171":"#4ade80");
    const diffText = isReceita
      ? (diff<=0?`✓ Meta atingida (+${mask(fmt(Math.abs(diff)))})`:`⚠️ Faltam ${mask(fmt(diff))}`)
      : (over?`⚠️ Estourou ${mask(fmt(Math.abs(diff)))}`:`✓ Faltam ${mask(fmt(diff))}`);
    return (
      <div key={cat.id} style={bStyles.row}>
        <div style={bStyles.rowTop}>
          <span style={bStyles.catName}>{cat.icon} {cat.label}</span>
          <div style={bStyles.amounts}>
            <span style={{...bStyles.spent,color:isReceita?(diff<=0?"#4ade80":"#e2e8f0"):(over?"#f87171":"#e2e8f0")}}>{mask(fmt(realSpent))}</span>
            <span style={bStyles.sep}>/</span>
            <span style={bStyles.budget}>{mask(fmt(budgeted))}</span>
          </div>
        </div>
        {budgeted>0&&<div style={bStyles.barBg}><div style={{...bStyles.barFill,width:`${pct}%`,background:barColor}}/></div>}
        <div style={bStyles.diffRow}>
          {budgeted>0
            ? <span style={{fontSize:11,fontWeight:700,color:diffColor}}>{diffText}</span>
            : <span style={{fontSize:11,color:"#334155"}}>Sem orçamento definido</span>}
        </div>
      </div>
    );
  };

  const totalOrcReceita = CATEGORIES.receita.reduce((s,c)=>s+getBudgetAmount(budgets,c.id,selectedYear,selectedMonth),0);
  const totalOrcDespesa = CATEGORIES.despesa.reduce((s,c)=>s+getBudgetAmount(budgets,c.id,selectedYear,selectedMonth),0);
  const budgetRowsReceita = budgetRows.filter(r=>r.isReceita);
  const budgetRowsDespesa = budgetRows.filter(r=>!r.isReceita);

  const dashboardJSX = (
    <>
      {dbLoading&&<div style={S.loadingBar}>Sincronizando com a nuvem...</div>}
      <div style={{...S.balanceCard,background:saldo>=0?"#0f172a":"#1a0a0a"}}>
        <span style={S.balanceLabel}>Saldo do mês</span>
        <span style={{...S.balanceValue,color:saldo>=0?"#4ade80":"#f87171"}}>{mask(fmt(saldo))}</span>
      </div>
      <div style={{...S.investSummaryCard,background:"linear-gradient(135deg,#052e16,#14532d)",border:"1px solid #166534"}} onClick={openBudget}>
        <div style={S.investSummaryLeft}>
          <span style={{...S.investSummaryLabel,color:"#4ade80"}}>🎯 Orçamento do Mês — Receitas</span>
          <div style={{display:"flex",gap:16,marginTop:4}}>
            <div>
              <div style={{fontSize:10,color:"#e2e8f0",fontWeight:600,textTransform:"uppercase",letterSpacing:0.8}}>Orçado</div>
              <div style={{fontSize:16,fontWeight:800,color:"#4ade80"}}>{mask(fmt(totalOrcReceita))}</div>
            </div>
            <div>
              <div style={{fontSize:10,color:"#e2e8f0",fontWeight:600,textTransform:"uppercase",letterSpacing:0.8}}>Realizado</div>
              <div style={{fontSize:16,fontWeight:800,color:"#e2e8f0"}}>{mask(fmt(totalReceita))}</div>
            </div>
          </div>
          <span style={{fontSize:11,color:"#166534",marginTop:4}}>Toque para editar orçamentos</span>
        </div>
        <span style={{fontSize:24,color:"#4ade80"}}>›</span>
      </div>
      <div style={{...S.investSummaryCard,background:"linear-gradient(135deg,#1a0800,#3b0f00)",border:"1px solid #92400e"}} onClick={openBudget}>
        <div style={S.investSummaryLeft}>
          <span style={{...S.investSummaryLabel,color:"#f59e0b"}}>🎯 Orçamento do Mês — Despesas</span>
          <div style={{display:"flex",gap:16,marginTop:4}}>
            <div>
              <div style={{fontSize:10,color:"#e2e8f0",fontWeight:600,textTransform:"uppercase",letterSpacing:0.8}}>Orçado</div>
              <div style={{fontSize:16,fontWeight:800,color:"#fbbf24"}}>{mask(fmt(totalOrcDespesa))}</div>
            </div>
            <div>
              <div style={{fontSize:10,color:"#e2e8f0",fontWeight:600,textTransform:"uppercase",letterSpacing:0.8}}>Realizado</div>
              <div style={{fontSize:16,fontWeight:800,color:totalDespesa>0?"#f87171":"#e2e8f0"}}>{mask(fmt(totalDespesa))}</div>
            </div>
          </div>
          <span style={{fontSize:11,color:"#78350f",marginTop:4}}>Toque para editar orçamentos</span>
        </div>
        <span style={{fontSize:24,color:"#f59e0b"}}>›</span>
      </div>
      {budgetRowsReceita.length>0&&(
        <div style={S.section}>
          <div style={{...S.sectionTitle,color:"#4ade80"}}>🎯 Orçamento × Realizado — Receitas</div>
          {budgetRowsReceita.map(row=>budgetBarJSX(row))}
        </div>
      )}
      {budgetRowsDespesa.length>0&&(
        <div style={S.section}>
          <div style={{...S.sectionTitle,color:"#f59e0b"}}>🎯 Orçamento × Realizado — Despesas</div>
          {budgetRowsDespesa.map(row=>budgetBarJSX(row))}
        </div>
      )}
      <div style={S.section}>
        <div style={S.sectionTitle}>Receitas & Despesas {selectedYear}</div>
        <ResponsiveContainer width="100%" height={220}>
          <BarChart data={yearlyData} barSize={isDesktop?14:8} barGap={2}>
            <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" vertical={false}/>
            <XAxis dataKey="name" tick={{fill:"#64748b",fontSize:11}} axisLine={false} tickLine={false}/>
            <YAxis hide/>
            <Tooltip contentStyle={{background:"#0f172a",border:"1px solid #1e293b",borderRadius:8,color:"#e2e8f0",fontSize:12}} formatter={v=>hidden?"••••••":fmt(v)}/>
            <Bar dataKey="Receita" fill="#4ade80" radius={[4,4,0,0]}/>
            <Bar dataKey="Despesa" fill="#f87171" radius={[4,4,0,0]}/>
          </BarChart>
        </ResponsiveContainer>
        <div style={S.legend}>
          <span style={S.legendItem}><span style={{...S.dot,background:"#4ade80"}}/>Receita</span>
          <span style={S.legendItem}><span style={{...S.dot,background:"#f87171"}}/>Despesa</span>
        </div>
      </div>
      {monthlyTx.length>0&&(
        <div style={S.section}>
          <div style={S.sectionTitle}>Lançamentos do mês</div>
          {monthlyTx.map(t=>(
            <div key={t.id} style={S.txRow}>
              <div style={{flex:1}}>
                <div style={S.txCat}>{getCatLabel(t.category)}</div>
                {t.description&&<div style={S.txDesc}>{t.description}</div>}
                <div style={S.txDate}>{new Date(t.date+"T12:00:00").toLocaleDateString("pt-BR")}</div>
              </div>
              <div style={S.txRight}>
                <span style={{...S.txAmt,color:t.type==="receita"?"#4ade80":"#f87171"}}>
                  {t.type==="receita"?"+":"-"}{mask(fmt(t.amount))}
                </span>
                <button style={S.editBtn} onClick={()=>openEditTx(t)}>✏️</button>
                <button style={S.delBtn} onClick={()=>deleteTransaction(t.id)}>✕</button>
              </div>
            </div>
          ))}
        </div>
      )}
      {editingTx&&(
        <div style={eStyles.overlay} onClick={()=>setEditingTx(null)}>
          <div style={eStyles.modal} onClick={e=>e.stopPropagation()}>
            <div style={eStyles.title}>Editar Lançamento</div>
            <div style={eStyles.subtitle}>{getCatLabel(editingTx.category)} · {new Date(editingTx.date+"T12:00:00").toLocaleDateString("pt-BR")}</div>
            <div style={{marginBottom:14}}>
              <label style={eStyles.label}>Valor (R$)</label>
              <input style={eStyles.input} type="text" inputMode="decimal" value={editAmount}
                onChange={e=>setEditAmount(e.target.value)} autoComplete="off"/>
            </div>
            <div style={{marginBottom:20}}>
              <label style={eStyles.label}>Descrição</label>
              <input style={eStyles.input} type="text" value={editDesc}
                onChange={e=>setEditDesc(e.target.value)} placeholder="Opcional"/>
            </div>
            <div style={eStyles.btns}>
              <button style={eStyles.cancelBtn} onClick={()=>setEditingTx(null)}>Cancelar</button>
              <button style={eStyles.saveBtn} onClick={saveEditTx} disabled={syncing}>{syncing?"Salvando...":"Salvar"}</button>
            </div>
            <button style={eStyles.deleteBtn} onClick={async()=>{await deleteTransaction(editingTx.id);setEditingTx(null);}}>
              🗑️ Excluir lançamento
            </button>
          </div>
        </div>
      )}
      {monthlyTx.length===0&&!dbLoading&&(
        <div style={S.empty}>
          <div style={S.emptyIcon}>📊</div>
          <div style={S.emptyText}>Nenhum lançamento em {MONTHS_FULL[selectedMonth]}</div>
          <div style={S.emptyHint}>Clique em Lançar para adicionar</div>
        </div>
      )}
      <div style={S.exportBanner} onClick={handleExport}>
        <span style={{fontSize:28}}>📥</span>
        <div>
          <div style={S.exportBannerTitle}>Exportar para Excel</div>
          <div style={S.exportBannerSub}>Balance_{selectedYear}.xlsx</div>
        </div>
        <span style={{marginLeft:"auto",fontSize:22,color:"#1e3a5f"}}>›</span>
      </div>
    </>
  );

  const addFormJSX = (
    <div style={S.form}>
      <div style={S.typeToggle}>
        <button style={{...S.typeBtn,...(type==="receita"?S.typeBtnActiveGreen:{})}} onClick={()=>{setType("receita");setCategory("");}}>💚 Receita</button>
        <button style={{...S.typeBtn,...(type==="despesa"?S.typeBtnActiveRed:{})}} onClick={()=>{setType("despesa");setCategory("");}}>❤️ Despesa</button>
      </div>
      <div style={S.field}>
        <label style={S.label}>Valor (R$)</label>
        <input style={S.input} type="text" inputMode="decimal" placeholder="0,00" value={amount} onChange={e=>setAmount(e.target.value)} autoComplete="off"/>
      </div>
      <div style={S.field}>
        <label style={S.label}>Categoria</label>
        <div style={S.catGrid}>
          {CATEGORIES[type].map(c=>(
            <button key={c.id} style={{...S.catBtn,...(category===c.id?S.catBtnActive:{})}} onClick={()=>setCategory(c.id)}>
              <span style={S.catBtnIcon}>{c.icon}</span>
              <span style={S.catBtnLabel}>{c.label}</span>
            </button>
          ))}
        </div>
      </div>
      <div style={S.field}>
        <label style={S.label}>Descrição (opcional)</label>
        <input style={S.input} type="text" placeholder="Ex: Supermercado Extra" value={desc} onChange={e=>setDesc(e.target.value)}/>
      </div>
      <div style={S.field}>
        <label style={S.label}>Data</label>
        <input style={S.input} type="date" value={date} onChange={e=>setDate(e.target.value)}/>
      </div>
      <button style={{...S.saveBtn,opacity:syncing?0.6:1}} onClick={addTransaction} disabled={syncing}>
        {syncing?"Salvando...":"Salvar Lançamento"}
      </button>
    </div>
  );

  const budgetFormJSX = (
    <div style={S.form}>
      <div style={{...S.investHeader}}>
        <div style={{...S.investTitle,color:"#fbbf24"}}>🎯 Orçamento</div>
        <div style={S.investSubtitle}>{MONTHS_FULL[selectedMonth]} {selectedYear}</div>
      </div>
      <div style={S.typeToggle}>
        <button style={{...S.typeBtn,...(budgetScope==="month"?{borderColor:"#d97706",color:"#fbbf24",background:"#1a1000"}:{})}}
          onClick={()=>setBudgetScope("month")}>📅 Este mês</button>
        <button style={{...S.typeBtn,...(budgetScope==="base"?{borderColor:"#d97706",color:"#fbbf24",background:"#1a1000"}:{})}}
          onClick={()=>setBudgetScope("base")}>📋 Base anual</button>
      </div>
      <div style={{fontSize:11,color:"#475569",marginBottom:16,lineHeight:1.5}}>
        {budgetScope==="month"
          ?`Valores aplicados apenas em ${MONTHS_FULL[selectedMonth]}. Sobrescreve o base anual.`
          :`Valores padrão para todos os meses de ${selectedYear} sem ajuste específico.`}
      </div>
      <div style={{...S.sectionTitle,marginBottom:12}}>💚 Receitas</div>
      {CATEGORIES.receita.map(c=>(
        <div key={c.id} style={S.field}>
          <label style={{...S.label,color:"#4ade80"}}>{c.icon} {c.label}</label>
          <input style={S.input} type="text" inputMode="decimal"
            placeholder={`Base: ${fmt(DEFAULT_BUDGETS[c.id]||0)}`}
            value={budgetEdits[c.id]??""} onChange={e=>setBudgetEdits(p=>({...p,[c.id]:e.target.value}))}/>
        </div>
      ))}
      <div style={{...S.sectionTitle,marginBottom:12,marginTop:8}}>❤️ Despesas</div>
      {CATEGORIES.despesa.map(c=>(
        <div key={c.id} style={S.field}>
          <label style={{...S.label,color:"#f87171"}}>{c.icon} {c.label}</label>
          <input style={S.input} type="text" inputMode="decimal"
            placeholder={`Base: ${fmt(DEFAULT_BUDGETS[c.id]||0)}`}
            value={budgetEdits[c.id]??""} onChange={e=>setBudgetEdits(p=>({...p,[c.id]:e.target.value}))}/>
        </div>
      ))}
      <button style={{...S.saveBtn,background:"linear-gradient(135deg,#d97706,#f59e0b)",opacity:syncing?0.6:1}}
        onClick={saveBudgets} disabled={syncing}>
        {syncing?"Salvando...":"Salvar Orçamento"}
      </button>
    </div>
  );

  const investFormJSX = (
    <div style={S.form}>
      <div style={S.investHeader}>
        <div style={S.investTitle}>💎 Investimentos</div>
        <div style={S.investSubtitle}>Saldo em {MONTHS_FULL[selectedMonth]}</div>
      </div>
      {INV_BANKS.map(b=>(
        <div key={b.id} style={S.field}>
          <label style={{...S.label,color:b.color}}>{b.icon} {b.label}</label>
          <input style={{...S.input,borderColor:invInputs[b.id]?b.color+"55":"#0f172a"}}
            type="text" inputMode="decimal" placeholder="R$ 0,00"
            value={invInputs[b.id]} onChange={e=>setInvInputs(p=>({...p,[b.id]:e.target.value}))}/>
        </div>
      ))}
      <div style={S.investLiveTotal}>
        <span style={S.investLiveTotalLabel}>Total calculado</span>
        <span style={S.investLiveTotalValue}>
          {mask(fmt(INV_BANKS.reduce((s,b)=>{const v=parseFloat(invInputs[b.id]?.replace(",","."));return s+(isNaN(v)?0:v);},0)))}
        </span>
      </div>
      <button style={{...S.saveBtn,background:"linear-gradient(135deg,#7c3aed,#a855f7)",opacity:syncing?0.6:1}}
        onClick={saveInvestMonth} disabled={syncing}>
        {syncing?"Salvando...":"Salvar Investimentos"}
      </button>
      <div style={{...S.section,marginTop:20}}>
        <div style={S.sectionTitle}>Evolução {selectedYear}</div>
        <ResponsiveContainer width="100%" height={200}>
          <BarChart data={yearlyInvestData} barSize={isDesktop?18:14}>
            <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" vertical={false}/>
            <XAxis dataKey="name" tick={{fill:"#64748b",fontSize:11}} axisLine={false} tickLine={false}/>
            <YAxis hide/>
            <Tooltip contentStyle={{background:"#0f172a",border:"1px solid #1e293b",borderRadius:8,color:"#e2e8f0",fontSize:12}} formatter={v=>hidden?"••••••":fmt(v)}/>
            <Bar dataKey="Total" fill="#a855f7" radius={[4,4,0,0]}/>
          </BarChart>
        </ResponsiveContainer>
      </div>
      <div style={S.section}>
        <div style={S.sectionTitle}>Por banco — mês a mês</div>
        {MONTHS.map((m,i)=>{
          const snap=getInvestMonth(i,selectedYear);
          const total=INV_BANKS.reduce((s,b)=>s+(snap[b.id]||0),0);
          if(total===0) return null;
          const prevSnap=getInvestMonth(i-1,selectedYear);
          const prevTotal=INV_BANKS.reduce((s,b)=>s+(prevSnap[b.id]||0),0);
          const variation=total-prevTotal;
          const hasVariation=i>0&&prevTotal>0;
          return (
            <div key={i} style={S.catRow}>
              <span style={S.catName}>{m}</span>
              <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:2}}>
                {INV_BANKS.map(b=>snap[b.id]>0&&(
                  <span key={b.id} style={{fontSize:11,color:b.color,fontWeight:600}}>{b.icon} {mask(fmt(snap[b.id]))}</span>
                ))}
                <span style={{fontSize:13,color:"#a855f7",fontWeight:800}}>{mask(fmt(total))}</span>
                {hasVariation&&(
                  <span style={{fontSize:11,fontWeight:700,color:variation>=0?"#4ade80":"#f87171"}}>
                    {variation>=0?"+":""}{mask(fmt(variation))}
                  </span>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );

  // ─── Render ────────────────────────────────────────────────────────────────

  return (
    <div style={S.root}>
      <style>{css}</style>
      {toast&&<div style={{...S.toast,background:toast.ok?"#22c55e":"#ef4444"}}>{toast.msg}</div>}
      {syncing&&<div style={S.syncDot}>☁️</div>}

      {isDesktop ? (
        <div style={S.desktopWrapper}>
          <aside style={S.sidebar}>
            <div style={S.sidebarLogo}>₿alance</div>
            <div style={S.sidebarEmail}>{session.user.email}</div>
            <div style={S.monthPicker}>
              <button style={S.arrowBtn} onClick={()=>setSelectedMonth(m=>(m-1+12)%12)}>‹</button>
              <span style={S.monthLabel}>{MONTHS_FULL[selectedMonth]}</span>
              <button style={S.arrowBtn} onClick={()=>setSelectedMonth(m=>(m+1)%12)}>›</button>
            </div>
            <nav style={S.sideNav}>
              <NavButton id="dashboard" icon="📊" label="Resumo"/>
              <NavButton id="add"       icon="✏️"  label="Lançar"/>
              <NavButton id="budget"    icon="🎯" label="Orçamento" onClick={openBudget}/>
              <NavButton id="invest"    icon="💎" label="Investimentos" onClick={openInvest}/>
            </nav>
            <div style={S.sideActions}>
              <button style={S.sideActionBtn} onClick={()=>setHidden(h=>!h)}>
                {hidden?"🙈":"👁️"} {hidden?"Mostrar valores":"Ocultar valores"}
              </button>
              <button style={S.sideActionBtn} onClick={handleExport}>📥 Exportar Excel</button>
              <button style={S.sideActionBtn} onClick={handle2FAEnroll}>🔐 Configurar 2FA</button>
              <button style={{...S.sideActionBtn,color:"#f87171",borderColor:"#7f1d1d"}} onClick={handleLogout}>🚪 Sair</button>
            </div>
            <div style={S.sideFooter}>🔒 Dados criptografados no Supabase com Row Level Security</div>
          </aside>
          <main style={S.desktopMain}>
            <div style={S.desktopContent}>
              {view==="dashboard"&&dashboardJSX}
              {view==="add"      &&addFormJSX}
              {view==="budget"   &&budgetFormJSX}
              {view==="invest"   &&investFormJSX}
            </div>
          </main>
        </div>
      ) : (
        <>
          <div style={S.header} className="mobile-header">
            <span style={S.logo}>₿alance</span>
            <div style={S.headerRight}>
              <div style={S.monthPicker}>
                <button style={S.arrowBtn} onClick={()=>setSelectedMonth(m=>(m-1+12)%12)}>‹</button>
                <span style={S.monthLabel}>{MONTHS_FULL[selectedMonth]}</span>
                <button style={S.arrowBtn} onClick={()=>setSelectedMonth(m=>(m+1)%12)}>›</button>
              </div>
              <button style={S.iconBtn} onClick={handleExport}>📥</button>
              <button style={S.iconBtn} onClick={()=>setHidden(h=>!h)}>{hidden?"🙈":"👁️"}</button>
              <button style={S.iconBtn} onClick={handle2FAEnroll}>🔐</button>
              <button style={S.iconBtn} onClick={handleLogout}>🚪</button>
            </div>
          </div>
          <div style={S.content}>
            {view==="dashboard"&&dashboardJSX}
            {view==="add"      &&addFormJSX}
            {view==="budget"   &&budgetFormJSX}
            {view==="invest"   &&investFormJSX}
          </div>
          <div style={S.nav} className="mobile-nav">
            <NavButton id="dashboard" icon="📊" label="Resumo"/>
            <button style={S.addBtn} onClick={()=>setView("add")}>
              <span style={{fontSize:26,lineHeight:1}}>+</span>
            </button>
            <NavButton id="budget" icon="🎯" label="Orçamento" onClick={openBudget}/>
            <NavButton id="invest" icon="💎" label="Invest." onClick={openInvest}/>
          </div>
        </>
      )}
      {/* 2FA Enrollment Modal */}
      {show2FA&&(
        <div style={eStyles.overlay} onClick={()=>setShow2FA(false)}>
          <div style={eStyles.modal} onClick={e=>e.stopPropagation()}>
            <div style={eStyles.title}>🔐 Ativar Autenticação em 2 Fatores</div>
            <div style={{...eStyles.subtitle,marginBottom:16}}>
              Escaneie o QR code com Google Authenticator ou Authy
            </div>
            {totp2FAQr&&<img src={totp2FAQr} alt="QR 2FA" style={{width:160,height:160,margin:"0 auto 16px",display:"block",borderRadius:12,background:"#fff",padding:8}}/>}
            <div style={{marginBottom:20}}>
              <label style={eStyles.label}>Confirme com o código gerado</label>
              <input style={{...eStyles.input,textAlign:"center",fontSize:22,letterSpacing:8}}
                type="text" inputMode="numeric" maxLength={6} placeholder="000000"
                value={totp2FACode} onChange={e=>setTotp2FACode(e.target.value.replace(/\D/g,""))}/>
            </div>
            <div style={eStyles.btns}>
              <button style={eStyles.cancelBtn} onClick={()=>setShow2FA(false)}>Cancelar</button>
              <button style={eStyles.saveBtn} onClick={handle2FAConfirm} disabled={syncing}>
                {syncing?"Ativando...":"Ativar 2FA"}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── Budget bar styles ────────────────────────────────────────────────────────

const bStyles = {
  row:    { marginBottom:14, paddingBottom:14, borderBottom:"1px solid #0f172a" },
  rowTop: { display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:6 },
  catName:{ fontSize:13, fontWeight:600, color:"#cbd5e1" },
  amounts:{ display:"flex", alignItems:"center", gap:4 },
  spent:  { fontSize:14, fontWeight:800 },
  sep:    { fontSize:12, color:"#334155" },
  budget: { fontSize:12, color:"#e2e8f0", fontWeight:600 },
  barBg:  { height:6, background:"#0f172a", borderRadius:99, overflow:"hidden", marginBottom:4 },
  barFill:{ height:"100%", borderRadius:99, transition:"width 0.4s ease" },
  diffRow:{ display:"flex", justifyContent:"flex-end" },
};

// ─── Styles ───────────────────────────────────────────────────────────────────

const shared = {
  section:        {background:"#0a1628",borderRadius:16,padding:"14px 16px",marginBottom:14,border:"1px solid #0f2040"},
  sectionTitle:   {fontSize:12,fontWeight:700,color:"#60a5fa",textTransform:"uppercase",letterSpacing:1,marginBottom:12},
  cards:          {display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10},
  card:           {borderRadius:14,padding:"14px 16px"},
  cardGreen:      {background:"linear-gradient(135deg,#052e16,#14532d)",border:"1px solid #166534"},
  cardRed:        {background:"linear-gradient(135deg,#1c0a0a,#450a0a)",border:"1px solid #7f1d1d"},
  cardLabel:      {fontSize:11,fontWeight:600,color:"#94a3b8",textTransform:"uppercase",letterSpacing:1,marginBottom:6},
  cardValue:      {fontSize:18,fontWeight:800,color:"#f1f5f9"},
  balanceCard:    {borderRadius:14,padding:"16px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10,border:"1px solid #1e293b"},
  balanceLabel:   {fontSize:13,color:"#e2e8f0",fontWeight:600},
  balanceValue:   {fontSize:24,fontWeight:800},
  investSummaryCard:  {background:"linear-gradient(135deg,#1a0a2e,#2e1065)",border:"1px solid #4c1d95",borderRadius:14,padding:"14px 16px",marginBottom:10,display:"flex",alignItems:"center",justifyContent:"space-between",cursor:"pointer"},
  investSummaryLeft:  {display:"flex",flexDirection:"column",gap:6},
  investSummaryLabel: {fontSize:11,fontWeight:700,color:"#a855f7",textTransform:"uppercase",letterSpacing:1},
  investSummaryValue: {fontSize:20,fontWeight:800,color:"#e9d5ff"},
  investBankRow:      {display:"flex",gap:6,flexWrap:"wrap"},
  investBankChip:     {fontSize:10,fontWeight:600,color:"#94a3b8",border:"1px solid",borderRadius:6,padding:"2px 6px"},
  investArrow:        {fontSize:24,color:"#7c3aed"},
  catRow:     {display:"flex",justifyContent:"space-between",alignItems:"center",paddingBottom:10,marginBottom:10,borderBottom:"1px solid #0f172a"},
  catName:    {fontSize:14,color:"#cbd5e1"},
  catAmounts: {display:"flex",gap:8},
  amtGreen:   {fontSize:13,fontWeight:700,color:"#4ade80"},
  amtRed:     {fontSize:13,fontWeight:700,color:"#f87171"},
  legend:     {display:"flex",justifyContent:"center",gap:20,marginTop:8},
  legendItem: {display:"flex",alignItems:"center",gap:6,fontSize:12,color:"#64748b"},
  dot:        {width:8,height:8,borderRadius:"50%",display:"inline-block"},
  txRow:   {display:"flex",justifyContent:"space-between",alignItems:"flex-start",padding:"10px 0",borderBottom:"1px solid #0f172a"},
  txCat:   {fontSize:14,fontWeight:600,color:"#e2e8f0"},
  txDesc:  {fontSize:12,color:"#64748b",marginTop:2},
  txDate:  {fontSize:11,color:"#e2e8f0",marginTop:4},
  txRight: {display:"flex",alignItems:"center",gap:10},
  txAmt:   {fontSize:15,fontWeight:700},
  delBtn:  {background:"none",border:"none",color:"#e2e8f0",fontSize:14,cursor:"pointer",padding:4},
  editBtn: {background:"none",border:"none",color:"#475569",fontSize:14,cursor:"pointer",padding:4},
  empty:     {textAlign:"center",padding:"48px 20px"},
  emptyIcon: {fontSize:48,marginBottom:12},
  emptyText: {color:"#475569",fontWeight:600,fontSize:15,marginBottom:6},
  emptyHint: {color:"#334155",fontSize:13},
  exportBanner:      {background:"#0a1628",border:"1px solid #1e3a5f",borderRadius:14,padding:"14px 16px",display:"flex",alignItems:"center",gap:14,cursor:"pointer",marginTop:4},
  exportBannerTitle: {fontSize:14,fontWeight:700,color:"#93c5fd"},
  exportBannerSub:   {fontSize:11,color:"#475569",marginTop:2},
  loadingBar:  {background:"#0f172a",border:"1px solid #1e293b",borderRadius:10,padding:"8px 14px",fontSize:12,color:"#60a5fa",marginBottom:12,textAlign:"center"},
  form:       {padding:"4px 0"},
  typeToggle: {display:"flex",gap:10,marginBottom:16},
  typeBtn:    {flex:1,padding:"12px 0",borderRadius:12,border:"2px solid #0f172a",background:"#0a1628",color:"#64748b",fontWeight:700,fontSize:14,cursor:"pointer"},
  typeBtnActiveGreen: {borderColor:"#16a34a",color:"#4ade80",background:"#052e16"},
  typeBtnActiveRed:   {borderColor:"#dc2626",color:"#f87171",background:"#1c0a0a"},
  field:  {marginBottom:16},
  label:  {display:"block",fontSize:12,fontWeight:700,color:"#60a5fa",textTransform:"uppercase",letterSpacing:0.8,marginBottom:8},
  input:  {width:"100%",background:"#0a1628",border:"1.5px solid #0f172a",color:"#e2e8f0",borderRadius:10,padding:"12px 14px",fontSize:16,boxSizing:"border-box",outline:"none",fontFamily:"inherit"},
  catGrid:      {display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8},
  catBtn:       {background:"#0a1628",border:"1.5px solid #0f172a",borderRadius:10,padding:"10px 4px",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:4},
  catBtnActive: {borderColor:"#3b82f6",background:"#0f2040"},
  catBtnIcon:   {fontSize:20},
  catBtnLabel:  {fontSize:10,color:"#94a3b8",fontWeight:600,textAlign:"center"},
  saveBtn:      {width:"100%",padding:"15px",borderRadius:14,border:"none",background:"linear-gradient(135deg,#3b82f6,#6366f1)",color:"#fff",fontWeight:800,fontSize:16,cursor:"pointer",boxShadow:"0 4px 20px rgba(59,130,246,0.35)",marginTop:8},
  investHeader:       {marginBottom:24,textAlign:"center"},
  investTitle:        {fontSize:22,fontWeight:800,color:"#e9d5ff",marginBottom:4},
  investSubtitle:     {fontSize:13,color:"#7c3aed"},
  investLiveTotal:    {background:"linear-gradient(135deg,#1a0a2e,#2e1065)",border:"1px solid #4c1d95",borderRadius:12,padding:"14px 16px",display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16},
  investLiveTotalLabel: {fontSize:12,color:"#a855f7",fontWeight:700,textTransform:"uppercase",letterSpacing:0.8},
  investLiveTotalValue: {fontSize:18,fontWeight:800,color:"#e9d5ff"},
  monthPicker: {display:"flex",alignItems:"center",gap:8},
  arrowBtn:    {background:"#0f172a",border:"none",color:"#94a3b8",fontSize:20,width:30,height:30,borderRadius:8,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"},
  monthLabel:  {fontSize:13,fontWeight:600,color:"#e2e8f0",minWidth:80,textAlign:"center"},
  toast:   {position:"fixed",top:16,left:"50%",transform:"translateX(-50%)",color:"#fff",fontWeight:700,fontSize:14,padding:"10px 22px",borderRadius:99,zIndex:999,boxShadow:"0 4px 24px rgba(0,0,0,0.4)",animation:"fadeIn 0.2s ease"},
  syncDot: {position:"fixed",top:16,right:16,fontSize:18,zIndex:998,animation:"pulse 1.5s infinite"},
  navBtn:      {display:"flex",flexDirection:"column",alignItems:"center",gap:3,background:"none",border:"none",cursor:"pointer",padding:"6px 10px"},
  navBtnActive:{},
  navIcon:     {fontSize:20},
  navLabel:    {fontSize:10,fontWeight:600},
};

const mobileStyles = {
  ...shared,
  root:    {fontFamily:"'DM Sans','Helvetica Neue',sans-serif",background:"#020817",color:"#e2e8f0",minHeight:"100vh",maxWidth:430,margin:"0 auto",display:"flex",flexDirection:"column"},
  header:  {display:"flex",alignItems:"center",justifyContent:"space-between",padding:"12px 14px",borderBottom:"1px solid #0f172a",background:"#020817",position:"sticky",top:0,zIndex:10},
  headerRight: {display:"flex",alignItems:"center",gap:6},
  logo:    {fontSize:20,fontWeight:800,background:"linear-gradient(135deg,#60a5fa,#a78bfa)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent",letterSpacing:"-0.5px"},
  iconBtn: {background:"#0f172a",border:"1px solid #1e3a5f",borderRadius:8,fontSize:16,width:32,height:32,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"},
  content: {flex:1,overflowY:"auto",padding:"14px 14px 90px"},
  nav:     {position:"fixed",bottom:0,left:"50%",transform:"translateX(-50%)",width:"100%",maxWidth:430,background:"#020817",borderTop:"1px solid #0f172a",display:"flex",alignItems:"center",justifyContent:"space-around",padding:"8px 0 20px",zIndex:20},
  addBtn:  {width:50,height:50,borderRadius:"50%",background:"linear-gradient(135deg,#3b82f6,#6366f1)",border:"none",color:"#fff",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",boxShadow:"0 4px 20px rgba(59,130,246,0.4)",fontSize:26,lineHeight:1,marginBottom:4},
};

const desktopStyles = {
  ...shared,
  root:           {fontFamily:"'DM Sans','Helvetica Neue',sans-serif",background:"#020817",color:"#e2e8f0",minHeight:"100vh"},
  desktopWrapper: {display:"flex",minHeight:"100vh"},
  sidebar:        {width:240,background:"#04091a",borderRight:"1px solid #0f172a",display:"flex",flexDirection:"column",padding:"24px 16px",position:"fixed",top:0,left:0,height:"100vh",zIndex:10},
  sidebarLogo:    {fontSize:26,fontWeight:800,background:"linear-gradient(135deg,#60a5fa,#a78bfa)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent",letterSpacing:"-0.5px",marginBottom:4},
  sidebarEmail:   {fontSize:11,color:"#334155",marginBottom:20,wordBreak:"break-all"},
  sideNav:        {display:"flex",flexDirection:"column",gap:4,marginTop:8,flex:1},
  navBtn:         {display:"flex",alignItems:"center",gap:12,background:"none",border:"none",cursor:"pointer",padding:"11px 14px",borderRadius:10,textAlign:"left",width:"100%"},
  navBtnActive:   {background:"#0f172a"},
  navIcon:        {fontSize:20},
  navLabel:       {fontSize:14,fontWeight:600},
  sideActions:    {display:"flex",flexDirection:"column",gap:8,marginTop:16},
  sideActionBtn:  {background:"#0f172a",border:"1px solid #1e293b",borderRadius:10,color:"#94a3b8",fontSize:13,fontWeight:600,padding:"10px 14px",cursor:"pointer",textAlign:"left"},
  sideFooter:     {fontSize:10,color:"#1e293b",marginTop:16,lineHeight:1.5},
  desktopMain:    {marginLeft:240,flex:1,overflowY:"auto",minHeight:"100vh"},
  desktopContent: {maxWidth:820,margin:"0 auto",padding:"32px 32px 48px"},
  monthPicker:    {display:"flex",alignItems:"center",gap:8,marginBottom:24},
  arrowBtn:       {background:"#0f172a",border:"none",color:"#94a3b8",fontSize:20,width:32,height:32,borderRadius:8,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"},
  monthLabel:     {fontSize:15,fontWeight:700,color:"#e2e8f0",minWidth:120,textAlign:"center"},
  cardValue:      {fontSize:22,fontWeight:800,color:"#f1f5f9"},
  balanceValue:   {fontSize:28,fontWeight:800},
  catGrid:        {display:"grid",gridTemplateColumns:"repeat(6,1fr)",gap:8},
  toast:  {position:"fixed",top:20,left:"50%",transform:"translateX(-50%)",color:"#fff",fontWeight:700,fontSize:14,padding:"10px 22px",borderRadius:99,zIndex:999,boxShadow:"0 4px 24px rgba(0,0,0,0.4)",animation:"fadeIn 0.2s ease"},
  syncDot:{position:"fixed",top:20,right:24,fontSize:20,zIndex:998,animation:"pulse 1.5s infinite"},
};

const eStyles = {
  overlay:   {position:"fixed",top:0,left:0,right:0,bottom:0,background:"rgba(0,0,0,0.7)",zIndex:500,display:"flex",alignItems:"flex-end",justifyContent:"center"},
  modal:     {background:"#0a1628",border:"1px solid #1e293b",borderRadius:"20px 20px 0 0",padding:"24px 20px 40px",width:"100%",maxWidth:480},
  title:     {fontSize:18,fontWeight:800,color:"#e2e8f0",marginBottom:4},
  subtitle:  {fontSize:12,color:"#475569",marginBottom:20},
  label:     {display:"block",fontSize:11,fontWeight:700,color:"#60a5fa",textTransform:"uppercase",letterSpacing:0.8,marginBottom:6},
  input:     {width:"100%",background:"#020817",border:"1.5px solid #1e293b",color:"#e2e8f0",borderRadius:10,padding:"12px 14px",fontSize:16,boxSizing:"border-box",outline:"none",fontFamily:"inherit"},
  btns:      {display:"flex",gap:10,marginBottom:12},
  cancelBtn: {flex:1,padding:"13px",borderRadius:12,border:"1.5px solid #1e293b",background:"none",color:"#64748b",fontWeight:700,fontSize:15,cursor:"pointer"},
  saveBtn:   {flex:2,padding:"13px",borderRadius:12,border:"none",background:"linear-gradient(135deg,#3b82f6,#6366f1)",color:"#fff",fontWeight:800,fontSize:15,cursor:"pointer"},
  deleteBtn: {width:"100%",padding:"12px",borderRadius:12,border:"1.5px solid #7f1d1d",background:"none",color:"#f87171",fontWeight:700,fontSize:14,cursor:"pointer"},
};

const css = `
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;600;700;800&display=swap');
  *{box-sizing:border-box;margin:0;padding:0;}
  body{background:#020817;}
  ::-webkit-scrollbar{display:none;}
  input::placeholder{color:#334155;}
  input[type="date"]::-webkit-calendar-picker-indicator{filter:invert(1) opacity(0.4);}
  button:hover{opacity:0.85;}
  button:disabled{opacity:0.6;cursor:not-allowed;}
  @keyframes fadeIn{from{opacity:0;transform:translateX(-50%) translateY(-8px)}to{opacity:1;transform:translateX(-50%) translateY(0)}}
  @keyframes pulse{0%,100%{opacity:1}50%{opacity:0.4}}
  .mobile-header{padding-top:max(12px, env(safe-area-inset-top)) !important;}
  .mobile-nav{padding-bottom:max(20px, env(safe-area-inset-bottom)) !important;}
`;
