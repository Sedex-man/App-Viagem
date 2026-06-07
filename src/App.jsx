import { useState, useMemo, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { getProductImage, imageCache } from './imageService';

// ─── COLORS ──────────────────────────────────────────────────────────────────
const C = {
  bg:"#F8FAFC",bgCard:"#FFFFFF",border:"#E5E7EB",borderLight:"#F1F5F9",
  primary:"#2563EB",primaryLight:"#EFF6FF",gradientA:"#2563EB",gradientB:"#06B6D4",
  success:"#10B981",successLight:"#ECFDF5",warning:"#F59E0B",warningLight:"#FFFBEB",
  danger:"#EF4444",dangerLight:"#FEF2F2",purple:"#8B5CF6",purpleLight:"#F5F3FF",
  text:"#0F172A",textMid:"#475569",textLight:"#94A3B8",textXLight:"#CBD5E1",
};

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
// IOF padrão: 3,5%
const INITIAL_SETTINGS = { dollarPago:5.62, iof:3.5, spread:0.99, taxa:6.5, pesoMax:23000, totalDolarViagem:3000 };
const LOJAS_SUGESTOES = ["Walmart","Basspro","Target","HomeGoods","Dollar Tree","Amazon","Marshalls","Ross","TJ Maxx","Tommy Hilfiger","Calvin Klein","The North Face","Sephora","Ulta","Best Buy","Costco","Newegg","Apple","Restaurante","Uber","Passeio","Outro"];
const PRIORIDADES = ["Alta","Média","Baixa"];
const CATEGORIAS_GASTO = ["🛍 Compras","🍔 Alimentação","🚗 Transporte","🎢 Passeio","🏨 Hospedagem","💊 Farmácia","🎁 Presente","💳 Outros"];

// ─── CALC ─────────────────────────────────────────────────────────────────────
const calcDolarAjustado = s => s.dollarPago * (1 + (s.iof + s.spread) / 100);
const calcUsdFinal = (usd, s) => usd * (1 + s.taxa / 100);
const calcBRL = (usd, s) => calcUsdFinal(usd, s) * calcDolarAjustado(s);
const calcBRLPago = (usd, s, dp) => calcUsdFinal(usd, s) * dp;
const pesoGramas = p => p.tipo === "liquido" ? (parseFloat(p.volume)||0)*28.3495 : parseFloat(p.peso)||0;

function calcMinhaParteUSD(gasto) {
  const totalUSD = parseFloat(gasto.usd) || 0;
  if (!gasto.divisao || gasto.divisao.length === 0) return totalUSD;
  return totalUSD / (1 + gasto.divisao.length);
}
function usdToBRL(usd, gasto, settings) {
  const cotacao = parseFloat(gasto.dolarPago) || settings.dollarPago;
  return usd * cotacao;
}
function calcTotalGastosUSD(gastos) {
  return gastos.reduce((a, g) => a + calcMinhaParteUSD(g), 0);
}

// ─── FORMATTERS com vírgula como separador decimal ────────────────────────────
const ptBR2 = v => parseFloat(v||0).toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2});
const ptBR0 = v => parseFloat(v||0).toLocaleString("pt-BR",{minimumFractionDigits:0,maximumFractionDigits:0});
const ptBR4 = v => parseFloat(v||0).toLocaleString("pt-BR",{minimumFractionDigits:4,maximumFractionDigits:4});
const ptBR3 = v => parseFloat(v||0).toLocaleString("pt-BR",{minimumFractionDigits:3,maximumFractionDigits:3});
const fBRL  = (v, dec=2) => `R$ ${parseFloat(v||0).toLocaleString("pt-BR",{minimumFractionDigits:dec,maximumFractionDigits:dec})}`;
const fUSD  = (v, dec=2) => `US$ ${parseFloat(v||0).toLocaleString("pt-BR",{minimumFractionDigits:dec,maximumFractionDigits:dec})}`;
const fKg   = v => `${(parseFloat(v||0)/1000).toLocaleString("pt-BR",{minimumFractionDigits:3,maximumFractionDigits:3})} kg`;

// ─── AWESOMEAPI COTAÇÃO ──────────────────────────────────────────────────────
async function fetchCotacao() {
  try {
    const res = await fetch("https://economia.awesomeapi.com.br/last/USD-BRL", { signal: AbortSignal.timeout(5000) });
    const data = await res.json();
    const bid = parseFloat(data?.USDBRL?.bid);
    if (!isNaN(bid)) return bid;
  } catch {}
  return null;
}

const LOJA_EMOJI = { Amazon:"📦","Best Buy":"🔵",Walmart:"🟡",Target:"🎯",Newegg:"💻",Apple:"🍎",Costco:"🏪",Basspro:"🎣",HomeGoods:"🏠","Dollar Tree":"🌳",Marshalls:"🏷",Ross:"🏷","TJ Maxx":"🏷","Tommy Hilfiger":"👔","Calvin Klein":"👔","The North Face":"🏔",Sephora:"💄",Ulta:"💄",Restaurante:"🍔",Uber:"🚗",Passeio:"🎢",Outro:"🛒" };

function ProductImage({ produto, style={}, iconSize=44 }) {
  const [imgUrl, setImgUrl] = useState(imageCache[String(produto.id)] || null);
  const [loading, setLoading] = useState(!imgUrl);
  const [err, setErr] = useState(false);

  useEffect(() => {
    let cancelled = false;
    const key = String(produto.id);
    if (imageCache[key]) { setImgUrl(imageCache[key]); setLoading(false); return; }
    setLoading(true); setErr(false);
    getProductImage(produto).then(url => {
      if (cancelled) return;
      if (url) { setImgUrl(url); } else { setErr(true); }
      setLoading(false);
    }).catch(() => { if (cancelled) return; setErr(true); setLoading(false); });
    return () => { cancelled = true; };
  }, [produto.id, produto.link, produto.imagem]);

  if (!imgUrl || err) {
    return (
      <div style={{width:"100%",height:"100%",display:"flex",alignItems:"center",justifyContent:"center",background:`linear-gradient(135deg,${C.primaryLight},${C.purpleLight})`,...style}}>
        {loading ? <div className="spinner"/> : <span style={{fontSize:iconSize}}>{LOJA_EMOJI[produto.loja]||"🛍"}</span>}
      </div>
    );
  }
  return (
    <div style={{width:"100%",height:"100%",...style}}>
      <img src={imgUrl} alt={produto.nome} style={{width:"100%",height:"100%",objectFit:"cover"}} onError={()=>setErr(true)}/>
    </div>
  );
}

// ─── XLSX PARSER ─────────────────────────────────────────────────────────────
function parseSheet(wb, sheetName) {
  const ws = wb.Sheets[sheetName]; if (!ws) return [];
  const rows = XLSX.utils.sheet_to_json(ws,{header:1,defval:null});
  let hi = rows.findIndex(r=>r&&r.some(c=>c==="PRODUTO")); if (hi<0) return [];
  const hd=rows[hi], io=n=>hd.findIndex(h=>h===n);
  const [iL,iP,iC,iW,iLj,iV]=[io("Link"),io("PRODUTO"),io("Classificação"),io("Peso/oz"),io("LOCAL"),io("VALOR")];
  return rows.slice(hi+1).filter(r=>r[iP]&&typeof r[iP]==="string").map((r,i)=>{
    const tipo=(r[iC]||"").toString().toLowerCase().includes("líquido")||(r[iC]||"").toString().toLowerCase().includes("liquido")?"liquido":"solido";
    const oz=parseFloat(r[iW])||0;
    return { id:Date.now()+i+Math.random(), nome:r[iP].trim(), loja:(r[iLj]||"Outro").toString().trim(),
      usd:parseFloat(r[iV])||0, peso:tipo==="liquido"?oz*28.3495:oz, volume:tipo==="liquido"?oz:0,
      tipo, status:"pendente", prioridade:"Média", link:r[iL]?r[iL].toString().trim():"", imagem:"", dollarPago:null };
  });
}

// ─── SAMPLE DATA ─────────────────────────────────────────────────────────────
const SAMPLE_PRODUTOS = [
  {id:1,nome:"AirPods Pro 2",loja:"Apple",usd:249,peso:61,tipo:"solido",volume:0,status:"pendente",prioridade:"Alta",link:"https://www.amazon.com/Apple-AirPods-Pro-Cancellation-Transparency/dp/B0BDHWDR12",imagem:"",dollarPago:null},
  {id:2,nome:"Tide Pods 31ct",loja:"Walmart",usd:9.99,peso:640,tipo:"solido",volume:0,status:"comprado",prioridade:"Baixa",link:"https://www.walmart.com/ip/Tide-PODS-Laundry-Detergent-Packs-Original-Scent-31-Count/42351728",imagem:"",dollarPago:5.68},
  {id:3,nome:"Keychron K2",loja:"Amazon",usd:119,peso:870,tipo:"solido",volume:0,status:"pendente",prioridade:"Média",link:"",imagem:"",dollarPago:null},
];

const _initSettings    = INITIAL_SETTINGS;
const _initProdutos    = SAMPLE_PRODUTOS;
const _initItensLegais = [];
const _initGastos      = [];

// ─── APP ─────────────────────────────────────────────────────────────────────
export default function App() {
  const [tab, setTab] = useState(0);
  const [settings, setSettings] = useState(_initSettings);
  const [produtos, setProdutos] = useState(_initProdutos);
  const [itensLegais, setItensLegais] = useState(_initItensLegais);
  const [gastos, setGastos] = useState(_initGastos);
  const [showSettings, setShowSettings] = useState(false);
  const [showForm, setShowForm] = useState(false);
  const [showLegaisForm, setShowLegaisForm] = useState(false);
  const [showGastoForm, setShowGastoForm] = useState(false);
  const [editProd, setEditProd] = useState(null);
  const [editGasto, setEditGasto] = useState(null);
  const [notification, setNotification] = useState(null);
  const [cotacaoLoading, setCotacaoLoading] = useState(false);

  function notify(msg, type="success") { setNotification({msg,type}); setTimeout(()=>setNotification(null),2800); }

  function toggleStatus(id) {
    setProdutos(ps => ps.map(p => {
      if (p.id !== id) return p;
      const newStatus = p.status === "comprado" ? "pendente" : "comprado";
      if (newStatus === "comprado") {
        setGastos(gs => gs.some(g=>g.produtoId===id) ? gs : [...gs, {
          id:`prod_${id}`, produtoId:id, descricao:p.nome, loja:p.loja,
          usd:p.usd, dolarPago:p.dollarPago||settings.dollarPago,
          brl:null, categoria:"🛍 Compras", divisao:[], data:new Date().toLocaleDateString("pt-BR"), tipo:"produto"
        }]);
      } else {
        setGastos(gs => gs.filter(g => g.produtoId !== id));
      }
      return {...p, status: newStatus};
    }));
  }

  function saveGasto(g) {
    if (g.id) setGastos(gs=>gs.map(x=>x.id===g.id?g:x));
    else setGastos(gs=>[...gs,{...g,id:Date.now()}]);
    notify(g.id?"Gasto atualizado!":"Gasto adicionado!");
    setShowGastoForm(false); setEditGasto(null);
  }

  function deleteProd(id,list="produtos") {
    if(list==="legais") setItensLegais(ps=>ps.filter(p=>p.id!==id));
    else { setProdutos(ps=>ps.filter(p=>p.id!==id)); setGastos(gs=>gs.filter(g=>g.produtoId!==id)); }
    notify("Removido","error");
  }

  function saveProd(prod) {
    delete imageCache[String(prod.id)];
    if(prod._legais){ prod.id?setItensLegais(ps=>ps.map(p=>p.id===prod.id?prod:p)):setItensLegais(ps=>[...ps,{...prod,id:Date.now()}]); }
    else { prod.id?setProdutos(ps=>ps.map(p=>p.id===prod.id?prod:p)):setProdutos(ps=>[...ps,{...prod,id:Date.now()}]); }
    notify(prod.id?"Atualizado!":"Adicionado!");
    setShowForm(false); setShowLegaisForm(false); setEditProd(null);
  }

  function moveToList(item) {
    setProdutos(ps=>[...ps,{...item,_legais:undefined,status:"pendente",prioridade:"Média",id:Date.now()}]);
    setItensLegais(ps=>ps.filter(p=>p.id!==item.id)); notify("Movido para lista!");
  }

  function handleImport(compras,legais) {
    if(compras.length) setProdutos(prev=>{const e=new Set(prev.map(p=>p.nome.toLowerCase()));return [...prev,...compras.filter(p=>!e.has(p.nome.toLowerCase()))];});
    if(legais.length) setItensLegais(prev=>{const e=new Set(prev.map(p=>p.nome.toLowerCase()));return [...prev,...legais.filter(p=>!e.has(p.nome.toLowerCase()))];});
    notify(`✅ ${compras.length} compras + ${legais.length} itens importados!`);
    setShowSettings(false);
  }

  const stats = useMemo(()=>{
    const comprados=produtos.filter(p=>p.status==="comprado");
    const pesoTotal=produtos.reduce((a,p)=>a+pesoGramas(p),0);
    const valorTotalUSD=produtos.reduce((a,p)=>a+p.usd,0);
    const valorTotalBRL=produtos.reduce((a,p)=>a+calcBRL(p.usd,settings),0);
    const valorGasto=comprados.reduce((a,p)=>a+(p.dollarPago?calcBRLPago(p.usd,settings,p.dollarPago):calcBRL(p.usd,settings)),0);
    const totalMeusGastosUSD=calcTotalGastosUSD(gastos);
    return {total:produtos.length,comprados:comprados.length,pendentes:produtos.length-comprados.length,pesoTotal,valorTotalUSD,valorTotalBRL,valorGasto,lojas:new Set(produtos.map(p=>p.loja)).size,totalMeusGastosUSD};
  },[produtos,settings,gastos]);

  const pesoPercent=Math.min(100,(stats.pesoTotal/settings.pesoMax)*100);
  const pesoColor=pesoPercent<70?C.success:pesoPercent<90?C.warning:C.danger;
  const pesoBg=pesoPercent<70?C.successLight:pesoPercent<90?C.warningLight:C.dangerLight;

  const TABS=[{label:"Início",icon:"⊞"},{label:"Produtos",icon:"📦"},{label:"Galeria",icon:"▦"},{label:"Gastos",icon:"💸"},{label:"Stats",icon:"◈"},{label:"Calc",icon:"⟨⟩"}];

  return (
    <div style={S.app}>
      <style>{CSS}</style>
      {notification&&<div className={`notif notif-${notification.type}`}>{notification.msg}</div>}
      <div style={S.header}>
        <div style={S.headerLeft}>
          <div style={S.logoBox}>✈</div>
          <div><div style={S.headerTitle}>TravelShop</div><div style={S.headerSub}>Orlando 2027</div></div>
        </div>
        <button style={S.settingsBtn} onClick={()=>setShowSettings(true)}>
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke={C.textMid} strokeWidth="2" strokeLinecap="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>
        </button>
      </div>

      <div style={S.content}>
        {tab===0&&<DashboardTab stats={stats} settings={settings} pesoPercent={pesoPercent} pesoColor={pesoColor} pesoBg={pesoBg}/>}
        {tab===1&&<ProdutosTab produtos={produtos} itensLegais={itensLegais} settings={settings} onToggle={toggleStatus} onDelete={deleteProd} onEdit={p=>{setEditProd(p);setShowForm(true);}} onAdd={()=>{setEditProd(null);setShowForm(true);}} onAddLegais={()=>{setEditProd(null);setShowLegaisForm(true);}} onMoveToList={moveToList}/>}
        {tab===2&&<GaleriaTab produtos={produtos} itensLegais={itensLegais} settings={settings} onEdit={p=>{setEditProd(p);setShowForm(true);}}/>}
        {tab===3&&<GastosTab gastos={gastos} settings={settings} onAdd={()=>{setEditGasto(null);setShowGastoForm(true);}} onEdit={g=>{setEditGasto(g);setShowGastoForm(true);}} onDelete={id=>{setGastos(gs=>gs.filter(g=>g.id!==id));notify("Removido","error");}} onTogglePago={(gastoId,pessoaIdx)=>setGastos(gs=>gs.map(g=>g.id===gastoId?{...g,divisao:g.divisao.map((p,i)=>i===pessoaIdx?{...p,pago:!p.pago}:p)}:g))} produtos={produtos} onToggleStatus={toggleStatus}/>}
        {tab===4&&<StatsTab produtos={produtos} gastos={gastos} settings={settings}/>}
        {tab===5&&<CalcTab settings={settings}/>}
      </div>

      <nav style={S.nav}>
        {TABS.map((t,i)=>(
          <button key={i} style={{...S.navBtn,...(tab===i?S.navBtnActive:{})}} onClick={()=>setTab(i)}>
            <span style={{fontSize:18,lineHeight:1}}>{t.icon}</span>
            <span style={S.navLabel}>{t.label}</span>
            {tab===i&&<div style={S.navIndicator}/>}
          </button>
        ))}
      </nav>

      {tab===1&&<button style={S.fab} onClick={()=>{setEditProd(null);setShowForm(true);}}><span style={{fontSize:22,color:"white"}}>＋</span></button>}
      {tab===3&&<button style={S.fab} onClick={()=>{setEditGasto(null);setShowGastoForm(true);}}><span style={{fontSize:22,color:"white"}}>＋</span></button>}

      {showSettings&&<SettingsModal settings={settings} onSave={s=>{setSettings(s);notify("Configurações salvas!");}} onImport={handleImport} onClose={()=>setShowSettings(false)}/>}
      {showForm&&<ProdutoForm prod={editProd} isLegais={false} onSave={saveProd} onClose={()=>{setShowForm(false);setEditProd(null);}}/>}
      {showLegaisForm&&<ProdutoForm prod={null} isLegais={true} onSave={saveProd} onClose={()=>{setShowLegaisForm(false);}}/>}
      {showGastoForm&&<GastoForm gasto={editGasto} settings={settings} onSave={saveGasto} onClose={()=>{setShowGastoForm(false);setEditGasto(null);}}/>}
    </div>
  );
}

// ─── DASHBOARD ────────────────────────────────────────────────────────────────
function DashboardTab({stats,settings,pesoPercent,pesoColor,pesoBg}) {
  const dolarAj=calcDolarAjustado(settings);
  const pct=stats.total?Math.round(stats.comprados/stats.total*100):0;
  const usdRestante=settings.totalDolarViagem - stats.valorTotalUSD;
  return (
    <div style={S.page}>
      <div style={S.heroCard}>
        <div style={{fontSize:12,fontWeight:500,color:"rgba(255,255,255,0.75)",marginBottom:3}}>Total planejado</div>
        <div style={{fontSize:30,fontWeight:800,color:"#fff",letterSpacing:"-1px",lineHeight:1}}>{fBRL(stats.valorTotalBRL)}</div>
        <div style={{fontSize:13,color:"rgba(255,255,255,0.75)",marginTop:4}}>Dólar pago: R$ {ptBR4(settings.dollarPago)} · Ajustado: R$ {ptBR4(dolarAj)}</div>
        <div style={{display:"flex",gap:8,marginTop:14,flexWrap:"wrap"}}>
          {[`IOF ${ptBR2(settings.iof)}%`,`Spread ${ptBR2(settings.spread)}%`,`Taxa ${ptBR2(settings.taxa)}%`].map(t=>(
            <span key={t} style={{background:"rgba(255,255,255,0.15)",borderRadius:999,padding:"3px 10px",fontSize:11,color:"rgba(255,255,255,0.9)",fontWeight:500}}>{t}</span>
          ))}
        </div>
      </div>

      <div style={{...S.card,background:"linear-gradient(135deg,#F0FDF4,#ECFDF5)",border:`1px solid ${C.success}33`}}>
        <div style={{fontSize:12,fontWeight:600,color:C.success,marginBottom:8,textTransform:"uppercase",letterSpacing:"0.5px"}}>💵 Dólares na viagem</div>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end"}}>
          <div>
            <div style={{fontSize:28,fontWeight:800,color:C.success,fontFamily:"'DM Mono',monospace"}}>{fUSD(stats.valorTotalUSD)}</div>
            <div style={{fontSize:13,color:C.textMid,marginTop:2}}>planejado para compras</div>
          </div>
          <div style={{textAlign:"right"}}>
            <div style={{fontSize:16,fontWeight:700,color:usdRestante>=0?C.textMid:C.danger,fontFamily:"'DM Mono',monospace"}}>
              {usdRestante>=0?"Sobra":"Falta"} {fUSD(Math.abs(usdRestante))}
            </div>
            <div style={{fontSize:12,color:C.textLight}}>de {fUSD(settings.totalDolarViagem)} total</div>
          </div>
        </div>
        <div style={{height:6,background:"#D1FAE5",borderRadius:999,overflow:"hidden",marginTop:12}}>
          <div style={{width:`${Math.min(100,(stats.valorTotalUSD/settings.totalDolarViagem)*100)}%`,height:"100%",background:C.success,borderRadius:999}}/>
        </div>
      </div>

      <div style={S.grid4}>
        {[{label:"Produtos",value:stats.total,color:C.primary},{label:"Comprados",value:stats.comprados,color:C.success},{label:"Pendentes",value:stats.pendentes,color:C.warning},{label:"Lojas",value:stats.lojas,color:C.purple}].map(({label,value,color})=>(
          <div key={label} style={{...S.card,padding:"12px 8px",textAlign:"center",flex:1,marginBottom:0}}>
            <div style={{fontSize:20,fontWeight:800,color,fontFamily:"'DM Mono',monospace"}}>{value}</div>
            <div style={{fontSize:10,color:C.textLight,marginTop:2,fontWeight:500}}>{label}</div>
          </div>
        ))}
      </div>

      <div style={{...S.card,background:"linear-gradient(135deg,#FFF7ED,#FFFBEB)",border:`1px solid ${C.warning}33`}}>
        <div style={{fontSize:12,fontWeight:600,color:C.warning,marginBottom:8,textTransform:"uppercase",letterSpacing:"0.5px"}}>💸 Meus gastos reais</div>
        <div style={{fontSize:26,fontWeight:800,color:C.warning,fontFamily:"'DM Mono',monospace"}}>{fUSD(stats.totalMeusGastosUSD)}</div>
        <div style={{fontSize:13,color:C.textMid,marginTop:2}}>≈ {fBRL(stats.totalMeusGastosUSD*settings.dollarPago)} · apenas minha parte</div>
      </div>

      <div style={{...S.card,display:"flex",alignItems:"center",gap:20}}>
        <div style={{position:"relative",width:70,height:70,flexShrink:0}}>
          <svg width="70" height="70" viewBox="0 0 70 70">
            <circle cx="35" cy="35" r="27" fill="none" stroke={C.borderLight} strokeWidth="7"/>
            <circle cx="35" cy="35" r="27" fill="none" stroke={C.primary} strokeWidth="7" strokeDasharray={`${2*Math.PI*27}`} strokeDashoffset={`${2*Math.PI*27*(1-pct/100)}`} strokeLinecap="round" transform="rotate(-90 35 35)" style={{transition:"stroke-dashoffset 0.6s ease"}}/>
          </svg>
          <div style={{position:"absolute",inset:0,display:"flex",alignItems:"center",justifyContent:"center"}}>
            <div style={{fontSize:14,fontWeight:800,color:C.primary,fontFamily:"'DM Mono',monospace"}}>{pct}%</div>
          </div>
        </div>
        <div style={{flex:1}}>
          <div style={{fontWeight:700,fontSize:14,color:C.text,marginBottom:3}}>Progresso de compras</div>
          <div style={{fontSize:12,color:C.textMid,marginBottom:8}}>{stats.comprados}/{stats.total} itens comprados</div>
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
            <span style={{fontSize:11,color:C.textLight}}>⚖ Peso: {fKg(stats.pesoTotal)} / {fKg(settings.pesoMax)}</span>
            <span style={{fontSize:11,color:pesoColor,fontWeight:700}}>{pesoPercent.toFixed(0)}%</span>
          </div>
          <div style={{height:5,background:C.borderLight,borderRadius:999,overflow:"hidden"}}>
            <div style={{width:`${pesoPercent}%`,height:"100%",background:pesoColor,borderRadius:999}}/>
          </div>
        </div>
      </div>

      <div style={S.card}>
        <div style={{fontWeight:700,fontSize:14,color:C.text,marginBottom:12}}>💰 Resumo financeiro</div>
        {[
          {label:"USD total planejado",value:fUSD(stats.valorTotalUSD),color:C.primary},
          {label:"BRL previsto (c/ taxas)",value:fBRL(stats.valorTotalBRL),color:C.text},
          {label:"BRL já gasto",value:fBRL(stats.valorGasto),color:C.success},
          {label:"Meus gastos (USD)",value:fUSD(stats.totalMeusGastosUSD),color:C.warning},
        ].map(({label,value,color})=>(
          <div key={label} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"9px 0",borderBottom:`1px solid ${C.borderLight}`}}>
            <span style={{fontSize:13,color:C.textMid}}>{label}</span>
            <span style={{fontSize:14,fontWeight:700,color,fontFamily:"'DM Mono',monospace"}}>{value}</span>
          </div>
        ))}
      </div>
      {pesoPercent>=80&&<div style={{background:pesoBg,border:`1px solid ${pesoColor}33`,borderRadius:12,padding:"10px 14px",fontSize:13,color:pesoColor,fontWeight:600,marginBottom:12}}>⚠ Peso da mala em {pesoPercent.toFixed(0)}% do limite!</div>}
    </div>
  );
}

// ─── GASTOS TAB ───────────────────────────────────────────────────────────────
function GastosTab({gastos,settings,onAdd,onEdit,onDelete,onTogglePago,produtos,onToggleStatus}) {
  const [filtro,setFiltro]=useState("todos");
  const totalUSD=gastos.reduce((a,g)=>a+calcMinhaParteUSD(g),0);
  const aReceberUSD=gastos.reduce((a,g)=>{
    if(!g.divisao||g.divisao.length===0) return a;
    const part=(parseFloat(g.usd)||0)/(1+g.divisao.length);
    return a+g.divisao.filter(p=>!p.pago).length*part;
  },0);
  const filtrados=gastos.filter(g=>{
    if(filtro==="compras") return g.tipo==="produto";
    if(filtro==="livres") return g.tipo!=="produto";
    return true;
  }).sort((a,b)=>(b.id||0)-(a.id||0));

  return (
    <div style={S.page}>
      <div style={S.heroCard}>
        <div style={{fontSize:12,fontWeight:500,color:"rgba(255,255,255,0.75)",marginBottom:4}}>Meus gastos totais</div>
        <div style={{fontSize:30,fontWeight:800,color:"#fff",letterSpacing:"-1px"}}>{fUSD(totalUSD)}</div>
        <div style={{fontSize:13,color:"rgba(255,255,255,0.65)",marginTop:4}}>≈ {fBRL(totalUSD*settings.dollarPago)}</div>
        <div style={{display:"flex",gap:12,marginTop:12}}>
          <div style={{background:"rgba(255,255,255,0.15)",borderRadius:12,padding:"8px 14px",flex:1,textAlign:"center"}}>
            <div style={{fontSize:11,color:"rgba(255,255,255,0.7)",marginBottom:2}}>A receber</div>
            <div style={{fontSize:15,fontWeight:700,color:"#fff",fontFamily:"'DM Mono',monospace"}}>{fUSD(aReceberUSD)}</div>
          </div>
          <div style={{background:"rgba(255,255,255,0.15)",borderRadius:12,padding:"8px 14px",flex:1,textAlign:"center"}}>
            <div style={{fontSize:11,color:"rgba(255,255,255,0.7)",marginBottom:2}}>Gastos</div>
            <div style={{fontSize:15,fontWeight:700,color:"#fff",fontFamily:"'DM Mono',monospace"}}>{gastos.length}</div>
          </div>
        </div>
      </div>

      <div style={{display:"flex",gap:4,background:C.borderLight,borderRadius:12,padding:4,marginBottom:12}}>
        {[["todos","Todos"],["compras","🛍 Compras"],["livres","✏ Manuais"]].map(([v,l])=>(
          <button key={v} style={{flex:1,padding:"8px 4px",borderRadius:9,border:"none",cursor:"pointer",fontSize:12,fontWeight:600,background:filtro===v?C.bgCard:"transparent",color:filtro===v?C.primary:C.textMid,boxShadow:filtro===v?"0 1px 4px rgba(0,0,0,0.08)":"none"}} onClick={()=>setFiltro(v)}>{l}</button>
        ))}
      </div>

      {filtrados.length===0&&<Empty text="Nenhum gasto ainda. Marque produtos como comprados ou adicione gastos manualmente."/>}
      {filtrados.map(g=><GastoCard key={g.id} g={g} settings={settings} onEdit={()=>onEdit(g)} onDelete={()=>onDelete(g.id)} onTogglePago={onTogglePago}/>)}
    </div>
  );
}

function GastoCard({g,settings,onEdit,onDelete,onTogglePago}) {
  const [expanded,setExpanded]=useState(false);
  const totalUSD=parseFloat(g.usd)||0;
  const minhaUSD=calcMinhaParteUSD(g);
  const minhaBRL=usdToBRL(minhaUSD,g,settings);
  const totalBRL=usdToBRL(totalUSD,g,settings);
  const cotUsada=parseFloat(g.dolarPago)||settings.dollarPago;
  const temDivisao=g.divisao&&g.divisao.length>0;
  const totalPessoas=temDivisao?1+g.divisao.length:1;
  const aReceberUSD=temDivisao?g.divisao.filter(p=>!p.pago).length*(totalUSD/totalPessoas):0;

  return (
    <div style={{...S.card,marginBottom:10}}>
      <div style={{display:"flex",gap:12,alignItems:"flex-start"}}>
        <div style={{width:40,height:40,borderRadius:12,background:g.tipo==="produto"?C.primaryLight:C.purpleLight,display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0}}>
          {LOJA_EMOJI[g.loja]||g.categoria?.split(" ")[0]||"💳"}
        </div>
        <div style={{flex:1}} onClick={()=>setExpanded(e=>!e)}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
            <div>
              <div style={{fontWeight:600,fontSize:14,color:C.text,lineHeight:1.3}}>{g.descricao}</div>
              <div style={{fontSize:12,color:C.textLight,marginTop:2}}>{g.loja||g.categoria} · {g.data}</div>
            </div>
            <div style={{textAlign:"right",flexShrink:0}}>
              <div style={{fontSize:15,fontWeight:800,color:C.primary,fontFamily:"'DM Mono',monospace"}}>{fUSD(minhaUSD)}</div>
              <div style={{fontSize:11,color:C.textLight,fontFamily:"'DM Mono',monospace"}}>≈ {fBRL(minhaBRL)}</div>
              {temDivisao&&<div style={{fontSize:11,color:C.textLight}}>de {fUSD(totalUSD)} total</div>}
            </div>
          </div>
          {temDivisao&&(
            <div style={{display:"flex",gap:4,marginTop:8,flexWrap:"wrap"}}>
              {g.divisao.map((p,i)=>(
                <button key={i} onClick={e=>{e.stopPropagation();onTogglePago(g.id,i);}}
                  style={{display:"flex",alignItems:"center",gap:4,background:p.pago?C.successLight:C.dangerLight,border:`1px solid ${p.pago?C.success:C.danger}33`,borderRadius:999,padding:"3px 8px",fontSize:11,fontWeight:600,color:p.pago?C.success:C.danger,cursor:"pointer"}}>
                  {p.pago?"✓":"○"} {p.nome}
                </button>
              ))}
              {aReceberUSD>0&&<span style={{...S.tag,background:C.warningLight,color:C.warning,borderColor:C.warning+"33",fontSize:11}}>A receber {fUSD(aReceberUSD)}</span>}
            </div>
          )}
        </div>
      </div>
      {expanded&&(
        <div style={{marginTop:12,paddingTop:12,borderTop:`1px solid ${C.borderLight}`}}>
          {[
            {label:"Total USD",value:fUSD(totalUSD),color:C.primary},
            {label:"Minha parte USD",value:fUSD(minhaUSD),color:C.primary},
            {label:"Minha parte BRL",value:fBRL(minhaBRL),color:C.textMid},
            {label:"Cotação usada",value:`R$ ${ptBR4(cotUsada)}`},
            ...(temDivisao?[{label:"Dividido entre",value:`${totalPessoas} pessoas`}]:[]),
          ].map(({label,value,color})=>(
            <div key={label} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:`1px solid ${C.borderLight}`}}>
              <span style={{fontSize:13,color:C.textMid}}>{label}</span>
              <span style={{fontSize:13,fontWeight:700,color:color||C.text,fontFamily:"'DM Mono',monospace"}}>{value}</span>
            </div>
          ))}
          {temDivisao&&(
            <>
              <div style={{fontSize:12,fontWeight:600,color:C.textMid,marginTop:10,marginBottom:6}}>Status de pagamento:</div>
              {g.divisao.map((p,i)=>(
                <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:`1px solid ${C.borderLight}`}}>
                  <span style={{fontSize:13,color:C.text}}>{p.nome}</span>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <span style={{fontSize:13,fontFamily:"'DM Mono',monospace",color:C.textMid}}>{fUSD(totalUSD/totalPessoas)}</span>
                    <button onClick={()=>onTogglePago(g.id,i)} style={{background:p.pago?C.successLight:C.bg,border:`1px solid ${p.pago?C.success:C.border}`,borderRadius:8,padding:"3px 10px",fontSize:12,fontWeight:600,color:p.pago?C.success:C.textMid,cursor:"pointer"}}>
                      {p.pago?"✓ Pago":"Pendente"}
                    </button>
                  </div>
                </div>
              ))}
            </>
          )}
          {g.tipo!=="produto"&&(
            <div style={{display:"flex",gap:8,marginTop:12}}>
              <button style={S.btnOutline} onClick={onEdit}>✏ Editar</button>
              <button style={{...S.btnOutline,color:C.danger,borderColor:C.danger+"44"}} onClick={onDelete}>🗑 Remover</button>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

// ─── GASTO FORM ───────────────────────────────────────────────────────────────
function GastoForm({gasto,settings,onSave,onClose}) {
  const cotacaoUsada = gasto?.dolarPago || settings.dollarPago;
  const empty={descricao:"",loja:"",usd:"",dolarPago:settings.dollarPago,categoria:"🍔 Alimentação",divisao:[],data:new Date().toLocaleDateString("pt-BR")};
  const [f,setF]=useState(gasto?{...gasto,usd:(gasto.usd||"").toString(),dolarPago:gasto.dolarPago||settings.dollarPago}:empty);
  const [novaPessoa,setNovaPessoa]=useState("");

  const totalUSD=parseFloat(f.usd)||0;
  const cotacao=parseFloat(f.dolarPago)||settings.dollarPago;
  const totalBRL=totalUSD*cotacao;
  const nPessoas=1+f.divisao.length;
  const minhaParteUSD=nPessoas>0?totalUSD/nPessoas:totalUSD;
  const minhaParteBRL=minhaParteUSD*cotacao;

  function addPessoa(){if(!novaPessoa.trim())return;setF(p=>({...p,divisao:[...p.divisao,{nome:novaPessoa.trim(),pago:false}]}));setNovaPessoa("");}
  function removePessoa(i){setF(p=>({...p,divisao:p.divisao.filter((_,idx)=>idx!==i)}));}

  function handleSave(){
    if(!f.descricao)return alert("Informe a descrição");
    if(!f.usd)return alert("Informe o valor em USD");
    onSave({...f,usd:parseFloat(f.usd),brl:null,dolarPago:f.dolarPago||settings.dollarPago,tipo:"livre"});
  }

  return (
    <Modal title={gasto?.id?"Editar gasto":"Novo gasto"} onClose={onClose}>
      <label style={S.label}>Descrição *</label>
      <input style={S.input} placeholder="Ex: Almoço no McDonald's" value={f.descricao} onChange={e=>setF(p=>({...p,descricao:e.target.value}))}/>
      <label style={S.label}>Categoria</label>
      <select style={S.input} value={f.categoria} onChange={e=>setF(p=>({...p,categoria:e.target.value}))}>
        {CATEGORIAS_GASTO.map(c=><option key={c}>{c}</option>)}
      </select>
      <label style={S.label}>Local / Loja (opcional)</label>
      <input style={S.input} placeholder="Ex: McDonald's International Drive" value={f.loja} onChange={e=>setF(p=>({...p,loja:e.target.value}))}/>
      <label style={S.label}>Valor em USD *</label>
      <input style={S.input} type="number" step="0.01" placeholder="Ex: 45.90" value={f.usd} onChange={e=>setF(p=>({...p,usd:e.target.value}))}/>
      <div style={{background:C.primaryLight,border:`1px solid ${C.primary}22`,borderRadius:9,padding:"8px 12px",fontSize:12,color:C.textMid,marginBottom:14,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <span>Cotação do dólar pago: <strong style={{color:C.primary}}>R$ {ptBR4(parseFloat(f.dolarPago)||settings.dollarPago)}</strong></span>
        <span style={{color:C.textLight,fontSize:11}}>automática das configurações</span>
      </div>
      <label style={S.label}>Data</label>
      <input style={S.input} placeholder="DD/MM/AAAA" value={f.data} onChange={e=>setF(p=>({...p,data:e.target.value}))}/>

      {totalUSD>0&&(
        <div style={{background:C.primaryLight,borderRadius:10,padding:"10px 14px",marginBottom:14}}>
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
            <span style={{fontSize:13,color:C.textMid}}>Total USD</span>
            <span style={{fontSize:14,fontWeight:800,color:C.primary,fontFamily:"'DM Mono',monospace"}}>{fUSD(totalUSD)}</span>
          </div>
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
            <span style={{fontSize:13,color:C.textMid}}>≈ Total BRL</span>
            <span style={{fontSize:13,fontWeight:600,color:C.textMid,fontFamily:"'DM Mono',monospace"}}>{fBRL(totalBRL)}</span>
          </div>
          <div style={{display:"flex",justifyContent:"space-between"}}>
            <span style={{fontSize:13,color:C.textMid}}>Minha parte ({nPessoas}p)</span>
            <span style={{fontSize:14,fontWeight:700,color:C.success,fontFamily:"'DM Mono',monospace"}}>{fUSD(minhaParteUSD)} · {fBRL(minhaParteBRL)}</span>
          </div>
        </div>
      )}

      <div style={{borderTop:`1px solid ${C.borderLight}`,paddingTop:14,marginBottom:14}}>
        <div style={{fontWeight:700,fontSize:14,color:C.text,marginBottom:4}}>Dividir com outras pessoas</div>
        <div style={{fontSize:12,color:C.textLight,marginBottom:10}}>Adicione quem vai dividir este gasto. O total será dividido igualmente entre você + as pessoas abaixo.</div>
        <div style={{display:"flex",gap:8,marginBottom:8}}>
          <input style={{...S.input,marginBottom:0,flex:1}} placeholder="Nome da pessoa" value={novaPessoa} onChange={e=>setNovaPessoa(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addPessoa()}/>
          <button style={{...S.btnOutline,whiteSpace:"nowrap",color:C.primary,borderColor:C.primary+"44"}} onClick={addPessoa}>＋ Add</button>
        </div>
        {f.divisao.map((p,i)=>(
          <div key={i} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"8px 12px",background:C.bg,borderRadius:10,marginBottom:6,border:`1px solid ${C.border}`}}>
            <div style={{display:"flex",alignItems:"center",gap:8}}>
              <div style={{width:28,height:28,borderRadius:"50%",background:C.primaryLight,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:700,color:C.primary}}>{p.nome[0].toUpperCase()}</div>
              <span style={{fontSize:13,fontWeight:600,color:C.text}}>{p.nome}</span>
            </div>
            <div style={{display:"flex",alignItems:"center",gap:8}}>
              {totalUSD>0&&<span style={{fontSize:12,color:C.textMid,fontFamily:"'DM Mono',monospace"}}>{fUSD(totalUSD/nPessoas)}</span>}
              <button onClick={()=>removePessoa(i)} style={{background:C.dangerLight,border:"none",borderRadius:6,width:22,height:22,cursor:"pointer",color:C.danger,fontSize:14}}>×</button>
            </div>
          </div>
        ))}
        {f.divisao.length>0&&(
          <div style={{background:C.successLight,border:`1px solid ${C.success}33`,borderRadius:10,padding:"8px 12px",fontSize:13,color:C.success,fontWeight:600}}>
            ✓ Dividindo entre {nPessoas} pessoas · Sua parte: {fUSD(minhaParteUSD)} · {fBRL(minhaParteBRL)}
          </div>
        )}
      </div>
      <button style={S.btnPrimary} onClick={handleSave}>{gasto?.id?"Salvar alterações":"Adicionar gasto"}</button>
    </Modal>
  );
}

// ─── PRODUTOS TAB ─────────────────────────────────────────────────────────────
function ProdutosTab({produtos,itensLegais,settings,onToggle,onDelete,onEdit,onAdd,onAddLegais,onMoveToList}) {
  const [subTab,setSubTab]=useState("compras");
  const [filterLoja,setFilterLoja]=useState("Todas");
  const [filterStatus,setFilterStatus]=useState("Todos");
  const [busca,setBusca]=useState("");
  const lista=subTab==="legais"?itensLegais:produtos;
  const filtered=useMemo(()=>lista.filter(p=>{
    if(filterLoja!=="Todas"&&p.loja!==filterLoja)return false;
    if(subTab!=="legais"&&filterStatus!=="Todos"&&p.status!==filterStatus)return false;
    if(busca&&!p.nome.toLowerCase().includes(busca.toLowerCase()))return false;
    return true;
  }),[lista,filterLoja,filterStatus,busca,subTab]);
  return (
    <div style={S.page}>
      <div style={{display:"flex",gap:4,background:C.borderLight,borderRadius:12,padding:4,marginBottom:16}}>
        {[["compras","🛒 Lista de compras"],["legais",`✨ Legais${itensLegais.length>0?` (${itensLegais.length})`:""}`]].map(([v,l])=>(
          <button key={v} style={{flex:1,padding:"9px 8px",borderRadius:9,border:"none",cursor:"pointer",fontSize:13,fontWeight:600,background:subTab===v?C.bgCard:"transparent",color:subTab===v?C.primary:C.textMid,boxShadow:subTab===v?"0 1px 4px rgba(0,0,0,0.08)":"none"}} onClick={()=>setSubTab(v)}>{l}</button>
        ))}
      </div>

      {/* ── Botão adicionar item legal ── */}
      {subTab==="legais"&&(
        <button
          style={{width:"100%",background:C.purpleLight,border:`1px solid ${C.purple}44`,color:C.purple,borderRadius:12,padding:"11px",fontSize:14,fontWeight:600,cursor:"pointer",marginBottom:12,display:"flex",alignItems:"center",justifyContent:"center",gap:6}}
          onClick={onAddLegais}
        >
          ✨ Adicionar item interessante
        </button>
      )}

      <input style={S.searchInput} placeholder="🔍 Buscar produto..." value={busca} onChange={e=>setBusca(e.target.value)}/>
      <div style={S.filterRow}>
        {["Todas",...new Set(lista.map(p=>p.loja))].map(l=><button key={l} style={{...S.chip,...(filterLoja===l?S.chipActive:{})}} onClick={()=>setFilterLoja(l)}>{l}</button>)}
      </div>
      {subTab==="compras"&&(
        <div style={S.filterRow}>
          {[["Todos","Todos"],["pendente","⏳ Pendentes"],["comprado","✅ Comprados"]].map(([v,l])=><button key={v} style={{...S.chip,...(filterStatus===v?S.chipActive:{})}} onClick={()=>setFilterStatus(v)}>{l}</button>)}
        </div>
      )}
      <div style={{fontSize:12,color:C.textLight,marginBottom:10,fontWeight:500}}>{filtered.length} item(s)</div>
      {filtered.map(p=><ProdutoCard key={p.id} p={p} settings={settings} onToggle={subTab==="compras"?()=>onToggle(p.id):null} onDelete={()=>onDelete(p.id,subTab==="legais"?"legais":"produtos")} onEdit={()=>onEdit({...p,_legais:subTab==="legais"})} onMoveToList={subTab==="legais"?()=>onMoveToList(p):null} isLegais={subTab==="legais"}/>)}
      {filtered.length===0&&lista.length>0&&<Empty text="Nenhum item encontrado"/>}
      {subTab==="legais"&&itensLegais.length===0&&<Empty text="Nenhum item interessante ainda. Toque no botão acima para adicionar!"/>}
    </div>
  );
}

function ProdutoCard({p,settings,onToggle,onDelete,onEdit,onMoveToList,isLegais}) {
  const [expanded,setExpanded]=useState(false);
  const brl=calcBRL(p.usd,settings);
  const brlPago=p.dollarPago?calcBRLPago(p.usd,settings,p.dollarPago):null;
  const peso=pesoGramas(p);
  const prioColors={Alta:{color:C.danger,bg:C.dangerLight},Média:{color:C.warning,bg:C.warningLight},Baixa:{color:C.primary,bg:C.primaryLight}};
  const pc=prioColors[p.prioridade]||prioColors["Média"];
  return (
    <div style={{...S.card,marginBottom:10,opacity:p.status==="comprado"?0.75:1}}>
      <div style={{display:"flex",gap:10,alignItems:"flex-start"}}>
        <div style={{width:50,height:50,borderRadius:12,overflow:"hidden",flexShrink:0,border:`1px solid ${C.border}`}}>
          <ProductImage produto={p} iconSize={22}/>
        </div>
        {!isLegais&&onToggle&&<button style={{...S.checkbox,...(p.status==="comprado"?S.checkboxDone:{})}} onClick={onToggle}>{p.status==="comprado"&&<svg width="12" height="12" viewBox="0 0 12 12"><polyline points="2,6 5,9 10,3" stroke="white" strokeWidth="2" fill="none" strokeLinecap="round"/></svg>}</button>}
        <div style={{flex:1}} onClick={()=>setExpanded(e=>!e)}>
          <div style={{display:"flex",justifyContent:"space-between",gap:8}}>
            <div style={{fontWeight:600,fontSize:14,color:p.status==="comprado"?C.textLight:C.text,textDecoration:p.status==="comprado"?"line-through":"none",lineHeight:1.3,flex:1}}>{p.nome}</div>
            <div style={{textAlign:"right",flexShrink:0}}>
              <div style={{fontSize:14,fontWeight:800,color:C.primary,fontFamily:"'DM Mono',monospace"}}>{fUSD(p.usd)}</div>
              <div style={{fontSize:11,color:C.textLight,fontFamily:"'DM Mono',monospace"}}>{fBRL(brl,0)}</div>
            </div>
          </div>
          <div style={{display:"flex",gap:5,marginTop:6,flexWrap:"wrap"}}>
            <span style={S.tag}>{p.loja}</span>
            {!isLegais&&<span style={{...S.tag,background:pc.bg,color:pc.color,borderColor:pc.color+"33"}}>{p.prioridade}</span>}
            <span style={S.tag}>{fKg(peso)}</span>
            {p.status==="comprado"&&<span style={{...S.tag,background:C.successLight,color:C.success,borderColor:C.success+"33"}}>✓ Comprado</span>}
          </div>
        </div>
      </div>
      {expanded&&(
        <div style={{marginTop:12,paddingTop:12,borderTop:`1px solid ${C.borderLight}`}}>
          {[{label:"BRL previsto",value:fBRL(brl)},...(brlPago?[{label:"BRL pago",value:fBRL(brlPago),color:C.success},{label:"Diferença",value:fBRL(Math.abs(brl-brlPago)),color:brl>brlPago?C.success:C.danger}]:[]),{label:"USD c/ taxa",value:fUSD(calcUsdFinal(p.usd,settings)),color:C.textMid}].map(({label,value,color})=>(
            <div key={label} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:`1px solid ${C.borderLight}`}}>
              <span style={{fontSize:13,color:C.textMid}}>{label}</span>
              <span style={{fontSize:13,fontWeight:700,color:color||C.text,fontFamily:"'DM Mono',monospace"}}>{value}</span>
            </div>
          ))}
          {p.link&&<a href={p.link} target="_blank" rel="noreferrer" style={{display:"block",fontSize:13,color:C.primary,marginTop:8}}>🔗 Ver produto</a>}
          <div style={{display:"flex",gap:8,marginTop:10}}>
            <button style={S.btnOutline} onClick={onEdit}>✏ Editar</button>
            {onMoveToList&&<button style={{...S.btnOutline,color:C.success,borderColor:C.success+"44"}} onClick={onMoveToList}>📋 Mover p/ lista</button>}
            <button style={{...S.btnOutline,color:C.danger,borderColor:C.danger+"44"}} onClick={onDelete}>🗑</button>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── GALERIA TAB ──────────────────────────────────────────────────────────────
function GaleriaTab({produtos,itensLegais,settings,onEdit}) {
  const [subTab,setSubTab]=useState("compras");
  const lista=subTab==="legais"?itensLegais:produtos;
  return (
    <div style={S.page}>
      <div style={{display:"flex",gap:4,background:C.borderLight,borderRadius:12,padding:4,marginBottom:12}}>
        {[["compras","🛒 Compras"],["legais","✨ Legais"]].map(([v,l])=>(
          <button key={v} style={{flex:1,padding:"9px 8px",borderRadius:9,border:"none",cursor:"pointer",fontSize:13,fontWeight:600,background:subTab===v?C.bgCard:"transparent",color:subTab===v?C.primary:C.textMid,boxShadow:subTab===v?"0 1px 4px rgba(0,0,0,0.08)":"none"}} onClick={()=>setSubTab(v)}>{l}</button>
        ))}
      </div>
      <div style={{...S.card,background:"#F0F9FF",border:"1px solid #BAE6FD",padding:"10px 14px",marginBottom:12}}>
        <div style={{fontSize:12,color:"#0369A1"}}>🔍 Imagens buscadas via og:image do link ou DuckDuckGo. Adicione o link do produto para melhor resultado.</div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        {lista.map(p=>(
          <div key={p.id} style={{...S.card,padding:0,overflow:"hidden",cursor:"pointer"}} onClick={()=>onEdit({...p,_legais:subTab==="legais"})} className="galeria-card">
            <div style={{height:120,position:"relative",overflow:"hidden"}}>
              <ProductImage produto={p} iconSize={40}/>
              {p.status==="comprado"&&<div style={{position:"absolute",top:8,right:8,background:C.success,borderRadius:999,padding:"2px 8px",fontSize:10,color:"white",fontWeight:700}}>✓ Comprado</div>}
              {p.prioridade==="Alta"&&p.status!=="comprado"&&<div style={{position:"absolute",top:8,left:8,background:C.danger,borderRadius:999,padding:"2px 8px",fontSize:10,color:"white",fontWeight:700}}>Alta</div>}
            </div>
            <div style={{padding:"10px 12px"}}>
              <div style={{fontSize:12,fontWeight:600,color:C.text,lineHeight:1.3,marginBottom:5,overflow:"hidden",display:"-webkit-box",WebkitLineClamp:2,WebkitBoxOrient:"vertical"}}>{p.nome}</div>
              <div style={{fontSize:14,fontWeight:800,color:C.primary,fontFamily:"'DM Mono',monospace"}}>{fUSD(p.usd)}</div>
              <div style={{fontSize:11,color:C.textLight,fontFamily:"'DM Mono',monospace"}}>{fBRL(calcBRL(p.usd,settings),0)}</div>
            </div>
          </div>
        ))}
      </div>
      {lista.length===0&&<Empty text="Nenhum item ainda"/>}
    </div>
  );
}

// ─── STATS TAB ────────────────────────────────────────────────────────────────
function StatsTab({produtos,gastos,settings}) {
  const [secao,setSecao]=useState("compras");
  const barColors=[C.primary,C.purple,C.success,C.warning,C.danger,"#06B6D4","#F97316","#EC4899"];

  const porLoja=useMemo(()=>{
    const map={};
    produtos.forEach(p=>{
      if(!map[p.loja])map[p.loja]={total:0,comprados:0,usd:0,brl:0,peso:0};
      map[p.loja].total++; if(p.status==="comprado")map[p.loja].comprados++;
      map[p.loja].usd+=p.usd; map[p.loja].brl+=calcBRL(p.usd,settings); map[p.loja].peso+=pesoGramas(p);
    });
    return Object.entries(map).sort((a,b)=>b[1].usd-a[1].usd);
  },[produtos,settings]);
  const totalUSDCompras=porLoja.reduce((a,[,v])=>a+v.usd,0);

  const porCategoria=useMemo(()=>{
    const map={};
    gastos.forEach(g=>{
      const cat=g.categoria||"💳 Outros";
      if(!map[cat])map[cat]={total:0,usd:0,minhaUSD:0,aReceber:0};
      const usd=parseFloat(g.usd)||0; const minha=calcMinhaParteUSD(g);
      const nP=1+(g.divisao?.length||0);
      const aRec=g.divisao?g.divisao.filter(p=>!p.pago).length*(usd/nP):0;
      map[cat].total++; map[cat].usd+=usd; map[cat].minhaUSD+=minha; map[cat].aReceber+=aRec;
    });
    return Object.entries(map).sort((a,b)=>b[1].usd-a[1].usd);
  },[gastos]);

  const porPessoa=useMemo(()=>{
    const map={};
    gastos.forEach(g=>{
      if(!g.divisao||g.divisao.length===0) return;
      const usd=parseFloat(g.usd)||0; const nP=1+g.divisao.length; const part=usd/nP;
      g.divisao.forEach(p=>{
        if(!map[p.nome])map[p.nome]={totalUSD:0,pago:0,pendente:0,gastos:0};
        map[p.nome].totalUSD+=part; map[p.nome].gastos++;
        if(p.pago) map[p.nome].pago+=part; else map[p.nome].pendente+=part;
      });
    });
    return Object.entries(map).sort((a,b)=>b[1].totalUSD-a[1].totalUSD);
  },[gastos]);

  const totalGastosUSD=gastos.reduce((a,g)=>a+(parseFloat(g.usd)||0),0);
  const totalMeuUSD=gastos.reduce((a,g)=>a+calcMinhaParteUSD(g),0);
  const totalAReceberUSD=gastos.reduce((a,g)=>{
    if(!g.divisao||g.divisao.length===0) return a;
    const usd=parseFloat(g.usd)||0; const nP=1+g.divisao.length;
    return a+g.divisao.filter(p=>!p.pago).length*(usd/nP);
  },0);

  return (
    <div style={S.page}>
      <div style={{display:"flex",gap:4,background:C.borderLight,borderRadius:12,padding:4,marginBottom:14}}>
        {[["compras","🛒 Compras"],["gastos","💸 Gastos"]].map(([v,l])=>(
          <button key={v} style={{flex:1,padding:"9px 8px",borderRadius:9,border:"none",cursor:"pointer",fontSize:13,fontWeight:600,background:secao===v?C.bgCard:"transparent",color:secao===v?C.primary:C.textMid,boxShadow:secao===v?"0 1px 4px rgba(0,0,0,0.08)":"none",transition:"all 0.2s"}} onClick={()=>setSecao(v)}>{l}</button>
        ))}
      </div>

      {secao==="compras"&&(
        <>
          <div style={{display:"flex",gap:8,marginBottom:12}}>
            {[{label:"Total USD",value:fUSD(totalUSDCompras,0),color:C.primary},{label:"Comprados",value:`${produtos.filter(p=>p.status==="comprado").length}/${produtos.length}`,color:C.success}].map(({label,value,color})=>(
              <div key={label} style={{...S.card,flex:1,textAlign:"center",padding:"12px 8px",marginBottom:0}}>
                <div style={{fontSize:16,fontWeight:800,color,fontFamily:"'DM Mono',monospace"}}>{value}</div>
                <div style={{fontSize:11,color:C.textLight,marginTop:2}}>{label}</div>
              </div>
            ))}
          </div>
          <div style={S.card}>
            <div style={{fontWeight:700,fontSize:14,color:C.text,marginBottom:14}}>Compras por loja</div>
            {porLoja.map(([loja,d],i)=>{
              const pct=totalUSDCompras?(d.usd/totalUSDCompras*100):0;
              const color=barColors[i%barColors.length];
              return(
                <div key={loja} style={{marginBottom:14}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}>
                    <div style={{display:"flex",alignItems:"center",gap:8}}><div style={{width:10,height:10,borderRadius:"50%",background:color,flexShrink:0}}/><span style={{fontSize:13,fontWeight:600,color:C.text}}>{loja}</span></div>
                    <div style={{textAlign:"right"}}><span style={{fontSize:13,fontWeight:700,color,fontFamily:"'DM Mono',monospace"}}>{fUSD(d.usd,0)}</span><span style={{fontSize:11,color:C.textLight,marginLeft:6}}>{ptBR0(pct)}%</span></div>
                  </div>
                  <div style={{height:6,background:C.borderLight,borderRadius:999,overflow:"hidden",marginBottom:4}}><div style={{width:`${pct}%`,height:"100%",background:color,borderRadius:999,transition:"width 0.5s"}}/></div>
                  <div style={{fontSize:11,color:C.textLight}}>{d.comprados}/{d.total} comprados · {fKg(d.peso)}</div>
                </div>
              );
            })}
            {porLoja.length===0&&<div style={{fontSize:13,color:C.textLight,textAlign:"center",padding:"16px 0"}}>Nenhum produto ainda</div>}
          </div>
          <div style={S.card}>
            <div style={{fontWeight:700,fontSize:14,color:C.text,marginBottom:12}}>Ranking de lojas</div>
            <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr",gap:4,marginBottom:8}}>
              {["Loja","Itens","USD","Concl."].map(h=><div key={h} style={{fontSize:10,fontWeight:700,color:C.textLight,textTransform:"uppercase",letterSpacing:"0.4px"}}>{h}</div>)}
            </div>
            {porLoja.map(([loja,d])=>(
              <div key={loja} style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr",gap:4,padding:"9px 0",borderTop:`1px solid ${C.borderLight}`}}>
                <div style={{fontSize:12,fontWeight:600,color:C.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{loja}</div>
                <div style={{fontSize:12,color:C.textMid,fontFamily:"'DM Mono',monospace"}}>{d.total}</div>
                <div style={{fontSize:12,color:C.primary,fontFamily:"'DM Mono',monospace",fontWeight:700}}>{fUSD(d.usd,0)}</div>
                <div style={{fontSize:12,color:d.total&&d.comprados/d.total>=1?C.success:C.textMid,fontFamily:"'DM Mono',monospace"}}>{d.total?Math.round(d.comprados/d.total*100):0}%</div>
              </div>
            ))}
          </div>
        </>
      )}

      {secao==="gastos"&&(
        <>
          <div style={S.heroCard}>
            <div style={{fontSize:12,color:"rgba(255,255,255,0.7)",marginBottom:4}}>Total de gastos (bruto)</div>
            <div style={{fontSize:28,fontWeight:800,color:"#fff",fontFamily:"'DM Mono',monospace",letterSpacing:"-0.5px"}}>{fUSD(totalGastosUSD)}</div>
            <div style={{display:"flex",gap:8,marginTop:12}}>
              <div style={{background:"rgba(255,255,255,0.15)",borderRadius:10,padding:"8px 12px",flex:1,textAlign:"center"}}>
                <div style={{fontSize:10,color:"rgba(255,255,255,0.65)",marginBottom:2}}>Minha parte</div>
                <div style={{fontSize:14,fontWeight:800,color:"#fff",fontFamily:"'DM Mono',monospace"}}>{fUSD(totalMeuUSD)}</div>
              </div>
              <div style={{background:"rgba(255,255,255,0.15)",borderRadius:10,padding:"8px 12px",flex:1,textAlign:"center"}}>
                <div style={{fontSize:10,color:"rgba(255,255,255,0.65)",marginBottom:2}}>A receber</div>
                <div style={{fontSize:14,fontWeight:800,color:"#fff",fontFamily:"'DM Mono',monospace"}}>{fUSD(totalAReceberUSD)}</div>
              </div>
              <div style={{background:"rgba(255,255,255,0.15)",borderRadius:10,padding:"8px 12px",flex:1,textAlign:"center"}}>
                <div style={{fontSize:10,color:"rgba(255,255,255,0.65)",marginBottom:2}}>Itens</div>
                <div style={{fontSize:14,fontWeight:800,color:"#fff",fontFamily:"'DM Mono',monospace"}}>{gastos.length}</div>
              </div>
            </div>
          </div>

          {porCategoria.length>0&&(
            <div style={S.card}>
              <div style={{fontWeight:700,fontSize:14,color:C.text,marginBottom:14}}>Por categoria</div>
              {porCategoria.map(([cat,d],i)=>{
                const pct=totalGastosUSD?(d.usd/totalGastosUSD*100):0;
                const color=barColors[i%barColors.length];
                return(
                  <div key={cat} style={{marginBottom:14}}>
                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}>
                      <div style={{display:"flex",alignItems:"center",gap:8}}><div style={{width:10,height:10,borderRadius:"50%",background:color,flexShrink:0}}/><span style={{fontSize:13,fontWeight:600,color:C.text}}>{cat}</span></div>
                      <div style={{textAlign:"right"}}>
                        <div style={{fontSize:13,fontWeight:700,color,fontFamily:"'DM Mono',monospace"}}>{fUSD(d.usd)}</div>
                        <div style={{fontSize:11,color:C.textLight}}>meu: {fUSD(d.minhaUSD)} · {ptBR0(pct)}%</div>
                      </div>
                    </div>
                    <div style={{height:6,background:C.borderLight,borderRadius:999,overflow:"hidden"}}><div style={{width:`${pct}%`,height:"100%",background:color,borderRadius:999,transition:"width 0.5s"}}/></div>
                  </div>
                );
              })}
            </div>
          )}

          {porPessoa.length>0&&(
            <div style={S.card}>
              <div style={{fontWeight:700,fontSize:14,color:C.text,marginBottom:4}}>Divisão por pessoa</div>
              <div style={{fontSize:12,color:C.textLight,marginBottom:12}}>Quanto cada pessoa deve ao total dos gastos divididos</div>
              {porPessoa.map(([nome,d])=>(
                <div key={nome} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 0",borderBottom:`1px solid ${C.borderLight}`}}>
                  <div style={{display:"flex",alignItems:"center",gap:10}}>
                    <div style={{width:32,height:32,borderRadius:"50%",background:C.primaryLight,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,fontWeight:700,color:C.primary}}>{nome[0].toUpperCase()}</div>
                    <div><div style={{fontSize:13,fontWeight:600,color:C.text}}>{nome}</div><div style={{fontSize:11,color:C.textLight}}>{d.gastos} gasto(s)</div></div>
                  </div>
                  <div style={{textAlign:"right"}}>
                    <div style={{fontSize:14,fontWeight:800,color:C.primary,fontFamily:"'DM Mono',monospace"}}>{fUSD(d.totalUSD)}</div>
                    <div style={{display:"flex",gap:6,justifyContent:"flex-end",marginTop:3}}>
                      {d.pago>0&&<span style={{fontSize:11,color:C.success,fontWeight:600}}>✓ {fUSD(d.pago)}</span>}
                      {d.pendente>0&&<span style={{fontSize:11,color:C.danger,fontWeight:600}}>⏳ {fUSD(d.pendente)}</span>}
                    </div>
                  </div>
                </div>
              ))}
              {totalAReceberUSD>0&&(
                <div style={{marginTop:12,background:C.warningLight,border:`1px solid ${C.warning}33`,borderRadius:10,padding:"10px 14px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <span style={{fontSize:13,fontWeight:600,color:C.warning}}>⏳ Total a receber</span>
                  <span style={{fontSize:15,fontWeight:800,color:C.warning,fontFamily:"'DM Mono',monospace"}}>{fUSD(totalAReceberUSD)}</span>
                </div>
              )}
            </div>
          )}

          <div style={S.card}>
            <div style={{fontWeight:700,fontSize:14,color:C.text,marginBottom:12}}>Todos os gastos</div>
            <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr",gap:4,marginBottom:8}}>
              {["Descrição","USD total","Meu USD","Divisão"].map(h=><div key={h} style={{fontSize:10,fontWeight:700,color:C.textLight,textTransform:"uppercase",letterSpacing:"0.3px"}}>{h}</div>)}
            </div>
            {gastos.length===0&&<div style={{fontSize:13,color:C.textLight,textAlign:"center",padding:"16px 0"}}>Nenhum gasto ainda</div>}
            {gastos.map(g=>{
              const usd=parseFloat(g.usd)||0; const minha=calcMinhaParteUSD(g); const nP=1+(g.divisao?.length||0);
              return(
                <div key={g.id} style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr",gap:4,padding:"9px 0",borderTop:`1px solid ${C.borderLight}`,alignItems:"center"}}>
                  <div><div style={{fontSize:12,fontWeight:600,color:C.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{g.descricao}</div><div style={{fontSize:10,color:C.textLight}}>{g.loja||g.categoria}</div></div>
                  <div style={{fontSize:12,color:C.textMid,fontFamily:"'DM Mono',monospace"}}>{fUSD(usd)}</div>
                  <div style={{fontSize:12,color:C.primary,fontFamily:"'DM Mono',monospace",fontWeight:700}}>{fUSD(minha)}</div>
                  <div style={{fontSize:11,color:C.textMid}}>{nP>1?`÷${nP}p`:"-"}</div>
                </div>
              );
            })}
            {gastos.length>0&&(
              <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr",gap:4,padding:"10px 0",borderTop:`2px solid ${C.border}`,marginTop:4}}>
                <div style={{fontSize:12,fontWeight:700,color:C.text}}>Total</div>
                <div style={{fontSize:12,fontWeight:700,color:C.textMid,fontFamily:"'DM Mono',monospace"}}>{fUSD(totalGastosUSD)}</div>
                <div style={{fontSize:12,fontWeight:700,color:C.primary,fontFamily:"'DM Mono',monospace"}}>{fUSD(totalMeuUSD)}</div>
                <div/>
              </div>
            )}
          </div>
        </>
      )}
    </div>
  );
}

// ─── BCB RATE COMPONENT (somente na calculadora) ─────────────────────────────
function BcbRate() {
  const [rate, setRate] = useState(null);
  const [loading, setLoading] = useState(false);
  const [lastFetch, setLastFetch] = useState(null);
  async function fetch_() {
    setLoading(true);
    const val = await fetchCotacao();
    if (val) { setRate(val); setLastFetch(new Date().toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"})); }
    setLoading(false);
  }
  return (
    <div style={{marginTop:12,paddingTop:12,borderTop:`1px solid ${C.borderLight}`}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div>
          <div style={{fontSize:12,fontWeight:600,color:C.textMid}}>Cotação de mercado (BCB)</div>
          {rate&&<div style={{fontSize:11,color:C.textLight}}>atualizado às {lastFetch}</div>}
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          {rate&&<span style={{fontSize:14,fontWeight:800,color:C.success,fontFamily:"'DM Mono',monospace"}}>R$ {ptBR4(rate)}</span>}
          <button onClick={fetch_} style={{background:C.primaryLight,border:`1px solid ${C.primary}33`,borderRadius:8,padding:"5px 10px",fontSize:12,fontWeight:600,color:C.primary,cursor:"pointer"}}>{loading?"...":"↻ BCB"}</button>
        </div>
      </div>
      {rate&&<div style={{fontSize:12,color:C.textLight,marginTop:4}}>Use este valor como referência de mercado. O dólar pago (configurações) é o que realmente pagou.</div>}
    </div>
  );
}

// ─── CALC TAB ─────────────────────────────────────────────────────────────────
function CalcTab({settings}) {
  const [usd,setUsd]=useState(""); const [dc,setDc]=useState(""); const [lbs,setLbs]=useState(""); const [oz,setOz]=useState("");
  const dolarAj=calcDolarAjustado(settings); const usdN=parseFloat(usd)||0;
  const usdF=calcUsdFinal(usdN,settings); const brlP=usdF*dolarAj; const brlC=dc?usdF*parseFloat(dc):null;
  return (
    <div style={S.page}>
      <div style={S.sectionLabel}>Conversor USD → BRL</div>
      <div style={S.card}>
        <label style={S.label}>Valor em USD</label>
        <input style={S.input} type="number" placeholder="Ex: 199" value={usd} onChange={e=>setUsd(e.target.value)}/>
        {usdN>0&&[{label:"USD c/ taxa",value:fUSD(usdF),color:C.textMid},{label:"Dólar ajustado",value:`R$ ${ptBR4(dolarAj)}`,color:C.textMid},{label:"Valor em BRL",value:fBRL(brlP),color:C.primary}].map(({label,value,color})=>(
          <div key={label} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${C.borderLight}`}}>
            <span style={{fontSize:13,color:C.textMid}}>{label}</span>
            <span style={{fontSize:13,fontWeight:700,color,fontFamily:"'DM Mono',monospace"}}>{value}</span>
          </div>
        ))}
        <div style={{height:14}}/>
        <label style={S.label}>Dólar que você pagou (opcional)</label>
        <input style={S.input} type="number" step="0.01" placeholder="Ex: 5.71" value={dc} onChange={e=>setDc(e.target.value)}/>
        {brlC&&usdN>0&&[{label:"Com seu dólar",value:fBRL(brlC),color:C.success},{label:"Diferença",value:fBRL(Math.abs(brlP-brlC)),color:brlP>brlC?C.success:C.danger}].map(({label,value,color})=>(
          <div key={label} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${C.borderLight}`}}>
            <span style={{fontSize:13,color:C.textMid}}>{label}</span>
            <span style={{fontSize:13,fontWeight:700,color,fontFamily:"'DM Mono',monospace"}}>{value}</span>
          </div>
        ))}
      </div>
      <div style={S.sectionLabel}>Conversor de Peso</div>
      <div style={S.card}>
        <label style={S.label}>Libras (lbs)</label>
        <input style={S.input} type="number" placeholder="Ex: 2.5" value={lbs} onChange={e=>setLbs(e.target.value)}/>
        {lbs&&[[`Gramas`,`${(parseFloat(lbs)*453.592).toLocaleString("pt-BR",{minimumFractionDigits:1,maximumFractionDigits:1})} g`],[`Kg`,fKg(parseFloat(lbs)*453.592)],["Oz",`${(parseFloat(lbs)*16).toLocaleString("pt-BR",{minimumFractionDigits:1,maximumFractionDigits:1})} oz`]].map(([l,v])=>(
          <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:`1px solid ${C.borderLight}`}}><span style={{fontSize:13,color:C.textMid}}>{l}</span><span style={{fontSize:13,fontWeight:700,color:C.text,fontFamily:"'DM Mono',monospace"}}>{v}</span></div>
        ))}
        <div style={{height:12}}/>
        <label style={S.label}>Onças (oz)</label>
        <input style={S.input} type="number" placeholder="Ex: 3.4" value={oz} onChange={e=>setOz(e.target.value)}/>
        {oz&&[["Gramas",`${(parseFloat(oz)*28.3495).toLocaleString("pt-BR",{minimumFractionDigits:1,maximumFractionDigits:1})} g`],["Kg",fKg(parseFloat(oz)*28.3495)],["Libras",`${(parseFloat(oz)/16).toLocaleString("pt-BR",{minimumFractionDigits:3,maximumFractionDigits:3})} lbs`]].map(([l,v])=>(
          <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:`1px solid ${C.borderLight}`}}><span style={{fontSize:13,color:C.textMid}}>{l}</span><span style={{fontSize:13,fontWeight:700,color:C.text,fontFamily:"'DM Mono',monospace"}}>{v}</span></div>
        ))}
      </div>
      <div style={S.card}>
        <div style={{fontWeight:700,fontSize:13,color:C.text,marginBottom:10}}>Taxas e cotações</div>
        {[["Dólar pago",`R$ ${ptBR4(settings.dollarPago)}`],["IOF",`${ptBR2(settings.iof)}%`],["Spread",`${ptBR2(settings.spread)}%`],["Taxa compra",`${ptBR2(settings.taxa)}%`],["Dólar ajustado",`R$ ${ptBR4(dolarAj)}`]].map(([l,v])=>(
          <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:`1px solid ${C.borderLight}`}}><span style={{fontSize:13,color:C.textMid}}>{l}</span><span style={{fontSize:13,fontWeight:700,color:C.primary,fontFamily:"'DM Mono',monospace"}}>{v}</span></div>
        ))}
        <BcbRate/>
      </div>
    </div>
  );
}

// ─── SETTINGS MODAL ───────────────────────────────────────────────────────────
function SettingsModal({settings,onSave,onImport,onClose}) {
  const [s,setS]=useState({...settings});
  const [tab,setTab]=useState("config");
  const [preview,setPreview]=useState(null);
  const [loading,setLoading]=useState(false);
  const [error,setError]=useState(null);
  const fileRef=useRef();
  function processFile(file) {
    if(!file)return; setLoading(true); setError(null); setPreview(null);
    const reader=new FileReader();
    reader.onload=e=>{
      try { const data=new Uint8Array(e.target.result); const wb=XLSX.read(data,{type:"array"}); const sheets=wb.SheetNames;
        const comprasSheet=sheets.find(s=>s.toLowerCase().includes("compras")&&!s.toLowerCase().includes("parcelas"));
        const legaisSheet=sheets.find(s=>s.toLowerCase().includes("legais")||s.toLowerCase().includes("legal"));
        const compras=comprasSheet?parseSheet(wb,comprasSheet):[];
        const legais=legaisSheet?parseSheet(wb,legaisSheet):[];
        setPreview({compras,legais,comprasSheet,legaisSheet}); setLoading(false);
      } catch(err){setError("Erro: "+err.message); setLoading(false);}
    };
    reader.readAsArrayBuffer(file);
  }
  return (
    <Modal title="Configurações" onClose={onClose}>
      <div style={{display:"flex",gap:4,background:C.borderLight,borderRadius:10,padding:4,marginBottom:20}}>
        {[["config","⚙ Config"],["import","📂 Excel"]].map(([v,l])=>(
          <button key={v} style={{flex:1,padding:"8px",borderRadius:8,border:"none",cursor:"pointer",fontSize:13,fontWeight:600,background:tab===v?C.bgCard:"transparent",color:tab===v?C.primary:C.textMid,boxShadow:tab===v?"0 1px 4px rgba(0,0,0,0.08)":"none"}} onClick={()=>setTab(v)}>{l}</button>
        ))}
      </div>
      {tab==="config"&&(
        <>
          <div style={{fontSize:11,fontWeight:700,color:C.textLight,textTransform:"uppercase",letterSpacing:"0.6px",marginBottom:10}}>💵 Dólar pago</div>
          <label style={S.label}>Quanto você pagou pelo dólar (R$)</label>
          <input style={S.input} type="number" step="0.0001" placeholder="Ex: 5.6200" value={s.dollarPago} onChange={e=>setS(p=>({...p,dollarPago:parseFloat(e.target.value)||0}))}/>
          <div style={{fontSize:11,fontWeight:700,color:C.textLight,textTransform:"uppercase",letterSpacing:"0.6px",marginBottom:10,marginTop:4}}>📊 Taxas (para cálculo do custo real)</div>
          {[["IOF (%)","iof","0.01"],["Spread (%)","spread","0.01"],["Taxa de compra (%)","taxa","0.1"]].map(([l,f,st])=>(
            <div key={f} style={{marginBottom:12}}>
              <label style={S.label}>{l}</label>
              <input style={S.input} type="number" step={st} value={s[f]} onChange={e=>setS(p=>({...p,[f]:parseFloat(e.target.value)||0}))}/>
            </div>
          ))}
          <div style={{fontSize:11,fontWeight:700,color:C.textLight,textTransform:"uppercase",letterSpacing:"0.6px",marginBottom:10}}>✈ Viagem</div>
          <div style={{marginBottom:12}}>
            <label style={S.label}>Total de dólares que está levando (US$)</label>
            <input style={S.input} type="number" step="100" value={s.totalDolarViagem} onChange={e=>setS(p=>({...p,totalDolarViagem:parseFloat(e.target.value)||0}))}/>
          </div>
          <div style={{marginBottom:12}}>
            <label style={S.label}>Peso máximo da mala (kg)</label>
            <input style={S.input} type="number" step="0.5" placeholder="23" value={(s.pesoMax/1000).toLocaleString("pt-BR",{maximumFractionDigits:1})} onChange={e=>setS(p=>({...p,pesoMax:Math.round((parseFloat(e.target.value.replace(",","."))||0)*1000)}))}/>
          </div>
          <div style={{padding:"10px 14px",background:C.primaryLight,borderRadius:10,fontSize:13,color:C.textMid,marginBottom:8}}>
            Dólar ajustado: <strong style={{color:C.primary}}>R$ {ptBR4(calcDolarAjustado(s))}</strong>
          </div>
          <button style={S.btnPrimary} onClick={()=>onSave(s)}>Salvar configurações</button>
        </>
      )}
      {tab==="import"&&(
        <>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e=>processFile(e.target.files[0])}/>
          <div style={{border:`2px dashed ${loading?C.primary:C.border}`,borderRadius:16,textAlign:"center",padding:"28px 20px",cursor:"pointer",background:loading?C.primaryLight:C.bg,transition:"all 0.2s"}} onClick={()=>fileRef.current.click()} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();processFile(e.dataTransfer.files[0]);}}>
            <div style={{fontSize:36,marginBottom:10}}>{loading?"⏳":"📊"}</div>
            <div style={{fontWeight:700,fontSize:14,color:C.text,marginBottom:4}}>{loading?"Processando...":"Selecionar planilha"}</div>
            <div style={{fontSize:13,color:C.textLight}}>Arraste ou toque · .xlsx</div>
          </div>
          {error&&<div style={{marginTop:10,background:C.dangerLight,border:`1px solid ${C.danger}33`,borderRadius:10,padding:"10px",fontSize:13,color:C.danger}}>{error}</div>}
          {preview&&(
            <div style={{marginTop:14}}>
              <div style={{fontWeight:700,fontSize:13,color:C.text,marginBottom:10}}>Prévia</div>
              {[["Aba Compras",preview.comprasSheet||"—"],["Aba Legais",preview.legaisSheet||"—"],["Produtos",`${preview.compras.length} itens`],["Legais",`${preview.legais.length} itens`]].map(([l,v])=>(
                <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${C.borderLight}`}}>
                  <span style={{fontSize:13,color:C.textMid}}>{l}</span><span style={{fontSize:13,fontWeight:600,color:C.text}}>{v}</span>
                </div>
              ))}
              {preview.compras.slice(0,3).map((p,i)=>(
                <div key={i} style={{fontSize:12,color:C.textMid,padding:"4px 0",borderBottom:`1px solid ${C.borderLight}`}}><span style={{fontWeight:600,color:C.text}}>{p.nome}</span> · {p.loja} · {fUSD(p.usd)}</div>
              ))}
              <button style={{...S.btnPrimary,marginTop:12}} onClick={()=>{onImport(preview.compras,preview.legais);setPreview(null);}}>✅ Confirmar importação</button>
            </div>
          )}
        </>
      )}
    </Modal>
  );
}

// ─── PRODUTO FORM ─────────────────────────────────────────────────────────────
function ProdutoForm({prod,isLegais,onSave,onClose}) {
  const isL = isLegais || prod?._legais === true;
  const empty={nome:"",loja:"Walmart",usd:"",peso:"",tipo:"solido",volume:"",status:"pendente",prioridade:"Média",link:"",imagem:"",dollarPago:"",_legais:isL};
  const [f,setF]=useState(prod?.id?{...prod,usd:prod.usd.toString(),peso:(prod.peso||"").toString(),dollarPago:(prod.dollarPago||"").toString(),_legais:prod._legais||isL}:empty);
  function save(){if(!f.nome||!f.usd)return alert("Preencha nome e USD");onSave({...f,usd:parseFloat(f.usd),peso:parseFloat(f.peso)||0,volume:parseFloat(f.volume)||0,dollarPago:f.dollarPago?parseFloat(f.dollarPago):null,_legais:isL});}
  return (
    <Modal title={prod?.id?"Editar produto":isL?"✨ Item interessante":"Novo produto"} onClose={onClose}>
      <label style={S.label}>Nome *</label>
      <input style={S.input} placeholder="Ex: AirPods Pro" value={f.nome} onChange={e=>setF(p=>({...p,nome:e.target.value}))}/>
      <label style={S.label}>Loja</label>
      <input style={S.input} list="lojas-list" placeholder="Digite ou escolha uma loja..." value={f.loja} onChange={e=>setF(p=>({...p,loja:e.target.value}))}/>
      <datalist id="lojas-list">{LOJAS_SUGESTOES.map(l=><option key={l} value={l}/>)}</datalist>
      <label style={S.label}>Preço USD *</label>
      <input style={S.input} type="number" placeholder="Ex: 199" value={f.usd} onChange={e=>setF(p=>({...p,usd:e.target.value}))}/>
      <label style={S.label}>Tipo</label>
      <div style={{display:"flex",gap:8,marginBottom:14}}>{[["solido","📦 Sólido"],["liquido","💧 Líquido"]].map(([v,l])=><button key={v} style={{...S.chipSel,flex:1,...(f.tipo===v?S.chipSelActive:{})}} onClick={()=>setF(p=>({...p,tipo:v}))}>{l}</button>)}</div>
      {f.tipo==="solido"
        ?<><label style={S.label}>Peso (gramas)</label><input style={S.input} type="number" placeholder="Ex: 250" value={f.peso} onChange={e=>setF(p=>({...p,peso:e.target.value}))}/>{f.peso&&<div style={{fontSize:12,color:C.textLight,marginTop:-8,marginBottom:12}}>= {fKg(parseFloat(f.peso)||0)}</div>}</>
        :<><label style={S.label}>Volume (oz)</label><input style={S.input} type="number" step="0.1" placeholder="Ex: 3.4" value={f.volume} onChange={e=>setF(p=>({...p,volume:e.target.value}))}/>{f.volume&&<div style={{fontSize:12,color:C.textLight,marginTop:-8,marginBottom:12}}>= {fKg((parseFloat(f.volume)||0)*28.3495)}</div>}</>
      }
      {!isL&&(<>
        <label style={S.label}>Prioridade</label>
        <div style={{display:"flex",gap:8,marginBottom:14}}>{PRIORIDADES.map(pr=><button key={pr} style={{...S.chipSel,flex:1,...(f.prioridade===pr?S.chipSelActive:{})}} onClick={()=>setF(p=>({...p,prioridade:pr}))}>{pr}</button>)}</div>
        <label style={S.label}>Status</label>
        <div style={{display:"flex",gap:8,marginBottom:14}}>{[["pendente","⏳ Pendente"],["comprado","✅ Comprado"]].map(([v,l])=><button key={v} style={{...S.chipSel,flex:1,...(f.status===v?S.chipSelActive:{})}} onClick={()=>setF(p=>({...p,status:v}))}>{l}</button>)}</div>
        {f.status==="comprado"&&<><label style={S.label}>Dólar pago (R$)</label><input style={S.input} type="number" step="0.01" placeholder="5.71" value={f.dollarPago} onChange={e=>setF(p=>({...p,dollarPago:e.target.value}))}/></>}
      </>)}
      <label style={S.label}>Link (para buscar imagem)</label>
      <input style={S.input} type="url" placeholder="https://amazon.com/..." value={f.link} onChange={e=>setF(p=>({...p,link:e.target.value}))}/>
      <label style={S.label}>URL da imagem (opcional)</label>
      <input style={S.input} type="url" placeholder="https://..." value={f.imagem} onChange={e=>setF(p=>({...p,imagem:e.target.value}))}/>
      <button style={S.btnPrimary} onClick={save}>{prod?.id?"Salvar":isL?"Adicionar item":"Adicionar produto"}</button>
    </Modal>
  );
}

// ─── SHARED ───────────────────────────────────────────────────────────────────
function Modal({title,onClose,children}) {
  return (
    <div style={S.modalOverlay} onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div style={S.modal}>
        <div style={S.modalHeader}><div style={S.modalTitle}>{title}</div><button style={S.closeBtn} onClick={onClose}>✕</button></div>
        <div style={S.modalBody}>{children}</div>
      </div>
    </div>
  );
}
function Empty({text}) {
  return <div style={{textAlign:"center",padding:"40px 0"}}><div style={{fontSize:36,marginBottom:10}}>📭</div><div style={{fontSize:13,color:C.textLight,fontWeight:500}}>{text}</div></div>;
}

// ─── STYLES ───────────────────────────────────────────────────────────────────
const S={
  app:{background:C.bg,minHeight:"100vh",maxWidth:430,margin:"0 auto",position:"relative",fontFamily:"'Inter',sans-serif",color:C.text,display:"flex",flexDirection:"column"},
  header:{background:C.bgCard,borderBottom:`1px solid ${C.border}`,padding:"12px 16px",display:"flex",justifyContent:"space-between",alignItems:"center",position:"sticky",top:0,zIndex:50,boxShadow:"0 1px 3px rgba(0,0,0,0.04)"},
  headerLeft:{display:"flex",gap:10,alignItems:"center"},
  logoBox:{width:36,height:36,background:`linear-gradient(135deg,${C.gradientA},${C.gradientB})`,borderRadius:10,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,color:"white"},
  headerTitle:{fontSize:15,fontWeight:700,color:C.text,letterSpacing:"-0.3px"},
  headerSub:{fontSize:11,color:C.textLight,fontWeight:500},
  settingsBtn:{background:C.bg,border:`1px solid ${C.border}`,width:36,height:36,borderRadius:10,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"},
  content:{flex:1,overflowY:"auto",paddingBottom:90},
  page:{padding:"14px 14px 8px"},
  nav:{position:"fixed",bottom:0,left:"50%",transform:"translateX(-50%)",width:"100%",maxWidth:430,background:C.bgCard,borderTop:`1px solid ${C.border}`,display:"flex",zIndex:50,padding:"6px 0 14px",boxShadow:"0 -2px 10px rgba(0,0,0,0.06)"},
  navBtn:{flex:1,background:"none",border:"none",color:C.textLight,cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:2,padding:"4px 1px",position:"relative"},
  navBtnActive:{color:C.primary},
  navLabel:{fontSize:9,fontWeight:600,letterSpacing:"0.2px"},
  navIndicator:{position:"absolute",bottom:-6,left:"50%",transform:"translateX(-50%)",width:4,height:4,borderRadius:"50%",background:C.primary},
  fab:{position:"fixed",bottom:80,right:"calc(50% - 215px + 14px)",background:`linear-gradient(135deg,${C.gradientA},${C.gradientB})`,border:"none",width:52,height:52,borderRadius:16,cursor:"pointer",boxShadow:`0 6px 20px ${C.primary}55`,display:"flex",alignItems:"center",justifyContent:"center",zIndex:60},
  heroCard:{background:`linear-gradient(135deg,${C.gradientA} 0%,${C.gradientB} 100%)`,borderRadius:18,padding:"20px 18px",marginBottom:12,boxShadow:`0 8px 24px ${C.primary}33`},
  card:{background:C.bgCard,border:`1px solid ${C.border}`,borderRadius:14,padding:"14px",marginBottom:10,boxShadow:"0 1px 3px rgba(0,0,0,0.04)"},
  grid4:{display:"flex",gap:8,marginBottom:10},
  checkbox:{width:22,height:22,borderRadius:7,border:`2px solid ${C.border}`,background:C.bg,cursor:"pointer",flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center"},
  checkboxDone:{background:C.success,borderColor:C.success},
  searchInput:{width:"100%",background:C.bgCard,border:`1px solid ${C.border}`,borderRadius:10,padding:"10px 12px",color:C.text,fontSize:14,marginBottom:8,boxSizing:"border-box",outline:"none",fontFamily:"'Inter',sans-serif"},
  filterRow:{display:"flex",gap:6,overflowX:"auto",paddingBottom:4,marginBottom:8},
  chip:{background:C.bgCard,border:`1px solid ${C.border}`,color:C.textMid,fontSize:12,padding:"4px 10px",borderRadius:999,cursor:"pointer",whiteSpace:"nowrap",fontWeight:600},
  chipActive:{background:C.primaryLight,border:`1px solid ${C.primary}44`,color:C.primary},
  chipSel:{background:C.bg,border:`1px solid ${C.border}`,color:C.textMid,fontSize:13,padding:"9px 8px",borderRadius:10,cursor:"pointer",fontWeight:600,fontFamily:"'Inter',sans-serif"},
  chipSelActive:{background:C.primaryLight,border:`1px solid ${C.primary}66`,color:C.primary},
  tag:{background:C.bg,border:`1px solid ${C.border}`,color:C.textMid,fontSize:11,padding:"2px 7px",borderRadius:999,fontWeight:600},
  sectionLabel:{fontSize:11,fontWeight:700,color:C.textLight,textTransform:"uppercase",letterSpacing:"0.8px",marginBottom:8,marginTop:4},
  btnPrimary:{width:"100%",background:`linear-gradient(135deg,${C.gradientA},${C.gradientB})`,border:"none",color:"white",borderRadius:12,padding:"13px",fontSize:14,fontWeight:700,cursor:"pointer",fontFamily:"'Inter',sans-serif",boxShadow:`0 4px 14px ${C.primary}44`},
  btnOutline:{background:C.bg,border:`1px solid ${C.border}`,color:C.textMid,borderRadius:9,padding:"7px 12px",fontSize:12,cursor:"pointer",fontWeight:600,fontFamily:"'Inter',sans-serif"},
  modalOverlay:{position:"fixed",inset:0,background:"rgba(15,23,42,0.35)",zIndex:100,display:"flex",alignItems:"flex-end",justifyContent:"center",backdropFilter:"blur(4px)"},
  modal:{background:C.bgCard,borderRadius:"20px 20px 0 0",width:"100%",maxWidth:430,maxHeight:"92vh",boxShadow:"0 -4px 32px rgba(0,0,0,0.12)"},
  modalHeader:{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"18px 18px 14px",borderBottom:`1px solid ${C.border}`},
  modalTitle:{fontSize:16,fontWeight:700,color:C.text},
  closeBtn:{background:C.bg,border:`1px solid ${C.border}`,color:C.textMid,width:28,height:28,borderRadius:7,cursor:"pointer",fontSize:13,fontFamily:"'Inter',sans-serif"},
  modalBody:{padding:"16px 18px 34px",overflowY:"auto",maxHeight:"calc(92vh - 60px)"},
  label:{fontSize:12,color:C.textMid,fontWeight:600,display:"block",marginBottom:5},
  input:{width:"100%",background:C.bg,border:`1px solid ${C.border}`,borderRadius:9,padding:"10px 11px",color:C.text,fontSize:14,marginBottom:12,boxSizing:"border-box",outline:"none",fontFamily:"'Inter',sans-serif"},
};

const CSS=`
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=DM+Mono:wght@400;500&display=swap');
  *{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent;}
  body{background:#F8FAFC;}
  select option{background:#fff;}
  ::-webkit-scrollbar{width:3px;height:3px;}
  ::-webkit-scrollbar-thumb{background:#CBD5E1;border-radius:2px;}
  .notif{position:fixed;top:68px;left:50%;transform:translateX(-50%);z-index:200;padding:9px 18px;border-radius:999px;font-size:13px;font-weight:600;animation:slideDown 0.3s ease;white-space:nowrap;box-shadow:0 4px 14px rgba(0,0,0,0.1);}
  .notif-success{background:#ECFDF5;border:1px solid #10B98133;color:#059669;}
  .notif-error{background:#FEF2F2;border:1px solid #EF444433;color:#DC2626;}
  @keyframes slideDown{from{opacity:0;transform:translateX(-50%) translateY(-8px);}to{opacity:1;transform:translateX(-50%) translateY(0);}}
  .galeria-card:active{transform:scale(0.98);}
  .spinner{width:24px;height:24px;border:3px solid #E5E7EB;border-top-color:#2563EB;border-radius:50%;animation:spin 0.7s linear infinite;display:inline-block;}
  @keyframes spin{to{transform:rotate(360deg);}}
`;
