import { useState, useMemo, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { getProductImage, imageCache } from './imageService';
import { db } from "./firebase";
import { doc, onSnapshot, setDoc, serverTimestamp, enableNetwork, disableNetwork } from "firebase/firestore";
import { getAuth, onAuthStateChanged, signInWithEmailAndPassword, createUserWithEmailAndPassword, signOut } from "firebase/auth";


// ─── FIREBASE AUTH / FIRESTORE PERSISTENCE ────────────────────────────────
// Cada usuário logado tem sua própria "caixinha" no Firestore.
// Os dados ficam em: usuarios_pwa/{uid}
const auth = getAuth();

const normalizeCloudState = data => ({
  settings: data?.settings || INITIAL_SETTINGS,
  produtos: Array.isArray(data?.produtos) ? data.produtos : SAMPLE_PRODUTOS,
  itensLegais: Array.isArray(data?.itensLegais) ? data.itensLegais : [],
  gastos: Array.isArray(data?.gastos) ? data.gastos : [],
  parcelas: Array.isArray(data?.parcelas) ? data.parcelas : [],
});

async function saveCloudState(userDocRef, state) {
  await setDoc(userDocRef, {
    ...state,
    updatedAt: serverTimestamp(),
  }, { merge: true });
}

// ─── OFFLINE / LOCALSTORAGE ──────────────────────────────────────────────────
const LS_KEY = "travelshop_offline_v1";
const LS_PENDING = "travelshop_pending_sync";

function lsSave(uid, state) {
  try { localStorage.setItem(`${LS_KEY}_${uid}`, JSON.stringify({ ...state, _savedAt: Date.now() })); } catch {}
}

function lsLoad(uid) {
  try {
    const raw = localStorage.getItem(`${LS_KEY}_${uid}`);
    return raw ? JSON.parse(raw) : null;
  } catch { return null; }
}

function lsSetPending(uid, hasPending) {
  try {
    if (hasPending) localStorage.setItem(`${LS_PENDING}_${uid}`, "1");
    else localStorage.removeItem(`${LS_PENDING}_${uid}`);
  } catch {}
}

function lsHasPending(uid) {
  try { return !!localStorage.getItem(`${LS_PENDING}_${uid}`); } catch { return false; }
}

// ─── COLORS ──────────────────────────────────────────────────────────────────
const C = {
  bg:"#F8FAFC",bgCard:"#FFFFFF",border:"#E5E7EB",borderLight:"#F1F5F9",
  primary:"#2563EB",primaryLight:"#EFF6FF",gradientA:"#2563EB",gradientB:"#06B6D4",
  success:"#10B981",successLight:"#ECFDF5",warning:"#F59E0B",warningLight:"#FFFBEB",
  danger:"#EF4444",dangerLight:"#FEF2F2",purple:"#8B5CF6",purpleLight:"#F5F3FF",
  text:"#0F172A",textMid:"#475569",textLight:"#94A3B8",textXLight:"#CBD5E1",
};

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
const INITIAL_SETTINGS = { dollarPago:5.62, iof:3.38, spread:0.99, taxa:6.5, pesoMax:23000, totalDolarViagem:3000 };
const LOJAS_SUGESTOES = [
  // --- Lista Original ---
  "Walmart", "Basspro", "Target", "HomeGoods", "Dollar Tree", "Amazon", 
  "Marshalls", "Ross", "TJ Maxx", "Tommy Hilfiger", "Calvin Klein", 
  "The North Face", "Sephora", "Ulta", "Best Buy", "Costco", "GameStop", 
  "Apple", "Restaurante", "Uber", "Passeio", "Outro",

  // --- Novas Sugestões Adicionadas ---
  // Lojas de Departamento, Variedades e Descontos
  "Macy's", "Nordstrom", "Burlington", "Five Below", "Sam's Club",
  
  // Vestuário, Moda e Esportes
  "Nike", "Adidas", "Under Armour", "Zara", "H&M", "Levi's", "Gap", "Cabela's",
  
  // Beleza e Cuidados Pessoais
  "Bath & Body Works", "CVS", "Walgreens",
  
  // Eletrônicos e Entretenimento
  "Micro Center", "Barnes & Noble",
  
  // Casa, Decoração e Ferramentas
  "Home Depot", "Lowe's", "Bed Bath & Beyond", "IKEA",
  
  // Serviços e Transporte (Para acompanhar Uber/Passeio)
  "Lyft", "Gas Station", "Supermarket", "Hotel"
];const PRIORIDADES = ["Alta","Média","Baixa"];
const CATEGORIAS_GASTO = ["🛍 Compras","🍔 Alimentação","🚗 Transporte","🎢 Passeio","🏨 Hospedagem","💊 Farmácia","🎁 Presente","💳 Outros"];

// ─── CALC ─────────────────────────────────────────────────────────────────────
const calcDolarAjustado = s => s.dollarPago * (1 + (s.iof + s.spread) / 100);
const calcUsdFinal = (usd, s) => usd * (1 + s.taxa / 100);
const calcBRL = (usd, s) => calcUsdFinal(usd, s) * calcDolarAjustado(s);
const calcBRLPago = (usd, s, dp) => calcUsdFinal(usd, s) * dp;
const pesoGramas = p => p.tipo === "liquido" ? (parseFloat(p.volume)||0)*28.3495 : parseFloat(p.peso)||0;

// ─── FORMATAÇÃO ─────────────────────────────────────────────────────────────
// Formata número com vírgula como separador decimal (padrão pt-BR)
const fmtUSD = (v, dec=2) => `US$ ${Number(v).toLocaleString("pt-BR",{minimumFractionDigits:dec,maximumFractionDigits:dec})}`;
const fmtBRL = (v, dec=2) => `R$ ${Number(v).toLocaleString("pt-BR",{minimumFractionDigits:dec,maximumFractionDigits:dec})}`;
const fmtN   = (v, dec=2) => Number(v).toLocaleString("pt-BR",{minimumFractionDigits:dec,maximumFractionDigits:dec});

// tudo em USD — dolarPago = cotação usada na compra
function calcMinhaParteUSD(gasto) {
  const totalUSD = parseFloat(gasto.usd) || 0;
  if (!gasto.divisao || gasto.divisao.length === 0) return totalUSD;
  const somaDivisao = gasto.divisao.reduce((acc, p) => acc + (parseFloat(p.valor) || 0), 0);
  return Math.max(0, totalUSD - somaDivisao);
}
function usdToBRL(usd, gasto, settings) {
  const cotacao = parseFloat(gasto.dolarPago) || settings.dollarPago;
  return usd * cotacao;
}
function calcTotalGastosUSD(gastos) {
  return gastos.reduce((a, g) => a + calcMinhaParteUSD(g), 0);
}

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

    if (imageCache[key]) {
      setImgUrl(imageCache[key]);
      setLoading(false);
      return;
    }

    setLoading(true);
    setErr(false);

    getProductImage(produto).then(url => {
      if (cancelled) return;
      if (url) {
        setImgUrl(url);
      } else {
        setErr(true);
      }
      setLoading(false);
    }).catch(() => {
      if (cancelled) return;
      setErr(true);
      setLoading(false);
    });

    return () => {
      cancelled = true;
    };
  }, [produto.id, produto.link, produto.imagem]);

  if (!imgUrl || err) {
    return (
      <div style={{
        width: "100%",
        height: "100%",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        background: `linear-gradient(135deg, ${C.primaryLight}, ${C.purpleLight})`,
        ...style
      }}>
        {loading ? (
          <div className="spinner" />
        ) : (
          <span style={{ fontSize: iconSize }}>
            {LOJA_EMOJI[produto.loja] || "🛍"}
          </span>
        )}
      </div>
    );
  }

  return (
    <div style={{ width: "100%", height: "100%", ...style }}>
      <img
        src={imgUrl}
        alt={produto.nome}
        style={{ width: "100%", height: "100%", objectFit: "cover" }}
        onError={() => setErr(true)}
      />
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

// ─── PARCELAS PARSER ─────────────────────────────────────────────────────────
function parseParcelasSheet(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  if (!ws) return [];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  // Encontrar linha de cabeçalho: deve ter "Descrição" e "Parcelas"
  const hi = rows.findIndex(r =>
    r && r.some(c => c && c.toString().toLowerCase().includes("descrição"))
       && r.some(c => c && c.toString().toLowerCase().includes("parcela"))
  );
  if (hi < 0) return [];

  const hd = rows[hi];
  const idx = name => hd.findIndex(h => h && h.toString().toLowerCase().includes(name.toLowerCase()));

  const iDesc   = idx("descrição");
  const iQtd    = idx("parcelas");
  const iValP   = idx("valor / parcela");   // "Valor / Parcela (R$)"
  const iValT   = idx("total da compra");   // "Total da Compra (R$)"
  const iFatura = idx("primeira fatura");   // "Primeira Fatura" — usado como cartão/referência

  const result = [];
  for (let i = hi + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r) continue;
    const desc = r[iDesc];
    // Parar ao encontrar linha vazia ou seção nova (ex: "DISTRIBUIÇÃO")
    if (!desc || typeof desc !== "string" || desc.toString().trim() === "") break;
    if (desc.toString().toUpperCase().includes("DISTRIBUIÇÃO")) break;

    const qt = parseInt(r[iQtd]) || 0;
    const valorParcela = parseFloat(
      r[iValP] !== null && r[iValP] !== undefined
        ? r[iValP].toString().replace(/[R$\s.]/g, "").replace(",", ".")
        : 0
    ) || 0;
    const valorTotal = parseFloat(
      r[iValT] !== null && r[iValT] !== undefined
        ? r[iValT].toString().replace(/[R$\s.]/g, "").replace(",", ".")
        : 0
    ) || 0;
    const cartaoRaw = r[iFatura] ? r[iFatura].toString().trim() : "";
    // "Primeira Fatura" na planilha é o mês inicial (ex: "abr/26") — salvar como primeiraFatura
    // Normalizar para formato "mmm/aa" em minúsculas
    let primeiraFatura = "";
    const pfMatch = cartaoRaw.match(/([a-záàãâéêíóôõúç]{3})[\/\-](\d{2,4})/i);
    if (pfMatch) {
      const mn = pfMatch[1].toLowerCase().substring(0,3);
      const an = pfMatch[2].slice(-2);
      primeiraFatura = `${mn}/${an}`;
    }
    const cartao = primeiraFatura; // reusa como cartão por enquanto

    if (qt <= 0 || !desc.trim()) continue;

    result.push({
      id: Date.now() + i + Math.random(),
      descricao: desc.toString().trim(),
      valorTotal: valorTotal || parseFloat((valorParcela * qt).toFixed(2)),
      quantidadeParcelas: qt,
      valorParcela,
      cartao,
      primeiraFatura,
      statusMensal: Array(qt).fill(false),
    });
  }
  return result;
}

// ─── SAMPLE DATA ─────────────────────────────────────────────────────────────
const SAMPLE_PRODUTOS = [
  {id:1,nome:"AirPods Pro 2",loja:"Apple",usd:249,peso:61,tipo:"solido",volume:0,status:"pendente",prioridade:"Alta",link:"https://www.amazon.com/Apple-AirPods-Pro-Cancellation-Transparency/dp/B0BDHWDR12",imagem:"",dollarPago:null},
  {id:2,nome:"Tide Pods 31ct",loja:"Walmart",usd:9.99,peso:640,tipo:"solido",volume:0,status:"comprado",prioridade:"Baixa",link:"https://www.walmart.com/ip/Tide-PODS-Laundry-Detergent-Packs-Original-Scent-31-Count/42351728",imagem:"",dollarPago:5.68},
  {id:3,nome:"Keychron K2",loja:"Amazon",usd:119,peso:870,tipo:"solido",volume:0,status:"pendente",prioridade:"Média",link:"",imagem:"",dollarPago:null},
];

// Estado inicial temporário até o primeiro snapshot do Firestore chegar.
const _initSettings = INITIAL_SETTINGS;
const _initProdutos = SAMPLE_PRODUTOS;
const _initItensLegais = [];
const _initGastos = [];


// ─── LOGIN ───────────────────────────────────────────────────────────────────
function LoginScreen() {
  const [email, setEmail] = useState("");
  const [senha, setSenha] = useState("");
  const [loading, setLoading] = useState(false);
  const [erro, setErro] = useState("");

  async function entrar() {
    if (!email || !senha) { setErro("Informe e-mail e senha."); return; }
    setLoading(true); setErro("");
    try {
      await signInWithEmailAndPassword(auth, email.trim(), senha);
    } catch (e) {
      setErro(traduzErroAuth(e));
    } finally {
      setLoading(false);
    }
  }

  async function criarConta() {
    if (!email || !senha) { setErro("Informe e-mail e senha."); return; }
    if (senha.length < 6) { setErro("A senha precisa ter pelo menos 6 caracteres."); return; }
    setLoading(true); setErro("");
    try {
      await createUserWithEmailAndPassword(auth, email.trim(), senha);
    } catch (e) {
      setErro(traduzErroAuth(e));
    } finally {
      setLoading(false);
    }
  }

  return (
    <div style={{minHeight:"100vh",background:`linear-gradient(135deg, ${C.gradientA}, ${C.gradientB})`,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{width:"100%",maxWidth:390,background:C.bgCard,borderRadius:24,padding:26,boxShadow:"0 24px 70px rgba(15,23,42,0.25)",border:"1px solid rgba(255,255,255,0.35)"}}>
        <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:22}}>
          <div style={{width:48,height:48,borderRadius:16,background:`linear-gradient(135deg, ${C.gradientA}, ${C.gradientB})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:25,color:"white"}}>✈</div>
          <div>
            <div style={{fontSize:24,fontWeight:850,color:C.text,letterSpacing:"-0.7px"}}>TravelShop</div>
            <div style={{fontSize:13,color:C.textMid}}>Entre para sincronizar PC e celular</div>
          </div>
        </div>

        <label style={S.label}>E-mail</label>
        <input style={S.input} type="email" placeholder="seuemail@exemplo.com" value={email} onChange={e=>setEmail(e.target.value)} autoComplete="email" />

        <label style={S.label}>Senha</label>
        <input style={S.input} type="password" placeholder="mínimo 6 caracteres" value={senha} onChange={e=>setSenha(e.target.value)} autoComplete="current-password" onKeyDown={e=>{if(e.key==="Enter")entrar();}} />

        {erro&&<div style={{background:C.dangerLight,color:C.danger,border:`1px solid ${C.danger}33`,borderRadius:12,padding:"10px 12px",fontSize:13,fontWeight:600,marginBottom:12}}>{erro}</div>}

        <button style={{...S.btnPrimary,opacity:loading?0.7:1}} disabled={loading} onClick={entrar}>{loading?"Aguarde...":"Entrar"}</button>
        <button style={{...S.btnOutline,width:"100%",justifyContent:"center",marginTop:10,padding:"12px"}} disabled={loading} onClick={criarConta}>Criar Conta</button>

        <div style={{fontSize:12,color:C.textLight,textAlign:"center",marginTop:16,lineHeight:1.4}}>
          Seus produtos, gastos e configurações serão salvos no Firebase dentro do seu usuário.
        </div>
      </div>
    </div>
  );
}

function traduzErroAuth(e) {
  const code = e?.code || "";
  if (code.includes("invalid-email")) return "E-mail inválido.";
  if (code.includes("user-not-found") || code.includes("wrong-password") || code.includes("invalid-credential")) return "E-mail ou senha incorretos.";
  if (code.includes("email-already-in-use")) return "Esse e-mail já tem uma conta. Use Entrar.";
  if (code.includes("weak-password")) return "A senha precisa ter pelo menos 6 caracteres.";
  if (code.includes("network-request-failed")) return "Falha de internet. Verifique a conexão.";
  return "Não foi possível autenticar. Tente novamente.";
}

// ─── APP ─────────────────────────────────────────────────────────────────────
export default function App() {
  const [tab, setTab] = useState(0);
  const [settings, setSettings] = useState(_initSettings);
  const [produtos, setProdutos] = useState(_initProdutos);
  const [itensLegais, setItensLegais] = useState(_initItensLegais);
  const [gastos, setGastos] = useState(_initGastos); // gastos livres + produtos comprados espelhados
  const [parcelas, setParcelas] = useState([]);
  const [showSettings, setShowSettings] = useState(false);
  const [showForm, setShowForm] = useState(false);
  const [showGastoForm, setShowGastoForm] = useState(false);
  const [prodSubTab, setProdSubTab] = useState("compras");
  const [editProd, setEditProd] = useState(null);
  const [editGasto, setEditGasto] = useState(null);
  const [notification, setNotification] = useState(null);
  const [cotacaoLoading, setCotacaoLoading] = useState(false);
  const [user, setUser] = useState(null);
  const [authReady, setAuthReady] = useState(false);
  const [cloudReady, setCloudReady] = useState(false);
  const [isOnline, setIsOnline] = useState(navigator.onLine);
  const [pendingSync, setPendingSync] = useState(false);
  const skipNextCloudSave = useRef(false);
  const userDocRef = useMemo(() => user ? doc(db, "usuarios_pwa", user.uid) : null, [user]);

  // Detectar online/offline e sincronizar quando voltar
  useEffect(() => {
    function handleOnline() {
      setIsOnline(true);
      // Quando voltar online, reabilitar Firestore (ele mesmo sincroniza)
      enableNetwork(db).catch(() => {});
    }
    function handleOffline() {
      setIsOnline(false);
      // Desabilitar Firestore para evitar erros de timeout
      disableNetwork(db).catch(() => {});
    }
    window.addEventListener("online", handleOnline);
    window.addEventListener("offline", handleOffline);
    return () => {
      window.removeEventListener("online", handleOnline);
      window.removeEventListener("offline", handleOffline);
    };
  }, []);

  // Verifica login/logout pelo Firebase Authentication.
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (firebaseUser) => {
      setUser(firebaseUser);
      setAuthReady(true);
      setCloudReady(false);
    });
    return () => unsubscribe();
  }, []);

  // Escuta em tempo real os dados do usuário logado.
  useEffect(() => {
    if (!userDocRef || !user) return;

    const unsubscribe = onSnapshot(userDocRef, async (snap) => {
      if (!snap.exists()) {
        skipNextCloudSave.current = true;
        await saveCloudState(userDocRef, {
          settings: INITIAL_SETTINGS,
          produtos: SAMPLE_PRODUTOS,
          itensLegais: [],
          gastos: [],
          parcelas: [],
        });
        setSettings(INITIAL_SETTINGS);
        setProdutos(SAMPLE_PRODUTOS);
        setItensLegais([]);
        setGastos([]);
        setParcelas([]);
        setCloudReady(true);
        return;
      }

      const cloudState = normalizeCloudState(snap.data());
      skipNextCloudSave.current = true;
      setSettings(cloudState.settings);
      setProdutos(cloudState.produtos);
      setItensLegais(cloudState.itensLegais);
      setGastos(cloudState.gastos);
      setParcelas(cloudState.parcelas || []);
      setCloudReady(true);
    }, (error) => {
      console.error("Erro ao sincronizar com Firestore:", error);
      notify("Erro ao sincronizar com a nuvem", "error");
      setCloudReady(true);
    });

    return () => unsubscribe();
  }, [userDocRef, user]);

  // Salva no localStorage imediatamente + Firestore quando online
  useEffect(() => {
    if (!cloudReady || !user) return;
    if (skipNextCloudSave.current) {
      skipNextCloudSave.current = false;
      return;
    }
    const state = { settings, produtos, itensLegais, gastos, parcelas, planejamento, checklist, comprasDolar };
    // Sempre salvar no localStorage (funciona offline)
    lsSave(user.uid, state);
    const timer = setTimeout(() => {
      if (!navigator.onLine) {
        lsSetPending(user.uid, true);
        setPendingSync(true);
        return;
      }
      saveCloudState(userDocRef, state)
        .then(() => { lsSetPending(user.uid, false); setPendingSync(false); })
        .catch((error) => {
          console.error("Erro ao salvar no Firestore:", error);
          lsSetPending(user.uid, true);
          setPendingSync(true);
        });
    }, 350);
    return () => clearTimeout(timer);
  }, [settings, produtos, itensLegais, gastos, parcelas, planejamento, checklist, comprasDolar, cloudReady, user, userDocRef]);

  // Quando voltar online com dados pendentes, sincronizar
  useEffect(() => {
    if (!isOnline || !pendingSync || !userDocRef || !user || !cloudReady) return;
    const state = { settings, produtos, itensLegais, gastos, parcelas, planejamento, checklist, comprasDolar };
    saveCloudState(userDocRef, state)
      .then(() => { lsSetPending(user.uid, false); setPendingSync(false); notify("✅ Sincronizado com a nuvem!"); })
      .catch(err => console.error("Erro no sync:", err));
  }, [isOnline]);

  function notify(msg, type="success") { setNotification({msg,type}); setTimeout(()=>setNotification(null),2800); }

  async function handleLogout() {
    try {
      await signOut(auth);
      setUser(null);
      setSettings(INITIAL_SETTINGS);
      setProdutos(SAMPLE_PRODUTOS);
      setItensLegais([]);
      setGastos([]);
      setParcelas([]);
      setCloudReady(false);
    } catch (e) {
      console.error("Erro ao sair:", e);
      notify("Erro ao sair", "error");
    }
  }

  function toggleStatus(id) {
    setProdutos(ps => ps.map(p => {
      if (p.id !== id) return p;
      const newStatus = p.status === "comprado" ? "pendente" : "comprado";
      // sync gastos: add or remove produto entry
      if (newStatus === "comprado") {
        const brl = calcBRL(p.usd, settings);
        setGastos(gs => gs.some(g=>g.produtoId===id) ? gs : [...gs, {
          id: `prod_${id}`, produtoId:id, descricao:p.nome, loja:p.loja,
          usd:p.usd, dolarPago:p.dollarPago||settings.dollarPago,
          brl: null, imagem:p.imagem||"",
          categoria:"🛍 Compras", divisao:[], data: new Date().toLocaleDateString("pt-BR"), tipo:"produto"
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
    notify(prod.id?"Atualizado!":"Adicionado!"); setShowForm(false); setEditProd(null);
  }

  function moveToList(item) {
    setProdutos(ps=>[...ps,{...item,_legais:undefined,status:"pendente",prioridade:"Média",id:Date.now()}]);
    setItensLegais(ps=>ps.filter(p=>p.id!==item.id)); notify("Movido para lista!");
  }

  function handleImport(compras,legais,parcelasImp=[]) {
    if(compras.length) setProdutos(prev=>{const e=new Set(prev.map(p=>p.nome.toLowerCase()));return [...prev,...compras.filter(p=>!e.has(p.nome.toLowerCase()))];});
    if(legais.length) setItensLegais(prev=>{const e=new Set(prev.map(p=>p.nome.toLowerCase()));return [...prev,...legais.filter(p=>!e.has(p.nome.toLowerCase()))];});
    if(parcelasImp.length) setParcelas(prev=>{const e=new Set(prev.map(p=>p.descricao.toLowerCase()));return [...prev,...parcelasImp.filter(p=>!e.has(p.descricao.toLowerCase()))];});
    const parts=[compras.length?`${compras.length} compras`:"",legais.length?`${legais.length} legais`:"",parcelasImp.length?`${parcelasImp.length} parcelas`:""].filter(Boolean).join(" + ");
    notify(`✅ ${parts} importados!`);
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

  const TABS=[{label:"Início",icon:"⊞"},{label:"Produtos",icon:"📦"},{label:"Galeria",icon:"▦"},{label:"Gastos",icon:"💸"},{label:"Parcelas",icon:"💳"},{label:"Stats",icon:"◈"},{label:"Calc",icon:"⟨⟩"}];


  if (!authReady) {
    return (
      <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:C.bg,color:C.textMid,fontWeight:700}}>
        Carregando...
      </div>
    );
  }

  if (!user) return <LoginScreen />;

  return (
    <div style={S.app}>
      <style>{CSS}</style>
      {notification&&<div className={`notif notif-${notification.type}`}>{notification.msg}</div>}
      <div style={S.header}>
        {/* Banner offline/sync */}
        {(!isOnline || pendingSync) && (
          <div style={{
            position:"fixed",top:0,left:"50%",transform:"translateX(-50%)",
            width:"100%",maxWidth:430,zIndex:200,
            background: !isOnline ? "#1E293B" : C.warning,
            color: !isOnline ? "#94A3B8" : "#fff",
            fontSize:12,fontWeight:600,textAlign:"center",
            padding:"5px 16px",display:"flex",alignItems:"center",justifyContent:"center",gap:6
          }}>
            {!isOnline ? (
              <><span>📶</span> Sem internet — alterações salvas localmente</>
            ) : (
              <><span>⏳</span> Sincronizando com a nuvem...</>
            )}
          </div>
        )}
        <div style={{...S.headerLeft,marginTop:(!isOnline||pendingSync)?24:0}}>
          <div style={S.logoBox}>✈</div>
          <div>
            <div style={S.headerTitle}>TravelShop</div>
            <div style={S.headerSub}>Orlando 2027</div>
          </div>
        </div>
        <div style={{display:"flex",gap:8,alignItems:"center"}}>
          <button style={{...S.settingsBtn,width:"auto",padding:"0 10px",fontSize:12,fontWeight:700,color:C.textMid}} onClick={handleLogout}>Sair</button>
          <button style={S.settingsBtn} onClick={()=>setShowSettings(true)}>
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke={C.textMid} strokeWidth="2" strokeLinecap="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>
          </button>
        </div>
      </div>

      <div style={S.content}>
        {tab===0&&<DashboardTab stats={stats} settings={settings} pesoPercent={pesoPercent} pesoColor={pesoColor} pesoBg={pesoBg}/>}
        {tab===1&&<ProdutosTab produtos={produtos} itensLegais={itensLegais} settings={settings} onToggle={toggleStatus} onDelete={deleteProd} onEdit={p=>{setEditProd(p);setShowForm(true);}} onAdd={()=>{setEditProd(null);setShowForm(true);}} onMoveToList={moveToList} onSubTabChange={setProdSubTab}/>}
        {tab===2&&<GaleriaTab produtos={produtos} itensLegais={itensLegais} settings={settings} onEdit={p=>{setEditProd(p);setShowForm(true);}}/>}
        {tab===3&&<GastosTab gastos={gastos} settings={settings} onAdd={()=>{setEditGasto(null);setShowGastoForm(true);}} onEdit={g=>{setEditGasto(g);setShowGastoForm(true);}} onDelete={id=>{ setGastos(gs=>gs.filter(g=>g.id!==id)); notify("Removido","error"); }} onTogglePago={(gastoId,pessoaIdx)=>setGastos(gs=>gs.map(g=>g.id===gastoId?{...g,divisao:g.divisao.map((p,i)=>i===pessoaIdx?{...p,pago:!p.pago}:p)}:g))} produtos={produtos} onToggleStatus={toggleStatus}/>}
        {tab===4&&<ParcelasTab parcelas={parcelas} setParcelas={setParcelas}/>}
        {tab===5&&<StatsTab produtos={produtos} gastos={gastos} settings={settings}/>}
        {tab===6&&<CalcTab settings={settings}/>}
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

      {tab===1&&<button style={S.fab} onClick={()=>{setEditProd({_legais:prodSubTab==="legais"});setShowForm(true);}}><span style={{fontSize:22,color:"white"}}>＋</span></button>}
      {tab===3&&<button style={S.fab} onClick={()=>{setEditGasto(null);setShowGastoForm(true);}}><span style={{fontSize:22,color:"white"}}>＋</span></button>}

      {showSettings&&<SettingsModal settings={settings} onSave={s=>{setSettings(s);notify("Configurações salvas!");}} onImport={handleImport} onClose={()=>setShowSettings(false)}/>}
      {showForm&&<ProdutoForm prod={editProd} onSave={saveProd} onClose={()=>{setShowForm(false);setEditProd(null);}}/>}
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
      {/* Hero */}
      <div style={S.heroCard}>
        <div style={{fontSize:12,fontWeight:500,color:"rgba(255,255,255,0.75)",marginBottom:3}}>Total planejado</div>
        <div style={{fontSize:30,fontWeight:800,color:"#fff",letterSpacing:"-1px",lineHeight:1}}>R$ {stats.valorTotalBRL.toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2})}</div>
        <div style={{fontSize:13,color:"rgba(255,255,255,0.75)",marginTop:4}}>Dólar pago: {fmtBRL(settings.dollarPago,4)} · Ajustado: {fmtBRL(calcDolarAjustado(settings),4)}</div>
        <div style={{display:"flex",gap:8,marginTop:14,flexWrap:"wrap"}}>
          {[`IOF ${settings.iof}%`,`Spread ${settings.spread}%`,`Taxa ${settings.taxa}%`].map(t=>(
            <span key={t} style={{background:"rgba(255,255,255,0.15)",borderRadius:999,padding:"3px 10px",fontSize:11,color:"rgba(255,255,255,0.9)",fontWeight:500}}>{t}</span>
          ))}
        </div>
      </div>

      {/* Dólar levando */}
      <div style={{...S.card,background:"linear-gradient(135deg,#F0FDF4,#ECFDF5)",border:`1px solid ${C.success}33`}}>
        <div style={{fontSize:12,fontWeight:600,color:C.success,marginBottom:8,textTransform:"uppercase",letterSpacing:"0.5px"}}>💵 Dólares na viagem</div>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end"}}>
          <div>
            <div style={{fontSize:28,fontWeight:800,color:C.success,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(stats.valorTotalUSD,2)}</div>
            <div style={{fontSize:13,color:C.textMid,marginTop:2}}>planejado para compras</div>
          </div>
          <div style={{textAlign:"right"}}>
            <div style={{fontSize:16,fontWeight:700,color:usdRestante>=0?C.textMid:C.danger,fontFamily:"'DM Mono',monospace"}}>
              {usdRestante>=0?"Sobra":"Falta"} US${fmtN(Math.abs(usdRestante),2)}
            </div>
            <div style={{fontSize:12,color:C.textLight}}>de US$ {settings.totalDolarViagem} total</div>
          </div>
        </div>
        <div style={{height:6,background:"#D1FAE5",borderRadius:999,overflow:"hidden",marginTop:12}}>
          <div style={{width:`${Math.min(100,(stats.valorTotalUSD/settings.totalDolarViagem)*100)}%`,height:"100%",background:C.success,borderRadius:999}}/>
        </div>
      </div>

      {/* 4 mini cards */}
      <div style={S.grid4}>
        {[{label:"Produtos",value:stats.total,color:C.primary},{label:"Comprados",value:stats.comprados,color:C.success},{label:"Pendentes",value:stats.pendentes,color:C.warning},{label:"Lojas",value:stats.lojas,color:C.purple}].map(({label,value,color})=>(
          <div key={label} style={{...S.card,padding:"12px 8px",textAlign:"center",flex:1,marginBottom:0}}>
            <div style={{fontSize:20,fontWeight:800,color,fontFamily:"'DM Mono',monospace"}}>{value}</div>
            <div style={{fontSize:10,color:C.textLight,marginTop:2,fontWeight:500}}>{label}</div>
          </div>
        ))}
      </div>

      {/* Meus gastos reais */}
      <div style={{...S.card,background:"linear-gradient(135deg,#FFF7ED,#FFFBEB)",border:`1px solid ${C.warning}33`}}>
        <div style={{fontSize:12,fontWeight:600,color:C.warning,marginBottom:8,textTransform:"uppercase",letterSpacing:"0.5px"}}>💸 Meus gastos reais</div>
        <div style={{fontSize:26,fontWeight:800,color:C.warning,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(stats.totalMeusGastosUSD,2)}</div>
        <div style={{fontSize:13,color:C.textMid,marginTop:2}}>≈ R$ {(stats.totalMeusGastosUSD*settings.dollarPago).toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2})} · apenas minha parte</div>
      </div>

      {/* Progresso + peso */}
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
            <span style={{fontSize:11,color:C.textLight}}>⚖ Peso: {(stats.pesoTotal/1000).toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2})}kg / {(settings.pesoMax/1000).toLocaleString("pt-BR",{minimumFractionDigits:1,maximumFractionDigits:1})}kg</span>
            <span style={{fontSize:11,color:pesoColor,fontWeight:700}}>{fmtN(pesoPercent,0)}%</span>
          </div>
          <div style={{height:5,background:C.borderLight,borderRadius:999,overflow:"hidden"}}>
            <div style={{width:`${pesoPercent}%`,height:"100%",background:pesoColor,borderRadius:999}}/>
          </div>
        </div>
      </div>

      {/* Financeiro */}
      <div style={S.card}>
        <div style={{fontWeight:700,fontSize:14,color:C.text,marginBottom:12}}>💰 Resumo financeiro</div>
        {[{label:"USD total planejado",value:`${fmtUSD(stats.valorTotalUSD,2)}`,color:C.primary},{label:"BRL previsto (c/ taxas)",value:`${fmtBRL(stats.valorTotalBRL,2)}`,color:C.text},{label:"BRL já gasto",value:`${fmtBRL(stats.valorGasto,2)}`,color:C.success},{label:"Meus gastos (USD)",value:`${fmtUSD(stats.totalMeusGastosUSD,2)}`,color:C.warning}].map(({label,value,color})=>(
          <div key={label} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"9px 0",borderBottom:`1px solid ${C.borderLight}`}}>
            <span style={{fontSize:13,color:C.textMid}}>{label}</span>
            <span style={{fontSize:14,fontWeight:700,color,fontFamily:"'DM Mono',monospace"}}>{value}</span>
          </div>
        ))}
      </div>
      {pesoPercent>=80&&<div style={{background:pesoBg,border:`1px solid ${pesoColor}33`,borderRadius:12,padding:"10px 14px",fontSize:13,color:pesoColor,fontWeight:600,marginBottom:12}}>⚠ Peso da mala em {fmtN(pesoPercent,0)}% do limite!</div>}
    </div>
  );
}

// ─── GASTOS TAB ───────────────────────────────────────────────────────────────
function GastosTab({gastos,settings,onAdd,onEdit,onDelete,onTogglePago,produtos,onToggleStatus}) {
  const [filtro,setFiltro]=useState("todos");

  const totalUSD=gastos.reduce((a,g)=>a+calcMinhaParteUSD(g),0);
  const aReceberUSD=gastos.reduce((a,g)=>{
    if(!g.divisao||g.divisao.length===0) return a;
    return a+g.divisao.filter(p=>!p.pago).reduce((s,p)=>s+(parseFloat(p.valor)||0),0);
  },0);

  const filtrados=gastos.filter(g=>{
    if(filtro==="compras") return g.tipo==="produto";
    if(filtro==="livres") return g.tipo!=="produto";
    return true;
  }).sort((a,b)=>(b.id||0)-(a.id||0));

  return (
    <div style={S.page}>
      {/* Resumo topo */}
      <div style={S.heroCard}>
        <div style={{fontSize:12,fontWeight:500,color:"rgba(255,255,255,0.75)",marginBottom:4}}>Meus gastos totais</div>
        <div style={{fontSize:30,fontWeight:800,color:"#fff",letterSpacing:"-1px"}}>{fmtUSD(totalUSD,2)}</div>
        <div style={{fontSize:13,color:"rgba(255,255,255,0.65)",marginTop:4}}>≈ R$ {(totalUSD*settings.dollarPago).toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2})}</div>
        <div style={{display:"flex",gap:12,marginTop:12}}>
          <div style={{background:"rgba(255,255,255,0.15)",borderRadius:12,padding:"8px 14px",flex:1,textAlign:"center"}}>
            <div style={{fontSize:11,color:"rgba(255,255,255,0.7)",marginBottom:2}}>A receber</div>
            <div style={{fontSize:15,fontWeight:700,color:"#fff",fontFamily:"'DM Mono',monospace"}}>{fmtUSD(aReceberUSD,2)}</div>
          </div>
          <div style={{background:"rgba(255,255,255,0.15)",borderRadius:12,padding:"8px 14px",flex:1,textAlign:"center"}}>
            <div style={{fontSize:11,color:"rgba(255,255,255,0.7)",marginBottom:2}}>Gastos</div>
            <div style={{fontSize:15,fontWeight:700,color:"#fff",fontFamily:"'DM Mono',monospace"}}>{gastos.length}</div>
          </div>
        </div>
      </div>

      {/* Filtros */}
      <div style={{display:"flex",gap:4,background:C.borderLight,borderRadius:12,padding:4,marginBottom:12}}>
        {[["todos","Todos"],["compras","🛍 Compras"],["livres","✏ Manuais"]].map(([v,l])=>(
          <button key={v} style={{flex:1,padding:"8px 4px",borderRadius:9,border:"none",cursor:"pointer",fontSize:12,fontWeight:600,background:filtro===v?C.bgCard:"transparent",color:filtro===v?C.primary:C.textMid,boxShadow:filtro===v?"0 1px 4px rgba(0,0,0,0.08)":"none"}} onClick={()=>setFiltro(v)}>{l}</button>
        ))}
      </div>

      {filtrados.length===0&&<Empty text="Nenhum gasto ainda. Marque produtos como comprados ou adicione gastos manualmente."/>}

      {filtrados.map(g=><GastoCard key={g.id} g={g} settings={settings} onEdit={()=>onEdit(g)} onDelete={()=>onDelete(g.id)} onTogglePago={onTogglePago} produtos={produtos}/>)}
    </div>
  );
}

function GastoCard({g,settings,onEdit,onDelete,onTogglePago,produtos}) {
  const [expanded,setExpanded]=useState(false);
  const totalUSD=parseFloat(g.usd)||0;
  const minhaUSD=calcMinhaParteUSD(g);
  const minhaBRL=usdToBRL(minhaUSD,g,settings);
  const totalBRL=usdToBRL(totalUSD,g,settings);
  const cotUsada=parseFloat(g.dolarPago)||settings.dollarPago;
  const imgSrc = g.imagem || (g.produtoId&&produtos?.find(p=>p.id===g.produtoId)?.imagem) || "";
  const temDivisao=g.divisao&&g.divisao.length>0;
  const totalPessoas=temDivisao?1+g.divisao.length:1;
  const aReceberUSD=temDivisao?g.divisao.filter(p=>!p.pago).reduce((s,p)=>s+(parseFloat(p.valor)||0),0):0;

  return (
    <div style={{...S.card,marginBottom:10}}>
      <div style={{display:"flex",gap:12,alignItems:"flex-start"}}>
        <div style={{width:40,height:40,borderRadius:12,background:g.tipo==="produto"?C.primaryLight:C.purpleLight,display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0,overflow:"hidden"}}>
          {imgSrc
            ? <img src={imgSrc} alt="" style={{width:"100%",height:"100%",objectFit:"cover",borderRadius:12}} onError={e=>{e.target.style.display="none";e.target.nextSibling.style.display="flex";}} />
            : null}
          <span style={{display:imgSrc?"none":"flex",alignItems:"center",justifyContent:"center",width:"100%",height:"100%"}}>
            {LOJA_EMOJI[g.loja]||g.categoria?.split(" ")[0]||"💳"}
          </span>
        </div>
        <div style={{flex:1}} onClick={()=>setExpanded(e=>!e)}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
            <div>
              <div style={{fontWeight:600,fontSize:14,color:C.text,lineHeight:1.3}}>{g.descricao}</div>
              <div style={{fontSize:12,color:C.textLight,marginTop:2}}>{g.loja||g.categoria} · {g.data}</div>
            </div>
            <div style={{textAlign:"right",flexShrink:0}}>
              <div style={{fontSize:15,fontWeight:800,color:C.primary,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(minhaUSD,2)}</div>
              <div style={{fontSize:11,color:C.textLight,fontFamily:"'DM Mono',monospace"}}>≈ {fmtBRL(minhaBRL,2)}</div>
              {temDivisao&&<div style={{fontSize:11,color:C.textLight}}>de {fmtUSD(totalUSD,2)} total</div>}
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
              {aReceberUSD>0&&<span style={{...S.tag,background:C.warningLight,color:C.warning,borderColor:C.warning+"33",fontSize:11}}>A receber {fmtUSD(aReceberUSD,2)}</span>}
            </div>
          )}
        </div>
      </div>
      {expanded&&(
        <div style={{marginTop:12,paddingTop:12,borderTop:`1px solid ${C.borderLight}`}}>
          {[
            {label:"Total USD",value:`${fmtUSD(totalUSD,2)}`,color:C.primary},
            {label:"Minha parte USD",value:`${fmtUSD(minhaUSD,2)}`,color:C.primary},
            {label:"Minha parte BRL",value:`${fmtBRL(minhaBRL,2)}`,color:C.textMid},
            {label:"Cotação usada",value:`${fmtBRL(cotUsada,4)}`},
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
                    <span style={{fontSize:13,fontFamily:"'DM Mono',monospace",color:C.textMid}}>{fmtUSD((parseFloat(p.valor)||0),2)}</span>
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
  const somaDivisao=f.divisao.reduce((a,p)=>a+(parseFloat(p.valor)||0),0);
  const minhaParteUSD=Math.max(0,totalUSD-somaDivisao);
  const minhaParteBRL=minhaParteUSD*cotacao;

  function addPessoa(){
    if(!novaPessoa.trim())return;
    const qtd=f.divisao.length+2; // +1 nova pessoa +1 eu
    const valorPadrao=totalUSD>0?parseFloat((totalUSD/qtd).toFixed(2)):0;
    setF(p=>({...p,divisao:[...p.divisao,{nome:novaPessoa.trim(),pago:false,valor:valorPadrao}]}));
    setNovaPessoa("");
  }
  function removePessoa(i){setF(p=>({...p,divisao:p.divisao.filter((_,idx)=>idx!==i)}));}
  function updateValor(i,val){setF(p=>({...p,divisao:p.divisao.map((item,idx)=>idx===i?{...item,valor:parseFloat(val)||0}:item)}));}

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
        <span>Cotação do dólar pago: <strong style={{color:C.primary}}>{fmtBRL((parseFloat(f.dolarPago)||settings.dollarPago),4)}</strong></span>
        <span style={{color:C.textLight,fontSize:11}}>automática das configurações</span>
      </div>
      <label style={S.label}>Data</label>
      <input style={S.input} placeholder="DD/MM/AAAA" value={f.data} onChange={e=>setF(p=>({...p,data:e.target.value}))}/>

      {/* Preview valor */}
      {totalUSD>0&&(
        <div style={{background:C.primaryLight,borderRadius:10,padding:"10px 14px",marginBottom:14}}>
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
            <span style={{fontSize:13,color:C.textMid}}>Total USD</span>
            <span style={{fontSize:14,fontWeight:800,color:C.primary,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(totalUSD,2)}</span>
          </div>
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
            <span style={{fontSize:13,color:C.textMid}}>≈ Total BRL</span>
            <span style={{fontSize:13,fontWeight:600,color:C.textMid,fontFamily:"'DM Mono',monospace"}}>{fmtBRL(totalBRL,2)}</span>
          </div>
          <div style={{display:"flex",justifyContent:"space-between"}}>
            <span style={{fontSize:13,color:C.textMid}}>Minha parte ({nPessoas}p)</span>
            <span style={{fontSize:14,fontWeight:700,color:C.success,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(minhaParteUSD,2)} · {fmtBRL(minhaParteBRL,2)}</span>
          </div>
        </div>
      )}

      {/* Divisão de pessoas */}
      <div style={{borderTop:`1px solid ${C.borderLight}`,paddingTop:14,marginBottom:14}}>
        <div style={{fontWeight:700,fontSize:14,color:C.text,marginBottom:4}}>Dividir com outras pessoas</div>
        <div style={{fontSize:12,color:C.textLight,marginBottom:10}}>Adicione quem vai dividir este gasto. O total será dividido igualmente entre você + as pessoas abaixo.</div>
        <div style={{display:"flex",gap:8,marginBottom:8}}>
          <input style={{...S.input,marginBottom:0,flex:1}} placeholder="Nome da pessoa" value={novaPessoa} onChange={e=>setNovaPessoa(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addPessoa()}/>
          <button style={{...S.btnOutline,whiteSpace:"nowrap",color:C.primary,borderColor:C.primary+"44"}} onClick={addPessoa}>＋ Add</button>
        </div>
        {f.divisao.map((p,i)=>(
          <div key={i} style={{display:"flex",alignItems:"center",gap:8,padding:"8px 10px",background:C.bg,borderRadius:10,marginBottom:6,border:`1px solid ${C.border}`}}>
            <div style={{width:28,height:28,borderRadius:"50%",background:C.primaryLight,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:700,color:C.primary,flexShrink:0}}>{p.nome[0].toUpperCase()}</div>
            <span style={{fontSize:13,fontWeight:600,color:C.text,flex:1}}>{p.nome}</span>
            <div style={{display:"flex",alignItems:"center",gap:4}}>
              <span style={{fontSize:11,color:C.textLight}}>US$</span>
              <input
                type="number" step="0.01"
                value={p.valor||""}
                onChange={e=>updateValor(i,e.target.value)}
                style={{width:76,background:C.bgCard,border:`1px solid ${C.border}`,borderRadius:7,padding:"5px 7px",fontSize:13,color:C.text,fontFamily:"'DM Mono',monospace",outline:"none",boxSizing:"border-box"}}
              />
            </div>
            <button onClick={()=>removePessoa(i)} style={{background:C.dangerLight,border:"none",borderRadius:6,width:22,height:22,cursor:"pointer",color:C.danger,fontSize:14,flexShrink:0}}>×</button>
          </div>
        ))}
        {f.divisao.length>0&&(
          <div style={{background:C.successLight,border:`1px solid ${C.success}33`,borderRadius:10,padding:"8px 12px",fontSize:13,color:C.success,fontWeight:600}}>
            ✓ Dividindo entre {nPessoas} pessoas · Sua parte: {fmtUSD(minhaParteUSD,2)} · {fmtBRL(minhaParteBRL,2)}
            {Math.abs(somaDivisao+minhaParteUSD-totalUSD)>0.02&&totalUSD>0&&(
              <div style={{fontSize:11,color:C.warning,marginTop:4,fontWeight:500}}>⚠ Soma das partes ({fmtUSD((somaDivisao+minhaParteUSD),2)}) ≠ total ({fmtUSD(totalUSD,2)})</div>
            )}
          </div>
        )}
      </div>
      <button style={S.btnPrimary} onClick={handleSave}>{gasto?.id?"Salvar alterações":"Adicionar gasto"}</button>
    </Modal>
  );
}


// ─── PARCELAS TAB ─────────────────────────────────────────────────────────────
const MESES_LABELS = ["1ª","2ª","3ª","4ª","5ª","6ª","7ª","8ª","9ª","10ª","11ª","12ª","13ª","14ª","15ª","16ª","17ª","18ª","19ª","20ª","21ª","22ª","23ª","24ª"];
const MESES_NOMES = ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"];
const MESES_SELECT = ["jan/25","fev/25","mar/25","abr/25","mai/25","jun/25","jul/25","ago/25","set/25","out/25","nov/25","dez/25","jan/26","fev/26","mar/26","abr/26","mai/26","jun/26","jul/26","ago/26","set/26","out/26","nov/26","dez/26","jan/27","fev/27","mar/27","abr/27","mai/27","jun/27","jul/27","ago/27","set/27","out/27","nov/27","dez/27"];

// Dado "abr/26" e offset 0,1,2... retorna "abr/26","mai/26","jun/26"...
function addMeses(mesAno, offset) {
  if (!mesAno) return "";
  const [m, a] = mesAno.split("/");
  const mi = MESES_NOMES.indexOf(m.toLowerCase());
  if (mi < 0) return mesAno;
  const total = mi + offset;
  const mes = MESES_NOMES[total % 12];
  const ano = (parseInt("20" + a) + Math.floor(total / 12)).toString().slice(-2);
  return `${mes}/${ano}`;
}

// Comparar dois mesAno: retorna -1,0,1
function compareMesAno(a, b) {
  const parse = s => { const [m,y]=s.split("/"); return parseInt("20"+y)*12+MESES_NOMES.indexOf(m); };
  return parse(a) - parse(b);
}

function parcelaVazia() {
  return {
    id: Date.now(),
    descricao: "",
    valorTotal: 0,
    quantidadeParcelas: 10,
    valorParcela: 0,
    cartao: "",
    statusMensal: Array(10).fill(false),
    nPessoas: "",       // opcional: dividido entre N pessoas
    minhaParte: "",     // minha parte do valorTotal
    primeiraFatura: "", // ex: "abr/26"
  };
}

function ParcelasTab({ parcelas, setParcelas }) {
  const [showForm, setShowForm] = useState(false);
  const [editItem, setEditItem] = useState(null);

  const minhaParcelaMensal = p => {
    const vt = parseFloat(p.valorTotal) || 0;
    const mp = parseFloat(p.minhaParte) || 0;
    const vp = parseFloat(p.valorParcela) || 0;
    if (mp > 0 && vt > 0) return parseFloat(((mp / vt) * vp).toFixed(2));
    return vp;
  };
  const totalMensal = parcelas.reduce((a, p) => a + minhaParcelaMensal(p), 0);
  const totalPago = parcelas.reduce((a, p) => {
    const pagas = (p.statusMensal || []).filter(Boolean).length;
    return a + pagas * minhaParcelaMensal(p);
  }, 0);
  const totalRestante = parcelas.reduce((a, p) => {
    const qt = parseInt(p.quantidadeParcelas) || 0;
    const pagas = (p.statusMensal || []).filter(Boolean).length;
    return a + Math.max(0, qt - pagas) * minhaParcelaMensal(p);
  }, 0);

  function abrirNova() { setEditItem(null); setShowForm(true); }
  function abrirEditar(p) { setEditItem(p); setShowForm(true); }

  function salvarParcela(item) {
    if (item.id && parcelas.some(p => p.id === item.id)) {
      setParcelas(ps => ps.map(p => p.id === item.id ? item : p));
    } else {
      setParcelas(ps => [...ps, { ...item, id: Date.now() }]);
    }
    setShowForm(false);
    setEditItem(null);
  }

  function excluirParcela(id) {
    setParcelas(ps => ps.filter(p => p.id !== id));
  }

  function toggleMes(parcelaId, mesIdx) {
    setParcelas(ps => ps.map(p => {
      if (p.id !== parcelaId) return p;
      const sm = [...(p.statusMensal || Array(p.quantidadeParcelas).fill(false))];
      sm[mesIdx] = !sm[mesIdx];
      return { ...p, statusMensal: sm };
    }));
  }

  return (
    <div style={S.page}>
      {/* Hero */}
      <div style={S.heroCard}>
        <div style={{fontSize:12,fontWeight:500,color:"rgba(255,255,255,0.75)",marginBottom:3}}>Parcelas da viagem</div>
        <div style={{fontSize:28,fontWeight:800,color:"#fff",letterSpacing:"-0.5px"}}>R$ {totalMensal.toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2})}<span style={{fontSize:13,fontWeight:500,opacity:0.75}}>/mês</span></div>
        <div style={{display:"flex",gap:10,marginTop:12}}>
          <div style={{background:"rgba(255,255,255,0.15)",borderRadius:12,padding:"8px 12px",flex:1,textAlign:"center"}}>
            <div style={{fontSize:10,color:"rgba(255,255,255,0.65)",marginBottom:2}}>Itens</div>
            <div style={{fontSize:16,fontWeight:700,color:"#fff"}}>{parcelas.length}</div>
          </div>
          <div style={{background:"rgba(255,255,255,0.15)",borderRadius:12,padding:"8px 12px",flex:1,textAlign:"center"}}>
            <div style={{fontSize:10,color:"rgba(255,255,255,0.65)",marginBottom:2}}>Pago</div>
            <div style={{fontSize:13,fontWeight:700,color:"#fff",fontFamily:"'DM Mono',monospace"}}>R$ {totalPago.toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2})}</div>
          </div>
          <div style={{background:"rgba(255,255,255,0.15)",borderRadius:12,padding:"8px 12px",flex:1,textAlign:"center"}}>
            <div style={{fontSize:10,color:"rgba(255,255,255,0.65)",marginBottom:2}}>Restante</div>
            <div style={{fontSize:13,fontWeight:700,color:"#fff",fontFamily:"'DM Mono',monospace"}}>R$ {totalRestante.toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2})}</div>
          </div>
        </div>
      </div>

      <button style={{...S.btnPrimary,marginBottom:14}} onClick={abrirNova}>＋ Nova parcela</button>

      {parcelas.length === 0 && <Empty text="Nenhuma parcela cadastrada ainda."/>}

      {parcelas.map(p => (
        <ParcelaCard
          key={p.id}
          p={p}
          onEditar={() => abrirEditar(p)}
          onExcluir={() => excluirParcela(p.id)}
          onToggleMes={(i) => toggleMes(p.id, i)}
        />
      ))}

      {/* Dashboard — Distribuição por Fatura */}
      {parcelas.some(p => p.primeiraFatura) && (
        <DistribuicaoFatura parcelas={parcelas} minhaParcelaMensal={minhaParcelaMensal} toggleMes={toggleMes}/>
      )}

      {showForm && (
        <ParcelaForm
          parcela={editItem}
          onSalvar={salvarParcela}
          onClose={() => { setShowForm(false); setEditItem(null); }}
        />
      )}
    </div>
  );
}

// ─── DISTRIBUIÇÃO POR FATURA ──────────────────────────────────────────────────
function DistribuicaoFatura({ parcelas, minhaParcelaMensal, toggleMes }) {
  // Coletar todos os meses do range de todas as parcelas
  const allMeses = new Set();
  parcelas.forEach(p => {
    if (!p.primeiraFatura) return;
    const qt = parseInt(p.quantidadeParcelas) || 0;
    for (let i = 0; i < qt; i++) allMeses.add(addMeses(p.primeiraFatura, i));
  });

  // Ordenar cronologicamente
  const mesesOrdenados = [...allMeses].sort(compareMesAno);
  if (mesesOrdenados.length === 0) return null;

  // Parcelas que têm primeiraFatura
  const parcelasComData = parcelas.filter(p => p.primeiraFatura);

  return (
    <div style={{marginTop:8,marginBottom:16}}>
      <div style={{fontSize:13,fontWeight:700,color:C.text,marginBottom:10,paddingTop:4}}>📅 Distribuição por Fatura</div>
      <div style={{overflowX:"auto",borderRadius:14,border:`1px solid ${C.border}`}}>
        <table style={{borderCollapse:"collapse",minWidth:"100%",fontSize:12}}>
          <thead>
            <tr style={{background:C.bg}}>
              <th style={{padding:"8px 12px",textAlign:"left",fontWeight:700,color:C.textMid,borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap",position:"sticky",left:0,background:C.bg,zIndex:1}}>Fatura</th>
              {parcelasComData.map(p => (
                <th key={p.id} style={{padding:"8px 10px",textAlign:"center",fontWeight:700,color:C.purple,borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap",minWidth:90,fontSize:11}}>
                  {p.descricao.length > 10 ? p.descricao.slice(0,10)+"…" : p.descricao}
                </th>
              ))}
              <th style={{padding:"8px 10px",textAlign:"right",fontWeight:700,color:C.text,borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap"}}>Total/Fatura</th>
            </tr>
          </thead>
          <tbody>
            {mesesOrdenados.map((mes, ri) => {
              let totalFatura = 0;
              return (
                <tr key={mes} style={{background:ri%2===0?C.bgCard:"transparent"}}>
                  <td style={{padding:"7px 12px",fontWeight:600,color:C.text,borderBottom:`1px solid ${C.borderLight}`,whiteSpace:"nowrap",position:"sticky",left:0,background:ri%2===0?C.bgCard:"#fff",zIndex:1}}>{mes}</td>
                  {parcelasComData.map(p => {
                    const qt = parseInt(p.quantidadeParcelas) || 0;
                    const idx = mesesOrdenados.indexOf(mes);
                    // Qual índice da parcela corresponde a este mês?
                    const parcelaIdx = (() => {
                      for (let i = 0; i < qt; i++) {
                        if (addMeses(p.primeiraFatura, i) === mes) return i;
                      }
                      return -1;
                    })();
                    if (parcelaIdx < 0) {
                      return <td key={p.id} style={{padding:"7px 10px",textAlign:"center",color:C.textLight,borderBottom:`1px solid ${C.borderLight}`}}>—</td>;
                    }
                    const statusMensal = p.statusMensal || Array(qt).fill(false);
                    const pago = statusMensal[parcelaIdx] || false;
                    const val = minhaParcelaMensal(p);
                    totalFatura += val;
                    return (
                      <td key={p.id} style={{padding:"7px 10px",textAlign:"center",borderBottom:`1px solid ${C.borderLight}`}}>
                        <button
                          onClick={() => toggleMes(p.id, parcelaIdx)}
                          style={{background:pago?C.successLight:C.warningLight,color:pago?C.success:C.warning,border:`1px solid ${pago?C.success:C.warning}44`,borderRadius:8,padding:"3px 7px",fontSize:11,fontWeight:700,cursor:"pointer",whiteSpace:"nowrap",fontFamily:"'DM Mono',monospace"}}
                        >
                          {pago?"✓ ":""}{fmtBRL(val,2)}
                        </button>
                      </td>
                    );
                  })}
                  <td style={{padding:"7px 10px",textAlign:"right",fontWeight:700,color:C.primary,borderBottom:`1px solid ${C.borderLight}`,fontFamily:"'DM Mono',monospace",whiteSpace:"nowrap"}}>{fmtBRL(totalFatura,2)}</td>
                </tr>
              );
            })}
            {/* Linha de total geral */}
            <tr style={{background:C.primaryLight}}>
              <td style={{padding:"8px 12px",fontWeight:700,color:C.primary,position:"sticky",left:0,background:C.primaryLight,zIndex:1}}>Total</td>
              {parcelasComData.map(p => (
                <td key={p.id} style={{padding:"8px 10px",textAlign:"center",fontWeight:700,color:C.purple,fontFamily:"'DM Mono',monospace",fontSize:11}}>
                  {fmtBRL(parseFloat(p.minhaParte)||parseFloat(p.valorTotal)||0,2)}
                </td>
              ))}
              <td style={{padding:"8px 10px",textAlign:"right",fontWeight:800,color:C.primary,fontFamily:"'DM Mono',monospace"}}>
                {fmtBRL(parcelasComData.reduce((a,p)=>a+(parseFloat(p.minhaParte)||parseFloat(p.valorTotal)||0),0),2)}
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

function ParcelaCard({ p, onEditar, onExcluir, onToggleMes }) {
  const [expanded, setExpanded] = useState(false);
  const qt = parseInt(p.quantidadeParcelas) || 0;
  const statusMensal = p.statusMensal || Array(qt).fill(false);
  const pagas = statusMensal.filter(Boolean).length;
  const restante = qt - pagas;
  const pct = qt > 0 ? Math.round(pagas / qt * 100) : 0;
  const valorParcela = parseFloat(p.valorParcela) || 0;

  return (
    <div style={{...S.card,marginBottom:10}}>
      <div style={{display:"flex",gap:12,alignItems:"flex-start"}}>
        <div style={{width:40,height:40,borderRadius:12,background:C.purpleLight,display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0}}>💳</div>
        <div style={{flex:1}} onClick={() => setExpanded(e => !e)}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
            <div>
              <div style={{fontWeight:600,fontSize:14,color:C.text,lineHeight:1.3}}>{p.descricao||"Sem descrição"}</div>
              <div style={{fontSize:12,color:C.textLight,marginTop:2}}>{p.cartao||"—"} · {qt}x{p.primeiraFatura?` · desde ${p.primeiraFatura}`:""}</div>
            </div>
            <div style={{textAlign:"right",flexShrink:0}}>
              <div style={{fontSize:15,fontWeight:800,color:C.purple,fontFamily:"'DM Mono',monospace"}}>{fmtBRL(valorParcela,2)}<span style={{fontSize:10,fontWeight:500,color:C.textLight}}>/mês</span></div>
              <div style={{fontSize:11,color:C.textLight,fontFamily:"'DM Mono',monospace"}}>Total: {fmtBRL(parseFloat(p.valorTotal)||0)}</div>
              {p.minhaParte>0&&<div style={{fontSize:11,color:C.purple,fontFamily:"'DM Mono',monospace",fontWeight:600}}>Minha parte: {fmtBRL(p.minhaParte)}</div>}
            </div>
          </div>
          {/* Barra de progresso */}
          <div style={{marginTop:8}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
              <span style={{fontSize:11,color:C.textLight}}>{pagas}/{qt} pagas</span>
              <span style={{fontSize:11,fontWeight:700,color:pct===100?C.success:C.purple}}>{pct}%</span>
            </div>
            <div style={{height:5,background:C.borderLight,borderRadius:999,overflow:"hidden"}}>
              <div style={{width:`${pct}%`,height:"100%",background:pct===100?C.success:C.purple,borderRadius:999,transition:"width 0.4s"}}/>
            </div>
          </div>
        </div>
      </div>

      {expanded && (
        <div style={{marginTop:14,paddingTop:12,borderTop:`1px solid ${C.borderLight}`}}>
          <div style={{fontSize:12,fontWeight:700,color:C.textLight,textTransform:"uppercase",letterSpacing:"0.6px",marginBottom:10}}>Status das parcelas</div>
          <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:12}}>
            {Array.from({length:qt}).map((_,i) => {
              const label = p.primeiraFatura ? addMeses(p.primeiraFatura, i) : (MESES_LABELS[i]||`${i+1}ª`);
              const pago = statusMensal[i]||false;
              return (
                <label key={i} style={{display:"flex",flexDirection:"column",alignItems:"center",gap:3,minWidth:44,cursor:"pointer",background:pago?C.successLight:C.borderLight,borderRadius:8,padding:"6px 4px",border:`1px solid ${pago?C.success+"44":C.border}`}}>
                  <input
                    type="checkbox"
                    checked={pago}
                    onChange={() => onToggleMes(i)}
                    style={{width:15,height:15,accentColor:C.success,cursor:"pointer"}}
                  />
                  <span style={{fontSize:10,fontWeight:700,color:pago?C.success:C.textMid,textAlign:"center",lineHeight:1.2}}>{label}</span>
                </label>
              );
            })}
          </div>
          {restante > 0 && (
            <div style={{background:C.warningLight,border:`1px solid ${C.warning}33`,borderRadius:10,padding:"8px 12px",fontSize:13,color:C.warning,fontWeight:600,marginBottom:12}}>
              ⏳ {restante} parcela(s) restante(s) · {fmtBRL((restante*valorParcela),2)}
            </div>
          )}
          {restante === 0 && qt > 0 && (
            <div style={{background:C.successLight,border:`1px solid ${C.success}33`,borderRadius:10,padding:"8px 12px",fontSize:13,color:C.success,fontWeight:600,marginBottom:12}}>
              ✅ Todas as parcelas pagas!
            </div>
          )}
          <div style={{display:"flex",gap:8}}>
            <button style={S.btnOutline} onClick={onEditar}>✏ Editar</button>
            <button style={{...S.btnOutline,color:C.danger,borderColor:C.danger+"44"}} onClick={onExcluir}>🗑 Remover</button>
          </div>
        </div>
      )}
    </div>
  );
}

function ParcelaForm({ parcela, onSalvar, onClose }) {
  const [f, setF] = useState(() => parcela ? { ...parcela } : parcelaVazia());

  // Auto-calcular valorParcela quando mudam valorTotal ou quantidadeParcelas
  useEffect(() => {
    const vt = parseFloat(f.valorTotal) || 0;
    const qt = parseInt(f.quantidadeParcelas) || 0;
    if (vt > 0 && qt > 0) {
      setF(p => ({ ...p, valorParcela: parseFloat((vt / qt).toFixed(2)) }));
    }
  }, [f.valorTotal, f.quantidadeParcelas]);

  // Auto-calcular minhaParte quando mudam valorTotal ou nPessoas
  useEffect(() => {
    const vt = parseFloat(f.valorTotal) || 0;
    const np = parseInt(f.nPessoas) || 0;
    if (vt > 0 && np > 1) {
      setF(p => ({ ...p, minhaParte: parseFloat((vt / np).toFixed(2)) }));
    }
  }, [f.valorTotal, f.nPessoas]);

  function salvar() {
    if (!f.descricao.trim()) return alert("Informe a descrição");
    if (!(parseInt(f.quantidadeParcelas) > 0)) return alert("Informe a quantidade de parcelas");
    const qt = parseInt(f.quantidadeParcelas);
    // Garantir que statusMensal tem o tamanho certo
    const sm = Array.from({length:qt}, (_, i) => (f.statusMensal||[])[i] || false);
    onSalvar({ ...f, quantidadeParcelas: qt, valorTotal: parseFloat(f.valorTotal)||0, valorParcela: parseFloat(f.valorParcela)||0, nPessoas: parseInt(f.nPessoas)||0, minhaParte: parseFloat(f.minhaParte)||0, primeiraFatura: f.primeiraFatura||'', statusMensal: sm });
  }

  return (
    <Modal title={parcela ? "Editar parcela" : "Nova parcela"} onClose={onClose}>
      <label style={S.label}>Descrição / Item *</label>
      <input style={S.input} placeholder="Ex: Passagem aérea, Hotel, Ingressos..." value={f.descricao} onChange={e => setF(p => ({...p, descricao: e.target.value}))}/>

      <label style={S.label}>Valor Total (R$)</label>
      <input style={S.input} type="number" step="0.01" placeholder="Ex: 2500.00" value={f.valorTotal||""} onChange={e => setF(p => ({...p, valorTotal: e.target.value}))}/>

      <label style={S.label}>Quantidade de Parcelas *</label>
      <input style={S.input} type="number" step="1" min="1" max="24" placeholder="Ex: 10" value={f.quantidadeParcelas||""} onChange={e => setF(p => ({...p, quantidadeParcelas: e.target.value}))}/>

      <label style={S.label}>Valor da Parcela (R$)</label>
      <input style={S.input} type="number" step="0.01" placeholder="Calculado automaticamente" value={f.valorParcela||""} onChange={e => setF(p => ({...p, valorParcela: e.target.value}))}/>

      <label style={S.label}>Cartão / Forma de Pagamento</label>
      <input style={S.input} placeholder="Ex: Nubank, Itaú, C6..." value={f.cartao||""} onChange={e => setF(p => ({...p, cartao: e.target.value}))}/>

      <label style={S.label}>Primeira Fatura</label>
      <select style={S.input} value={f.primeiraFatura||""} onChange={e => setF(p => ({...p, primeiraFatura: e.target.value}))}>
        <option value="">— Selecione o mês inicial —</option>
        {MESES_SELECT.map(m => <option key={m} value={m}>{m}</option>)}
      </select>

      <div style={{fontSize:11,fontWeight:700,color:C.textLight,textTransform:"uppercase",letterSpacing:"0.6px",marginBottom:8,marginTop:4}}>👥 Divisão (opcional)</div>
      <div style={{display:"flex",gap:8,marginBottom:14}}>
        <div style={{flex:1}}>
          <label style={S.label}>Dividido entre (pessoas)</label>
          <input style={S.input} type="number" min="2" step="1" placeholder="Ex: 2" value={f.nPessoas||""} onChange={e => setF(p => ({...p, nPessoas: e.target.value}))}/>
        </div>
        <div style={{flex:1}}>
          <label style={S.label}>Minha parte (R$)</label>
          <input style={S.input} type="number" step="0.01" placeholder="Calculado auto." value={f.minhaParte||""} onChange={e => setF(p => ({...p, minhaParte: e.target.value}))}/>
        </div>
      </div>

      {f.quantidadeParcelas > 0 && parseFloat(f.valorParcela) > 0 && (
        <div style={{background:C.purpleLight,border:`1px solid ${C.purple}33`,borderRadius:10,padding:"10px 14px",marginBottom:14,fontSize:13,color:C.purple,fontWeight:600}}>
          💳 {f.quantidadeParcelas}x de {fmtBRL(parseFloat(f.valorParcela),2)} = {fmtBRL((f.quantidadeParcelas*parseFloat(f.valorParcela)),2)}
        </div>
      )}

      <button style={S.btnPrimary} onClick={salvar}>{parcela ? "Salvar alterações" : "Adicionar parcela"}</button>
    </Modal>
  );
}

// ─── PRODUTOS TAB ─────────────────────────────────────────────────────────────
function ProdutosTab({produtos,itensLegais,settings,onToggle,onDelete,onEdit,onAdd,onMoveToList,onSubTabChange}) {
  const [subTab,setSubTab]=useState("compras");
  function changeSubTab(v){setSubTab(v);onSubTabChange&&onSubTabChange(v);}
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
          <button key={v} style={{flex:1,padding:"9px 8px",borderRadius:9,border:"none",cursor:"pointer",fontSize:13,fontWeight:600,background:subTab===v?C.bgCard:"transparent",color:subTab===v?C.primary:C.textMid,boxShadow:subTab===v?"0 1px 4px rgba(0,0,0,0.08)":"none"}} onClick={()=>changeSubTab(v)}>{l}</button>
        ))}
      </div>
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
              <div style={{fontSize:14,fontWeight:800,color:C.primary,fontFamily:"'DM Mono',monospace"}}>US${p.usd}</div>
              <div style={{fontSize:11,color:C.textLight,fontFamily:"'DM Mono',monospace"}}>{fmtBRL(brl,0)}</div>
            </div>
          </div>
          <div style={{display:"flex",gap:5,marginTop:6,flexWrap:"wrap"}}>
            <span style={S.tag}>{p.loja}</span>
            {!isLegais&&<span style={{...S.tag,background:pc.bg,color:pc.color,borderColor:pc.color+"33"}}>{p.prioridade}</span>}
            <span style={S.tag}>{(peso/1000).toLocaleString("pt-BR",{minimumFractionDigits:3,maximumFractionDigits:3})}kg</span>
            {p.status==="comprado"&&<span style={{...S.tag,background:C.successLight,color:C.success,borderColor:C.success+"33"}}>✓ Comprado</span>}
          </div>
        </div>
      </div>
      {expanded&&(
        <div style={{marginTop:12,paddingTop:12,borderTop:`1px solid ${C.borderLight}`}}>
          {[{label:"BRL previsto",value:`${fmtBRL(brl,2)}`},...(brlPago?[{label:"BRL pago",value:`${fmtBRL(brlPago,2)}`,color:C.success},{label:"Diferença",value:`${fmtBRL((brl-brlPago),2)}`,color:brl>brlPago?C.success:C.danger}]:[]),{label:"USD c/ taxa",value:`${fmtUSD(calcUsdFinal(p.usd,settings),2)}`,color:C.textMid}].map(({label,value,color})=>(
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
              <div style={{fontSize:14,fontWeight:800,color:C.primary,fontFamily:"'DM Mono',monospace"}}>US${p.usd}</div>
              <div style={{fontSize:11,color:C.textLight,fontFamily:"'DM Mono',monospace"}}>{fmtBRL(calcBRL(p.usd,settings),0)}</div>
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

  // ── Compras stats ──────────────────────────────────────────────────────────
  const porLoja=useMemo(()=>{
    const map={};
    produtos.forEach(p=>{
      if(!map[p.loja])map[p.loja]={total:0,comprados:0,usd:0,brl:0,peso:0};
      map[p.loja].total++;
      if(p.status==="comprado")map[p.loja].comprados++;
      map[p.loja].usd+=p.usd;
      map[p.loja].brl+=calcBRL(p.usd,settings);
      map[p.loja].peso+=pesoGramas(p);
    });
    return Object.entries(map).sort((a,b)=>b[1].usd-a[1].usd);
  },[produtos,settings]);
  const totalUSDCompras=porLoja.reduce((a,[,v])=>a+v.usd,0);

  // ── Gastos stats ───────────────────────────────────────────────────────────
  const porCategoria=useMemo(()=>{
    const map={};
    gastos.forEach(g=>{
      const cat=g.categoria||"💳 Outros";
      if(!map[cat])map[cat]={total:0,usd:0,minhaUSD:0,aReceber:0};
      const usd=parseFloat(g.usd)||0;
      const minha=calcMinhaParteUSD(g);
      const aRec=g.divisao?g.divisao.filter(p=>!p.pago).reduce((s,p)=>s+(parseFloat(p.valor)||0),0):0;
      map[cat].total++;
      map[cat].usd+=usd;
      map[cat].minhaUSD+=minha;
      map[cat].aReceber+=aRec;
    });
    return Object.entries(map).sort((a,b)=>b[1].usd-a[1].usd);
  },[gastos]);

  const porPessoa=useMemo(()=>{
    const map={};
    gastos.forEach(g=>{
      if(!g.divisao||g.divisao.length===0) return;
      g.divisao.forEach(p=>{
        const part=parseFloat(p.valor)||0;
        if(!map[p.nome])map[p.nome]={totalUSD:0,pago:0,pendente:0,gastos:0};
        map[p.nome].totalUSD+=part;
        map[p.nome].gastos++;
        if(p.pago) map[p.nome].pago+=part;
        else map[p.nome].pendente+=part;
      });
    });
    return Object.entries(map).sort((a,b)=>b[1].totalUSD-a[1].totalUSD);
  },[gastos]);

  const totalGastosUSD=gastos.reduce((a,g)=>a+(parseFloat(g.usd)||0),0);
  const totalMeuUSD=gastos.reduce((a,g)=>a+calcMinhaParteUSD(g),0);
  const totalAReceberUSD=gastos.reduce((a,g)=>{
    if(!g.divisao||g.divisao.length===0) return a;
    return a+g.divisao.filter(p=>!p.pago).reduce((s,p)=>s+(parseFloat(p.valor)||0),0);
  },0);

  return (
    <div style={S.page}>
      {/* Sub-tabs */}
      <div style={{display:"flex",gap:4,background:C.borderLight,borderRadius:12,padding:4,marginBottom:14}}>
        {[["compras","🛒 Compras"],["gastos","💸 Gastos"]].map(([v,l])=>(
          <button key={v} style={{flex:1,padding:"9px 8px",borderRadius:9,border:"none",cursor:"pointer",fontSize:13,fontWeight:600,background:secao===v?C.bgCard:"transparent",color:secao===v?C.primary:C.textMid,boxShadow:secao===v?"0 1px 4px rgba(0,0,0,0.08)":"none",transition:"all 0.2s"}} onClick={()=>setSecao(v)}>{l}</button>
        ))}
      </div>

      {/* ── COMPRAS ── */}
      {secao==="compras"&&(
        <>
          {/* Resumo compras */}
          <div style={{display:"flex",gap:8,marginBottom:12}}>
            {[{label:"Total USD",value:`${fmtUSD(totalUSDCompras,0)}`,color:C.primary},{label:"Comprados",value:`${produtos.filter(p=>p.status==="comprado").length}/${produtos.length}`,color:C.success}].map(({label,value,color})=>(
              <div key={label} style={{...S.card,flex:1,textAlign:"center",padding:"12px 8px",marginBottom:0}}>
                <div style={{fontSize:16,fontWeight:800,color,fontFamily:"'DM Mono',monospace"}}>{value}</div>
                <div style={{fontSize:11,color:C.textLight,marginTop:2}}>{label}</div>
              </div>
            ))}
          </div>

          {/* Gastos por loja */}
          <div style={S.card}>
            <div style={{fontWeight:700,fontSize:14,color:C.text,marginBottom:14}}>Compras por loja</div>
            {porLoja.map(([loja,d],i)=>{
              const pct=totalUSDCompras?(d.usd/totalUSDCompras*100):0;
              const color=barColors[i%barColors.length];
              return(
                <div key={loja} style={{marginBottom:14}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}>
                    <div style={{display:"flex",alignItems:"center",gap:8}}>
                      <div style={{width:10,height:10,borderRadius:"50%",background:color,flexShrink:0}}/>
                      <span style={{fontSize:13,fontWeight:600,color:C.text}}>{loja}</span>
                    </div>
                    <div style={{textAlign:"right"}}>
                      <span style={{fontSize:13,fontWeight:700,color,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(d.usd,0)}</span>
                      <span style={{fontSize:11,color:C.textLight,marginLeft:6}}>{fmtN(pct,0)}%</span>
                    </div>
                  </div>
                  <div style={{height:6,background:C.borderLight,borderRadius:999,overflow:"hidden",marginBottom:4}}>
                    <div style={{width:`${pct}%`,height:"100%",background:color,borderRadius:999,transition:"width 0.5s"}}/>
                  </div>
                  <div style={{fontSize:11,color:C.textLight}}>{d.comprados}/{d.total} comprados · {(d.peso/1000).toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2})}kg</div>
                </div>
              );
            })}
            {porLoja.length===0&&<div style={{fontSize:13,color:C.textLight,textAlign:"center",padding:"16px 0"}}>Nenhum produto ainda</div>}
          </div>

          {/* Ranking tabela */}
          <div style={S.card}>
            <div style={{fontWeight:700,fontSize:14,color:C.text,marginBottom:12}}>Ranking de lojas</div>
            <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr",gap:4,marginBottom:8}}>
              {["Loja","Itens","USD","Concl."].map(h=><div key={h} style={{fontSize:10,fontWeight:700,color:C.textLight,textTransform:"uppercase",letterSpacing:"0.4px"}}>{h}</div>)}
            </div>
            {porLoja.map(([loja,d])=>(
              <div key={loja} style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr",gap:4,padding:"9px 0",borderTop:`1px solid ${C.borderLight}`}}>
                <div style={{fontSize:12,fontWeight:600,color:C.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{loja}</div>
                <div style={{fontSize:12,color:C.textMid,fontFamily:"'DM Mono',monospace"}}>{d.total}</div>
                <div style={{fontSize:12,color:C.primary,fontFamily:"'DM Mono',monospace",fontWeight:700}}>{fmtUSD(d.usd,0)}</div>
                <div style={{fontSize:12,color:d.total&&d.comprados/d.total>=1?C.success:C.textMid,fontFamily:"'DM Mono',monospace"}}>{d.total?Math.round(d.comprados/d.total*100):0}%</div>
              </div>
            ))}
          </div>
        </>
      )}

      {/* ── GASTOS ── */}
      {secao==="gastos"&&(
        <>
          {/* Resumo gastos */}
          <div style={S.heroCard}>
            <div style={{fontSize:12,color:"rgba(255,255,255,0.7)",marginBottom:4}}>Total de gastos (bruto)</div>
            <div style={{fontSize:28,fontWeight:800,color:"#fff",fontFamily:"'DM Mono',monospace",letterSpacing:"-0.5px"}}>{fmtUSD(totalGastosUSD,2)}</div>
            <div style={{display:"flex",gap:8,marginTop:12}}>
              <div style={{background:"rgba(255,255,255,0.15)",borderRadius:10,padding:"8px 12px",flex:1,textAlign:"center"}}>
                <div style={{fontSize:10,color:"rgba(255,255,255,0.65)",marginBottom:2}}>Minha parte</div>
                <div style={{fontSize:14,fontWeight:800,color:"#fff",fontFamily:"'DM Mono',monospace"}}>{fmtUSD(totalMeuUSD,2)}</div>
              </div>
              <div style={{background:"rgba(255,255,255,0.15)",borderRadius:10,padding:"8px 12px",flex:1,textAlign:"center"}}>
                <div style={{fontSize:10,color:"rgba(255,255,255,0.65)",marginBottom:2}}>A receber</div>
                <div style={{fontSize:14,fontWeight:800,color:"#fff",fontFamily:"'DM Mono',monospace"}}>{fmtUSD(totalAReceberUSD,2)}</div>
              </div>
              <div style={{background:"rgba(255,255,255,0.15)",borderRadius:10,padding:"8px 12px",flex:1,textAlign:"center"}}>
                <div style={{fontSize:10,color:"rgba(255,255,255,0.65)",marginBottom:2}}>Itens</div>
                <div style={{fontSize:14,fontWeight:800,color:"#fff",fontFamily:"'DM Mono',monospace"}}>{gastos.length}</div>
              </div>
            </div>
          </div>

          {/* Por categoria */}
          {porCategoria.length>0&&(
            <div style={S.card}>
              <div style={{fontWeight:700,fontSize:14,color:C.text,marginBottom:14}}>Por categoria</div>
              {porCategoria.map(([cat,d],i)=>{
                const pct=totalGastosUSD?(d.usd/totalGastosUSD*100):0;
                const color=barColors[i%barColors.length];
                return(
                  <div key={cat} style={{marginBottom:14}}>
                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}>
                      <div style={{display:"flex",alignItems:"center",gap:8}}>
                        <div style={{width:10,height:10,borderRadius:"50%",background:color,flexShrink:0}}/>
                        <span style={{fontSize:13,fontWeight:600,color:C.text}}>{cat}</span>
                      </div>
                      <div style={{textAlign:"right"}}>
                        <div style={{fontSize:13,fontWeight:700,color,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(d.usd,2)}</div>
                        <div style={{fontSize:11,color:C.textLight}}>meu: {fmtUSD(d.minhaUSD,2)} · {fmtN(pct,0)}%</div>
                      </div>
                    </div>
                    <div style={{height:6,background:C.borderLight,borderRadius:999,overflow:"hidden"}}>
                      <div style={{width:`${pct}%`,height:"100%",background:color,borderRadius:999,transition:"width 0.5s"}}/>
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {/* Por pessoa — quem me deve */}
          {porPessoa.length>0&&(
            <div style={S.card}>
              <div style={{fontWeight:700,fontSize:14,color:C.text,marginBottom:4}}>Divisão por pessoa</div>
              <div style={{fontSize:12,color:C.textLight,marginBottom:12}}>Quanto cada pessoa deve ao total dos gastos divididos</div>
              {porPessoa.map(([nome,d])=>(
                <div key={nome} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 0",borderBottom:`1px solid ${C.borderLight}`}}>
                  <div style={{display:"flex",alignItems:"center",gap:10}}>
                    <div style={{width:32,height:32,borderRadius:"50%",background:C.primaryLight,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,fontWeight:700,color:C.primary}}>{nome[0].toUpperCase()}</div>
                    <div>
                      <div style={{fontSize:13,fontWeight:600,color:C.text}}>{nome}</div>
                      <div style={{fontSize:11,color:C.textLight}}>{d.gastos} gasto(s)</div>
                    </div>
                  </div>
                  <div style={{textAlign:"right"}}>
                    <div style={{fontSize:14,fontWeight:800,color:C.primary,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(d.totalUSD,2)}</div>
                    <div style={{display:"flex",gap:6,justifyContent:"flex-end",marginTop:3}}>
                      {d.pago>0&&<span style={{fontSize:11,color:C.success,fontWeight:600}}>✓ {fmtUSD(d.pago,2)}</span>}
                      {d.pendente>0&&<span style={{fontSize:11,color:C.danger,fontWeight:600}}>⏳ {fmtUSD(d.pendente,2)}</span>}
                    </div>
                  </div>
                </div>
              ))}
              {/* Total a receber */}
              {totalAReceberUSD>0&&(
                <div style={{marginTop:12,background:C.warningLight,border:`1px solid ${C.warning}33`,borderRadius:10,padding:"10px 14px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <span style={{fontSize:13,fontWeight:600,color:C.warning}}>⏳ Total a receber</span>
                  <span style={{fontSize:15,fontWeight:800,color:C.warning,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(totalAReceberUSD,2)}</span>
                </div>
              )}
            </div>
          )}

          {/* Tabela completa de gastos */}
          <div style={S.card}>
            <div style={{fontWeight:700,fontSize:14,color:C.text,marginBottom:12}}>Todos os gastos</div>
            <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr",gap:4,marginBottom:8}}>
              {["Descrição","USD total","Meu USD","Divisão"].map(h=><div key={h} style={{fontSize:10,fontWeight:700,color:C.textLight,textTransform:"uppercase",letterSpacing:"0.3px"}}>{h}</div>)}
            </div>
            {gastos.length===0&&<div style={{fontSize:13,color:C.textLight,textAlign:"center",padding:"16px 0"}}>Nenhum gasto ainda</div>}
            {gastos.map(g=>{
              const usd=parseFloat(g.usd)||0;
              const minha=calcMinhaParteUSD(g);
              const nP=1+(g.divisao?.length||0);
              return(
                <div key={g.id} style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr",gap:4,padding:"9px 0",borderTop:`1px solid ${C.borderLight}`,alignItems:"center"}}>
                  <div>
                    <div style={{fontSize:12,fontWeight:600,color:C.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{g.descricao}</div>
                    <div style={{fontSize:10,color:C.textLight}}>{g.loja||g.categoria}</div>
                  </div>
                  <div style={{fontSize:12,color:C.textMid,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(usd,2)}</div>
                  <div style={{fontSize:12,color:C.primary,fontFamily:"'DM Mono',monospace",fontWeight:700}}>{fmtUSD(minha,2)}</div>
                  <div style={{fontSize:11,color:C.textMid}}>{nP>1?`÷${nP}p`:"-"}</div>
                </div>
              );
            })}
            {gastos.length>0&&(
              <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr",gap:4,padding:"10px 0",borderTop:`2px solid ${C.border}`,marginTop:4}}>
                <div style={{fontSize:12,fontWeight:700,color:C.text}}>Total</div>
                <div style={{fontSize:12,fontWeight:700,color:C.textMid,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(totalGastosUSD,2)}</div>
                <div style={{fontSize:12,fontWeight:700,color:C.primary,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(totalMeuUSD,2)}</div>
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
          {rate&&<span style={{fontSize:14,fontWeight:800,color:C.success,fontFamily:"'DM Mono',monospace"}}>{fmtBRL(rate,4)}</span>}
          <button onClick={fetch_} style={{background:C.primaryLight,border:`1px solid ${C.primary}33`,borderRadius:8,padding:"5px 10px",fontSize:12,fontWeight:600,color:C.primary,cursor:"pointer"}}>
            {loading?"...":"↻ BCB"}
          </button>
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
        {usdN>0&&[{label:"USD c/ taxa",value:`${fmtUSD(usdF,2)}`,color:C.textMid},{label:"Dólar ajustado",value:`${fmtBRL(dolarAj,4)}`,color:C.textMid},{label:"Valor em BRL",value:`${fmtBRL(brlP,2)}`,color:C.primary}].map(({label,value,color})=>(
          <div key={label} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${C.borderLight}`}}>
            <span style={{fontSize:13,color:C.textMid}}>{label}</span>
            <span style={{fontSize:13,fontWeight:700,color,fontFamily:"'DM Mono',monospace"}}>{value}</span>
          </div>
        ))}
        <div style={{height:14}}/>
        <label style={S.label}>Dólar que você pagou (opcional)</label>
        <input style={S.input} type="number" step="0.01" placeholder="Ex: 5.71" value={dc} onChange={e=>setDc(e.target.value)}/>
        {brlC&&usdN>0&&[{label:"Com seu dólar",value:`${fmtBRL(brlC,2)}`,color:C.success},{label:"Diferença",value:`${fmtBRL((brlP-brlC),2)}`,color:brlP>brlC?C.success:C.danger}].map(({label,value,color})=>(
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
        {lbs&&[[`Gramas`,`${(parseFloat(lbs)*453.592).toLocaleString("pt-BR",{minimumFractionDigits:1,maximumFractionDigits:1})}g`],[`Kg`,`${(parseFloat(lbs)*453.592/1000).toLocaleString("pt-BR",{minimumFractionDigits:3,maximumFractionDigits:3})}kg`],["Oz",`${(parseFloat(lbs)*16).toLocaleString("pt-BR",{minimumFractionDigits:1,maximumFractionDigits:1})} oz`]].map(([l,v])=>(
          <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:`1px solid ${C.borderLight}`}}><span style={{fontSize:13,color:C.textMid}}>{l}</span><span style={{fontSize:13,fontWeight:700,color:C.text,fontFamily:"'DM Mono',monospace"}}>{v}</span></div>
        ))}
        <div style={{height:12}}/>
        <label style={S.label}>Onças (oz)</label>
        <input style={S.input} type="number" placeholder="Ex: 3.4" value={oz} onChange={e=>setOz(e.target.value)}/>
        {oz&&[["Gramas",`${(parseFloat(oz)*28.3495).toLocaleString("pt-BR",{minimumFractionDigits:1,maximumFractionDigits:1})}g`],["Kg",`${(parseFloat(oz)*28.3495/1000).toLocaleString("pt-BR",{minimumFractionDigits:3,maximumFractionDigits:3})}kg`],["Libras",`${(parseFloat(oz)/16).toLocaleString("pt-BR",{minimumFractionDigits:3,maximumFractionDigits:3})} lbs`]].map(([l,v])=>(
          <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:`1px solid ${C.borderLight}`}}><span style={{fontSize:13,color:C.textMid}}>{l}</span><span style={{fontSize:13,fontWeight:700,color:C.text,fontFamily:"'DM Mono',monospace"}}>{v}</span></div>
        ))}
      </div>
      <div style={S.card}>
        <div style={{fontWeight:700,fontSize:13,color:C.text,marginBottom:10}}>Taxas e cotações</div>
        {[["Dólar pago",`${fmtBRL(settings.dollarPago,4)}`],["IOF",`${settings.iof}%`],["Spread",`${settings.spread}%`],["Taxa compra",`${settings.taxa}%`],["Dólar ajustado",`${fmtBRL(dolarAj,4)}`]].map(([l,v])=>(
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
        const parcelasSheet=sheets.find(s=>s.toLowerCase().includes("parcela"));
        const compras=comprasSheet?parseSheet(wb,comprasSheet):[];
        const legais=legaisSheet?parseSheet(wb,legaisSheet):[];
        const parcelas=parcelasSheet?parseParcelasSheet(wb,parcelasSheet):[];
        setPreview({compras,legais,comprasSheet,legaisSheet,parcelas,parcelasSheet}); setLoading(false);
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
          <button style={S.btnPrimary} onClick={()=>onSave(s)}>Salvar configurações</button>
        </>
      )}
      {tab==="import"&&(
        <>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e=>processFile(e.target.files[0])}/>
          <div style={{border:`2px dashed ${loading?C.primary:C.border}`,borderRadius:16,textAlign:"center",padding:"28px 20px",cursor:"pointer",background:loading?C.primaryLight:C.bg}} onClick={()=>fileRef.current.click()} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();processFile(e.dataTransfer.files[0]);}}>
            <div style={{fontSize:36,marginBottom:10}}>{loading?"⏳":"📊"}</div>
            <div style={{fontWeight:700,fontSize:14,color:C.text,marginBottom:4}}>{loading?"Processando...":"Selecionar planilha"}</div>
            <div style={{fontSize:13,color:C.textLight}}>Arraste ou toque · .xlsx</div>
          </div>
          {error&&<div style={{marginTop:10,background:C.dangerLight,border:`1px solid ${C.danger}33`,borderRadius:10,padding:"10px",fontSize:13,color:C.danger}}>{error}</div>}
          {preview&&(
            <div style={{marginTop:14}}>
              <div style={{fontWeight:700,fontSize:13,color:C.text,marginBottom:10}}>Prévia</div>
              {[["Aba Compras",preview.comprasSheet||"—"],["Aba Legais",preview.legaisSheet||"—"],["Aba Parcelas",preview.parcelasSheet||"—"],["Produtos",`${preview.compras.length} itens`],["Legais",`${preview.legais.length} itens`],["Parcelas",`${(preview.parcelas||[]).length} itens`]].map(([l,v])=>(
                <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${C.borderLight}`}}>
                  <span style={{fontSize:13,color:C.textMid}}>{l}</span><span style={{fontSize:13,fontWeight:600,color:C.text}}>{v}</span>
                </div>
              ))}
              {preview.compras.slice(0,3).map((p,i)=>(
                <div key={i} style={{fontSize:12,color:C.textMid,padding:"4px 0",borderBottom:`1px solid ${C.borderLight}`}}><span style={{fontWeight:600,color:C.text}}>{p.nome}</span> · {p.loja} · US${p.usd}</div>
              ))}
              <button style={{...S.btnPrimary,marginTop:12}} onClick={()=>{onImport(preview.compras,preview.legais,preview.parcelas||[]);setPreview(null);}}>✅ Confirmar importação</button>
            </div>
          )}
        </>
      )}
    </Modal>
  );
}

// ─── PRODUTO FORM ─────────────────────────────────────────────────────────────
function ProdutoForm({prod,onSave,onClose}) {
  const isL=prod?._legais===true;
  const empty={nome:"",loja:"Walmart",usd:"",peso:"",tipo:"solido",volume:"",status:"pendente",prioridade:"Média",link:"",imagem:"",dollarPago:"",_legais:isL};
  const [f,setF]=useState(prod?.id?{...prod,usd:prod.usd.toString(),peso:(prod.peso||"").toString(),dollarPago:(prod.dollarPago||"").toString(),_legais:prod._legais||isL}:empty);
  function save(){if(!f.nome||!f.usd)return alert("Preencha nome e USD");onSave({...f,usd:parseFloat(f.usd),peso:parseFloat(f.peso)||0,volume:parseFloat(f.volume)||0,dollarPago:f.dollarPago?parseFloat(f.dollarPago):null});}
  return (
    <Modal title={prod?.id?"Editar produto":isL?"✨ Item legal":"Novo produto"} onClose={onClose}>
      <label style={S.label}>Nome *</label>
      <input style={S.input} placeholder="Ex: AirPods Pro" value={f.nome} onChange={e=>setF(p=>({...p,nome:e.target.value}))}/>
      <label style={S.label}>Loja</label>
      <input style={S.input} list="lojas-list" placeholder="Digite ou escolha uma loja..." value={f.loja} onChange={e=>setF(p=>({...p,loja:e.target.value}))}/>
      <datalist id="lojas-list">{LOJAS_SUGESTOES.map(l=><option key={l} value={l}/>)}</datalist>
      <label style={S.label}>Preço USD *</label>
      <input style={S.input} type="number" placeholder="Ex: 199" value={f.usd} onChange={e=>setF(p=>({...p,usd:e.target.value}))}/>
      <label style={S.label}>Tipo</label>
      <div style={{display:"flex",gap:8,marginBottom:14}}>{[["solido","📦 Sólido"],["liquido","💧 Líquido"]].map(([v,l])=><button key={v} style={{...S.chipSel,flex:1,...(f.tipo===v?S.chipSelActive:{})}} onClick={()=>setF(p=>({...p,tipo:v}))}>{l}</button>)}</div>
      {f.tipo==="solido"?<><label style={S.label}>Peso (gramas)</label><input style={S.input} type="number" placeholder="Ex: 250" value={f.peso} onChange={e=>setF(p=>({...p,peso:e.target.value}))}/>{f.peso&&<div style={{fontSize:12,color:C.textLight,marginTop:-8,marginBottom:12}}>= {((parseFloat(f.peso)||0)/1000).toLocaleString("pt-BR",{minimumFractionDigits:3})} kg</div>}</>:<><label style={S.label}>Volume (oz)</label><input style={S.input} type="number" step="0.1" placeholder="Ex: 3.4" value={f.volume} onChange={e=>setF(p=>({...p,volume:e.target.value}))}/>{f.volume&&<div style={{fontSize:12,color:C.textLight,marginTop:-8,marginBottom:12}}>= {((parseFloat(f.volume)||0)*28.3495/1000).toLocaleString("pt-BR",{minimumFractionDigits:3})} kg</div>}</>}
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
