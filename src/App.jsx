import { useState, useMemo, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { getProductImage, imageCache } from './imageService';
import { db } from "./firebase";
import { doc, onSnapshot, setDoc, serverTimestamp } from "firebase/firestore";
import { getAuth, onAuthStateChanged, signInWithEmailAndPassword, createUserWithEmailAndPassword, signOut, setPersistence, browserLocalPersistence } from "firebase/auth";


// ─── FIREBASE AUTH / FIRESTORE PERSISTENCE ────────────────────────────────
// Cada usuário logado tem sua própria "caixinha" no Firestore.
// Os dados ficam em: usuarios_pwa/{uid}
const auth = getAuth();
// Mantém o login salvo no aparelho, mesmo após fechar o navegador/app (login offline)
setPersistence(auth, browserLocalPersistence).catch(err => console.error("Erro ao definir persistência do Auth:", err));

const normalizeCloudState = data => ({
  settings: data?.settings || INITIAL_SETTINGS,
  produtos: Array.isArray(data?.produtos) ? data.produtos : [],
  itensLegais: Array.isArray(data?.itensLegais) ? data.itensLegais : [],
  gastos: Array.isArray(data?.gastos) ? data.gastos : [],
  parcelas: Array.isArray(data?.parcelas) ? data.parcelas : [],
  anotacoes: typeof data?.anotacoes === "string" ? data.anotacoes : "",
  planejamento: data?.planejamento || { dataInicio:"", dataFim:"", eventos:[] },
  checklist: Array.isArray(data?.checklist) ? data.checklist : [],
  comprasDolar: Array.isArray(data?.comprasDolar) ? data.comprasDolar : [],
});

async function saveCloudState(userDocRef, state, { force=false } = {}) {
  try {
    // TRAVA DE SEGURANÇA: se a lista de produtos estiver vazia, cancela o salvamento
    // para não sobrescrever dados reais com um estado vazio (ex: carregamento ainda em curso).
    // `force` é usado apenas na inicialização de um usuário novo (documento ainda não existe).
    if (!force && (!state.produtos || state.produtos.length === 0)) {
      console.warn("⚠️ [Segurança] Tentativa de salvar lista de produtos vazia abortada. Nuvem protegida.");
      return;
    }
    await setDoc(userDocRef, {
      ...state,
      updatedAt: serverTimestamp(),
    }, { merge: true });
  } catch (error) {
    console.error("Erro ao salvar dados na nuvem:", error);
  }
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
// Taxa de imposto local (Orlando 6.5%, Kissimmee 7.5%, isento 0%)
const taxaLocal = p => p.localTaxa === 'orlando' ? 0.065 : p.localTaxa === 'kissimmee' ? 0.075 : 0;
// USD com imposto local aplicado
const usdComTaxa = p => (parseFloat(p.usd) || 0) * (1 + taxaLocal(p));
// BRL total do produto (qtd × usd × taxaLocal × câmbio)
const calcBRLProduto = (p, s) => usdComTaxa(p) * prodQtd(p) * calcDolarAjustado(s);
// Converte o peso/volume cadastrado para gramas, suportando 3 unidades:
// 'g' (gramas), 'oz_peso' (onça de massa, 28.3495g), 'oz_liquido' (fluid oz, 29.5735g equiv.)
// Mantém compatibilidade com itens antigos (tipo "liquido" + campo volume em oz de massa)
const pesoGramas = p => {
  if (p.tipoPeso) {
    const v = parseFloat(p.peso) || 0;
    if (p.tipoPeso === 'oz_peso') return v * 28.3495;
    if (p.tipoPeso === 'oz_liquido') return v * 29.5735;
    return v; // 'g'
  }
  // Fallback para itens antigos sem tipoPeso
  return p.tipo === "liquido" ? (parseFloat(p.volume)||0)*28.3495 : parseFloat(p.peso)||0;
};
// Quantidade comprada: lê qtdComprada se comprado, senão 0 (para cálculos de gastos)
const prodQtd = p => p.status === "comprado" ? Math.max(1, parseInt(p.qtdComprada) || 1) : 0;
// Quantidade cadastrada no produto (para exibição e peso estimado)
const prodQtdCad = p => Math.max(1, parseInt(p.quantidade) || 1);
// USD total considerando quantidade comprada
const prodUSD = p => (parseFloat(p.usd) || 0) * Math.max(1, parseInt(p.qtdComprada) || (p.status === "comprado" ? 1 : 0));
// Peso total considerando quantidade cadastrada
const prodPeso = p => pesoGramas(p) * prodQtdCad(p);
// USD planeado (usa quantidade cadastrada, independente do status)
const prodUSDPlanejado = p => (parseFloat(p.usd) || 0) * prodQtdCad(p);
// BRL planeado (usa quantidade cadastrada, independente do status)
const calcBRLProdutoPlanejado = (p, s) => usdComTaxa(p) * prodQtdCad(p) * calcDolarAjustado(s);

// ─── FORMATAÇÃO ─────────────────────────────────────────────────────────────
// Formata número com vírgula como separador decimal (padrão pt-BR)
const fmtUSD = (v, dec=2) => `US$ ${Number(v).toLocaleString("pt-BR",{minimumFractionDigits:dec,maximumFractionDigits:dec})}`;
const fmtBRL = (v, dec=2) => `R$ ${Number(v).toLocaleString("pt-BR",{minimumFractionDigits:dec,maximumFractionDigits:dec})}`;
const fmtN   = (v, dec=2) => Number(v).toLocaleString("pt-BR",{minimumFractionDigits:dec,maximumFractionDigits:dec});

// tudo em USD — dolarPago = cotação usada na compra
function calcMinhaParteUSD(gasto, produtosArr) {
  // Para gastos de produto: usd é unitário, multiplicar por qtd e taxa local
  // Se usd=0 (gasto antigo), tentar recuperar do array de produtos
  let usdUnit = parseFloat(gasto.usd) || 0;
  if (usdUnit === 0 && gasto.produtoId && produtosArr) {
    const pai = produtosArr.find(p => p.id === gasto.produtoId);
    usdUnit = parseFloat(pai?.usd) || 0;
  }
  const qtd = parseInt(gasto.qtdComprada) || 1;
  const taxa = gasto.localTaxa === "orlando" ? 0.065 : gasto.localTaxa === "kissimmee" ? 0.075 : 0;
  const totalUSD = gasto.tipo === "produto" ? usdUnit * (1 + taxa) * qtd : usdUnit;
  if (!gasto.divisao || gasto.divisao.length === 0) return totalUSD;
  const somaDivisao = gasto.divisao.reduce((acc, p) => acc + (parseFloat(p.valor) || 0), 0);
  return Math.max(0, totalUSD - somaDivisao);
}
function usdToBRL(usd, gasto, settings) {
  // Se o gasto tem cotação específica registrada, usar ela (já inclui o custo real pago)
  // Se não, usar o dólar ajustado com IOF+spread para refletir o custo real
  const cotacao = parseFloat(gasto.dolarPago) || calcDolarAjustado(settings);
  return usd * cotacao;
}
function calcTotalGastosUSD(gastos, produtosArr) {
  return gastos.reduce((a, g) => a + calcMinhaParteUSD(g, produtosArr), 0);
}

// ─── AWESOMEAPI COTAÇÃO ──────────────────────────────────────────────────────
async function fetchCotacao() {
  const YAHOO = "https://query1.finance.yahoo.com/v8/finance/chart/USDBRL=X?interval=1m&range=1d";
  // 1. Tentar Yahoo Finance via dois proxies CORS
  for (const proxy of [
    `https://corsproxy.io/?${encodeURIComponent(YAHOO)}`,
    `https://api.allorigins.win/raw?url=${encodeURIComponent(YAHOO)}`,
  ]) {
    try {
      const res = await fetch(proxy, { signal: AbortSignal.timeout(4000) });
      const d   = await res.json();
      const meta = d?.chart?.result?.[0]?.meta;
      const bid  = parseFloat(meta?.regularMarketPrice);
      const prev = parseFloat(meta?.chartPreviousClose);
      if (!isNaN(bid) && bid > 1) {
        const pct = !isNaN(prev) ? parseFloat(((bid - prev) / prev * 100).toFixed(2)) : null;
        return { bid, pct };
      }
    } catch {}
  }
  // 2. Fallback: AwesomeAPI
  try {
    const res = await fetch("https://economia.awesomeapi.com.br/last/USD-BRL", { signal: AbortSignal.timeout(5000) });
    const d   = await res.json();
    const bid = parseFloat(d?.USDBRL?.bid);
    const pct = parseFloat(d?.USDBRL?.pctChange);
    if (!isNaN(bid)) return { bid, pct: !isNaN(pct) ? pct : null };
  } catch {}
  return null;
}

const LOJA_EMOJI = { Amazon:"📦","Best Buy":"🔵",Walmart:"🟡",Target:"🎯",Newegg:"💻",Apple:"🍎",Costco:"🏪",Basspro:"🎣",HomeGoods:"🏠","Dollar Tree":"🌳",Marshalls:"🏷",Ross:"🏷","TJ Maxx":"🏷","Tommy Hilfiger":"👔","Calvin Klein":"👔","The North Face":"🏔",Sephora:"💄",Ulta:"💄",Restaurante:"🍔",Uber:"🚗",Passeio:"🎢",Outro:"🛒" };

function ProductImage({ produto, style={}, iconSize=44, fit="cover" }) {
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
        style={{ width: "100%", height: "100%", objectFit: fit }}
        onError={() => {
          // Se a cópia offline (blob) falhar, tenta a URL direta cadastrada
          if (imgUrl !== produto.imagem && produto.imagem) {
            setImgUrl(produto.imagem);
          } else {
            setErr(true);
          }
        }}
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
    if (!navigator.onLine) { setErro("📡 Sem internet. Conecte-se à rede para fazer login."); return; }
    setLoading(true); setErro("");
    try {
      await signInWithEmailAndPassword(auth, email.trim(), senha);
    } catch (e) {
      if (e?.code === "auth/network-request-failed" || e?.message?.includes("network")) {
        setErro("📡 Falha na rede. Verifique sua conexão para fazer login.");
      } else {
        setErro(traduzErroAuth(e));
      }
    } finally {
      setLoading(false);
    }
  }

  async function criarConta() {
    if (!email || !senha) { setErro("Informe e-mail e senha."); return; }
    if (senha.length < 6) { setErro("A senha precisa ter pelo menos 6 caracteres."); return; }
    if (!navigator.onLine) { setErro("📡 Sem internet. Conecte-se à rede para criar a conta."); return; }
    setLoading(true); setErro("");
    try {
      await createUserWithEmailAndPassword(auth, email.trim(), senha);
    } catch (e) {
      if (e?.code === "auth/network-request-failed" || e?.message?.includes("network")) {
        setErro("📡 Falha na rede. Verifique sua conexão para criar a conta.");
      } else {
        setErro(traduzErroAuth(e));
      }
    } finally {
      setLoading(false);
    }
  }

  return (
    <div style={{minHeight:"100vh",background:`linear-gradient(135deg, ${C.gradientA}, ${C.gradientB})`,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{width:"100%",maxWidth:390,background:C.bgCard,borderRadius:24,padding:26,boxShadow:"0 24px 70px rgba(15,23,42,0.25)",border:"1px solid rgba(255,255,255,0.35)"}}>
        <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:22}}>
          <img src="/image_a375cf.png" alt="TravelShop" style={{width:48,height:48,borderRadius:16,objectFit:"cover"}}/>
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
  const [planejamento, setPlanejamento] = useState({ dataInicio:"", dataFim:"", eventos:[] });
  const [checklist, setChecklist] = useState([]);
  const [comprasDolar, setComprasDolar] = useState([]);
  const [anotacoes, setAnotacoes] = useState(`WALMART:
Mais afastado: 3101 W Princeton St Orlando, FL 32808, EUA
Com itens da Disney: 8990 Turkey Lak RD, Orlando FL 32819, EUA

TARGET:
4750 Millenia Plazza Way, Orlando, FL 32839, EUA

TJ MAXX:
5748 Hamlin Groves Trail Winter Garden, FL 34787, EUA

DOLLAR TREE:
13605 SApopka Vineland Rd Ste 103A, Orlando, FL 32821`);
  const [showSettings, setShowSettings] = useState(false);
  const [showForm, setShowForm] = useState(false);
  const [showGastoForm, setShowGastoForm] = useState(false);
  const [prodSubTab, setProdSubTab] = useState("compras");
  const [calcSubTab, setCalcSubTab] = useState("conversor");
  const [editProd, setEditProd] = useState(null);
  const [editGasto, setEditGasto] = useState(null);
  const [notification, setNotification] = useState(null);
  const [cotacaoLoading, setCotacaoLoading] = useState(false);
  const [user, setUser] = useState(null);
  const [authReady, setAuthReady] = useState(false);
  const [cloudReady, setCloudReady] = useState(false);
  const skipNextCloudSave = useRef(false);
  const skipNextSnapshot = useRef(false);
  const userDocRef = useMemo(() => user ? doc(db, "usuarios_pwa", user.uid) : null, [user]);

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
          planejamento: { dataInicio:"", dataFim:"", eventos:[] },
          checklist: [],
          comprasDolar: [],
        }, { force: true });
        setSettings(INITIAL_SETTINGS);
        setProdutos(SAMPLE_PRODUTOS);
        setItensLegais([]);
        setGastos([]);
        setParcelas([]);
        setPlanejamento({ dataInicio:"", dataFim:"", eventos:[] });
        setChecklist([]);
        setComprasDolar([]);
        setCloudReady(true);
        return;
      }

      const cloudState = normalizeCloudState(snap.data());
      // Se foi o próprio app que salvou, não sobrescrever o state local
      if (skipNextSnapshot.current) {
        skipNextSnapshot.current = false;
        setCloudReady(true);
        return;
      }
      skipNextCloudSave.current = true;
      setSettings(cloudState.settings);
      setProdutos(cloudState.produtos);
      setItensLegais(cloudState.itensLegais);
      setGastos(cloudState.gastos);
      setParcelas(cloudState.parcelas || []);
      setPlanejamento(cloudState.planejamento || { dataInicio:"", dataFim:"", eventos:[] });
      setChecklist(cloudState.checklist || []);
      setComprasDolar(cloudState.comprasDolar || []);
      setAnotacoes(cloudState.anotacoes || 'WALMART:\n\nMais afastado: 3101 W Princeton St, Orlando, FL 32808, EUA\n\nCom itens da Disney: 8990 Turkey Lake Rd, Orlando FL 32819, EUA\n\nTARGET:\n\n4750 Millenia Plaza Way, Orlando, FL 32839, EUA\n\nTJ MAXX:\n\n5748 Hamlin Groves Trail, Winter Garden, FL 34787, EUA\n\nDOLLAR TREE:\n\n13605 S Apopka Vineland Rd Ste 103A, Orlando, FL 32821');
      setCloudReady(true);
    }, (error) => {
      console.error("Erro ao sincronizar com Firestore:", error);
      notify("Erro ao sincronizar com a nuvem", "error");
      setCloudReady(true);
    });

    return () => unsubscribe();
  }, [userDocRef, user]);

  // Salva alterações locais direto no Firestore. Não usa mais LocalStorage nem ID pela URL.
  useEffect(() => {
    if (!userDocRef || !cloudReady) return;
    if (skipNextCloudSave.current) {
      skipNextCloudSave.current = false;
      return;
    }

    const timer = setTimeout(() => {
      skipNextSnapshot.current = true;
      saveCloudState(userDocRef, { settings, produtos, itensLegais, gastos, parcelas, planejamento, checklist, comprasDolar, anotacoes })
        .catch((error) => {
          console.error("Erro ao salvar no Firestore:", error);
          notify("Erro ao salvar na nuvem", "error");
          skipNextSnapshot.current = false;
        });
    }, 350);

    return () => clearTimeout(timer);
  }, [settings, produtos, itensLegais, gastos, parcelas, planejamento, checklist, comprasDolar, anotacoes, cloudReady, userDocRef]);

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
      setPlanejamento({ dataInicio:"", dataFim:"", eventos:[] });
      setChecklist([]);
      setComprasDolar([]);
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
      const novaQtd = newStatus === "comprado" ? 1 : 0;
      if (newStatus === "comprado") {
        setGastos(gs => {
          if (gs.some(g => g.produtoId === id)) return gs;
          return [...gs, {
            id: `prod_${id}`, produtoId: id, descricao: p.nome, loja: p.loja || "Não especificada",
            usd: parseFloat(p.usd) || 0,
            qtdComprada: novaQtd,
            localTaxa: "isento",
            dolarPago: p.dollarPago || settings.dollarPago,
            brl: null, imagem: p.imagem || "",
            categoria: p.categoria || "🛍 Compras", divisao: [], data: new Date().toLocaleDateString("pt-BR"), tipo: "produto"
          }];
        });
      } else {
        setGastos(gs => gs.filter(g => g.produtoId !== id));
      }
      return {...p, status: newStatus, qtdComprada: novaQtd, localTaxa: p.localTaxa || "isento"};
    }));
  }

  function toggleStatusItemLegal(id) {
    setItensLegais(ps => ps.map(p => {
      if (p.id !== id) return p;
      const newStatus = p.status === "comprado" ? "pendente" : "comprado";
      const novaQtd = newStatus === "comprado" ? (parseInt(p.qtdComprada) || 1) : 0;
      if (newStatus === "comprado") {
        setGastos(gs => {
          if (gs.some(g => g.produtoId === id)) return gs;
          return [...gs, {
            id: `legal_${id}`, produtoId: id, descricao: p.nome, loja: p.loja || "Não especificada",
            usd: parseFloat(p.usd) || 0,
            qtdComprada: novaQtd,
            localTaxa: p.localTaxa || "isento",
            dolarPago: p.dollarPago || settings.dollarPago,
            brl: null, imagem: p.imagem || "",
            categoria: "⚖️ Itens Legais", divisao: [], data: new Date().toLocaleDateString("pt-BR"), tipo: "produto"
          }];
        });
      } else {
        setGastos(gs => gs.filter(g => g.produtoId !== id));
      }
      return {...p, status: newStatus, qtdComprada: novaQtd, localTaxa: p.localTaxa || "isento"};
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

  function updateProduto(id, campos) {
    setProdutos(ps => {
      const novosProdutos = ps.map(p => p.id === id ? {...p, ...campos} : p);
      const prod = novosProdutos.find(p => p.id === id);
      if (!prod) return novosProdutos;

      if ('qtdComprada' in campos) {
        const novaQtd = campos.qtdComprada;
        if (novaQtd === 0) {
          setGastos(gs => gs.filter(g => g.produtoId !== id));
        } else {
          // Se não estava comprado, adicionar como gasto
          setGastos(gs => {
            if (!gs.some(g => g.produtoId === id)) {
              return [...gs, {
                id: `prod_${id}`, produtoId: id, descricao: prod.nome, loja: prod.loja,
                usd: parseFloat(prod.usd) || 0,
                qtdComprada: novaQtd,
                localTaxa: prod.localTaxa || "isento",
                dolarPago: prod.dollarPago || settings.dollarPago,
                brl: null, imagem: prod.imagem || "",
                categoria: "🛍 Compras", divisao: [], data: new Date().toLocaleDateString("pt-BR"), tipo: "produto"
              }];
            }
            return gs.map(g => g.produtoId === id ? {...g, qtdComprada: novaQtd} : g);
          });
        }
      }
      if ('localTaxa' in campos) {
        setGastos(gs => gs.map(g => g.produtoId === id ? {...g, localTaxa: campos.localTaxa} : g));
      }
      return novosProdutos;
    });
  }

  function updateItemLegal(id, campos) {
    setItensLegais(ps => {
      const novosItens = ps.map(p => p.id === id ? {...p, ...campos} : p);
      const item = novosItens.find(p => p.id === id);
      if (!item) return novosItens;

      if ('qtdComprada' in campos) {
        const novaQtd = campos.qtdComprada;
        if (novaQtd === 0) {
          setGastos(gs => gs.filter(g => g.produtoId !== id));
        } else {
          setGastos(gs => {
            if (!gs.some(g => g.produtoId === id)) {
              return [...gs, {
                id: `legal_${id}`, produtoId: id, descricao: item.nome, loja: item.loja,
                usd: parseFloat(item.usd) || 0,
                qtdComprada: novaQtd,
                localTaxa: item.localTaxa || "isento",
                dolarPago: item.dollarPago || settings.dollarPago,
                brl: null, imagem: item.imagem || "",
                categoria: "⚖️ Itens Legais", divisao: [], data: new Date().toLocaleDateString("pt-BR"), tipo: "produto"
              }];
            }
            return gs.map(g => g.produtoId === id ? {...g, qtdComprada: novaQtd} : g);
          });
        }
      }
      if ('localTaxa' in campos) {
        setGastos(gs => gs.map(g => g.produtoId === id ? {...g, localTaxa: campos.localTaxa} : g));
      }
      return novosItens;
    });
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

  function exportarParaExcel() {
    const workbook = XLSX.utils.book_new();
    let possuiDados = false;

    // 1. COMPRAS
    if (produtos && produtos.length > 0) {
      const dados = produtos.map(p => ({
        "Nome": p.nome || "",
        "Loja": p.loja || "",
        "Preço USD": p.usd || 0,
        "Quantidade Planejada": p.quantidade || 1,
        "Categoria": p.categoria || "",
        "Prioridade": p.prioridade || "",
        "Peso Gramas": p.pesoGramas || 0,
        "Status": p.status || "pendente",
        "Local Taxa": p.localTaxa || "isento",
        "Qtd Comprada": p.qtdComprada || 0,
        "Link/Ref": p.link || ""
      }));
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(dados), "Compras");
      possuiDados = true;
    }

    // 2. LEGAIS
    if (itensLegais && itensLegais.length > 0) {
      const dados = itensLegais.map(p => ({
        "Nome": p.nome || "",
        "Loja": p.loja || "",
        "Preço USD": p.usd || 0,
        "Quantidade Planejada": p.quantidade || 1,
        "Peso Gramas": p.pesoGramas || 0,
        "Status": p.status || "pendente",
        "Local Taxa": p.localTaxa || "isento",
        "Qtd Comprada": p.qtdComprada || 0,
        "Link/Ref": p.link || ""
      }));
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(dados), "Legais");
      possuiDados = true;
    }

    // 3. GASTOS REAIS
    if (gastos && gastos.length > 0) {
      const dados = gastos.map(g => ({
        "Descrição": g.descricao || "",
        "Loja": g.loja || "",
        "Valor USD": g.usd || 0,
        "Qtd Comprada": g.qtdComprada || 1,
        "Taxa Local": g.localTaxa || "isento",
        "Categoria": g.categoria || "",
        "Data": g.data || "",
        "Tipo Registro": g.tipo || "avulso"
      }));
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(dados), "Gastos Reais");
      possuiDados = true;
    }

    // 4. PARCELAS
    if (parcelas && parcelas.length > 0) {
      const dados = parcelas.map(p => ({
        "Descrição": p.descricao || "",
        "Valor Total BRL": p.valorTotal || 0,
        "Qtd Parcelas": p.quantidadeParcelas || 1,
        "Valor Parcela": p.valorParcela || 0,
        "Minha Parte": p.minhaParte || "",
        "Primeira Fatura": p.primeiraFatura || "",
        "Cartão": p.cartao || ""
      }));
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(dados), "Parcelas");
      possuiDados = true;
    }

    // 5. CHECKLIST
    if (checklist && checklist.length > 0) {
      const dados = checklist.map(c => ({
        "Tarefa/Item": c.texto || "",
        "Concluído": c.feito ? "Sim" : "Não",
        "Categoria": c.cat || "Geral"
      }));
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(dados), "Checklist");
      possuiDados = true;
    }

    // 6. CÂMBIO E DÓLAR
    if (comprasDolar && comprasDolar.length > 0) {
      const dados = comprasDolar.map(d => ({
        "Data da Compra": d.data || "",
        "Valor USD": d.quantidade || 0,
        "Cotação Paga (BRL)": d.cotacao || 0,
        "Observação": d.obs || ""
      }));
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(dados), "Câmbio e Dólar");
      possuiDados = true;
    }

    // 7. RESUMO E NOTAS
    const infoGeral = [];
    if (anotacoes && anotacoes.trim() !== "") {
      infoGeral.push({ "Bloco de Informação": "Minhas Anotações de Viagem", "Conteúdo / Detalhes": anotacoes });
    }
    if (planejamento) {
      infoGeral.push({ "Bloco de Informação": "Data de Início da Viagem", "Conteúdo / Detalhes": planejamento.dataInicio || "Não informada" });
      infoGeral.push({ "Bloco de Informação": "Data de Término da Viagem", "Conteúdo / Detalhes": planejamento.dataFim || "Não informada" });
      if (planejamento.eventos && planejamento.eventos.length > 0) {
        infoGeral.push({ "Bloco de Informação": "Total de Eventos no Roteiro", "Conteúdo / Detalhes": `${planejamento.eventos.length} evento(s) cadastrado(s)` });
      }
    }
    if (settings) {
      infoGeral.push({ "Bloco de Informação": "Dólar Pago (R$)", "Conteúdo / Detalhes": settings.dollarPago || 0 });
      infoGeral.push({ "Bloco de Informação": "Total de Dólares da Viagem (US$)", "Conteúdo / Detalhes": settings.totalDolarViagem || 0 });
    }
    if (infoGeral.length > 0) {
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(infoGeral), "Resumo e Notas");
      possuiDados = true;
    }

    if (!possuiDados) {
      notify("Não há dados cadastrados para exportar.","error");
      return;
    }

    const dataHoje = new Date().toLocaleDateString('pt-BR').replace(/\//g,'-');
    XLSX.writeFile(workbook, `Backup_Completo_TravelShop_${dataHoje}.xlsx`);
    notify("Planilha exportada!");
  }


  const stats = useMemo(()=>{
    const comprados=produtos.filter(p=>p.status==="comprado");
    const pesoTotal=produtos.reduce((a,p)=>a+(prodPeso(p)||0),0);
    const valorTotalUSD=produtos.reduce((a,p)=>a+(prodUSDPlanejado(p)||0),0);
    const valorTotalBRL=produtos.reduce((a,p)=>a+(calcBRLProdutoPlanejado(p,settings)||0),0);
    const valorGasto=comprados.reduce((a,p)=>{
      const v=p.dollarPago?calcBRLPago(p.usd,settings,p.dollarPago):calcBRL(p.usd,settings);
      return a+(isNaN(v)?0:v);
    },0);
    let totalMeusGastosUSD=0;
    if(Array.isArray(gastos)) gastos.forEach(g=>{
      // Fallback para produto pai se usd=0 (gastos legados criados com bug)
      const pai=g.produtoId?produtos.find(p=>p.id===g.produtoId):null;
      const uUnit=(parseFloat(g.usd)||0)||(parseFloat(pai?.usd)||0);
      let qtd=1;
      if(g.tipo==="produto"){
        qtd=parseInt(g.qtdComprada);
        if(isNaN(qtd)||qtd<=0) qtd=1;
      }
      const taxa=g.localTaxa==="orlando"?0.065:g.localTaxa==="kissimmee"?0.075:0;
      let subtotal;
      if(g.divisao&&g.divisao.length>0){
        const soma=g.divisao.reduce((a,p)=>a+(parseFloat(p.valor)||0),0);
        subtotal=Math.max(0,uUnit*(1+taxa)*qtd-soma);
      } else {
        subtotal=uUnit*(1+taxa)*qtd;
      }
      totalMeusGastosUSD+=isNaN(subtotal)?0:subtotal;
    });
    return {
      total:produtos.length,
      comprados:comprados.length,
      pendentes:produtos.length-comprados.length,
      pesoTotal:isNaN(pesoTotal)?0:pesoTotal,
      valorTotalUSD:isNaN(valorTotalUSD)?0:valorTotalUSD,
      valorTotalBRL:isNaN(valorTotalBRL)?0:valorTotalBRL,
      valorGasto:isNaN(valorGasto)?0:valorGasto,
      lojas:new Set(produtos.map(p=>p.loja||"Não especificada")).size,
      totalMeusGastosUSD:isNaN(totalMeusGastosUSD)?0:totalMeusGastosUSD
    };
  },[produtos,settings,gastos]);

  const pesoPercent=Math.min(100,(stats.pesoTotal/settings.pesoMax)*100);
  const pesoColor=pesoPercent<70?C.success:pesoPercent<90?C.warning:C.danger;
  const pesoBg=pesoPercent<70?C.successLight:pesoPercent<90?C.warningLight:C.dangerLight;

  const TABS=["Início","Produtos","Galeria","Gastos","Parcelas","Roteiro","Stats","Dólar","Calc"];


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
        <div style={S.headerLeft}>
          <img src="/image_a375cf.png" alt="TravelShop" style={{width:36,height:36,borderRadius:10,objectFit:"cover"}}/>
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
        {tab===0&&<DashboardTab stats={stats} settings={settings} pesoPercent={pesoPercent} pesoColor={pesoColor} pesoBg={pesoBg} onTabChange={setTab} anotacoes={anotacoes} setAnotacoes={setAnotacoes}/>}
        {tab===1&&<ProdutosTab produtos={produtos} itensLegais={itensLegais} settings={settings} onToggle={toggleStatus} onDelete={deleteProd} onEdit={p=>{setEditProd(p);setShowForm(true);}} onAdd={()=>{setEditProd(null);setShowForm(true);}} onMoveToList={moveToList} onSubTabChange={setProdSubTab} onUpdate={updateProduto}/>}
        {tab===2&&<GaleriaTab produtos={produtos} itensLegais={itensLegais} settings={settings} onEdit={p=>{setEditProd(p);setShowForm(true);}} onToggle={toggleStatus} onToggleLegal={toggleStatusItemLegal} onUpdate={updateProduto} onUpdateLegal={updateItemLegal}/>}
        {tab===3&&<GastosTab gastos={gastos} settings={settings} onAdd={()=>{setEditGasto(null);setShowGastoForm(true);}} onEdit={g=>{setEditGasto(g);setShowGastoForm(true);}} onDelete={id=>{ setGastos(gs=>gs.filter(g=>g.id!==id)); notify("Removido","error"); }} onTogglePago={(gastoId,pessoaIdx)=>setGastos(gs=>gs.map(g=>g.id===gastoId?{...g,divisao:g.divisao.map((p,i)=>i===pessoaIdx?{...p,pago:!p.pago}:p)}:g))} produtos={produtos} onToggleStatus={toggleStatus} parcelas={parcelas}/>}
        {tab===4&&<ParcelasTab parcelas={parcelas} setParcelas={setParcelas}/>}
        {tab===5&&<RoteiroTab planejamento={planejamento} setPlanejamento={setPlanejamento}/>}
        {tab===6&&<StatsTab produtos={produtos} gastos={gastos} settings={settings} checklist={checklist} setChecklist={setChecklist}/>}
        {tab===7&&<HistoricoDolarTab comprasDolar={comprasDolar} setComprasDolar={setComprasDolar} settings={settings}/>}
        {tab===8&&<CalcTab settings={settings} gastos={gastos} produtos={produtos} parcelas={parcelas} comprasDolar={comprasDolar} setComprasDolar={setComprasDolar} checklist={checklist} setChecklist={setChecklist} initialSubTab={calcSubTab} onSubTabChange={setCalcSubTab}/>}
        {tab===9&&<div style={S.page}><button onClick={()=>setTab(0)} style={{...S.btnOutline,marginBottom:14,display:"flex",alignItems:"center",gap:6}}><span>←</span> Voltar</button><BagagemTab produtos={produtos} settings={settings}/></div>}
      </div>

      <nav style={S.nav}>
        {TABS.map((label,i)=>(
          <button key={i} style={{...S.navBtn,...(tab===i?S.navBtnActive:{})}} onClick={()=>setTab(i)}>
            <NavIcon name={label} active={tab===i}/>
            <span style={S.navLabel}>{label}</span>
            {tab===i&&<div style={S.navIndicator}/>}
          </button>
        ))}
      </nav>

      {tab===1&&<button style={S.fab} onClick={()=>{setEditProd({_legais:prodSubTab==="legais"});setShowForm(true);}}><span style={{fontSize:22,color:"white"}}>＋</span></button>}
      {tab===3&&<button style={S.fab} onClick={()=>{setEditGasto(null);setShowGastoForm(true);}}><span style={{fontSize:22,color:"white"}}>＋</span></button>}

      {showSettings&&<SettingsModal settings={settings} onSave={s=>{setSettings(s);notify("Configurações salvas!");}} onImport={handleImport} onExport={exportarParaExcel} onClose={()=>setShowSettings(false)}/>}
      {showForm&&<ProdutoForm prod={editProd} onSave={saveProd} onClose={()=>{setShowForm(false);setEditProd(null);}}/>}
      {showGastoForm&&<GastoForm gasto={editGasto} settings={settings} onSave={saveGasto} onClose={()=>{setShowGastoForm(false);setEditGasto(null);}}/>}
    </div>
  );
}


// ─── COTAÇÃO BCB CARD (Dashboard) ────────────────────────────────────────────
function CotacaoBcbCard({settings}) {
  const [rate, setRate] = useState(null);
  const [loading, setLoading] = useState(false);
  const [lastFetch, setLastFetch] = useState(null);
  const [variacao, setVariacao] = useState(null);

  async function buscarCotacao() {
    setLoading(true);
    const result = await fetchCotacao();
    if (result?.bid) {
      setRate(result.bid);
      setVariacao(result.pct);
      setLastFetch(new Date().toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"}));
    }
    setLoading(false);
  }

  useEffect(() => { buscarCotacao(); }, []);

  const comIOF = rate ? rate * (1 + (settings.iof + settings.spread) / 100) : null;
  const varPos = variacao >= 0;

  return (
    <div style={{...S.card,marginBottom:10,background:"linear-gradient(135deg,#F0FDF4,#ECFDF5)",border:`1px solid ${C.success}33`,padding:"12px 14px"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
        <div>
          <div style={{fontSize:11,fontWeight:700,color:C.success,textTransform:"uppercase",letterSpacing:"0.5px",marginBottom:4}}>💱 Dólar hoje (BCB)</div>
          {loading && <div style={{fontSize:13,color:C.textLight}}>Buscando...</div>}
          {!loading && rate && (
            <>
              <div style={{display:"flex",alignItems:"baseline",gap:6}}>
                <span style={{fontSize:22,fontWeight:800,color:C.text,fontFamily:"'DM Mono',monospace"}}>{fmtBRL(rate,4)}</span>
                {variacao !== null && (
                  <span style={{fontSize:12,fontWeight:700,color:varPos?C.danger:C.success}}>
                    {varPos?"▲":"▼"} {Math.abs(variacao).toFixed(2)}%
                  </span>
                )}
              </div>
              <div style={{fontSize:11,color:C.textLight,marginTop:1}}>mercado às {lastFetch}</div>
            </>
          )}
          {!loading && !rate && <div style={{fontSize:12,color:C.textLight}}>Toque para buscar</div>}
        </div>
        <div style={{textAlign:"right"}}>
          {comIOF && (
            <div>
              <div style={{fontSize:11,color:C.textMid,marginBottom:2}}>c/ IOF + spread</div>
              <div style={{fontSize:16,fontWeight:800,color:C.warning,fontFamily:"'DM Mono',monospace"}}>{fmtBRL(comIOF,4)}</div>
              <div style={{fontSize:10,color:C.textLight}}>{settings.iof}% IOF + {settings.spread}% spread</div>
            </div>
          )}
          <button onClick={buscarCotacao} style={{background:C.successLight,border:`1px solid ${C.success}44`,borderRadius:8,padding:"5px 10px",fontSize:11,fontWeight:700,color:C.success,cursor:"pointer",marginTop:6}}>
            {loading?"...":"↻ Atualizar"}
          </button>
        </div>
      </div>
      {rate && (
        <div style={{marginTop:8,paddingTop:8,borderTop:`1px solid ${C.success}22`,display:"flex",gap:10}}>
          {[
            {label:"Mercado",val:fmtBRL(rate,4),color:C.text},
            {label:"c/ IOF+spread",val:fmtBRL(comIOF,4),color:C.warning},
            {label:"Seu dólar",val:fmtBRL(settings.dollarPago,4),color:rate<=settings.dollarPago?C.danger:C.success},
          ].map(({label,val,color})=>(
            <div key={label} style={{flex:1,textAlign:"center",background:"rgba(255,255,255,0.6)",borderRadius:8,padding:"6px 4px"}}>
              <div style={{fontSize:9,color:C.textLight,marginBottom:2,fontWeight:600,textTransform:"uppercase"}}>{label}</div>
              <div style={{fontSize:12,fontWeight:800,color,fontFamily:"'DM Mono',monospace"}}>{val}</div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ─── DASHBOARD ────────────────────────────────────────────────────────────────
function DashboardTab({stats,settings,pesoPercent,pesoColor,pesoBg,onTabChange,anotacoes,setAnotacoes}) {
  const [showNotas,setShowNotas]=useState(false);
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

      {/* Cotação BCB */}
      <CotacaoBcbCard settings={settings}/>

      {/* Anotações */}
      <div style={{...S.card,padding:0,overflow:"hidden",marginBottom:10}}>
        <button onClick={()=>setShowNotas(v=>!v)} style={{width:"100%",background:"none",border:"none",cursor:"pointer",padding:"12px 14px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <span style={{fontSize:18}}>📝</span>
            <span style={{fontSize:14,fontWeight:700,color:C.text}}>Anotações</span>
            {anotacoes&&<span style={{fontSize:11,background:C.primaryLight,color:C.primary,borderRadius:999,padding:"1px 8px",fontWeight:600}}>{anotacoes.split("\n").filter(Boolean).length} linha(s)</span>}
          </div>
          <span style={{fontSize:12,color:C.textLight,fontWeight:600}}>{showNotas?"▲ Fechar":"▼ Abrir"}</span>
        </button>
        {showNotas&&(
          <div style={{borderTop:`1px solid ${C.border}`,padding:"12px 14px"}}>
            <textarea
              value={anotacoes}
              onChange={e=>setAnotacoes(e.target.value)}
              placeholder="Anote aqui o que quiser — lista de compras, lembretes, observações..."
              style={{width:"100%",minHeight:160,background:C.bg,border:`1px solid ${C.border}`,borderRadius:9,padding:"10px 12px",color:C.text,fontSize:14,fontFamily:"'Inter',sans-serif",resize:"vertical",outline:"none",boxSizing:"border-box",lineHeight:1.6}}
            />
            <div style={{fontSize:11,color:C.textLight,marginTop:4,textAlign:"right"}}>Salvo automaticamente</div>
          </div>
        )}
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
      <div style={{...S.card,display:"flex",alignItems:"center",gap:20,cursor:"pointer"}} onClick={()=>{onTabChange&&onTabChange(9);}}>
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
          <div style={{fontSize:10,color:C.textLight,marginTop:5,fontWeight:500}}>Toque para ver detalhes →</div>
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
function GastosTab({gastos,settings,onAdd,onEdit,onDelete,onTogglePago,produtos,onToggleStatus,parcelas}) {
  const [filtro,setFiltro]=useState("todos");
  const [subTab,setSubTab]=useState("gastos");

  const totalUSD=gastos.reduce((a,g)=>a+calcMinhaParteUSD(g,produtos),0);
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
      {/* Sub-abas */}
      <div style={{display:"flex",gap:4,background:C.borderLight,borderRadius:12,padding:4,marginBottom:12}}>
        {[["gastos","💸 Gastos"],["totais","📊 Gastos totais"]].map(([v,l])=>(
          <button key={v} style={{flex:1,padding:"9px 8px",borderRadius:9,border:"none",cursor:"pointer",fontSize:13,fontWeight:600,background:subTab===v?C.bgCard:"transparent",color:subTab===v?C.primary:C.textMid,boxShadow:subTab===v?"0 1px 4px rgba(0,0,0,0.08)":"none"}} onClick={()=>setSubTab(v)}>{l}</button>
        ))}
      </div>
      {subTab==="totais"&&<SimuladorTab settings={settings} gastos={gastos} parcelas={parcelas||[]} produtos={produtos||[]}/>}
      {subTab==="gastos"&&<>
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
      </>}
    </div>
  );
}

function GastoCard({g,settings,onEdit,onDelete,onTogglePago,produtos}) {
  const [expanded,setExpanded]=useState(false);
  // Se usd=0 e for produto, buscar preço do produto pai (gastos antigos podem ter usd:0)
  const prodPai = g.produtoId ? produtos?.find(p=>p.id===g.produtoId) : null;
  const usdUnit=parseFloat(g.usd)||parseFloat(prodPai?.usd)||0;
  const qtdG=parseInt(g.qtdComprada)||1;
  const taxaG=g.localTaxa==="orlando"?0.065:g.localTaxa==="kissimmee"?0.075:0;
  const totalUSD=g.tipo==="produto"?usdUnit*(1+taxaG)*qtdG:usdUnit;
  const minhaUSD=calcMinhaParteUSD(g,produtos);
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



// ─── NAV ICONS (SVG) ─────────────────────────────────────────────────────────
function NavIcon({ name, active }) {
  const col = active ? C.primary : C.textLight;
  const w = 22, h = 22;
  const icons = {
    "Início": (
      <svg width={w} height={h} viewBox="0 0 24 24" fill="none">
        <path d="M3 9.5L12 3l9 6.5V20a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V9.5z" stroke={col} strokeWidth="1.8" strokeLinejoin="round" fill={active?col+"22":"none"}/>
        <path d="M9 21V12h6v9" stroke={col} strokeWidth="1.8" strokeLinecap="round"/>
      </svg>
    ),
    "Produtos": (
      <svg width={w} height={h} viewBox="0 0 24 24" fill="none">
        <path d="M6 2L3 6v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6l-3-4z" stroke={col} strokeWidth="1.8" strokeLinejoin="round" fill={active?col+"22":"none"}/>
        <line x1="3" y1="6" x2="21" y2="6" stroke={col} strokeWidth="1.8"/>
        <path d="M16 10a4 4 0 0 1-8 0" stroke={col} strokeWidth="1.8" strokeLinecap="round"/>
      </svg>
    ),
    "Galeria": (
      <svg width={w} height={h} viewBox="0 0 24 24" fill="none">
        <rect x="3" y="3" width="8" height="8" rx="2" stroke={col} strokeWidth="1.8" fill={active?col+"22":"none"}/>
        <rect x="13" y="3" width="8" height="8" rx="2" stroke={col} strokeWidth="1.8" fill={active?col+"22":"none"}/>
        <rect x="3" y="13" width="8" height="8" rx="2" stroke={col} strokeWidth="1.8" fill={active?col+"22":"none"}/>
        <rect x="13" y="13" width="8" height="8" rx="2" stroke={col} strokeWidth="1.8" fill={active?col+"22":"none"}/>
      </svg>
    ),
    "Gastos": (
      <svg width={w} height={h} viewBox="0 0 24 24" fill="none">
        <circle cx="12" cy="12" r="9" stroke={col} strokeWidth="1.8" fill={active?col+"22":"none"}/>
        <path d="M12 6v1m0 10v1M9.5 9.5C9.5 8.12 10.62 7 12 7s2.5 1.12 2.5 2.5c0 1.5-1.5 2-2.5 2.5-1 .5-2.5 1-2.5 2.5C9.5 15.88 10.62 17 12 17s2.5-1.12 2.5-2.5" stroke={col} strokeWidth="1.8" strokeLinecap="round"/>
      </svg>
    ),
    "Parcelas": (
      <svg width={w} height={h} viewBox="0 0 24 24" fill="none">
        <rect x="2" y="5" width="20" height="14" rx="3" stroke={col} strokeWidth="1.8" fill={active?col+"22":"none"}/>
        <path d="M2 10h20" stroke={col} strokeWidth="1.8"/>
        <circle cx="7" cy="15" r="1.2" fill={col}/>
        <circle cx="12" cy="15" r="1.2" fill={col}/>
      </svg>
    ),
    "Roteiro": (
      <svg width={w} height={h} viewBox="0 0 24 24" fill="none">
        <rect x="3" y="4" width="18" height="18" rx="3" stroke={col} strokeWidth="1.8" fill={active?col+"22":"none"}/>
        <path d="M16 2v4M8 2v4M3 10h18" stroke={col} strokeWidth="1.8" strokeLinecap="round"/>
        <circle cx="8" cy="15" r="1.5" fill={col}/>
        <circle cx="12" cy="15" r="1.5" fill={col}/>
        <circle cx="16" cy="15" r="1.5" fill={col}/>
      </svg>
    ),
    "Stats": (
      <svg width={w} height={h} viewBox="0 0 24 24" fill="none">
        <path d="M18 20V10M12 20V4M6 20v-6" stroke={col} strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"/>
        {active && <rect x="4" y="14" width="4" height="6" rx="1" fill={col+"33"}/>}
        {active && <rect x="10" y="4" width="4" height="16" rx="1" fill={col+"33"}/>}
        {active && <rect x="16" y="10" width="4" height="10" rx="1" fill={col+"33"}/>}
      </svg>
    ),
    "Dólar": (
      <svg width={w} height={h} viewBox="0 0 24 24" fill="none">
        <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2z" stroke={col} strokeWidth="1.8" fill={active?col+"22":"none"}/>
        <path d="M12 6v1m0 10v1M9.5 9.5C9.5 8.12 10.62 7 12 7s2.5 1.12 2.5 2.5c0 1.5-1.5 2-2.5 2.5-1 .5-2.5 1-2.5 2.5C9.5 15.88 10.62 17 12 17s2.5-1.12 2.5-2.5" stroke={col} strokeWidth="1.8" strokeLinecap="round"/>
        <path d="M4 7l2 2M18 15l2 2M4 17l2-2M18 7l2-2" stroke={col} strokeWidth="1.4" strokeLinecap="round" opacity="0.5"/>
      </svg>
    ),
    "Calc": (
      <svg width={w} height={h} viewBox="0 0 24 24" fill="none">
        <rect x="4" y="2" width="16" height="20" rx="3" stroke={col} strokeWidth="1.8" fill={active?col+"22":"none"}/>
        <rect x="7" y="5" width="10" height="4" rx="1.5" fill={col} opacity="0.5"/>
        <circle cx="8" cy="13" r="1.2" fill={col}/>
        <circle cx="12" cy="13" r="1.2" fill={col}/>
        <circle cx="16" cy="13" r="1.2" fill={col}/>
        <circle cx="8" cy="17" r="1.2" fill={col}/>
        <circle cx="12" cy="17" r="1.2" fill={col}/>
        <rect x="14.5" y="15.5" width="3" height="3" rx="0.8" fill={col}/>
      </svg>
    ),
  };
  return icons[name] || <span style={{fontSize:18}}>{name[0]}</span>;
}

// ─── ROTEIRO TAB ──────────────────────────────────────────────────────────────
function RoteiroTab({ planejamento, setPlanejamento }) {
  const { dataInicio, dataFim, eventos } = planejamento;
  const [showEventoForm, setShowEventoForm] = useState(false);
  const [editEvento, setEditEvento] = useState(null);
  const [mesCalendario, setMesCalendario] = useState(() => {
    if (planejamento.dataInicio) return planejamento.dataInicio.slice(0, 7);
    return new Date().toISOString().slice(0, 7);
  });
  const hoje = new Date().toISOString().slice(0, 10);
  const eventosHoje = (eventos||[]).filter(e => e.data === hoje).sort((a,b) => (a.hora||"99:99").localeCompare(b.hora||"99:99"));
  const viagemAtiva = dataInicio && dataFim && hoje >= dataInicio && hoje <= dataFim;
  const diasRestantes = dataInicio && hoje < dataInicio ? Math.ceil((new Date(dataInicio) - new Date(hoje)) / 86400000) : null;

  function salvarConfig(campo, valor) { setPlanejamento(p => ({ ...p, [campo]: valor })); }
  function salvarEvento(ev) {
    if (ev.id) { setPlanejamento(p => ({ ...p, eventos: p.eventos.map(e => e.id===ev.id?ev:e) })); }
    else { setPlanejamento(p => ({ ...p, eventos: [...(p.eventos||[]), {...ev, id:Date.now()}] })); }
    setShowEventoForm(false); setEditEvento(null);
  }
  function excluirEvento(id) { setPlanejamento(p => ({ ...p, eventos: p.eventos.filter(e => e.id!==id) })); }

  const eventosPorData = useMemo(() => {
    const map = {};
    (eventos||[]).forEach(e => { if (!map[e.data]) map[e.data]=[]; map[e.data].push(e); });
    Object.values(map).forEach(arr => arr.sort((a,b) => (a.hora||"99:99").localeCompare(b.hora||"99:99")));
    return map;
  }, [eventos]);

  const datasViagem = useMemo(() => {
    if (!dataInicio || !dataFim) return new Set();
    const s = new Set(); const cur = new Date(dataInicio); const fim = new Date(dataFim);
    while (cur <= fim) { s.add(cur.toISOString().slice(0,10)); cur.setDate(cur.getDate()+1); }
    return s;
  }, [dataInicio, dataFim]);

  return (
    <div style={S.page}>
      <div style={S.heroCard}>
        <div style={{fontSize:12,fontWeight:500,color:"rgba(255,255,255,0.75)",marginBottom:3}}>Roteiro da viagem</div>
        {viagemAtiva ? (
          <div style={{fontSize:22,fontWeight:800,color:"#fff"}}>✈ Hoje é dia de viagem! 🎉</div>
        ) : diasRestantes !== null ? (
          <>
            <div style={{fontSize:28,fontWeight:800,color:"#fff",letterSpacing:"-0.5px"}}>{diasRestantes} dias</div>
            <div style={{fontSize:13,color:"rgba(255,255,255,0.8)",marginTop:2}}>para a viagem começar 🗓</div>
          </>
        ) : (
          <div style={{fontSize:20,fontWeight:700,color:"#fff"}}>Configure as datas abaixo</div>
        )}
        {eventosHoje.length > 0 && (
          <div style={{marginTop:12,background:"rgba(255,255,255,0.15)",borderRadius:12,padding:"10px 12px"}}>
            <div style={{fontSize:11,fontWeight:700,color:"rgba(255,255,255,0.8)",marginBottom:6,textTransform:"uppercase",letterSpacing:"0.5px"}}>Hoje</div>
            {eventosHoje.map(e => (
              <div key={e.id} style={{display:"flex",alignItems:"center",gap:8,marginBottom:4}}>
                <span style={{fontSize:16}}>{e.emoji||"📌"}</span>
                <span style={{fontSize:13,fontWeight:600,color:"#fff"}}>{e.titulo}</span>
                {e.hora&&<span style={{fontSize:11,color:"rgba(255,255,255,0.7)"}}>às {e.hora}</span>}
              </div>
            ))}
          </div>
        )}
      </div>

      <div style={S.card}>
        <div style={{fontWeight:700,fontSize:13,color:C.text,marginBottom:12}}>📅 Período da viagem</div>
        <div style={{display:"flex",gap:10}}>
          <div style={{flex:1}}><label style={S.label}>Ida</label><input style={S.input} type="date" value={dataInicio} onChange={e=>salvarConfig("dataInicio",e.target.value)}/></div>
          <div style={{flex:1}}><label style={S.label}>Volta</label><input style={S.input} type="date" value={dataFim} onChange={e=>salvarConfig("dataFim",e.target.value)}/></div>
        </div>
        {dataInicio&&dataFim&&dataInicio<=dataFim&&(
          <div style={{background:C.primaryLight,borderRadius:10,padding:"8px 12px",fontSize:13,color:C.primary,fontWeight:600,marginTop:-4}}>
            ✈ {Math.ceil((new Date(dataFim)-new Date(dataInicio))/86400000)+1} dias · {new Date(dataInicio+"T12:00:00").toLocaleDateString("pt-BR",{day:"2-digit",month:"short"})} → {new Date(dataFim+"T12:00:00").toLocaleDateString("pt-BR",{day:"2-digit",month:"short",year:"numeric"})}
          </div>
        )}
      </div>

      <CalendarioViagem mesCalendario={mesCalendario} onMesChange={setMesCalendario} datasViagem={datasViagem} eventosPorData={eventosPorData} hoje={hoje} dataInicio={dataInicio} dataFim={dataFim} onDiaClick={data=>{setEditEvento({data});setShowEventoForm(true);}}/>

      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
        <div style={{fontWeight:700,fontSize:14,color:C.text}}>Eventos ({(eventos||[]).length})</div>
        <button style={{...S.btnPrimary,padding:"7px 14px",fontSize:13,marginBottom:0,width:"auto"}} onClick={()=>{setEditEvento(null);setShowEventoForm(true);}}>＋ Evento</button>
      </div>

      {(eventos||[]).length===0&&<Empty text="Nenhum evento. Toque em ＋ ou clique em um dia no calendário."/>}

      {Object.keys(eventosPorData).sort().map(data=>(
        <div key={data} style={{marginBottom:12}}>
          <div style={{fontSize:11,fontWeight:700,color:C.textMid,marginBottom:6,textTransform:"uppercase",letterSpacing:"0.5px",display:"flex",alignItems:"center",gap:6}}>
            {new Date(data+"T12:00:00").toLocaleDateString("pt-BR",{weekday:"short",day:"2-digit",month:"short"})}
            {data===hoje&&<span style={{background:C.primary,color:"#fff",borderRadius:999,padding:"1px 8px",fontSize:10}}>Hoje</span>}
          </div>
          {eventosPorData[data].map(e=>(
            <div key={e.id} style={{...S.card,marginBottom:6,padding:"10px 14px"}}>
              <div style={{display:"flex",alignItems:"center",gap:10}}>
                <div style={{width:36,height:36,borderRadius:10,background:C.primaryLight,display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0}}>{e.emoji||"📌"}</div>
                <div style={{flex:1}}>
                  <div style={{fontWeight:600,fontSize:14,color:C.text}}>{e.titulo}</div>
                  <div style={{fontSize:12,color:C.textLight,marginTop:1}}>
                    {e.hora&&<span>🕐 {e.hora}{e.horaFim?` → ${e.horaFim}`:""}{e.local?" · ":""}</span>}
                    {e.local&&<span>📍 {e.local}</span>}
                  </div>
                  {e.notas&&<div style={{fontSize:12,color:C.textMid,marginTop:3}}>{e.notas}</div>}
                </div>
                <div style={{display:"flex",gap:6,flexShrink:0}}>
                  <button style={{background:C.borderLight,border:"none",borderRadius:7,width:28,height:28,cursor:"pointer",fontSize:13}} onClick={()=>{setEditEvento(e);setShowEventoForm(true);}}>✏</button>
                  <button style={{background:C.dangerLight,border:"none",borderRadius:7,width:28,height:28,cursor:"pointer",fontSize:13,color:C.danger}} onClick={()=>excluirEvento(e.id)}>🗑</button>
                </div>
              </div>
            </div>
          ))}
        </div>
      ))}

      {showEventoForm&&<EventoForm evento={editEvento} onSalvar={salvarEvento} onClose={()=>{setShowEventoForm(false);setEditEvento(null);}}/>}
    </div>
  );
}

function CalendarioViagem({mesCalendario,onMesChange,datasViagem,eventosPorData,hoje,dataInicio,dataFim,onDiaClick}) {
  const [ano,mes] = mesCalendario.split("-").map(Number);
  const primeiroDia = new Date(ano,mes-1,1).getDay();
  const diasNoMes = new Date(ano,mes,0).getDate();
  const nomeMes = new Date(ano,mes-1,1).toLocaleDateString("pt-BR",{month:"long",year:"numeric"});
  function navMes(dir){const d=new Date(ano,mes-1+dir,1);onMesChange(d.toISOString().slice(0,7));}
  const celulas=[];
  for(let i=0;i<primeiroDia;i++) celulas.push(null);
  for(let d=1;d<=diasNoMes;d++) celulas.push(d);
  return (
    <div style={{...S.card,marginBottom:14,padding:"14px 10px"}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
        <button onClick={()=>navMes(-1)} style={{background:"none",border:"none",fontSize:20,cursor:"pointer",color:C.textMid,padding:"4px 10px"}}>‹</button>
        <div style={{textAlign:"center"}}>
          <div style={{fontWeight:700,fontSize:14,color:C.text,textTransform:"capitalize"}}>{nomeMes}</div>
          {dataInicio&&mesCalendario!==dataInicio.slice(0,7)&&<button onClick={()=>onMesChange(dataInicio.slice(0,7))} style={{background:"none",border:"none",fontSize:11,color:C.primary,cursor:"pointer",fontWeight:600}}>Ir para viagem →</button>}
        </div>
        <button onClick={()=>navMes(1)} style={{background:"none",border:"none",fontSize:20,cursor:"pointer",color:C.textMid,padding:"4px 10px"}}>›</button>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",marginBottom:6}}>
        {["D","S","T","Q","Q","S","S"].map((d,i)=><div key={i} style={{textAlign:"center",fontSize:10,fontWeight:700,color:C.textLight}}>{d}</div>)}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:2}}>
        {celulas.map((d,i)=>{
          if(!d) return <div key={i}/>;
          const ds=`${ano}-${String(mes).padStart(2,"0")}-${String(d).padStart(2,"0")}`;
          const eV=datasViagem.has(ds),eH=ds===hoje,eI=ds===dataInicio,eF=ds===dataFim;
          const qtEvs=(eventosPorData[ds]||[]).length;
          return (
            <button key={i} onClick={()=>onDiaClick(ds)} style={{
              aspectRatio:"1",borderRadius:eI?"10px 4px 4px 10px":eF?"4px 10px 10px 4px":eV?"4px":"8px",
              border:eH?`2px solid ${C.primary}`:"2px solid transparent",
              background:eH?C.primary:eI||eF?C.primary:eV?C.primaryLight:"transparent",
              color:eH||eI||eF?"#fff":eV?C.primary:C.text,
              fontWeight:eH||eV?700:400,fontSize:13,cursor:"pointer",
              display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:"2px",
            }}>
              {d}
              {qtEvs>0&&<div style={{width:qtEvs>1?14:5,height:4,borderRadius:999,background:eH||eI||eF?"rgba(255,255,255,0.8)":C.primary,marginTop:1,fontSize:8,display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontWeight:700}}>{qtEvs>1?qtEvs:""}</div>}
            </button>
          );
        })}
      </div>
      <div style={{display:"flex",gap:12,marginTop:10,flexWrap:"wrap"}}>
        {[{bg:C.primary,label:"Hoje/Chegada/Volta"},{bg:C.primaryLight,label:"Período da viagem"}].map(({bg,label})=>(
          <div key={label} style={{display:"flex",alignItems:"center",gap:5}}><div style={{width:12,height:12,borderRadius:3,background:bg}}/><span style={{fontSize:10,color:C.textLight}}>{label}</span></div>
        ))}
      </div>
    </div>
  );
}

const EMOJIS_EVENTO=["📌","✈","🏨","🎡","🛍","🍔","🎢","🌊","🎭","🚗","⛵","🎠","🎪","🎟","🏖","🌴","🎆","🎇","🎑","🎈"];
function EventoForm({evento,onSalvar,onClose}) {
  const [f,setF]=useState(evento?.id?{...evento}:{titulo:"",data:evento?.data||"",hora:"",horaFim:"",local:"",notas:"",emoji:"📌"});
  function salvar(){if(!f.titulo.trim())return alert("Informe o título");if(!f.data)return alert("Informe a data");onSalvar({...f,titulo:f.titulo.trim()});}
  return (
    <Modal title={evento?.id?"Editar evento":"Novo evento"} onClose={onClose}>
      <label style={S.label}>Ícone</label>
      <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:14}}>
        {EMOJIS_EVENTO.map(e=><button key={e} onClick={()=>setF(p=>({...p,emoji:e}))} style={{width:36,height:36,borderRadius:8,border:`2px solid ${f.emoji===e?C.primary:C.border}`,background:f.emoji===e?C.primaryLight:"transparent",fontSize:18,cursor:"pointer"}}>{e}</button>)}
      </div>
      <label style={S.label}>Título *</label>
      <input style={S.input} placeholder="Ex: Universal Studios..." value={f.titulo} onChange={e=>setF(p=>({...p,titulo:e.target.value}))}/>
      <label style={S.label}>Data *</label>
      <input style={S.input} type="date" value={f.data} onChange={e=>setF(p=>({...p,data:e.target.value}))}/>
      <div style={{display:"flex",gap:10}}>
        <div style={{flex:1}}><label style={S.label}>Início (opcional)</label><input style={S.input} type="time" value={f.hora} onChange={e=>setF(p=>({...p,hora:e.target.value}))}/></div>
        <div style={{flex:1}}><label style={S.label}>Fim (opcional)</label><input style={S.input} type="time" value={f.horaFim} onChange={e=>setF(p=>({...p,horaFim:e.target.value}))}/></div>
      </div>
      <label style={S.label}>Local (opcional)</label>
      <input style={S.input} placeholder="Ex: Universal Orlando Resort" value={f.local} onChange={e=>setF(p=>({...p,local:e.target.value}))}/>
      <label style={S.label}>Notas (opcional)</label>
      <input style={S.input} placeholder="Ex: Reserva confirmada..." value={f.notas} onChange={e=>setF(p=>({...p,notas:e.target.value}))}/>
      <button style={S.btnPrimary} onClick={salvar}>{evento?.id?"Salvar alterações":"Adicionar evento"}</button>
    </Modal>
  );
}

// ─── CALC TAB (extendida: Simulador + Histórico Dólar + Checklist + Bagagem) ──
function CalcTab({settings, gastos, produtos, parcelas, comprasDolar, setComprasDolar, checklist, setChecklist, initialSubTab, onSubTabChange}) {
  const [subTab, setSubTab] = useState(initialSubTab||"conversor");
  function changeSubTab(v){setSubTab(v);onSubTabChange&&onSubTabChange(v);}
  const SUBTABS = [
    {id:"conversor",label:"💱 Câmbio"},



  ];
  return (
    <div style={S.page}>
      <div style={{display:"flex",gap:4,overflowX:"auto",paddingBottom:4,marginBottom:14}}>
        {SUBTABS.map(t=>(
          <button key={t.id} onClick={()=>changeSubTab(t.id)} style={{...S.chip,...(subTab===t.id?S.chipActive:{}),whiteSpace:"nowrap",flexShrink:0,fontSize:12,padding:"6px 12px"}}>{t.label}</button>
        ))}
      </div>
      {subTab==="conversor"&&<ConversorTab settings={settings}/>}

    </div>
  );
}

// ─── CONVERSOR (era CalcTab original) ────────────────────────────────────────
function ConversorTab({settings}) {
  const [usdN,setUsdN]=useState(""); const [brlN,setBrlN]=useState(""); const [dc,setDc]=useState(""); const [lbs,setLbs]=useState(""); const [oz,setOz]=useState("");
  const dolarAj=calcDolarAjustado(settings); const brlP=parseFloat(usdN)*(1+settings.taxa/100)*dolarAj; const brlC=dc&&parseFloat(usdN)>0?parseFloat(usdN)*(1+settings.taxa/100)*parseFloat(dc):null;
  return (
    <>
      <div style={S.sectionLabel}>💵 Conversor USD → BRL</div>
      <div style={S.card}>
        <label style={S.label}>Valor em USD</label>
        <input style={S.input} type="number" placeholder="Ex: 150" value={usdN} onChange={e=>setUsdN(e.target.value)}/>
        {usdN&&<div style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${C.borderLight}`}}><span style={{fontSize:13,color:C.textMid}}>BRL estimado</span><span style={{fontSize:13,fontWeight:700,color:C.primary,fontFamily:"'DM Mono',monospace"}}>{fmtBRL(brlP)}</span></div>}
        <label style={{...S.label,marginTop:8}}>Dólar que você pagou (opcional)</label>
        <input style={S.input} type="number" step="0.01" placeholder="Ex: 5.71" value={dc} onChange={e=>setDc(e.target.value)}/>
        {brlC&&parseFloat(usdN)>0&&[{label:"Com seu dólar",value:fmtBRL(brlC),color:C.success},{label:"Diferença",value:fmtBRL(brlP-brlC),color:brlP>brlC?C.success:C.danger}].map(({label,value,color})=>(
          <div key={label} style={{display:"flex",justifyContent:"space-between",padding:"7px 0",borderBottom:`1px solid ${C.borderLight}`}}><span style={{fontSize:13,color:C.textMid}}>{label}</span><span style={{fontSize:13,fontWeight:700,color,fontFamily:"'DM Mono',monospace"}}>{value}</span></div>
        ))}
      </div>
      <div style={S.sectionLabel}>⚖ Conversor de Peso</div>
      <div style={S.card}>
        <label style={S.label}>Libras (lbs)</label>
        <input style={S.input} type="number" placeholder="Ex: 2.5" value={lbs} onChange={e=>setLbs(e.target.value)}/>
        {lbs&&[[`Gramas`,`${fmtN(parseFloat(lbs)*453.592,1)}g`],[`Kg`,`${fmtN(parseFloat(lbs)*0.453592,3)}kg`],["Oz",`${fmtN(parseFloat(lbs)*16,1)} oz`]].map(([l,v])=>(
          <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:`1px solid ${C.borderLight}`}}><span style={{fontSize:13,color:C.textMid}}>{l}</span><span style={{fontSize:13,fontWeight:700,color:C.text,fontFamily:"'DM Mono',monospace"}}>{v}</span></div>
        ))}
        <div style={{height:12}}/>
        <label style={S.label}>Onças (oz)</label>
        <input style={S.input} type="number" placeholder="Ex: 3.4" value={oz} onChange={e=>setOz(e.target.value)}/>
        {oz&&[["Gramas",`${fmtN(parseFloat(oz)*28.3495,1)}g`],["Kg",`${fmtN(parseFloat(oz)*28.3495/1000,3)}kg`],["Libras",`${fmtN(parseFloat(oz)/16,3)} lbs`]].map(([l,v])=>(
          <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:`1px solid ${C.borderLight}`}}><span style={{fontSize:13,color:C.textMid}}>{l}</span><span style={{fontSize:13,fontWeight:700,color:C.text,fontFamily:"'DM Mono',monospace"}}>{v}</span></div>
        ))}
      </div>
      <div style={S.card}>
        <div style={{fontWeight:700,fontSize:13,color:C.text,marginBottom:10}}>Taxas e cotações</div>
        {[["Dólar pago",fmtBRL(settings.dollarPago,4)],["IOF",`${settings.iof}%`],["Spread",`${settings.spread}%`],["Taxa compra",`${settings.taxa}%`],["Dólar ajustado",fmtBRL(dolarAj,4)]].map(([l,v])=>(
          <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:`1px solid ${C.borderLight}`}}><span style={{fontSize:13,color:C.textMid}}>{l}</span><span style={{fontSize:13,fontWeight:700,color:C.primary,fontFamily:"'DM Mono',monospace"}}>{v}</span></div>
        ))}
        <BcbRate/>
      </div>
    </>
  );
}

// ─── SIMULADOR ───────────────────────────────────────────────────────────────
function SimuladorTab({settings, gastos, parcelas, produtos}) {
  const totalGastosUSD = calcTotalGastosUSD(gastos, produtos||[]);
  const totalParcelasMensal = parcelas.reduce((a,p)=>a+(parseFloat(p.minhaParte||p.valorParcela)||0)/parseInt(p.quantidadeParcelas||1),0);
  const totalParcelasBRL = parcelas.reduce((a,p)=>a+(parseFloat(p.minhaParte)||parseFloat(p.valorTotal)||0),0);
  const parcelasRestBRL = parcelas.reduce((a,p)=>{
    const qt=parseInt(p.quantidadeParcelas)||0; const pagas=(p.statusMensal||[]).filter(Boolean).length;
    return a+Math.max(0,qt-pagas)*(parseFloat(p.minhaParte||p.valorParcela)||0)/qt;
  },0)*parcelas.reduce((a,p)=>{const qt=parseInt(p.quantidadeParcelas)||0;const pagas=(p.statusMensal||[]).filter(Boolean).length;return a+Math.max(0,qt-pagas);},0)/Math.max(1,parcelas.reduce((a,p)=>{const qt=parseInt(p.quantidadeParcelas)||0;return a+qt;},0));
  const totalParcelasRestBRL = parcelas.reduce((a,p)=>{
    const qt=parseInt(p.quantidadeParcelas)||0;
    const pagas=(p.statusMensal||[]).filter(Boolean).length;
    const restante=Math.max(0,qt-pagas);
    // Se tem minhaParte: é o total da minha parte, então valor/parcela = minhaParte/qt
    // Se não tem: usar valorParcela inteiro
    const mp=parseFloat(p.minhaParte)||0;
    const vp=mp>0&&qt>0 ? mp/qt : (parseFloat(p.valorParcela)||0);
    return a+restante*vp;
  },0);
  const usdGastosEmBRL = totalGastosUSD * calcDolarAjustado(settings);
  const totalViagem = usdGastosEmBRL + totalParcelasRestBRL;

  return (
    <>
      <div style={{...S.heroCard,marginBottom:14}}>
        <div style={{fontSize:12,fontWeight:500,color:"rgba(255,255,255,0.75)",marginBottom:3}}>Custo total estimado da viagem</div>
        <div style={{fontSize:30,fontWeight:800,color:"#fff",letterSpacing:"-1px"}}>{fmtBRL(totalViagem)}</div>
        <div style={{fontSize:13,color:"rgba(255,255,255,0.75)",marginTop:3}}>gastos lá + parcelas restantes</div>
      </div>
      <div style={S.card}>
        <div style={{fontWeight:700,fontSize:13,color:C.text,marginBottom:12}}>Composição do custo</div>
        {[
          {label:"💸 Gastos na viagem (USD→BRL)",value:fmtBRL(usdGastosEmBRL),color:C.primary,sub:`${fmtUSD(totalGastosUSD)} × ${fmtBRL(calcDolarAjustado(settings),4)} (c/ IOF+spread)`},
          {label:"💳 Parcelas restantes (BRL)",value:fmtBRL(totalParcelasRestBRL),color:C.purple,sub:`${parcelas.filter(p=>(p.statusMensal||[]).some(s=>!s)).length} itens com parcelas a pagar`},
          {label:"📊 Total comprometido",value:fmtBRL(totalViagem),color:C.text,sub:""},
        ].map(({label,value,color,sub})=>(
          <div key={label} style={{padding:"10px 0",borderBottom:`1px solid ${C.borderLight}`}}>
            <div style={{display:"flex",justifyContent:"space-between"}}>
              <span style={{fontSize:13,color:C.textMid,fontWeight:600}}>{label}</span>
              <span style={{fontSize:14,fontWeight:800,color,fontFamily:"'DM Mono',monospace"}}>{value}</span>
            </div>
            {sub&&<div style={{fontSize:11,color:C.textLight,marginTop:2}}>{sub}</div>}
          </div>
        ))}
      </div>
      <div style={S.card}>
        <div style={{fontWeight:700,fontSize:13,color:C.text,marginBottom:8}}>💵 Dólar para levar</div>
        <div style={{fontSize:12,color:C.textMid,marginBottom:10}}>Baseado nos produtos pendentes e gastos estimados</div>
        <div style={{display:"flex",justifyContent:"space-between",padding:"8px 0"}}>
          <span style={{fontSize:13,color:C.textMid}}>Budget configurado</span>
          <span style={{fontSize:13,fontWeight:700,color:C.primary,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(settings.totalDolarViagem)}</span>
        </div>
        <div style={{display:"flex",justifyContent:"space-between",padding:"8px 0"}}>
          <span style={{fontSize:13,color:C.textMid}}>Já gastou</span>
          <span style={{fontSize:13,fontWeight:700,color:C.danger,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(totalGastosUSD)}</span>
        </div>
        <div style={{height:1,background:C.border,margin:"4px 0"}}/>
        <div style={{display:"flex",justifyContent:"space-between",padding:"8px 0"}}>
          <span style={{fontSize:13,fontWeight:700,color:C.text}}>Ainda disponível</span>
          <span style={{fontSize:15,fontWeight:800,color:settings.totalDolarViagem-totalGastosUSD>=0?C.success:C.danger,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(Math.abs(settings.totalDolarViagem-totalGastosUSD))}{settings.totalDolarViagem-totalGastosUSD<0?" (excedido)":""}</span>
        </div>
      </div>
    </>
  );
}

// ─── HISTÓRICO DÓLAR ─────────────────────────────────────────────────────────
function HistoricoDolarTab({comprasDolar, setComprasDolar, settings}) {
  const [showForm, setShowForm] = useState(false);
  const [f, setF] = useState({data:"",quantidade:"",cotacao:"",obs:""});
  const [cotacaoBCB, setCotacaoBCB] = useState(null);
  const totalUSD = comprasDolar.reduce((a,c)=>a+(parseFloat(c.quantidade)||0),0);
  const totalBRL = comprasDolar.reduce((a,c)=>a+(parseFloat(c.quantidade)||0)*(parseFloat(c.cotacao)||0),0);
  const custoMedio = totalUSD>0 ? totalBRL/totalUSD : 0;
  // custo médio já inclui IOF+spread — comparar com mercado puro (BCB sem taxas)
  const mercadoBase = cotacaoBCB || settings.dollarPago;
  // Custo médio - (mercado + IOF + spread): taxas somadas ao valor de mercado
  const cotacaoComTaxas = mercadoBase + mercadoBase * (settings.iof + settings.spread) / 100;

  useEffect(()=>{
    fetchCotacao().then(r=>{ if(r?.bid) setCotacaoBCB(r.bid); });
  },[]);

  function salvar(){
    if(!f.quantidade||!f.cotacao||!f.data) return alert("Preencha data, quantidade e cotação");
    setComprasDolar(ps=>[...ps,{...f,id:Date.now(),quantidade:parseFloat(f.quantidade),cotacao:parseFloat(f.cotacao)}]);
    setF({data:"",quantidade:"",cotacao:"",obs:""}); setShowForm(false);
  }

  return (
    <>
      <div style={{...S.heroCard,marginBottom:14}}>
        <div style={{fontSize:12,fontWeight:500,color:"rgba(255,255,255,0.75)",marginBottom:3}}>Dólar acumulado</div>
        <div style={{fontSize:28,fontWeight:800,color:"#fff",letterSpacing:"-0.5px"}}>{fmtUSD(totalUSD)}</div>
        <div style={{display:"flex",gap:10,marginTop:10}}>
          <div style={{background:"rgba(255,255,255,0.15)",borderRadius:12,padding:"8px 12px",flex:1,textAlign:"center"}}>
            <div style={{fontSize:10,color:"rgba(255,255,255,0.65)",marginBottom:2}}>Custo médio</div>
            <div style={{fontSize:14,fontWeight:700,color:"#fff",fontFamily:"'DM Mono',monospace"}}>{fmtBRL(custoMedio,4)}</div>
          </div>
          <div style={{background:"rgba(255,255,255,0.15)",borderRadius:12,padding:"8px 12px",flex:1,textAlign:"center"}}>
            <div style={{fontSize:10,color:"rgba(255,255,255,0.65)",marginBottom:2}}>Total investido</div>
            <div style={{fontSize:14,fontWeight:700,color:"#fff",fontFamily:"'DM Mono',monospace"}}>{fmtBRL(totalBRL)}</div>
          </div>
          <div style={{background:custoMedio>cotacaoComTaxas?"rgba(239,68,68,0.3)":"rgba(16,185,129,0.3)",borderRadius:12,padding:"8px 12px",flex:1,textAlign:"center"}}>
            <div style={{fontSize:10,color:"rgba(255,255,255,0.65)",marginBottom:2}}>{custoMedio>cotacaoComTaxas?"Acima":"Abaixo"} mercado+taxas</div>
            <div style={{fontSize:14,fontWeight:700,color:"#fff",fontFamily:"'DM Mono',monospace"}}>{custoMedio>0?`${custoMedio>cotacaoComTaxas?"+":"-"}${fmtBRL(Math.abs(custoMedio-cotacaoComTaxas),4)}`:"—"}</div>
            {cotacaoBCB&&<div style={{fontSize:9,color:"rgba(255,255,255,0.6)",marginTop:2}}>mercado BCB: {fmtBRL(cotacaoBCB,4)}</div>}
          </div>
        </div>
      </div>

      <button style={{...S.btnPrimary,marginBottom:14}} onClick={()=>setShowForm(s=>!s)}>
        {showForm?"Cancelar":"＋ Registrar compra de dólar"}
      </button>

      {showForm&&(
        <div style={{...S.card,marginBottom:14}}>
          <div style={{fontWeight:700,fontSize:13,color:C.text,marginBottom:12}}>Nova compra</div>
          <label style={S.label}>Data</label>
          <input style={S.input} type="date" value={f.data} onChange={e=>setF(p=>({...p,data:e.target.value}))}/>
          <div style={{display:"flex",gap:10}}>
            <div style={{flex:1}}><label style={S.label}>Quantidade (US$)</label><input style={S.input} type="number" step="50" placeholder="Ex: 500" value={f.quantidade} onChange={e=>setF(p=>({...p,quantidade:e.target.value}))}/></div>
            <div style={{flex:1}}><label style={S.label}>Cotação (R$)</label><input style={S.input} type="number" step="0.01" placeholder="Ex: 5.65" value={f.cotacao} onChange={e=>setF(p=>({...p,cotacao:e.target.value}))}/></div>
          </div>
          {f.quantidade&&f.cotacao&&<div style={{background:C.primaryLight,borderRadius:10,padding:"8px 12px",fontSize:13,color:C.primary,fontWeight:600,marginBottom:12}}>Total: {fmtBRL(parseFloat(f.quantidade)*parseFloat(f.cotacao))}</div>}
          <label style={S.label}>Observação (opcional)</label>
          <input style={S.input} placeholder="Ex: Wise, Banco do Brasil..." value={f.obs} onChange={e=>setF(p=>({...p,obs:e.target.value}))}/>
          <button style={S.btnPrimary} onClick={salvar}>Salvar</button>
        </div>
      )}

      {comprasDolar.length===0&&<Empty text="Nenhuma compra registrada ainda."/>}
      {[...comprasDolar].reverse().map(c=>(
        <div key={c.id} style={{...S.card,marginBottom:8,padding:"10px 14px"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
            <div>
              <div style={{fontWeight:600,fontSize:14,color:C.text}}>{fmtUSD(c.quantidade)}</div>
              <div style={{fontSize:12,color:C.textLight,marginTop:1}}>{c.data?new Date(c.data+"T12:00:00").toLocaleDateString("pt-BR"):""}{c.obs?` · ${c.obs}`:""}</div>
            </div>
            <div style={{textAlign:"right"}}>
              <div style={{fontSize:13,fontWeight:700,color:C.purple,fontFamily:"'DM Mono',monospace"}}>{fmtBRL(parseFloat(c.cotacao),4)}/US$</div>
              <div style={{fontSize:11,color:C.textLight,fontFamily:"'DM Mono',monospace"}}>{fmtBRL(c.quantidade*c.cotacao)}</div>
            </div>
            <button onClick={()=>setComprasDolar(ps=>ps.filter(p=>p.id!==c.id))} style={{background:C.dangerLight,border:"none",borderRadius:6,width:24,height:24,cursor:"pointer",color:C.danger,fontSize:12,marginLeft:8,flexShrink:0}}>✕</button>
          </div>
        </div>
      ))}
    </>
  );
}

// ─── BAGAGEM VISUAL ──────────────────────────────────────────────────────────
function BagagemTab({produtos, settings}) {
  const pesoMax = settings.pesoMax;
  const categorias = [
    {label:"📱 Eletrônicos",lojas:["Apple","Best Buy","Newegg"],cor:C.primary},
    {label:"👗 Roupas",lojas:["Tommy Hilfiger","Calvin Klein","The North Face","Marshalls","Ross","TJ Maxx"],cor:C.purple},
    {label:"💧 Líquidos",tipoPeso:["oz_liquido"],cor:C.success},
    {label:"🛍 Outros",cor:C.warning},
  ];
  const pesoTotal = produtos.filter(p=>p.status==="comprado").reduce((a,p)=>a+prodPeso(p),0);
  const pct = Math.min(100,(pesoTotal/pesoMax)*100);
  const cor = pct<70?C.success:pct<90?C.warning:C.danger;

  const porCategoria = categorias.map(cat=>{
    const itens = produtos.filter(p=>p.status==="comprado"&&(
      cat.tipoPeso ? cat.tipoPeso.includes(p.tipoPeso||"g") :
      cat.lojas ? cat.lojas.includes(p.loja) :
      true
    ));
    const peso = itens.reduce((a,p)=>a+prodPeso(p),0);
    return {...cat, peso, itens:itens.length};
  });
  // "Outros" = total - categorias específicas
  const pesoCategorizado = porCategoria.slice(0,-1).reduce((a,c)=>a+c.peso,0);
  porCategoria[porCategoria.length-1].peso = Math.max(0, pesoTotal - pesoCategorizado);

  return (
    <>
      <div style={{...S.heroCard,marginBottom:14}}>
        <div style={{fontSize:12,fontWeight:500,color:"rgba(255,255,255,0.75)",marginBottom:3}}>Peso da mala</div>
        <div style={{fontSize:28,fontWeight:800,color:"#fff"}}>{fmtN(pesoTotal/1000,2)} kg <span style={{fontSize:14,opacity:0.75}}>/ {fmtN(pesoMax/1000,0)} kg</span></div>
        <div style={{marginTop:12,background:"rgba(255,255,255,0.2)",borderRadius:999,height:12,overflow:"hidden"}}>
          <div style={{width:`${pct}%`,height:"100%",background:pct<70?"#34D399":pct<90?"#FBBF24":"#F87171",borderRadius:999,transition:"width 0.5s"}}/>
        </div>
        <div style={{display:"flex",justifyContent:"space-between",marginTop:4}}>
          <span style={{fontSize:11,color:"rgba(255,255,255,0.7)"}}>{fmtN(pct,0)}% usado</span>
          <span style={{fontSize:11,color:"rgba(255,255,255,0.7)"}}>Disponível: {fmtN((pesoMax-pesoTotal)/1000,2)} kg</span>
        </div>
      </div>

      {pct>=90&&<div style={{background:C.dangerLight,border:`1px solid ${C.danger}33`,borderRadius:12,padding:"10px 14px",fontSize:13,color:C.danger,fontWeight:600,marginBottom:12}}>⚠ Atenção! Mala com {fmtN(pct,0)}% da capacidade.</div>}

      <div style={S.card}>
        <div style={{fontWeight:700,fontSize:13,color:C.text,marginBottom:12}}>Por categoria (itens comprados)</div>
        {porCategoria.map(cat=>{
          const pCat = pesoTotal>0?cat.peso/pesoTotal*100:0;
          return (
            <div key={cat.label} style={{marginBottom:10}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                <span style={{fontSize:13,color:C.text,fontWeight:600}}>{cat.label}</span>
                <span style={{fontSize:12,color:C.textMid,fontFamily:"'DM Mono',monospace"}}>{fmtN(cat.peso/1000,2)} kg</span>
              </div>
              <div style={{height:6,background:C.borderLight,borderRadius:999,overflow:"hidden"}}>
                <div style={{width:`${pCat}%`,height:"100%",background:cat.cor,borderRadius:999,transition:"width 0.5s"}}/>
              </div>
            </div>
          );
        })}
      </div>

      <div style={S.card}>
        <div style={{fontWeight:700,fontSize:13,color:C.text,marginBottom:10}}>Top 5 mais pesados</div>
        {[...produtos.filter(p=>p.status==="comprado")].sort((a,b)=>pesoGramas(b)-pesoGramas(a)).slice(0,5).map(p=>(
          <div key={p.id} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:`1px solid ${C.borderLight}`}}>
            <span style={{fontSize:13,color:C.text,fontWeight:500}}>{p.nome}</span>
            <span style={{fontSize:12,color:C.textMid,fontFamily:"'DM Mono',monospace"}}>{fmtN(prodPeso(p)/1000,3)} kg</span>
          </div>
        ))}
        {produtos.filter(p=>p.status==="comprado").length===0&&<div style={{fontSize:13,color:C.textLight,textAlign:"center",padding:"12px 0"}}>Nenhum item comprado ainda</div>}
      </div>
    </>
  );
}

// ─── CHECKLIST ───────────────────────────────────────────────────────────────
const CHECKLIST_DEFAULTS = [
  {cat:"📋 Documentos",items:["Passaporte válido","Visto americano (ESTA)","Seguro viagem","Cartão de crédito internacional","Comprovante de reserva hotel","Passagens impressas"]},
  {cat:"💊 Saúde",items:["Vacinas em dia","Remédios prescritos","Protetor solar","Repelente"]},
  {cat:"💰 Financeiro",items:["Dólar em espécie","Avisar o banco sobre viagem","Baixar app do banco","Checar limite do cartão"]},
  {cat:"📱 Tecnologia",items:["Adaptador de tomada","Carregadores","Power bank","Câmera/memória","Chip internacional / eSIM"]},
  {cat:"🧳 Mala",items:["Roupas para o clima","Calçado confortável","Necessaire","Cadeado para mala"]},
];

function ChecklistTab({checklist, setChecklist}) {
  const [initialized, setInitialized] = useState(false);
  const [novoTexto, setNovoTexto] = useState("");
  const [novaCat, setNovaCat] = useState("");
  const [showAdd, setShowAdd] = useState(false);

  useEffect(()=>{
    if(checklist.length===0 && !initialized) {
      const defaults = CHECKLIST_DEFAULTS.flatMap(g=>g.items.map(item=>({id:Date.now()+Math.random(),texto:item,cat:g.cat,feito:false})));
      setChecklist(defaults);
      setInitialized(true);
    }
  },[]);

  function toggle(id) { setChecklist(ps=>ps.map(p=>p.id===id?{...p,feito:!p.feito}:p)); }
  function remover(id) { setChecklist(ps=>ps.filter(p=>p.id!==id)); }
  function adicionar() {
    if(!novoTexto.trim()) return;
    const cats = [...new Set(checklist.map(p=>p.cat))];
    const cat = novaCat.trim() || (cats[0] || "📋 Geral");
    setChecklist(ps=>[...ps,{id:Date.now(),texto:novoTexto.trim(),cat,feito:false}]);
    setNovoTexto(""); setShowAdd(false);
  }

  const cats = [...new Set(checklist.map(p=>p.cat))];
  const feitos = checklist.filter(p=>p.feito).length;
  const pct = checklist.length>0?Math.round(feitos/checklist.length*100):0;

  return (
    <>
      <div style={{...S.heroCard,marginBottom:14}}>
        <div style={{fontSize:12,fontWeight:500,color:"rgba(255,255,255,0.75)",marginBottom:3}}>Checklist pré-viagem</div>
        <div style={{fontSize:28,fontWeight:800,color:"#fff"}}>{feitos}/{checklist.length} <span style={{fontSize:14,opacity:0.75}}>itens</span></div>
        <div style={{marginTop:10,background:"rgba(255,255,255,0.2)",borderRadius:999,height:10,overflow:"hidden"}}>
          <div style={{width:`${pct}%`,height:"100%",background:pct===100?"#34D399":"#fff",borderRadius:999,transition:"width 0.5s"}}/>
        </div>
        {pct===100&&<div style={{marginTop:8,fontSize:14,fontWeight:700,color:"#fff"}}>✅ Tudo pronto para a viagem!</div>}
      </div>

      <div style={{display:"flex",gap:8,marginBottom:14}}>
        <button style={{...S.btnPrimary,padding:"9px 14px",fontSize:13,marginBottom:0}} onClick={()=>setShowAdd(s=>!s)}>
          {showAdd?"Cancelar":"＋ Novo item"}
        </button>
      </div>
      {showAdd&&(
        <div style={{...S.card,marginBottom:14}}>
          <label style={S.label}>Item *</label>
          <input style={S.input} placeholder="Ex: Renovar passaporte" value={novoTexto} onChange={e=>setNovoTexto(e.target.value)} onKeyDown={e=>e.key==="Enter"&&adicionar()}/>
          <label style={S.label}>Categoria</label>
          <input style={S.input} list="cats-list" placeholder="Selecione ou crie uma categoria" value={novaCat} onChange={e=>setNovaCat(e.target.value)}/>
          <datalist id="cats-list">{cats.map(c=><option key={c} value={c}/>)}</datalist>
          <button style={S.btnPrimary} onClick={adicionar}>Adicionar</button>
        </div>
      )}
      {cats.map(cat=>{
        const itens = checklist.filter(p=>p.cat===cat);
        const feitosCat = itens.filter(p=>p.feito).length;
        return (
          <div key={cat} style={{...S.card,marginBottom:10}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
              <div style={{fontWeight:700,fontSize:13,color:C.text}}>{cat}</div>
              <span style={{fontSize:11,color:C.textLight,fontWeight:600}}>{feitosCat}/{itens.length}</span>
            </div>
            {itens.map(item=>(
              <div key={item.id} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 0",borderBottom:`1px solid ${C.borderLight}`}}>
                <button onClick={()=>toggle(item.id)} style={{...S.checkbox,...(item.feito?S.checkboxDone:{})}}>
                  {item.feito&&<svg width="12" height="12" viewBox="0 0 12 12"><polyline points="2,6 5,9 10,3" stroke="white" strokeWidth="2" fill="none" strokeLinecap="round"/></svg>}
                </button>
                <span style={{fontSize:13,color:item.feito?C.textLight:C.text,textDecoration:item.feito?"line-through":"none",flex:1}}>{item.texto}</span>
                <button onClick={()=>remover(item.id)} style={{background:"none",border:"none",color:C.textXLight,cursor:"pointer",fontSize:14,padding:"0 4px"}}>✕</button>
              </div>
            ))}
          </div>
        );
      })}
    </>
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
    const qt = parseInt(p.quantidadeParcelas) || 0;
    const mp = parseFloat(p.minhaParte) || 0;
    const vp = parseFloat(p.valorParcela) || 0;
    // minhaParte = total da minha parte; valor mensal = minhaParte / qtd de parcelas
    if (mp > 0 && qt > 0) return parseFloat((mp / qt).toFixed(2));
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
                  {(()=>{const mp=parseFloat(p.minhaParte)||0;const qt=parseInt(p.quantidadeParcelas)||0;const vp=parseFloat(p.valorParcela)||0;return fmtBRL(mp>0?mp:vp*qt,2);})()}
                </td>
              ))}
              <td style={{padding:"8px 10px",textAlign:"right",fontWeight:800,color:C.primary,fontFamily:"'DM Mono',monospace"}}>
                {fmtBRL(parcelasComData.reduce((a,p)=>{ const mp=parseFloat(p.minhaParte)||0; const qt=parseInt(p.quantidadeParcelas)||0; const vp=parseFloat(p.valorParcela)||0; return a+(mp>0?mp:vp*qt); },0),2)}
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
  const mp = parseFloat(p.minhaParte) || 0;
  // Valor mensal da MINHA parte: se tem minhaParte, divide pelo total de parcelas
  const meuValorMensal = mp > 0 && qt > 0 ? parseFloat((mp / qt).toFixed(2)) : valorParcela;

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
              <div style={{fontSize:15,fontWeight:800,color:C.purple,fontFamily:"'DM Mono',monospace"}}>{fmtBRL(meuValorMensal,2)}<span style={{fontSize:10,fontWeight:500,color:C.textLight}}>/mês</span></div>
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
              ⏳ {restante} parcela(s) restante(s) · {fmtBRL((restante*meuValorMensal),2)}
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
function ProdutosTab({produtos,itensLegais,settings,onToggle,onDelete,onEdit,onAdd,onMoveToList,onSubTabChange,onUpdate}) {
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
      {filtered.map(p=><ProdutoCard key={p.id} p={p} settings={settings} onToggle={subTab==="compras"?()=>onToggle(p.id):null} onDelete={()=>onDelete(p.id,subTab==="legais"?"legais":"produtos")} onEdit={()=>onEdit({...p,_legais:subTab==="legais"})} onMoveToList={subTab==="legais"?()=>onMoveToList(p):null} isLegais={subTab==="legais"} onUpdate={onUpdate}/>)}
      {filtered.length===0&&lista.length>0&&<Empty text="Nenhum item encontrado"/>}
    </div>
  );
}

function ProdutoCard({p,settings,onToggle,onDelete,onEdit,onMoveToList,isLegais,onUpdate}) {
  const [expanded,setExpanded]=useState(false);
  const isComprado=p.status==="comprado";
  const qtdC=isComprado?(parseInt(p.qtdComprada)||1):0;
  const usdUnit=parseFloat(p.usd)||0;
  const usdTotal=usdComTaxa(p)*Math.max(1,qtdC);
  const brl=usdUnit*calcDolarAjustado(settings);
  const brlPago=p.dollarPago?calcBRLPago(usdUnit,settings,p.dollarPago):null;
  const peso=prodPeso(p);
  const prioColors={Alta:{color:C.danger,bg:C.dangerLight},Média:{color:C.warning,bg:C.warningLight},Baixa:{color:C.primary,bg:C.primaryLight}};
  const pc=prioColors[p.prioridade]||prioColors["Média"];
  return (
    <div style={{...S.card,marginBottom:10,border:p.status==="comprado"?`1px solid ${C.success}44`:undefined}}>
      <div style={{display:"flex",gap:10,alignItems:"flex-start"}}>
        <div style={{width:50,height:50,borderRadius:12,overflow:"hidden",flexShrink:0,border:`1px solid ${C.border}`}}>
          <ProductImage produto={p} iconSize={22}/>
        </div>
        {!isLegais&&onToggle&&<button style={{...S.checkbox,...(p.status==="comprado"?S.checkboxDone:{})}} onClick={onToggle}>{p.status==="comprado"&&<svg width="12" height="12" viewBox="0 0 12 12"><polyline points="2,6 5,9 10,3" stroke="white" strokeWidth="2" fill="none" strokeLinecap="round"/></svg>}</button>}
        <div style={{flex:1}} onClick={()=>setExpanded(e=>!e)}>
          <div style={{display:"flex",justifyContent:"space-between",gap:8}}>
            <div style={{fontWeight:600,fontSize:14,color:p.status==="comprado"?C.textLight:C.text,textDecoration:p.status==="comprado"?"line-through":"none",lineHeight:1.3,flex:1}}>{p.nome}</div>
            <div style={{textAlign:"right",flexShrink:0}}>
              <div style={{fontSize:14,fontWeight:800,color:C.primary,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(usdUnit,2)}</div>
              <div style={{fontSize:11,color:C.textLight,fontFamily:"'DM Mono',monospace"}}>{fmtBRL(brl,0)}</div>
            </div>
          </div>
          <div style={{display:"flex",gap:5,marginTop:6,flexWrap:"wrap"}}>
            <span style={S.tag}>{p.loja}</span>
            {!isLegais&&<span style={{...S.tag,background:pc.bg,color:pc.color,borderColor:pc.color+"33"}}>{p.prioridade}</span>}
            <span style={S.tag}>{(peso/1000).toLocaleString("pt-BR",{minimumFractionDigits:3,maximumFractionDigits:3})}kg</span>
            {p.peso>0&&<span style={S.tag}>{p.peso} {p.tipoPeso==="oz_peso"?"Peso Oz":p.tipoPeso==="oz_liquido"?"Líquido Oz":"G"}</span>}
            {p.status==="comprado"&&<span style={{...S.tag,background:C.successLight,color:C.success,borderColor:C.success+"33"}}>✓ Comprado</span>}
            {p.localTaxa&&p.localTaxa!=="isento"&&<span style={{...S.tag,background:C.purpleLight,color:C.purple,borderColor:C.purple+"33"}}>{p.localTaxa==="orlando"?"ORL 6,5%":"KIS 7,5%"}</span>}
          </div>
        </div>
      </div>

      {/* Controles visíveis apenas quando comprado */}
      {isComprado&&<div style={{marginTop:10,padding:"10px 12px",background:C.bg,borderRadius:10,border:`1px solid ${C.success}44`}}>
        {/* Seletor de taxa */}
        <div style={{marginBottom:8}}>
          <div style={{fontSize:10,fontWeight:700,color:C.textLight,textTransform:"uppercase",letterSpacing:"0.5px",marginBottom:6}}>💰 Taxa local</div>
          <div style={{display:"flex",gap:5}}>
            {[{v:"isento",l:"Sem taxa",c:"#64748B"},{v:"orlando",l:"Orlando 6,5%",c:C.primary},{v:"kissimmee",l:"Kissimmee 7,5%",c:C.purple}].map(({v,l,c})=>{
              const active=(p.localTaxa||"isento")===v;
              return <button key={v} onClick={e=>{e.stopPropagation();onUpdate&&onUpdate(p.id,{localTaxa:v});}} style={{flex:1,padding:"5px 2px",borderRadius:7,border:`1.5px solid ${active?c:C.border}`,background:active?c:"transparent",color:active?"#fff":C.textLight,fontSize:9,fontWeight:700,cursor:"pointer"}}>{l}</button>;
            })}
          </div>
        </div>
        {/* Controle de quantidade */}
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",paddingTop:8,borderTop:`1px dashed ${C.border}`}}>
          <span style={{fontSize:12,fontWeight:600,color:C.textMid}}>Qtd comprada:</span>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <button onClick={e=>{e.stopPropagation();const n=Math.max(0,qtdC-1);onUpdate&&onUpdate(p.id,{qtdComprada:n,status:n>0?"comprado":"pendente"});}} style={{width:30,height:30,borderRadius:"50%",border:"none",background:C.dangerLight,color:C.danger,fontSize:18,fontWeight:700,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>−</button>
            <span style={{fontSize:16,fontWeight:800,color:C.success,minWidth:24,textAlign:"center",fontFamily:"'DM Mono',monospace"}}>{qtdC}</span>
            <button onClick={e=>{e.stopPropagation();const n=qtdC+1;onUpdate&&onUpdate(p.id,{qtdComprada:n,status:"comprado"});}} style={{width:30,height:30,borderRadius:"50%",border:"none",background:C.successLight,color:C.success,fontSize:18,fontWeight:700,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>＋</button>
          </div>
          {qtdC>0&&usdTotal>0&&<span style={{fontSize:11,color:C.textMid,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(usdTotal,2)} total</span>}
        </div>
      </div>}
      {expanded&&(
        <div style={{marginTop:12,paddingTop:12,borderTop:`1px solid ${C.borderLight}`}}>
          {[{label:"BRL previsto",value:`${fmtBRL(brl,2)}`},...(brlPago?[{label:"BRL pago",value:`${fmtBRL(brlPago,2)}`,color:C.success},{label:"Diferença",value:`${fmtBRL((brl-brlPago),2)}`,color:brl>brlPago?C.success:C.danger}]:[]),...(qtdC>1?[{label:"USD unitário",value:fmtUSD(p.usd,2),color:C.textMid},{label:`USD total (×${qtdC})`,value:fmtUSD(usdTotal,2),color:C.primary}]:[{label:"USD c/ taxa",value:`${fmtUSD(calcUsdFinal(usdTotal,settings),2)}`,color:C.textMid}])].map(({label,value,color})=>(
            <div key={label} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:`1px solid ${C.borderLight}`}}>
              <span style={{fontSize:13,color:C.textMid}}>{label}</span>
              <span style={{fontSize:13,fontWeight:700,color:color||C.text,fontFamily:"'DM Mono',monospace"}}>{value}</span>
            </div>
          ))}
          {(p.quantidade||1)>1&&p.status==="comprado"&&(
            <div style={{marginTop:8,marginBottom:4}}>
              <div style={{fontSize:12,fontWeight:700,color:C.textMid,marginBottom:6}}>Quantos foram comprados?</div>
              <div style={{display:"flex",alignItems:"center",gap:8}}>
                <button onClick={e=>{e.stopPropagation();onEdit({...p,qtdComprada:Math.max(0,(p.qtdComprada||p.quantidade)-1)});}} style={{width:32,height:32,borderRadius:8,border:`1px solid ${C.border}`,background:C.bg,fontSize:16,cursor:"pointer"}}>−</button>
                <span style={{fontSize:16,fontWeight:700,color:C.primary,minWidth:60,textAlign:"center",fontFamily:"'DM Mono',monospace"}}>{p.qtdComprada||p.quantidade}/{p.quantidade}</span>
                <button onClick={e=>{e.stopPropagation();onEdit({...p,qtdComprada:Math.min(p.quantidade,(p.qtdComprada||p.quantidade)+1)});}} style={{width:32,height:32,borderRadius:8,border:`1px solid ${C.border}`,background:C.bg,fontSize:16,cursor:"pointer"}}>＋</button>
              </div>
            </div>
          )}
          {/* Seletor de taxa local */}
          {!isLegais&&<div style={{marginTop:10,padding:"10px 12px",background:C.bg,borderRadius:10,border:`1px solid ${C.border}`}}>
            <div style={{fontSize:11,fontWeight:700,color:C.textLight,marginBottom:8,textTransform:"uppercase",letterSpacing:"0.5px"}}>💰 Taxa local</div>
            <div style={{display:"flex",gap:6}}>
              {[{v:"isento",l:"Sem taxa",bg:C.borderLight,col:C.textMid,act:C.textMid,actBg:C.bg},{v:"orlando",l:"Orlando 6,5%",bg:C.primaryLight,col:C.primary,act:C.primary,actBg:C.primaryLight},{v:"kissimmee",l:"Kissimmee 7,5%",bg:C.purpleLight,col:C.purple,act:C.purple,actBg:C.purpleLight}].map(({v,l,actBg,act,bg,col})=>{
                const active=(p.localTaxa||"isento")===v;
                return <button key={v} onClick={e=>{e.stopPropagation();onUpdate&&onUpdate(p.id,{localTaxa:v});}} style={{flex:1,padding:"6px 4px",borderRadius:8,border:`1.5px solid ${active?act:C.border}`,background:active?actBg:C.bgCard,color:active?act:C.textLight,fontSize:10,fontWeight:700,cursor:"pointer",transition:"all 0.15s"}}>{l}</button>;
              })}
            </div>
            {(p.localTaxa==="orlando"||p.localTaxa==="kissimmee")&&<div style={{fontSize:11,color:C.textMid,marginTop:6,fontFamily:"'DM Mono',monospace"}}>
              US$ {fmtN(usdComTaxa(p),2)} c/ imposto · {fmtBRL(usdComTaxa(p)*prodQtd(p)*calcDolarAjustado(settings))} total
            </div>}
          </div>}
          {/* Controle de quantidade comprada */}
          {!isLegais&&(p.quantidade||1)>1&&<div style={{marginTop:8,display:"flex",alignItems:"center",justifyContent:"space-between",padding:"8px 12px",background:C.bg,borderRadius:10,border:`1px solid ${C.border}`}}>
            <span style={{fontSize:13,fontWeight:600,color:C.textMid}}>Qtd comprada:</span>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <button onClick={e=>{e.stopPropagation();const n=Math.max(0,(p.qtdComprada||0)-1);onUpdate&&onUpdate(p.id,{qtdComprada:n,status:n>0?"comprado":"pendente"});}} style={{width:30,height:30,borderRadius:"50%",border:"none",background:C.dangerLight,color:C.danger,fontSize:18,fontWeight:700,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>−</button>
              <span style={{fontSize:16,fontWeight:800,color:(p.qtdComprada||0)>0?C.success:C.textLight,minWidth:50,textAlign:"center",fontFamily:"'DM Mono',monospace"}}>{p.qtdComprada||0}/{p.quantidade}</span>
              <button onClick={e=>{e.stopPropagation();const n=Math.min(p.quantidade||1,(p.qtdComprada||0)+1);onUpdate&&onUpdate(p.id,{qtdComprada:n,status:"comprado"});}} style={{width:30,height:30,borderRadius:"50%",border:"none",background:C.successLight,color:C.success,fontSize:18,fontWeight:700,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>＋</button>
            </div>
          </div>}
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
function GaleriaTab({produtos,itensLegais,settings,onEdit,onToggle,onToggleLegal,onUpdate,onUpdateLegal}) {
  const [subTab,setSubTab]=useState("compras");
  const [selected,setSelected]=useState(null);
  const lista=subTab==="legais"?itensLegais:produtos;
  const isLegais=subTab==="legais";
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
          <div key={p.id} style={{...S.card,padding:0,overflow:"hidden",cursor:"pointer"}} onClick={()=>setSelected(p)} className="galeria-card">
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
      {selected&&<GaleriaDetailModal
        p={selected}
        settings={settings}
        isLegais={isLegais}
        onClose={()=>setSelected(null)}
        onToggle={()=>{ (isLegais?onToggleLegal:onToggle)(selected.id); setSelected(s=>s?{...s,status:s.status==="comprado"?"pendente":"comprado",qtdComprada:s.status==="comprado"?0:(parseInt(s.qtdComprada)||1)}:s); }}
        onUpdate={(id,campos)=>{ (isLegais?onUpdateLegal:onUpdate)(id,campos); setSelected(s=>s?{...s,...campos}:s); }}
        onEditFull={()=>{ onEdit({...selected,_legais:isLegais}); setSelected(null); }}
      />}
    </div>
  );
}

function GaleriaDetailModal({p,settings,isLegais,onClose,onToggle,onUpdate,onEditFull}) {
  const [link,setLink]=useState(p.link||"");
  const [fullscreen,setFullscreen]=useState(false);
  const isComprado=p.status==="comprado";
  const qtdC=isComprado?(parseInt(p.qtdComprada)||1):0;
  const usdUnit=parseFloat(p.usd)||0;
  const usdTotal=usdComTaxa(p)*Math.max(1,qtdC);
  const brl=usdUnit*calcDolarAjustado(settings);
  return (
    <Modal title={p.nome} onClose={onClose}>
      <div style={{borderRadius:14,overflow:"hidden",border:`1px solid ${C.border}`,marginBottom:14,height:220,cursor:"zoom-in"}} onClick={()=>setFullscreen(true)}>
        <ProductImage produto={p} iconSize={48}/>
      </div>

      {fullscreen&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.9)",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center"}} onClick={()=>setFullscreen(false)}>
          <button style={{position:"absolute",top:16,right:16,background:"rgba(255,255,255,0.15)",border:"none",color:"#fff",width:36,height:36,borderRadius:"50%",fontSize:16,cursor:"pointer"}} onClick={()=>setFullscreen(false)}>✕</button>
          <div style={{width:"92vw",height:"80vh",maxWidth:600}}>
            <ProductImage produto={p} iconSize={64} style={{background:"transparent"}} fit="contain"/>
          </div>
        </div>
      )}

      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
        <div>
          <div style={{fontSize:20,fontWeight:800,color:C.primary,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(usdUnit,2)}</div>
          <div style={{fontSize:12,color:C.textLight,fontFamily:"'DM Mono',monospace"}}>{fmtBRL(brl,2)}</div>
        </div>
        <button onClick={onToggle} style={{display:"flex",alignItems:"center",gap:8,padding:"9px 14px",borderRadius:10,border:`1.5px solid ${isComprado?C.success:C.border}`,background:isComprado?C.successLight:C.bg,color:isComprado?C.success:C.textMid,fontWeight:700,fontSize:13,cursor:"pointer"}}>
          <span style={{...S.checkbox,...(isComprado?S.checkboxDone:{}),width:18,height:18}}>{isComprado&&<svg width="10" height="10" viewBox="0 0 12 12"><polyline points="2,6 5,9 10,3" stroke="white" strokeWidth="2" fill="none" strokeLinecap="round"/></svg>}</span>
          {isComprado?"Comprado":"Marcar comprado"}
        </button>
      </div>

      {/* Quantidade */}
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"10px 12px",background:C.bg,borderRadius:10,border:`1px solid ${C.border}`,marginBottom:12}}>
        <span style={{fontSize:13,fontWeight:600,color:C.textMid}}>Qtd comprada:</span>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <button onClick={()=>{const n=Math.max(0,qtdC-1);onUpdate(p.id,{qtdComprada:n,status:n>0?"comprado":"pendente"});}} style={{width:32,height:32,borderRadius:"50%",border:"none",background:C.dangerLight,color:C.danger,fontSize:18,fontWeight:700,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>−</button>
          <span style={{fontSize:17,fontWeight:800,color:qtdC>0?C.success:C.textLight,minWidth:24,textAlign:"center",fontFamily:"'DM Mono',monospace"}}>{qtdC}</span>
          <button onClick={()=>{const n=qtdC+1;onUpdate(p.id,{qtdComprada:n,status:"comprado"});}} style={{width:32,height:32,borderRadius:"50%",border:"none",background:C.successLight,color:C.success,fontSize:18,fontWeight:700,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>＋</button>
        </div>
        {qtdC>0&&<span style={{fontSize:12,color:C.textMid,fontFamily:"'DM Mono',monospace"}}>{fmtUSD(usdTotal,2)} total</span>}
      </div>

      {/* Taxa local */}
      <div style={{marginBottom:14}}>
        <div style={S.label}>💰 Taxa local</div>
        <div style={{display:"flex",gap:6}}>
          {[{v:"isento",l:"Sem taxa",c:"#64748B",bg:C.borderLight},{v:"orlando",l:"Orlando 6,5%",c:C.primary,bg:C.primaryLight},{v:"kissimmee",l:"Kissimmee 7,5%",c:C.purple,bg:C.purpleLight}].map(({v,l,c,bg})=>{
            const active=(p.localTaxa||"isento")===v;
            return <button key={v} onClick={()=>onUpdate(p.id,{localTaxa:v})} style={{flex:1,padding:"8px 4px",borderRadius:9,border:`1.5px solid ${active?c:C.border}`,background:active?bg:C.bgCard,color:active?c:C.textLight,fontSize:11,fontWeight:700,cursor:"pointer"}}>{l}</button>;
          })}
        </div>
        {(p.localTaxa==="orlando"||p.localTaxa==="kissimmee")&&<div style={{fontSize:12,color:C.textMid,marginTop:6,fontFamily:"'DM Mono',monospace"}}>US$ {fmtN(usdComTaxa(p),2)} c/ imposto</div>}
      </div>

      {/* Link do produto */}
      <div style={{marginBottom:14}}>
        <div style={S.label}>🔗 Link do produto</div>
        <input style={{...S.input,marginBottom:6}} type="url" placeholder="https://amazon.com/..." value={link} onChange={e=>setLink(e.target.value)} onBlur={()=>onUpdate(p.id,{link})}/>
        {p.link&&<a href={p.link} target="_blank" rel="noreferrer" style={{fontSize:13,color:C.primary}}>Abrir link ↗</a>}
      </div>

      <div style={{display:"flex",gap:8}}>
        <button style={S.btnOutline} onClick={onEditFull}>✏ Editar tudo</button>
        <button style={{...S.btnOutline,flex:1}} onClick={onClose}>Fechar</button>
      </div>
    </Modal>
  );
}

// ─── STATS TAB ────────────────────────────────────────────────────────────────
function StatsTab({produtos,gastos,settings,checklist,setChecklist}) {
  const [secao,setSecao]=useState("compras");
  const barColors=[C.primary,C.purple,C.success,C.warning,C.danger,"#06B6D4","#F97316","#EC4899"];

  // ── Compras stats ──────────────────────────────────────────────────────────
  const porLoja=useMemo(()=>{
    const map={};
    produtos.forEach(p=>{
      if(!map[p.loja])map[p.loja]={total:0,comprados:0,usd:0,brl:0,peso:0};
      map[p.loja].total++;
      if(p.status==="comprado")map[p.loja].comprados++;
      map[p.loja].usd+=prodUSD(p);
      map[p.loja].brl+=calcBRLProduto(p,settings);
      map[p.loja].peso+=prodPeso(p);
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
      const minha=calcMinhaParteUSD(g,produtos);
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
  const totalMeuUSD=gastos.reduce((a,g)=>a+calcMinhaParteUSD(g,produtos),0);
  const totalAReceberUSD=gastos.reduce((a,g)=>{
    if(!g.divisao||g.divisao.length===0) return a;
    return a+g.divisao.filter(p=>!p.pago).reduce((s,p)=>s+(parseFloat(p.valor)||0),0);
  },0);

  return (
    <div style={S.page}>
      {/* Sub-tabs */}
      <div style={{display:"flex",gap:4,background:C.borderLight,borderRadius:12,padding:4,marginBottom:14}}>
        {[["compras","🛒 Compras"],["gastos","💸 Gastos"],["checklist","✅ Checklist"]].map(([v,l])=>(
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

      {/* ── CHECKLIST ── */}
      {secao==="checklist"&&<ChecklistTab checklist={checklist} setChecklist={setChecklist}/>}

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
              const minha=calcMinhaParteUSD(g,produtos);
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


// ─── SETTINGS MODAL ───────────────────────────────────────────────────────────
function SettingsModal({settings,onSave,onImport,onExport,onClose}) {
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
          <div style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:12,padding:14,marginBottom:14}}>
            <div style={{fontSize:13,fontWeight:700,color:C.text,marginBottom:8,display:"flex",alignItems:"center",gap:6}}>📊 Backup</div>
            <div style={{fontSize:12,color:C.textMid,marginBottom:10}}>Exporte seus dados atuais para uma planilha Excel.</div>
            <button style={S.btnPrimary} onClick={onExport}>📥 Exportar planilha atual</button>
          </div>
          <div style={{borderTop:`1px dashed ${C.border}`,margin:"4px 0 14px"}}></div>
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
  const empty={nome:"",loja:"Walmart",usd:"",peso:"",tipoPeso:"g",status:"pendente",prioridade:"Média",link:"",imagem:"",dollarPago:"",quantidade:1,_legais:isL};
  const [f,setF]=useState(prod?.id?{...prod,usd:prod.usd.toString(),peso:(prod.peso||"").toString(),tipoPeso:prod.tipoPeso||"g",dollarPago:(prod.dollarPago||"").toString(),quantidade:prod.quantidade||1,_legais:prod._legais||isL}:empty);
  function save(){if(!f.nome||!f.usd)return alert("Preencha nome e USD");onSave({...f,usd:parseFloat(f.usd),peso:parseFloat(f.peso)||0,tipoPeso:f.tipoPeso||"g",dollarPago:f.dollarPago?parseFloat(f.dollarPago):null,quantidade:parseInt(f.quantidade)||1});}
  return (
    <Modal title={prod?.id?"Editar produto":isL?"✨ Item legal":"Novo produto"} onClose={onClose}>
      <label style={S.label}>Nome *</label>
      <input style={S.input} placeholder="Ex: AirPods Pro" value={f.nome} onChange={e=>setF(p=>({...p,nome:e.target.value}))}/>
      <label style={S.label}>Quantidade</label>
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14}}>
        <button onClick={()=>setF(p=>({...p,quantidade:Math.max(1,(parseInt(p.quantidade)||1)-1)}))} style={{width:36,height:36,borderRadius:9,border:`1px solid ${C.border}`,background:C.bg,fontSize:18,cursor:"pointer",color:C.text,flexShrink:0,fontFamily:"'Inter',sans-serif"}}>−</button>
        <input style={{...S.input,marginBottom:0,textAlign:"center",fontWeight:700,fontSize:16}} type="number" min="1" value={f.quantidade||1} onChange={e=>setF(p=>({...p,quantidade:Math.max(1,parseInt(e.target.value)||1)}))}/>
        <button onClick={()=>setF(p=>({...p,quantidade:(parseInt(p.quantidade)||1)+1}))} style={{width:36,height:36,borderRadius:9,border:`1px solid ${C.border}`,background:C.bg,fontSize:18,cursor:"pointer",color:C.text,flexShrink:0,fontFamily:"'Inter',sans-serif"}}>＋</button>
      </div>
      <label style={S.label}>Loja</label>
      <input style={S.input} list="lojas-list" placeholder="Digite ou escolha uma loja..." value={f.loja} onChange={e=>setF(p=>({...p,loja:e.target.value}))}/>
      <datalist id="lojas-list">{LOJAS_SUGESTOES.map(l=><option key={l} value={l}/>)}</datalist>
      <label style={S.label}>Preço USD *</label>
      <input style={S.input} type="number" placeholder="Ex: 199" value={f.usd} onChange={e=>setF(p=>({...p,usd:e.target.value}))}/>
      <div style={{ marginBottom: 14 }}>
        <label style={S.label}>Peso / Volume</label>
        <div style={{ display: "flex", gap: 8 }}>
          <input
            style={{ ...S.input, marginBottom: 0, flex: 1 }}
            type="number"
            placeholder="0"
            value={f.peso}
            onChange={e=>setF(p=>({...p,peso:e.target.value}))}
          />
          <select
            value={f.tipoPeso || "g"}
            onChange={e=>setF(p=>({...p,tipoPeso:e.target.value}))}
            style={{ background:C.bg, border:`1px solid ${C.border}`, borderRadius:9, padding:"0 10px", color:C.text, fontSize:14, outline:"none", fontFamily:"'Inter',sans-serif" }}
          >
            <option value="g">Peso G</option>
            <option value="oz_peso">Peso Oz</option>
            <option value="oz_liquido">Líquido Oz</option>
          </select>
        </div>
        {f.peso&&<div style={{fontSize:12,color:C.textLight,marginTop:6}}>= {(pesoGramas({peso:f.peso,tipoPeso:f.tipoPeso||"g"})/1000).toLocaleString("pt-BR",{minimumFractionDigits:3})} kg</div>}
      </div>
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
