import { useState, useMemo, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { getProductImage, imageCache } from './imageService';
import { db } from "./firebase";
import { doc, onSnapshot, setDoc, collection, deleteDoc, serverTimestamp } from "firebase/firestore";
import { getAuth, onAuthStateChanged, signInWithEmailAndPassword, createUserWithEmailAndPassword, signOut } from "firebase/auth";

// ─── FIREBASE AUTH / FIRESTORE PERSISTENCE ────────────────────────────────
const auth = getAuth();

const normalizeCloudState = data => ({
  settings: data?.settings || INITIAL_SETTINGS,
  produtos: Array.isArray(data?.produtos) ? data.produtos : SAMPLE_PRODUTOS,
  itensLegais: Array.isArray(data?.itensLegais) ? data.itensLegais : [],
  gastos: Array.isArray(data?.gastos) ? data.gastos : [],
});

async function saveCloudState(userDocRef, state) {
  await setDoc(userDocRef, {
    ...state,
    updatedAt: serverTimestamp(),
  }, { merge: true });
}

// ─── COLORS ──────────────────────────────────────────────────────────────────
const C = {
  bg:"#F8FAFC",bgCard:\"#FFFFFF\",border:\"#E5E7EB\",borderLight:\"#F1F5F9\",
  primary:\"#2563EB\",primaryLight:\"#EFF6FF\",gradientA:\"#2563EB\",gradientB:\"#06B6D4\",
  success:\"#10B981\",successLight:\"#ECFDF5\",warning:\"#F59E0B\",warningLight:\"#FFFBEB\",
  danger:\"#EF4444\",dangerLight:\"#FEF2F2\",purple:\"#8B5CF6\",purpleLight:\"#F5F3FF\",
  text:\"#0F172A\",textMid:\"#475569\",textLight:\"#94A3B8\",textXLight:\"#E2E8F0\",
};

// ─── CONSTANTS & INITIAL STATES ──────────────────────────────────────────────
const CATEGORIES = [
  { id: "compras", label: "Compras/Lojas", icon: "🛍️", color: C.primary },
  { id: "alimentacao", label: "Alimentação", icon: "🍔", color: C.warning },
  { id: "hospedagem", label: "Hospedagem", icon: "🏨", color: C.purple },
  { id: "transporte", label: "Transporte/Gasolina", icon: "🚗", color: C.success },
  { id: "lazer", label: "Lazer/Ingressos", icon: "🎟️", color: C.gradientB },
  { id: "outros", label: "Outros", icon: "📦", color: C.textMid }
];

const INITIAL_SETTINGS = { taxRate: 6.5, exchangeRate: 5.80, budgetUSD: 2000, pyrUSD: 0 };

const SAMPLE_PRODUTOS = [
  { id: 1, nome: "Exemplo de Produto", local: "Walmart", valorUSD: 10, qtd: 1, categoria: "compras", link: "", nota: "" }
];

// ─── TRANSLATION MAP ─────────────────────────────────────────────────────────
const CAT_TRANSLATIONS = {
  "compras": "Compras", "alimentacao": "Alimentação", "hospedagem": "Hospedagem",
  "transporte": "Transporte", "lazer": "Lazer", "outros": "Outros"
};

export default function App() {
  // --- Auth & Sync States ---
  const [user, setUser] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [authEmail, setAuthEmail] = useState("");
  const [authPassword, setAuthPassword] = useState("");
  const [authIsSignUp, setAuthIsSignUp] = useState(false);
  const [authError, setAuthError] = useState("");

  // --- Core Application States ---
  const [activeTab, setActiveTab] = useState(0); // 0:Produtos, 1:Itens Legais, 2:Gastos, 3:Resumo, 4:Parcelas, 5:Dashboard, 6:Calculadora
  const [settings, setSettings] = useState(INITIAL_SETTINGS);
  const [produtos, setProdutos] = useState([]);
  const [itensLegais, setItensLegais] = useState([]);
  const [gastos, setGastos] = useState([]);
  const [parcelas, setParcelas] = useState([]);
  const [parcelasLoading, setParcelasLoading] = useState(true);

  // --- UI Control States ---
  const [search, setSearch] = useState("");
  const [catFilter, setCatFilter] = useState("all");
  const [lojaFilter, setLojaFilter] = useState("all");
  const [editingItem, setEditingItem] = useState(null);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [modalType, setModalType] = useState("produto"); // produto, itemLegal, gasto, parcela
  const [notification, setNotification] = useState(null);

  // Reference for the cloud document
  const userDocRefRef = useRef(null);
  const isInitialLoadRef = useRef(true);

  // --- Auth Observer ---
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (currentUser) => {
      setUser(currentUser);
      setAuthLoading(false);
      if (!currentUser) {
        userDocRefRef.current = null;
        isInitialLoadRef.current = true;
        setProdutos([]);
        setItensLegais([]);
        setGastos([]);
        setParcelas([]);
      }
    });
    return unsubscribe;
  }, []);

  // --- Firestore Real-time Sync (Core Data) ---
  useEffect(() => {
    if (!user) return;
    const docRef = doc(db, "usuarios_pwa", user.uid);
    userDocRefRef.current = docRef;

    const unsubscribe = onSnapshot(docRef, (docSnap) => {
      if (docSnap.exists()) {
        const cloudData = normalizeCloudState(docSnap.data());
        setSettings(cloudData.settings);
        setProdutos(cloudData.produtos);
        setItensLegais(cloudData.itensLegais);
        setGastos(cloudData.gastos);
      } else {
        setSettings(INITIAL_SETTINGS);
        setProdutos(SAMPLE_PRODUTOS);
        setItensLegais([]);
        setGastos([]);
        saveCloudState(docRef, {
          settings: INITIAL_SETTINGS,
          produtos: SAMPLE_PRODUTOS,
          itensLegais: [],
          gastos: []
        });
      }
      isInitialLoadRef.current = false;
    }, (error) => {
      console.error("Firestore sync error:", error);
      showNotification("Erro ao sincronizar com a nuvem", "danger");
    });

    return unsubscribe;
  }, [user]);

  // --- Firestore Real-time Sync (Parcelas) ---
  useEffect(() => {
    if (!user) return;
    const parcelasCollRef = collection(db, "usuarios_pwa", user.uid, "parcelas");

    const unsubscribe = onSnapshot(parcelasCollRef, (snap) => {
      const docs = snap.docs.map(d => ({ id: d.id, ...d.data() }));
      setParcelas(docs);
      setParcelasLoading(false);
    }, (error) => {
      console.error("Error fetching parcelas:", error);
      setParcelasLoading(false);
    });

    return unsubscribe;
  }, [user]);

  // --- Trigger Cloud Saves ---
  const triggerCloudSave = useCallback((updatedFields) => {
    if (!userDocRefRef.current || isInitialLoadRef.current) return;
    
    const currentState = {
      settings,
      produtos,
      itensLegais,
      gastos,
      ...updatedFields
    };
    
    saveCloudState(userDocRefRef.current, currentState).catch(err => {
      console.error("Cloud save failed:", err);
      showNotification("Erro ao salvar dados na nuvem", "danger");
    });
  }, [settings, produtos, itensLegais, gastos]);

  const saveParcelaToCloud = async (parcelaData) => {
    if (!user) return;
    const id = parcelaData.id || "par_" + Date.now();
    const docRef = doc(db, "usuarios_pwa", user.uid, "parcelas", id);
    await setDoc(docRef, { ...parcelaData, id }, { merge: true });
    showNotification("Parcela salva na nuvem!", "success");
  };

  const deleteParcelaFromCloud = async (id) => {
    if (!user) return;
    const docRef = doc(db, "usuarios_pwa", user.uid, "parcelas", id);
    await deleteDoc(docRef);
    showNotification("Parcela removida!", "warning");
  };

  // --- UI Notification Helper ---
  const showNotification = (text, type = "success") => {
    setNotification({ text, type });
    setTimeout(() => setNotification(null), 3000);
  };

  // --- Auth Handlers ---
  const handleAuthSubmit = async (e) => {
    e.preventDefault();
    setAuthError("");
    if (!authEmail || !authPassword) {
      setAuthError("Preencha todos os campos.");
      return;
    }
    try {
      if (authIsSignUp) {
        await createUserWithEmailAndPassword(auth, authEmail, authPassword);
        showNotification("Conta criada com sucesso!", "success");
      } else {
        await signInWithEmailAndPassword(auth, authEmail, authPassword);
        showNotification("Login realizado!", "success");
      }
    } catch (err) {
      console.error(err);
      if (err.code === "auth/wrong-password") setAuthError("Senha incorreta.");
      else if (err.code === "auth/user-not-found") setAuthError("Usuário não encontrado.");
      else if (err.code === "auth/email-already-in-use") setAuthError("E-mail já cadastrado.");
      else setAuthError("Erro na autenticação. Verifique os dados.");
    }
  };

  const handleLogout = async () => {
    try {
      await signOut(auth);
      showNotification("Você saiu do app.", "warning");
    } catch (err) {
      console.error(err);
    }
  };

  // --- Actions ---
  const updateSettings = (key, value) => {
    const val = parseFloat(value) || 0;
    const next = { ...settings, [key]: val };
    setSettings(next);
    if (userDocRefRef.current) {
      saveCloudState(userDocRefRef.current, { settings: next, produtos, itensLegais, gastos });
    }
  };

  const handleSave = (item) => {
    if (modalType === "produto") {
      let next;
      if (item.id) {
        next = produtos.map(p => p.id === item.id ? item : p);
      } else {
        next = [...produtos, { ...item, id: Date.now() }];
      }
      setProdutos(next);
      triggerCloudSave({ produtos: next });
      showNotification(item.id ? "Produto atualizado!" : "Produto adicionado!");
    } else if (modalType === "itemLegal") {
      let next;
      if (item.id) {
        next = itensLegais.map(i => i.id === item.id ? item : i);
      } else {
        next = [...itensLegais, { ...item, id: Date.now() }];
      }
      setItensLegais(next);
      triggerCloudSave({ itensLegais: next });
      showNotification(item.id ? "Item atualizado!" : "Item adicionado!");
    } else if (modalType === "gasto") {
      let next;
      if (item.id) {
        next = gastos.map(g => g.id === item.id ? item : g);
      } else {
        next = [...gastos, { ...item, id: Date.now() }];
      }
      setGastos(next);
      triggerCloudSave({ gastos: next });
      showNotification(item.id ? "Gasto atualizado!" : "Gasto adicionado!");
    } else if (modalType === "parcela") {
      saveParcelaToCloud(item);
    }
    setIsModalOpen(false);
    setEditingItem(null);
  };

  const handleDelete = (id, type) => {
    if (window.confirm("Deseja realmente excluir este item?")) {
      if (type === "produto") {
        const next = produtos.filter(p => p.id !== id);
        setProdutos(next);
        triggerCloudSave({ produtos: next });
        showNotification("Produto removido", "warning");
      } else if (type === "itemLegal") {
        const next = itensLegais.filter(i => i.id !== id);
        setItensLegais(next);
        triggerCloudSave({ itensLegais: next });
        showNotification("Item removido", "warning");
      } else if (type === "gasto") {
        const next = gastos.filter(g => g.id !== id);
        setGastos(next);
        triggerCloudSave({ gastos: next });
        showNotification("Gasto removido", "warning");
      } else if (type === "parcela") {
        deleteParcelaFromCloud(id);
      }
    }
  };

  // --- Financial Computations ---
  const taxMultiplier = useMemo(() => 1 + (settings.taxRate / 100), [settings.taxRate]);

  const totals = useMemo(() => {
    let prodUSD = 0;
    produtos.forEach(p => {
      const v = (p.valorUSD || 0) * (p.qtd || 1);
      prodUSD += p.categoria === "compras" ? v * taxMultiplier : v;
    });

    let legalUSD = 0;
    itensLegais.forEach(i => {
      const v = (i.valorUSD || 0) * (i.qtd || 1);
      legalUSD += i.categoria === "compras" ? v * taxMultiplier : v;
    });

    let gastosUSD = 0;
    gastos.forEach(g => {
      gastosUSD += (g.valorUSD || 0);
    });

    const totalUSD = prodUSD + legalUSD + gastosUSD;
    const totalBRL = totalUSD * settings.exchangeRate;
    const remainingUSD = settings.budgetUSD - totalUSD;

    return { prodUSD, legalUSD, gastosUSD, totalUSD, totalBRL, remainingUSD };
  }, [produtos, itensLegais, gastos, taxMultiplier, settings.exchangeRate, settings.budgetUSD]);

  // Unique stores list for filtering
  const lojasUnicas = useMemo(() => {
    const s = new Set();
    produtos.forEach(p => p.local && s.add(p.local));
    itensLegais.forEach(i => i.local && s.add(i.local));
    gastos.forEach(g => g.local && s.add(g.local));
    return Array.from(s);
  }, [produtos, itensLegais, gastos]);

  // --- Auth View Fallback ---
  if (authLoading) {
    return (
      <div style={{ display: "flex", height: "100vh", alignItems: "center", justifyContent: "center", background: C.bg, fontFamily: "'Inter', sans-serif" }}>
        <div style={{ textAlign: "center" }}>
          <div style={{ width: 40, height: 40, border: `4px solid ${C.border}`, borderTopColor: C.primary, borderRadius: "50%", animation: "spin 1s linear infinite", margin: "0 auto 15px" }} />
          <p style={{ color: C.textMid, fontSize: 14, fontWeight: 500 }}>Carregando TravelShop...</p>
          <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
        </div>
      </div>
    );
  }

  if (!user) {
    return (
      <div style={{ minHeight: "100vh", background: C.bg, display: "flex", alignItems: "center", justifyContent: "center", padding: 20, fontFamily: "'Inter', sans-serif" }}>
        <div style={{ background: C.bgCard, borderRadius: 24, padding: 32, width: "100%", maxWidth: 420, boxShadow: "0 10px 25px -5px rgba(0,0,0,0.05), 0 8px 10px -6px rgba(0,0,0,0.05)" }}>
          <div style={{ textAlign: "center", marginBottom: 32 }}>
            <span style={{ fontSize: 48 }}>✈️</span>
            <h1 style={{ fontSize: 24, fontWeight: 800, color: C.text, marginTop: 12 }}>TravelShop</h1>
            <p style={{ color: C.textLight, fontSize: 14, marginTop: 4 }}>Gerencie as compras e custos da sua viagem</p>
          </div>

          {authError && (
            <div style={{ background: C.dangerLight, color: C.danger, padding: "12px 16px", borderRadius: 12, fontSize: 13, fontWeight: 600, marginBottom: 20 }}>
              {authError}
            </div>
          )}

          <form onSubmit={handleAuthSubmit}>
            <div style={{ marginBottom: 16 }}>
              <label style={{ display: "block", fontSize: 12, fontWeight: 600, color: C.textMid, marginBottom: 6 }}>E-mail</label>
              <input type="email" value={authEmail} onChange={e => setAuthEmail(e.target.value)} placeholder="seu@email.com" style={{ width: "100%", padding: "12px 16px", borderRadius: 12, border: `1px solid ${C.border}`, fontSize: 14, outline: "none", boxSizing: "border-box" }} />
            </div>
            <div style={{ marginBottom: 24 }}>
              <label style={{ display: "block", fontSize: 12, fontWeight: 600, color: C.textMid, marginBottom: 6 }}>Senha</label>
              <input type="password" value={authPassword} onChange={e => setAuthPassword(e.target.value)} placeholder="••••••••" style={{ width: "100%", padding: "12px 16px", borderRadius: 12, border: `1px solid ${C.border}`, fontSize: 14, outline: "none", boxSizing: "border-box" }} />
            </div>

            <button type="submit" style={{ width: "100%", padding: "14px", borderRadius: 12, background: C.primary, color: "#FFF", border: "none", fontSize: 14, fontWeight: 700, cursor: "pointer", boxShadow: "0 4px 12px rgba(37,99,235,0.2)" }}>
              {authIsSignUp ? "Criar Minha Conta" : "Entrar no Aplicativo"}
            </button>
          </form>

          <div style={{ textAlign: "center", marginTop: 24 }}>
            <button onClick={() => { setAuthIsSignUp(!authIsSignUp); setAuthError(""); }} style={{ background: "none", border: "none", color: C.primary, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
              {authIsSignUp ? "Já tem uma conta? Faça login" : "Não tem conta? Cadastre-se grátis"}
            </button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={{ background: C.bg, minHeight: "100vh", paddingBottom: 90, fontFamily: "'Inter', sans-serif", color: C.text }}>
      <style>{CSS}</style>
      
      {notification && (
        <div className="notif" style={{
          background: notification.type === "success" ? C.success : notification.type === "danger" ? C.danger : C.warning,
          color: "#FFF"
        }}>
          {notification.text}
        </div>
      )}

      {/* --- HEADER CONTROL PANEL ─── */}
      <header style={{ background: C.bgCard, borderBottom: `1px solid ${C.border}`, padding: "16px 20px", sticky: "top", zIndex: 100 }}>
        <div style={{ maxWidth: 1200, margin: "0 auto", display: "flex", flexDirection: "column", gap: 16 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <div>
              <h1 style={{ fontSize: 20, fontWeight: 800, letterSpacing: "-0.5px", display: "flex", alignItems: "center", gap: 8 }}>
                <span>🇺🇸</span> TravelShop <span style={{ fontSize: 12, fontWeight: 500, color: C.textLight, background: C.borderLight, padding: "2px 8px", borderRadius: 6 }}>v1.2</span>
              </h1>
              <p style={{ fontSize: 11, color: C.textLight, marginTop: 2 }}>Usuário: {user.email}</p>
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <button onClick={handleLogout} style={{ padding: "6px 12px", borderRadius: 8, background: "none", border: `1px solid ${C.border}`, fontSize: 12, fontWeight: 600, color: C.danger, cursor: "pointer" }}>Sair</button>
              <button onClick={() => { setModalType(activeTab === 1 ? "itemLegal" : activeTab === 2 ? "gasto" : activeTab === 4 ? "parcela" : "produto"); setEditingItem(null); setIsModalOpen(true); }} style={{ background: `linear-gradient(135deg, ${C.gradientA}, ${C.gradientB})`, color: "#FFF", border: "none", padding: "10px 16px", borderRadius: 10, fontSize: 13, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", gap: 6, boxShadow: "0 4px 10px rgba(37,99,235,0.15)" }}>
                <span>＋</span> {activeTab === 1 ? "Item Legal" : activeTab === 2 ? "Gasto" : activeTab === 4 ? "Parcela" : "Produto"}
              </button>
            </div>
          </div>

          {/* Quick Settings Grid */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: 10 }}>
            <div style={{ background: C.bg, padding: "10px 12px", borderRadius: 12, border: `1px solid ${C.borderLight}` }}>
              <span style={{ fontSize: 10, color: C.textLight, fontWeight: 700, display: \"block\", transform: \"translateY(-2px)\" }}>TAXA FLÓRIDA</span>
              <div style={{ display: \"flex\", alignItems: \"center\" }}>
                <input type=\"number\" step=\"0.01\" value={settings.taxRate} onChange={e => updateSettings(\"taxRate\", e.target.value)} style={{ background: \"none\", border: \"none\", fontSize: 15, fontWeight: 800, color: C.text, width: \"100%\", outline: \"none\" }} />
                <span style={{ fontSize: 13, fontWeight: 700, color: C.textMid }}>%</span>
              </div>
            </div>
            <div style={{ background: C.bg, padding: \"10px 12px\", borderRadius: 12, border: `1px solid ${C.borderLight}` }}>
              <span style={{ fontSize: 10, color: C.textLight, fontWeight: 700, display: \"block\", transform: \"translateY(-2px)\" }}>CÂMBIO COM IOF</span>
              <div style={{ display: \"flex\", alignItems: \"center\" }}>
                <span style={{ fontSize: 13, fontWeight: 700, color: C.textMid, marginRight: 2 }}>R$</span>
                <input type=\"number\" step=\"0.01\" value={settings.exchangeRate} onChange={e => updateSettings(\"exchangeRate\", e.target.value)} style={{ background: \"none\", border: \"none\", fontSize: 15, fontWeight: 800, color: C.text, width: \"100%\", outline: \"none\" }} />
              </div>
            </div>
          </div>
        </div>
      </header>

      {/* --- FINANCIAL HERO CARD ─── */}
      <div style={{ padding: \"16px 20px 0\", maxWidth: 1200, margin: \"0 auto\" }}>
        <div style={{ background: `linear-gradient(135deg, ${C.gradientA}, ${C.gradientB})`, borderRadius: 20, padding: 20, color: \"#FFF\", boxShadow: \"0 10px 20px -5px rgba(37,99,235,0.2)\" }}>
          <div style={{ display: \"flex\", justifyContent: \"space-between\", opacity: 0.85, fontSize: 12, fontWeight: 600 }}>
            <span>TOTAL GERAL INVESTIDO</span>
            <span>CÂMBIO: R$ {settings.exchangeRate.toFixed(2)}</span>
          </div>
          <div style={{ fontSize: 32, fontWeight: 800, margin: \"4px 0 16px\", letterSpacing: \"-1px\" }}>
            USD {totals.totalUSD.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
            <span style={{ fontSize: 16, fontWeight: 500, marginLeft: 10, opacity: 0.9 }}>
              (R$ {totals.totalBRL.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })})
            </span>
          </div>
          <div style={{ display: \"grid\", gridTemplateColumns: \"repeat(3, 1fr)\", gap: 8, background: \"rgba(255,255,255,0.1)\", padding: 12, borderRadius: 14 }}>
            <div>
              <span style={{ display: \"block\", fontSize: 9, opacity: 0.75, fontWeight: 700 }}>COMPRAS</span>
              <span style={{ fontSize: 13, fontWeight: 700 }}>${totals.prodUSD.toFixed(1)}</span>
            </div>
            <div>
              <span style={{ display: \"block\", fontSize: 9, opacity: 0.75, fontWeight: 700 }}>LEGAIS</span>
              <span style={{ fontSize: 13, fontWeight: 700 }}>${totals.legalUSD.toFixed(1)}</span>
            </div>
            <div>
              <span style={{ display: \"block\", fontSize: 9, opacity: 0.75, fontWeight: 700 }}>GASTOS</span>
              <span style={{ fontSize: 13, fontWeight: 700 }}>${totals.gastosUSD.toFixed(1)}</span>
            </div>
          </div>
        </div>
      </div>

      {/* --- FILTERS DISPLAY (Tabs 0,1,2 Only) --- */}
      {[0, 1, 2].includes(activeTab) && (
        <div style={{ padding: \"12px 20px 0\", maxWidth: 1200, margin: \"0 auto\", display: \"flex\", flexDirection: \"column\", gap: 8 }}>
          <input type=\"text\" placeholder=\"🔍 Buscar item ou local...\" value={search} onChange={e => setSearch(e.target.value)} style={{ width: \"100%\", padding: \"11px 14px\", borderRadius: 12, border: `1px solid ${C.border}`, background: C.bgCard, fontSize: 13, outline: \"none\", boxSizing: \"border-box\" }} />
          <div style={{ display: \"grid\", gridTemplateColumns: \"1fr 1fr\", gap: 8 }}>
            <select value={catFilter} onChange={e => setCatFilter(e.target.value)} style={{ width: \"100%\", padding: 10, borderRadius: 12, border: `1px solid ${C.border}`, background: C.bgCard, fontSize: 12, fontWeight: 600, color: C.textMid, outline: \"none\" }}>
              <option value=\"all\">📂 Todas Categorias</option>
              {CATEGORIES.map(c => <option key={c.id} value={c.id}>{c.icon} {c.label}</option>)}
            </select>
            <select value={lojaFilter} onChange={e => setLojaFilter(e.target.value)} style={{ width: \"100%\", padding: 10, borderRadius: 12, border: `1px solid ${C.border}`, background: C.bgCard, fontSize: 12, fontWeight: 600, color: C.textMid, outline: \"none\" }}>
              <option value=\"all\">🏪 Todos Locais</option>
              {lojasUnicas.map(l => <option key={l} value={l}>{l}</option>)}
            </select>
          </div>
        </div>
      )}

      {/* --- MAIN TABS ROUTER --- */}
      <main style={{ padding: \"16px 20px\", maxWidth: 1200, margin: \"0 auto\" }}>
        {activeTab === 0 && <ProdutosTab produtos={produtos} search={search} catFilter={catFilter} lojaFilter={lojaFilter} taxMultiplier={taxMultiplier} onEdit={p => { setModalType(\"produto\"); setEditingItem(p); setIsModalOpen(true); }} onDelete={id => handleDelete(id, \"produto\")} />}
        {activeTab === 1 && <ItensLegaisTab itensLegais={itensLegais} search={search} catFilter={catFilter} lojaFilter={lojaFilter} taxMultiplier={taxMultiplier} onEdit={i => { setModalType(\"itemLegal\"); setEditingItem(i); setIsModalOpen(true); }} onDelete={id => handleDelete(id, \"itemLegal\")} />}
        {activeTab === 2 && <GastosTab gastos={gastos} search={search} catFilter={catFilter} lojaFilter={lojaFilter} onEdit={g => { setModalType(\"gasto\"); setEditingItem(g); setIsModalOpen(true); }} onDelete={id => handleDelete(id, \"gasto\")} />}
        {activeTab === 3 && <ResumoTab produtos={produtos} itensLegais={itensLegais} gastos={gastos} settings={settings} taxMultiplier={taxMultiplier} />}
        {activeTab === 4 && <ParcelasTab parcelas={parcelas} loading={parcelasLoading} onSave={saveParcelaToCloud} onDelete={id => handleDelete(id, \"parcela\")} />}
        {activeTab === 5 && <DashboardTab produtos={produtos} itensLegais={itensLegais} gastos={gastos} taxMultiplier={taxMultiplier} />}
        {activeTab === 6 && <CalculadoraTab settings={settings} updateSettings={updateSettings} />}
      </main>

      {/* --- NAV BAR ─── */}
      <nav style={{ position: \"fixed\", bottom: 0, left: 0, right: 0, background: \"rgba(255,255,255,0.92)\", backdropFilter: \"blur(12px)\", borderTop: `1px solid ${C.border}`, display: \"flex\", justifyContent: \"space-around\", padding: \"8px 4px 22px\", zIndex: 100 }}>
        <NavButton active={activeTab === 0} label=\"Produtos\" icon=\"🛍️\" onClick={() => setActiveTab(0)} />
        <NavButton active={activeTab === 1} label=\"Legais\" icon=\"✨\" onClick={() => setActiveTab(1)} />
        <NavButton active={activeTab === 2} label=\"Gastos\" icon=\"💸\" onClick={() => setActiveTab(2)} />
        <NavButton active={activeTab === 3} label=\"Resumo\" icon=\"📊\" onClick={() => setActiveTab(3)} />
        <NavButton active={activeTab === 4} label=\"Parcelas\" icon=\"💳\" onClick={() => setActiveTab(4)} />
        <NavButton active={activeTab === 5} label=\"Dash\" icon=\"📈\" onClick={() => setActiveTab(5)} />
        <NavButton active={activeTab === 6} label=\"Calc\" icon=\"🧮\" onClick={() => setActiveTab(6)} />
      </nav>

      {/* --- FORM MODAL LAYER ─── */}
      {isModalOpen && (
        <FormModal type={modalType} item={editingItem} onSave={handleSave} onClose={() => { setIsModalOpen(false); setEditingItem(null); }} />
      )}
    </div>
  );
}

// ─── NAV BUTTON COMPONENT ────────────────────────────────────────────────────
function NavButton({ active, label, icon, onClick }) {
  return (
    <button onClick={onClick} style={{ background: \"none\", border: \"none\", display: \"flex\", flexDirection: \"column\", alignItems: \"center\", gap: 4, cursor: \"pointer\", padding: \"6px 10px\", borderRadius: 10, minWidth: 52 }}>
      <span style={{ fontSize: 20, filter: active ? \"none\" : \"grayscale(40%)\", transform: active ? \"scale(1.1)\" : \"none\", transition: \"transform 0.2s\" }}>{icon}</span>
      <span style={{ fontSize: 9, fontWeight: active ? 800 : 500, color: active ? C.primary : C.textLight }}>{label}</span>
    </button>
  );
}

// ─── TAB 0: PRODUTOS ─────────────────────────────────────────────────────────
function ProdutosTab({ produtos, search, catFilter, lojaFilter, taxMultiplier, onEdit, onDelete }) {
  const filtered = useMemo(() => {
    return produtos.filter(p => {
      const matchS = p.nome.toLowerCase().includes(search.toLowerCase()) || (p.local && p.local.toLowerCase().includes(search.toLowerCase()));
      const matchC = catFilter === \"all\" || p.categoria === catFilter;
      const matchL = lojaFilter === \"all\" || p.local === lojaFilter;
      return matchS && matchC && matchL;
    });
  }, [produtos, search, catFilter, lojaFilter]);

  if (!filtered.length) return <EmptyState text=\"Nenhum produto encontrado.\" />;

  return (
    <div style={{ display: \"flex\", flexDirection: \"column\", gap: 12 }}>
      {filtered.map(p => {
        const costUSD = p.categoria === \"compras\" ? p.valorUSD * taxMultiplier : p.valorUSD;
        const totalUSD = costUSD * p.qtd;
        const cat = CATEGORIES.find(c => c.id === p.categoria) || CATEGORIES[5];

        return (
          <ItemCard key={p.id} icon={cat.icon} color={cat.color} title={p.nome} subtitle={`${p.local || 'Local não definido'} · Qtd: ${p.qtd}`} valA={`$ ${p.valorUSD.toFixed(2)}`} valB={`Total: $ ${totalUSD.toFixed(2)}`} link={p.link} nota={p.nota} pwaImageKey={p.nome} onEdit={() => onEdit(p)} onDelete={() => onDelete(p.id)} />
        );
      })}
    </div>
  );
}

// ─── TAB 1: ITENS LEGAIS ─────────────────────────────────────────────────────
function ItensLegaisTab({ itensLegais, search, catFilter, lojaFilter, taxMultiplier, onEdit, onDelete }) {
  const filtered = useMemo(() => {
    return itensLegais.filter(i => {
      const matchS = i.nome.toLowerCase().includes(search.toLowerCase()) || (i.local && i.local.toLowerCase().includes(search.toLowerCase()));
      const matchC = catFilter === \"all\" || i.categoria === catFilter;
      const matchL = lojaFilter === \"all\" || i.local === lojaFilter;
      return matchS && matchC && matchL;
    });
  }, [itensLegais, search, catFilter, lojaFilter]);

  if (!filtered.length) return <EmptyState text=\"Nenhum item legal cadastrado.\" />;

  return (
    <div style={{ display: \"flex\", flexDirection: \"column\", gap: 12 }}>
      {filtered.map(i => {
        const costUSD = i.categoria === \"compras\" ? i.valorUSD * taxMultiplier : i.valorUSD;
        const totalUSD = costUSD * i.qtd;
        const cat = CATEGORIES.find(c => c.id === i.categoria) || CATEGORIES[5];

        return (
          <ItemCard key={i.id} icon={cat.icon} color={cat.color} title={i.nome} subtitle={`${i.local || 'Local'} · Qtd: ${i.qtd}`} valA={`$ ${i.valorUSD.toFixed(2)}`} valB={`Total: $ ${totalUSD.toFixed(2)}`} link={i.link} nota={i.nota} onEdit={() => onEdit(i)} onDelete={() => onDelete(i.id)} />
        );
      })}
    </div>
  );
}

// ─── TAB 2: GASTOS DIÁRIOS ────────────────────────────────────────────────────
function GastosTab({ gastos, search, catFilter, lojaFilter, onEdit, onDelete }) {
  const filtered = useMemo(() => {
    return gastos.filter(g => {
      const matchS = g.nome.toLowerCase().includes(search.toLowerCase()) || (g.local && g.local.toLowerCase().includes(search.toLowerCase()));
      const matchC = catFilter === \"all\" || g.categoria === catFilter;
      const matchL = lojaFilter === \"all\" || g.local === lojaFilter;
      return matchS && matchC && matchL;
    });
  }, [gastos, search, catFilter, lojaFilter]);

  if (!filtered.length) return <EmptyState text=\"Nenhum gasto diário inserido.\" />;

  return (
    <div style={{ display: \"flex\", flexDirection: \"column\", gap: 12 }}>
      {filtered.map(g => {
        const cat = CATEGORIES.find(c => c.id === g.categoria) || CATEGORIES[5];
        const partsText = g.pessoas && g.pessoas.length > 0 
          ? `Dividido: ${g.pessoas.map(p => p.valorCustom ? `${p.nome} ($${p.valorCustom})` : p.nome).join(', ')}`
          : "Individual (Você)";

        return (
          <ItemCard key={g.id} icon={cat.icon} color={cat.color} title={g.nome} subtitle={`${g.local || 'Gasto'} · ${partsText}`} valA={`$ ${g.valorUSD.toFixed(2)}`} valB=\"\" nota={g.nota} onEdit={() => onEdit(g)} onDelete={() => onDelete(g.id)} />
        );
      })}
    </div>
  );
}

// ─── TAB 4: PARCELAS (NOVA ABA PLANILHA EDITÁVEL) ───────────────────────────
function ParcelasTab({ parcelas, loading, onSave, onDelete }) {
  const [isFormOpen, setIsFormOpen] = useState(false);
  const [editingParcela, setEditingParcela] = useState(null);

  const monthlyTotal = useMemo(() => {
    return parcelas.reduce((acc, curr) => acc + (parseFloat(curr.valorParcela) || 0), 0);
  }, [parcelas]);

  const grandRemainingTotal = useMemo(() => {
    return parcelas.reduce((acc, curr) => {
      const total = parseFloat(curr.valorTotal) || 0;
      const qtd = parseInt(curr.qtdParcelas) || 1;
      const vParc = parseFloat(curr.valorParcela) || 0;
      const paidCount = Array.isArray(curr.statusMeses) ? curr.statusMeses.filter(Boolean).length : 0;
      const remaining = Math.max(0, total - (paidCount * vParc));
      return acc + remaining;
    }, 0);
  }, [parcelas]);

  if (loading) {
    return (
      <div style={{ textAlign: "center", padding: 40 }}>
        <div style={{ width: 30, height: 30, border: `3px solid ${C.border}`, borderTopColor: C.primary, borderRadius: "50%", animation: "spin 1s linear infinite", margin: "0 auto 10px" }} />
        <p style={{ color: C.textMid, fontSize: 13 }}>Carregando parcelas em tempo real...</p>
      </div>
    );
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, background: C.bgCard, padding: 16, borderRadius: 16, border: `1px solid ${C.border}` }}>
        <div>
          <span style={{ fontSize: 11, fontWeight: 600, color: C.textLight }}>COMPROMISSO MENSAL</span>
          <div style={{ fontSize: 18, fontWeight: 800, color: C.text, marginTop: 2 }}>R$ {monthlyTotal.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</div>
        </div>
        <div>
          <span style={{ fontSize: 11, fontWeight: 600, color: C.textLight }}>RESTANTE TOTAL</span>
          <div style={{ fontSize: 18, fontWeight: 800, color: C.purple, marginTop: 2 }}>R$ {grandRemainingTotal.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</div>
        </div>
      </div>

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <h3 style={{ fontSize: 14, fontWeight: 700, color: C.textMid }}>Tabela de Parcelas</h3>
        <button onClick={() => { setEditingParcela(null); setIsFormOpen(true); }} style={{ background: C.primaryLight, color: C.primary, border: "none", padding: "6px 12px", borderRadius: 8, fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
          ＋ Nova Parcela
        </button>
      </div>

      {!parcelas.length ? (
        <EmptyState text="Nenhuma parcela cadastrada na nuvem." />
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          {parcelas.map(p => (
            <ParcelaCard key={p.id} parcela={p} onEdit={(item) => { setEditingParcela(item); setIsFormOpen(true); }} onDelete={onDelete} onToggleMonth={(item) => onSave(item)} />
          ))}
        </div>
      )}

      {isFormOpen && (
        <FormModal type="parcela" item={editingParcela} onSave={(item) => { onSave(item); setIsFormOpen(false); }} onClose={() => setIsFormOpen(false)} />
      )}
    </div>
  );
}

function ParcelaCard({ parcela, onEdit, onDelete, onToggleMonth }) {
  const [expanded, setExpanded] = useState(false);
  const MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
  
  const statusMeses = Array.isArray(parcela.statusMeses) ? parcela.statusMeses : Array(12).fill(false);
  const paidCount = statusMeses.filter(Boolean).length;
  const totalQty = parseInt(parcela.qtdParcelas) || 1;
  const isFullyPaid = paidCount >= totalQty;

  const handleMonthClick = (idx) => {
    const nextStatus = [...statusMeses];
    nextStatus[idx] = !nextStatus[idx];
    onToggleMonth({ ...parcela, statusMeses: nextStatus });
  };

  return (
    <div style={{ background: C.bgCard, borderRadius: 16, border: `1px solid ${C.border}`, padding: 14, display: "flex", flexDirection: "column", gap: 10 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }} onClick={() => setExpanded(!expanded)}>
        <div style={{ flex: 1, cursor: "pointer" }}>
          <h4 style={{ fontSize: 14, fontWeight: 700, color: C.text }}>{parcela.descricao}</h4>
          <span style={{ fontSize: 11, color: C.textLight, display: "block", marginTop: 2 }}>
            💳 {parcela.cartao || "Não inf."} · {parcela.qtdParcelas}x de R$ {parseFloat(parcela.valorParcela || 0).toFixed(2)}
          </span>
        </div>
        <div style={{ textAlign: "right" }}>
          <span style={{ fontSize: 14, fontWeight: 800, color: C.text }}>R$ {parseFloat(parcela.valorTotal || 0).toFixed(2)}</span>
          <span style={{ display: "block", fontSize: 10, fontWeight: 700, color: isFullyPaid ? C.success : C.warning, marginTop: 4, background: isFullyPaid ? C.successLight : C.warningLight, padding: "2px 6px", borderRadius: 6 }}>
            {isFullyPaid ? "Pago" : `${paidCount}/${totalQty} Meses`}
          </span>
        </div>
      </div>

      <div style={{ width: "100%", background: C.borderLight, height: 6, borderRadius: 3, overflow: "hidden" }}>
        <div style={{ width: `${Math.min(100, (paidCount / totalQty) * 100)}%`, background: isFullyPaid ? C.success : C.primary, height: "100%", borderRadius: 3, transition: "width 0.3s" }} />
      </div>

      {expanded && (
        <div style={{ borderTop: `1px solid ${C.borderLight}`, paddingTop: 10, marginTop: 4, display: "flex", flexDirection: "column", gap: 12 }}>
          <div>
            <span style={{ fontSize: 11, fontWeight: 600, color: C.textLight, display: \"block\", marginBottom: 6 }}>CONTROLE MENSAL DA PLANILHA (PAGO/PENDENTE):</span>
            <div style={{ display: \"grid\", gridTemplateColumns: \"repeat(4, 1fr)\", gap: 6 }}>
              {MESES.map((mes, i) => (
                <button key={mes} onClick={() => handleMonthClick(i)} style={{ padding: \"6px 2px\", borderRadius: 8, border: `1px solid ${statusMeses[i] ? C.success : C.border}`, background: statusMeses[i] ? C.successLight : \"none\", color: statusMeses[i] ? C.success : C.textMid, fontSize: 11, fontWeight: 600, cursor: \"pointer\" }}>
                  {mes} {statusMeses[i] ? \"✅\" : \"\"}
                </button>
              ))}
            </div>
          </div>
          <div style={{ display: \"flex\", justifyContent: \"flex-end\", gap: 8 }}>
            <button onClick={() => onEdit(parcela)} style={{ background: \"none\", border: `1px solid ${C.border}`, padding: \"6px 12px\", borderRadius: 8, fontSize: 11, fontWeight: 600, color: C.textMid, cursor: \"pointer\" }}>Editar</button>
            <button onClick={() => onDelete(parcela.id, \"parcela\")} style={{ background: \"none\", border: `1px solid ${C.danger}`, padding: \"6px 12px\", borderRadius: 8, fontSize: 11, fontWeight: 600, color: C.danger, cursor: \"pointer\" }}>Excluir</button>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── TAB 3: RESUMO DE QUEM DEVE QUEM ─────────────────────────────────────────
function ResumoTab({ produtos, itensLegais, gastos, settings, taxMultiplier }) {
  const debits = useMemo(() => {
    let enzoOwesPai = 0;
    let paiOwesEnzo = 0;

    // Calcular parcelas/compras se houvesse tags de divisão nos produtos, mas o foco solicitado foi na divisão unequal de Gastos Diários.
    // Analisar coleção de gastos diários
    gastos.forEach(g => {
      const amt = g.valorUSD || 0;
      if (!g.pessoas || g.pessoas.length === 0) return; // individual do enzo

      // Descobrir se há valor customizado preenchido
      const hasCustom = g.pessoas.some(p => p.valorCustom !== undefined && p.valorCustom !== \"\");

      if (hasCustom) {
        g.pessoas.forEach(p => {
          const valCustom = parseFloat(p.valorCustom) || 0;
          if (p.nome === \"Pai\") enzoOwesPai += valCustom;
          if (p.nome === \"Mãe\") enzoOwesPai += valCustom; // se houver outra pessoa
        });
      } else {
        // Divisão padrão igual
        const share = amt / (g.pessoas.length + 1);
        g.pessoas.forEach(p => {
          if (p.nome === \"Pai\") enzoOwesPai += share;
        });
      }
    });

    const netUSD = enzoOwesPai - paiOwesEnzo;
    const netBRL = netUSD * settings.exchangeRate;

    return { enzoOwesPai, paiOwesEnzo, netUSD, netBRL };
  }, [gastos, settings.exchangeRate]);

  return (
    <div style={{ display: \"flex\", flexDirection: \"column\", gap: 16 }}>
      <div style={{ background: C.bgCard, borderRadius: 16, border: `1px solid ${C.border}`, padding: 16, textAlign: \"center\" }}>
        <span style={{ fontSize: 12, fontWeight: 600, color: C.textLight }}>SALDO DA DIVISÃO DE GASTOS</span>
        <h2 style={{ fontSize: 24, fontWeight: 800, color: debits.netUSD >= 0 ? C.warning : C.success, margin: \"6px 0 2px\" }}>
          USD {Math.abs(debits.netUSD).toFixed(2)}
        </h2>
        <span style={{ fontSize: 13, fontWeight: 600, color: C.textMid, display: \"block\" }}>
          (R$ {Math.abs(debits.netBRL).toFixed(2)})
        </span>
        <p style={{ fontSize: 12, color: C.textLight, marginTop: 10, fontWeight: 500 }}>
          {debits.netUSD > 0 ? \"Enzo deve repassar esse valor para o Pai.\" : debits.netUSD < 0 ? \"O Pai deve repassar esse valor para o Enzo.\" : \"Todos os gastos divididos estão quitados!\"}
        </p>
      </div>

      <div style={{ background: C.bgCard, borderRadius: 16, border: `1px solid ${C.border}`, padding: 14 }}>
        <h3 style={{ fontSize: 13, fontWeight: 700, color: C.textMid, marginBottom: 12 }}>Detalhamento por Integrante</h3>
        <div style={{ display: \"flex\", flexDirection: \"column\", gap: 10 }}>
          <div style={{ display: \"flex\", justifyContent: \"space-between\", paddingBottom: 8, borderBottom: `1px solid ${C.borderLight}` }}>
            <span style={{ fontSize: 13, fontWeight: 600, color: C.text }}>Gastos Cobertos pelo Pai</span>
            <span style={{ fontSize: 13, fontWeight: 700, color: C.text }}>$ {debits.enzoOwesPai.toFixed(2)}</span>
          </div>
          <div style={{ display: \"flex\", justifyContent: \"space-between\", paddingBottom: 4 }}>
            <span style={{ fontSize: 13, fontWeight: 600, color: C.text }}>Gastos Cobertos pelo Enzo</span>
            <span style={{ fontSize: 13, fontWeight: 700, color: C.text }}>$ {debits.paiOwesEnzo.toFixed(2)}</span>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── TAB 5: DASHBOARD CONTROLE ───────────────────────────────────────────────
function DashboardTab({ produtos, itensLegais, gastos, taxMultiplier }) {
  const chartData = useMemo(() => {
    const sums = { compras: 0, alimentacao: 0, hospedagem: 0, transporte: 0, lazer: 0, outros: 0 };
    
    produtos.forEach(p => {
      const v = (p.valorUSD || 0) * (p.qtd || 1);
      sums[p.categoria || \"outros\"] += p.categoria === \"compras\" ? v * taxMultiplier : v;
    });
    itensLegais.forEach(i => {
      const v = (i.valorUSD || 0) * (i.qtd || 1);
      sums[i.categoria || \"outros\"] += i.categoria === \"compras\" ? v * taxMultiplier : v;
    });
    gastos.forEach(g => {
      sums[g.categoria || \"outros\"] += (g.valorUSD || 0);
    });

    return Object.keys(sums).map(k => ({
      id: k,
      label: CAT_TRANSLATIONS[k] || k,
      val: sums[k],
      cat: CATEGORIES.find(c => c.id === k) || CATEGORIES[5]
    })).filter(item => item.val > 0).sort((a,b) => b.val - a.val);
  }, [produtos, itensLegais, gastos, taxMultiplier]);

  const maxVal = chartData[0]?.val || 1;

  if (!chartData.length) return <EmptyState text=\"Nenhum dado financeiro para gerar gráficos.\" />;

  return (
    <div style={{ background: C.bgCard, borderRadius: 16, border: `1px solid ${C.border}`, padding: 16 }}>
      <h3 style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 16 }}>Custos por Categoria (USD)</h3>
      <div style={{ display: \"flex\", flexDirection: \"column\", gap: 14 }}>
        {chartData.map(c => {
          const pct = (c.val / maxVal) * 100;
          return (
            <div key={c.id}>
              <div style={{ display: \"flex\", justifyContent: \"space-between\", fontSize: 12, fontWeight: 600, marginBottom: 4 }}>
                <span style={{ display: \"flex\", alignItems: \"center\", gap: 4 }}><span>{c.cat.icon}</span> {c.label}</span>
                <span style={{ color: C.text }}>$ {c.val.toFixed(2)}</span>
              </div>
              <div style={{ width: \"100%\", background: C.borderLight, height: 8, borderRadius: 4, overflow: \"hidden\" }}>
                <div style={{ width: `${pct}%`, background: c.cat.color, height: \"100%\", borderRadius: 4 }} />
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ─── TAB 6: CALCULADORA COMPRAS ──────────────────────────────────────────────
function CalculadoraTab({ settings, updateSettings }) {
  const [calcVal, setCalcVal] = useState(\"\");
  const taxMult = 1 + (settings.taxRate / 100);

  const parsed = parseFloat(calcVal) || 0;
  const resUSD = parsed * taxMult;
  const resBRL = resUSD * settings.exchangeRate;

  return (
    <div style={{ background: C.bgCard, borderRadius: 16, border: `1px solid ${C.border}`, padding: 18, display: \"flex\", flexDirection: \"column\", gap: 14 }}>
      <h3 style={{ fontSize: 14, fontWeight: 700, color: C.text }}>Simulador de Imposto de Etiqueta</h3>
      <div>
        <label style={{ display: \"block\", fontSize: 11, fontWeight: 600, color: C.textMid, marginBottom: 5 }}>VALOR DA ETIQUETA (USD)</label>
        <input type=\"number\" placeholder=\"$ 0.00\" value={calcVal} onChange={e => setCalcVal(e.target.value)} style={{ width: \"100%\", padding: 12, borderRadius: 10, border: `1px solid ${C.border}`, fontSize: 15, fontWeight: 700, outline: \"none\", boxSizing: \"border-box\" }} />
      </div>

      <div style={{ background: C.bg, padding: 14, borderRadius: 12, display: \"flex\", flexDirection: \"column\", gap: 8, border: `1px solid ${C.borderLight}` }}>
        <div style={{ display: \"flex\", justifyContent: \"space-between\", fontSize: 13 }}>
          <span style={{ color: C.textMid, fontWeight: 500 }}>Com Taxa Flórida (${settings.taxRate}%):</span>
          <span style={{ fontWeight: 700, color: C.text }}>USD {resUSD.toFixed(2)}</span>
        </div>
        <div style={{ display: \"flex\", justifyContent: \"space-between\", fontSize: 14, borderTop: `1px solid ${C.borderLight}`, paddingTop: 8, marginTop: 2 }}>
          <span style={{ color: C.primary, fontWeight: 700 }}>Total Convertido em Reais:</span>
          <span style={{ fontWeight: 800, color: C.primary }}>R$ {resBRL.toFixed(2)}</span>
        </div>
      </div>
    </div>
  );
}

// ─── GENERIC UI CARD COMPONENT ────────────────────────────────────────────────
function ItemCard({ icon, color, title, subtitle, valA, valB, link, nota, pwaImageKey, onEdit, onDelete }) {
  const [expanded, setExpanded] = useState(false);
  const [imgUrl, setImgUrl] = useState(null);

  useEffect(() => {
    if (pwaImageKey) {
      if (imageCache[pwaImageKey]) {
        setImgUrl(imageCache[pwaImageKey]);
      } else {
        getProductImage(pwaImageKey).then(url => {
          if (url) setImgUrl(url);
        });
      }
    }
  }, [pwaImageKey]);

  return (
    <div style={{ background: C.bgCard, borderRadius: 16, border: `1px solid ${C.border}`, padding: 12, display: \"flex\", flexDirection: \"column\", gap: 8 }}>
      <div style={{ display: \"flex\", alignItems: \"center\", gap: 10 }}>
        {imgUrl ? (
          <img src={imgUrl} alt=\"\" style={{ width: 42, height: 42, borderRadius: 10, objectFit: \"cover\", border: `1px solid ${C.borderLight}` }} onClick={() => setExpanded(!expanded)} />
        ) : (
          <div style={{ width: 42, height: 42, borderRadius: 10, background: color + \"15\", display: \"flex\", alignItems: \"center\", justifyContent: \"center\", fontSize: 20, color }} onClick={() => setExpanded(!expanded)}>
            {icon}
          </div>
        )}

        <div style={{ flex: 1, minWidth: 0, cursor: \"pointer\" }} onClick={() => setExpanded(!expanded)}>
          <h4 style={{ fontSize: 13, fontWeight: 700, color: C.text, overflow: \"hidden\", textOverflow: \"ellipsis\", whiteSpace: \"nowrap\" }}>{title}</h4>
          <span style={{ fontSize: 11, color: C.textLight, overflow: \"hidden\", textOverflow: \"ellipsis\", whiteSpace: \"nowrap\", display: \"block\" }}>{subtitle}</span>
        </div>

        <div style={{ textAlign: \"right\", minWidth: 70, cursor: \"pointer\" }} onClick={() => setExpanded(!expanded)}>
          <span style={{ fontSize: 13, fontWeight: 800, color: C.text, display: \"block\" }}>{valA}</span>
          {valB && <span style={{ fontSize: 10, fontWeight: 600, color: C.textLight, display: \"block\", marginTop: 2 }}>{valB}</span>}
        </div>
      </div>

      {expanded && (
        <div style={{ borderTop: `1px solid ${C.borderLight}`, paddingTop: 10, display: \"flex\", flexDirection: \"column\", gap: 8, marginTop: 2 }}>
          {nota && <p style={{ fontSize: 11, color: C.textMid, background: C.bg, padding: 8, borderRadius: 8, borderLeft: `3px solid ${color}` }}>{nota}</p>}
          <div style={{ display: \"flex\", justifyContent: \"space-between\", alignItems: \"center\" }}>
            {link ? (
              <a href={link} target=\"_blank\" rel=\"noreferrer\" style={{ fontSize: 11, color: C.primary, fontWeight: 600, textDecoration: \"none\" }}>🌐 Ver Link Oficial</a>
            ) : <span />}
            <div style={{ display: \"flex\", gap: 6 }}>
              <button onClick={onEdit} style={{ background: \"none\", border: `1px solid ${C.border}`, padding: \"5px 10px\", borderRadius: 8, fontSize: 11, fontWeight: 600, color: C.textMid, cursor: \"pointer\" }}>Editar</button>
              <button onClick={onDelete} style={{ background: \"none\", border: `1px solid ${C.danger}`, padding: \"5px 10px\", borderRadius: 8, fontSize: 11, fontWeight: 600, color: C.danger, cursor: \"pointer\" }}>Excluir</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function EmptyState({ text }) {
  return (
    <div style={{ textAlign: \"center\", padding: \"32px 16px\", background: C.bgCard, borderRadius: 16, border: `1px solid ${C.border}` }}>
      <span style={{ fontSize: 24, display: \"block\", marginBottom: 6 }}>🍃</span>
      <p style={{ fontSize: 12, fontWeight: 500, color: C.textLight }}>{text}</p>
    </div>
  );
}

// ─── FULL MODAL WINDOW IMPLEMENTATION ─────────────────────────────────────────
function FormModal({ type, item, onSave, onClose }) {
  // Common Fields
  const [nome, setNome] = useState(item?.nome || \"\");
  const [local, setLocal] = useState(item?.local || \"\");
  const [valorUSD, setValorUSD] = useState(item?.valorUSD || \"\");
  const [qtd, setQtd] = useState(item?.qtd || 1);
  const [categoria, setCategoria] = useState(item?.categoria || \"compras\");
  const [link, setLink] = useState(item?.link || \"\");
  const [nota, setNota] = useState(item?.nota || \"\");

  // Gasto Integrantes Selection with Custom Unequal Splits
  const [gastoPessoas, setGastoPessoas] = useState(item?.pessoas || []);

  // Parcela Specific Fields
  const [descParcela, setDescParcela] = useState(item?.descricao || \"\");
  const [valTotalParcela, setValTotalParcela] = useState(item?.valorTotal || \"\");
  const [qtyParcelas, setQtyParcelas] = useState(item?.qtdParcelas || \"\");
  const [valParcela, setValParcela] = useState(item?.valorParcela || \"\");
  const [cardUsed, setCardUsed] = useState(item?.cartao || \"\");

  // Auto compute installment values for Parcela Form
  useEffect(() => {
    if (type === \"parcela\") {
      const tot = parseFloat(valTotalParcela) || 0;
      const qty = parseInt(qtyParcelas) || 1;
      if (tot > 0 && qty > 0) {
        setValParcela((tot / qty).toFixed(2));
      }
    }
  }, [valTotalParcela, qtyParcelas, type]);

  const handleTogglePessoa = (name) => {
    const exists = gastoPessoas.some(p => p.nome === name);
    if (exists) {
      const next = gastoPessoas.filter(p => p.nome !== name);
      // Redividir igualmente de forma automática ao remover
      const baseShare = (parseFloat(valorUSD) || 0) / (next.length + 1);
      setGastoPessoas(next.map(p => ({ ...p, valorCustom: baseShare.toFixed(2) })));
    } else {
      const next = [...gastoPessoas, { nome: name, valorCustom: \"\" }];
      // Redividir igualmente de forma automática ao adicionar
      const baseShare = (parseFloat(valorUSD) || 0) / (next.length + 1);
      setGastoPessoas(next.map(p => ({ ...p, valorCustom: baseShare.toFixed(2) })));
    }
  };

  const handleUpdateGastoCustomValue = (index, val) => {
    const next = [...gastoPessoas];
    next[index].valorCustom = val;
    setGastoPessoas(next);
  };

  const submit = (e) => {
    e.preventDefault();
    if (type === \"parcela\") {
      if (!descParcela || !valTotalParcela) return alert(\"Preencha descrição e valor total!\");
      onSave({
        id: item?.id,
        descricao: descParcela,
        valorTotal: parseFloat(valTotalParcela),
        qtdParcelas: parseInt(qtyParcelas) || 1,
        valorParcela: parseFloat(valParcela),
        cartao: cardUsed,
        statusMeses: item?.statusMeses || Array(12).fill(false)
      });
      return;
    }

    if (!nome || !valorUSD) return alert(\"Preencha o nome e o valor!\");

    const payload = {
      id: item?.id,
      nome,
      local,
      valorUSD: parseFloat(valorUSD),
      categoria,
      nota,
      ...(type !== \"gasto\" ? { qtd: parseInt(qtd) || 1, link } : { pessoas: gastoPessoas })
    };
    onSave(payload);
  };

  const titleText = item ? `Editar ${type}` : `Novo ${type === 'itemLegal' ? 'Item Legal' : type}`;

  return (
    <div style={{ position: \"fixed\", top: 0, left: 0, right: 0, bottom: 0, background: \"rgba(15,23,42,0.4)\", backdropFilter: \"blur(4px)\", display: \"flex\", alignItems: \"flex-end\", zIndex: 200 }}>
      <form onSubmit={submit} style={{ background: C.bgCard, width: \"100%\", maxHeigh: \"90vh\", overflowY: \"auto\", borderTopLeftRadius: 24, borderTopRightRadius: 24, padding: 20, boxShadow: \"0 -10px 25px rgba(0,0,0,0.1)\", display: \"flex\", flexDirection: \"column\", gap: 12 }}>
        <div style={{ display: \"flex\", justifyContent: \"space-between\", alignItems: \"center\", paddingBottom: 4 }}>
          <h3 style={{ fontSize: 15, fontWeight: 800, color: C.text, textTransform: \"capitalize\" }}>{titleText}</h3>
          <button type=\"button\" onClick={onClose} style={{ background: C.borderLight, border: \"none\", width: 28, height: 28, borderRadius: \"50%\", fontSize: 12, fontWeight: 700, color: C.textMid, cursor: \"pointer\" }}>✕</button>
        </div>

        {type === \"parcela\" ? (
          <>
            <div>
              <label style={{ display: \"block\", fontSize: 11, fontWeight: 600, color: C.textMid, marginBottom: 4 }}>DESCRIÇÃO COMPRA</label>
              <input type=\"text\" value={descParcela} onChange={e => setDescParcela(e.target.value)} placeholder=\"Ex: PASSAGEM, CASA, UNIVERSAL\" style={{ width: \"100%\", padding: 10, borderRadius: 10, border: `1px solid ${C.border}`, fontSize: 13, outline: \"none\", boxSizing: \"border-box\" }} />
            </div>
            <div style={{ display: \"grid\", gridTemplateColumns: \"1fr 1fr\", gap: 10 }}>
              <div>
                <label style={{ display: \"block\", fontSize: 11, fontWeight: 600, color: C.textMid, marginBottom: 4 }}>TOTAL COMPRA (R$)</label>
                <input type=\"number\" step=\"0.01\" value={valTotalParcela} onChange={e => setValTotalParcela(e.target.value)} placeholder=\"5134.21\" style={{ width: \"100%\", padding: 10, borderRadius: 10, border: `1px solid ${C.border}`, fontSize: 13, outline: \"none\", boxSizing: \"border-box\" }} />
              </div>
              <div>
                <label style={{ display: \"block\", fontSize: 11, fontWeight: 600, color: C.textMid, marginBottom: 4 }}>QTD PARCELAS</label>
                <input type=\"number\" value={qtyParcelas} onChange={e => setQtyParcelas(e.target.value)} placeholder=\"4\" style={{ width: \"100%\", padding: 10, borderRadius: 10, border: `1px solid ${C.border}`, fontSize: 13, outline: \"none\", boxSizing: \"border-box\" }} />
              </div>
            </div>
            <div style={{ display: \"grid\", gridTemplateColumns: \"1fr 1fr\", gap: 10 }}>
              <div>
                <label style={{ display: \"block\", fontSize: 11, fontWeight: 600, color: C.textMid, marginBottom: 4 }}>VALOR PARCELA (R$)</label>
                <input type=\"number\" step=\"0.01\" value={valParcela} onChange={e => setValParcela(e.target.value)} style={{ width: \"100%\", padding: 10, borderRadius: 10, border: `1px solid ${C.border}`, fontSize: 13, outline: \"none\", boxSizing: \"border-box\", background: C.bg }} readOnly />
              </div>
              <div>
                <label style={{ display: \"block\", fontSize: 11, fontWeight: 600, color: C.textMid, marginBottom: 4 }}>CARTÃO / BANCO</label>
                <input type=\"text\" value={cardUsed} onChange={e => setCardUsed(e.target.value)} placeholder=\"Ex: Itaú Personalité\" style={{ width: \"100%\", padding: 10, borderRadius: 10, border: `1px solid ${C.border}`, fontSize: 13, outline: \"none\", boxSizing: \"border-box\" }} />
              </div>
            </div>
          </>
        ) : (
          <>
            <div>
              <label style={{ display: \"block\", fontSize: 11, fontWeight: 600, color: C.textMid, marginBottom: 4 }}>NOME / IDENTIFICAÇÃO</label>
              <input type=\"text\" value={nome} onChange={e => setNome(e.target.value)} placeholder=\"Ex: Tênis Nike, Almoço, Gasolina\" style={{ width: \"100%\", padding: 10, borderRadius: 10, border: `1px solid ${C.border}`, fontSize: 13, outline: \"none\", boxSizing: \"border-box\" }} />
            </div>

            <div style={{ display: \"grid\", gridTemplateColumns: \"2fr 1fr\", gap: 10 }}>
              <div>
                <label style={{ display: \"block\", fontSize: 11, fontWeight: 600, color: C.textMid, marginBottom: 4 }}>LOJA / LOCAL</label>
                <input type=\"text\" value={local} onChange={e => setLocal(e.target.value)} placeholder=\"Ex: Walmart, Target\" style={{ width: \"100%\", padding: 10, borderRadius: 10, border: `1px solid ${C.border}`, fontSize: 13, outline: \"none\", boxSizing: \"border-box\" }} />
              </div>
              <div>
                <label style={{ display: \"block\", fontSize: 11, fontWeight: 600, color: C.textMid, marginBottom: 4 }}>VALOR (USD)</label>
                <input type=\"number\" step=\"0.01\" value={valorUSD} onChange={e => {
                  setValorUSD(e.target.value);
                  // Atualizar divisão igualitária ao redefinir o valor base
                  if (type === \"gasto\" && gastoPessoas.length > 0) {
                    const share = (parseFloat(e.target.value) || 0) / (gastoPessoas.length + 1);
                    setGastoPessoas(gastoPessoas.map(p => ({ ...p, valorCustom: share.toFixed(2) })));
                  }
                }} placeholder=\"0.00\" style={{ width: \"100%\", padding: 10, borderRadius: 10, border: `1px solid ${C.border}`, fontSize: 13, fontWeight: 700, outline: \"none\", boxSizing: \"border-box\" }} />
              </div>
            </div>

            <div style={{ display: \"grid\", gridTemplateColumns: \"1fr 1fr\", gap: 10 }}>
              <div>
                <label style={{ display: \"block\", fontSize: 11, fontWeight: 600, color: C.textMid, marginBottom: 4 }}>CATEGORIA</label>
                <select value={categoria} onChange={e => setCategoria(e.target.value)} style={{ width: \"100%\", padding: 10, borderRadius: 10, border: `1px solid ${C.border}`, background: C.bgCard, fontSize: 13, outline: \"none\" }}>
                  {CATEGORIES.map(c => <option key={c.id} value={c.id}>{c.icon} {c.label}</option>)}
                </select>
              </div>
              {type !== \"gasto\" ? (
                <div>
                  <label style={{ display: \"block\", fontSize: 11, fontWeight: 600, color: C.textMid, marginBottom: 4 }}>QUANTIDADE</label>
                  <input type=\"number\" value={qtd} onChange={e => setQtd(e.target.value)} style={{ width: \"100%\", padding: 10, borderRadius: 10, border: `1px solid ${C.border}`, fontSize: 13, outline: \"none\", boxSizing: \"border-box\" }} />
                </div>
              ) : (
                <div style={{ display: \"flex\", flexDirection: \"column\", justifyContent: \"flex-end\" }}>
                  <span style={{ fontSize: 11, fontWeight: 600, color: C.textLight, marginBottom: 2 }}>DIVISÃO DE GASTOS:</span>
                  <div style={{ display: \"flex\", gap: 8 }}>
                    <button type=\"button\" onClick={() => handleTogglePessoa(\"Pai\")} style={{ flex: 1, padding: 8, borderRadius: 8, border: `1px solid ${gastoPessoas.some(p=>p.nome===\"Pai\") ? C.primary : C.border}`, background: gastoPessoas.some(p=>p.nome===\"Pai\") ? C.primaryLight : \"none\", color: gastoPessoas.some(p=>p.nome===\"Pai\") ? C.primary : C.textMid, fontSize: 12, fontWeight: 700, cursor: \"pointer\" }}>
                      👨🏻 Pai
                    </button>
                  </div>
                </div>
              )}
            </div>

            {/* EXPANSÃO DA DIVISÃO DESIGUAL DE GASTOS */}
            {type === \"gasto\" && gastoPessoas.length > 0 && (
              <div style={{ background: C.bg, padding: 12, borderRadius: 12, border: `1px solid ${C.borderLight}`, marginTop: 4 }}>
                <span style={{ fontSize: 11, fontWeight: 700, color: C.textMid, display: \"block\", marginBottom: 8 }}>AJUSTAR SALDO POR PESSOA (EDITÁVEL):</span>
                <div style={{ display: \"flex\", flexDirection: \"column\", gap: 8 }}>
                  {gastoPessoas.map((p, idx) => (
                    <div key={p.nome} style={{ display: \"flex\", alignItems: \"center\", justifyContent: \"space-between\", gap: 10 }}>
                      <span style={{ fontSize: 12, fontWeight: 600, color: C.text }}>Share do {p.nome}:</span>
                      <div style={{ display: \"flex\", alignItems: \"center\", background: C.bgCard, border: `1px solid ${C.border}`, borderRadius: 8, padding: \"4px 8px\", width: 120 }}>
                        <span style={{ fontSize: 12, color: C.textLight, marginRight: 2 }}>$</span>
                        <input type=\"number\" step=\"0.01\" value={p.valorCustom} onChange={e => handleUpdateGastoCustomValue(idx, e.target.value)} style={{ border: \"none\", width: \"100%\", fontSize: 12, fontWeight: 700, outline: \"none\", color: C.text }} placeholder=\"0.00\" />
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {type !== \"gasto\" && (
              <div>
                <label style={{ display: \"block\", fontSize: 11, fontWeight: 600, color: C.textMid, marginBottom: 4 }}>LINK DO SITE (OPCIONAL)</label>
                <input type=\"url\" value={link} onChange={e => setLink(e.target.value)} placeholder=\"https://...\" style={{ width: \"100%\", padding: 10, borderRadius: 10, border: `1px solid ${C.border}`, fontSize: 13, outline: \"none\", boxSizing: \"border-box\" }} />
              </div>
            )}

            <div>
              <label style={{ display: \"block\", fontSize: 11, fontWeight: 600, color: C.textMid, marginBottom: 4 }}>OBSERVAÇÕES / NOTAS</label>
              <textarea value={nota} onChange={e => setNota(e.target.value)} placeholder=\"Cor, tamanho, detalhes adicionais...\" style={{ width: \"100%\", padding: 10, borderRadius: 10, border: `1px solid ${C.border}`, fontSize: 13, outline: \"none\", boxSizing: \"border-box\", height: 50, fontFamily: \"inherit\", resize: \"none\" }} />
            </div>
          </>
        )}

        <button type=\"submit\" style={{ width: \"100%\", padding: 12, background: `linear-gradient(135deg, ${C.gradientA}, ${C.gradientB})`, color: \"#FFF\", border: \"none\", borderRadius: 12, fontSize: 14, fontWeight: 700, cursor: \"pointer\", marginTop: 6, boxShadow: \"0 4px 12px rgba(37,99,235,0.15)\" }}>
          Salvar Alterações na Nuvem
        </button>
      </form>
    </div>
  );
}

// ─── STYLES VARIABLES INJECTION ──────────────────────────────────────────────
const CSS = `
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
  * { box-sizing: border-box; margin: 0; padding: 0; -webkit-tap-highlight-color: transparent; }
  body { background: #F8FAFC; }
  select option { background: #fff; }
  ::-webkit-scrollbar { width: 3px; height: 3px; }
  ::-webkit-scrollbar-thumb { background: #CBD5E1; border-radius: 2px; }
  .notif { position: fixed; top: 12px; left: 50%; transform: translateX(-50%); z-index: 300; padding: 10px 20px; border-radius: 12px; font-size: 13px; font-weight: 700; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); animation: slideDown 0.2s ease-out; }
  @keyframes slideDown { from { top: -40px; opacity: 0; } to { top: 12px; opacity: 1; } }
  @keyframes spin { to { transform: rotate(360deg); } }
`;
