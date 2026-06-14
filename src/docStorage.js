import { db, storage } from "./firebase";
import { collection, doc, setDoc, deleteDoc, onSnapshot, query, orderBy } from "firebase/firestore";
import { ref, uploadString, getDownloadURL, deleteObject } from "firebase/storage";

// ── IndexedDB: cache local dos arquivos para acesso 100% offline ──────────────
const DB_NAME = "travelshop_docs";
const DB_VERSION = 1;
const STORE_NAME = "documentos";

function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const idb = req.result;
      if (!idb.objectStoreNames.contains(STORE_NAME)) {
        idb.createObjectStore(STORE_NAME, { keyPath: "id" });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

// Salva { id, fileName, mimeType, data (base64) } localmente
async function cacheArquivoLocal(item) {
  const idb = await openDB();
  return new Promise((resolve, reject) => {
    const tx = idb.transaction(STORE_NAME, "readwrite");
    tx.objectStore(STORE_NAME).put(item);
    tx.oncomplete = () => resolve(item);
    tx.onerror = () => reject(tx.error);
  });
}

// Busca o arquivo (base64) salvo localmente pelo id
async function obterArquivoLocal(id) {
  const idb = await openDB();
  return new Promise((resolve, reject) => {
    const tx = idb.transaction(STORE_NAME, "readonly");
    const req = tx.objectStore(STORE_NAME).get(id);
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error);
  });
}

async function removerArquivoLocal(id) {
  const idb = await openDB();
  return new Promise((resolve, reject) => {
    const tx = idb.transaction(STORE_NAME, "readwrite");
    tx.objectStore(STORE_NAME).delete(id);
    tx.oncomplete = () => resolve(true);
    tx.onerror = () => reject(tx.error);
  });
}

// Converte um File em base64 (data URL)
export function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// ── Firestore: metadados sincronizados entre dispositivos ─────────────────────
function docsCollection(uid) {
  return collection(db, "usuarios_pwa", uid, "documentos");
}

// Escuta em tempo real a lista de documentos do usuário (sincroniza entre dispositivos)
export function ouvirDocumentos(uid, callback) {
  const q = query(docsCollection(uid), orderBy("criadoEm", "desc"));
  return onSnapshot(q, snap => {
    callback(snap.docs.map(d => ({ id: d.id, ...d.data() })));
  });
}

// Faz upload do arquivo para o Firebase Storage, salva os metadados no Firestore
// e guarda uma cópia local (base64) no IndexedDB para acesso offline.
export async function salvarDocumento(uid, { texto, file }) {
  const id = String(Date.now());
  let meta = { texto: texto || "", criadoEm: Date.now() };

  if (file) {
    const base64 = await fileToBase64(file);
    // Upload para o Storage (data URL no formato "data:<mime>;base64,<dados>")
    const storageRef = ref(storage, `usuarios_pwa/${uid}/documentos/${id}_${file.name}`);
    await uploadString(storageRef, base64, "data_url");
    const url = await getDownloadURL(storageRef);
    meta = { ...meta, fileName: file.name, mimeType: file.type, url, storagePath: storageRef.fullPath };
    // Cache local para abrir offline sem precisar baixar de novo
    await cacheArquivoLocal({ id, fileName: file.name, mimeType: file.type, data: base64 });
  }

  await setDoc(doc(docsCollection(uid), id), meta);
  return { id, ...meta };
}

// Remove o documento do Firestore, do Storage e do cache local
export async function removerDocumento(uid, item) {
  await deleteDoc(doc(docsCollection(uid), item.id));
  if (item.storagePath) {
    try { await deleteObject(ref(storage, item.storagePath)); } catch (e) { console.warn("Erro ao remover do Storage:", e); }
  }
  await removerArquivoLocal(item.id).catch(()=>{});
}

// Obtém o arquivo para visualização/download: usa o cache local se existir
// (funciona offline); caso contrário, baixa do Storage e guarda em cache.
export async function obterArquivo(item) {
  if (!item.url) return null;
  const local = await obterArquivoLocal(item.id).catch(()=>null);
  if (local?.data) return local.data;

  // Sem cópia local (ex: documento adicionado em outro dispositivo) — baixa do Storage
  const resp = await fetch(item.url);
  const blob = await resp.blob();
  const base64 = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
  await cacheArquivoLocal({ id: item.id, fileName: item.fileName, mimeType: item.mimeType, data: base64 }).catch(()=>{});
  return base64;
}
