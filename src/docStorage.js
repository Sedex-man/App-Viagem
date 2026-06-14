import { db } from "./firebase";
import { collection, doc, setDoc, deleteDoc, onSnapshot, query, orderBy, getDocs } from "firebase/firestore";

// ── Limites ─────────────────────────────────────────────────────────────────
// Firestore limita cada documento a ~1MB. Base64 já adiciona ~33% de overhead,
// então usamos um teto seguro por chunk e dividimos arquivos grandes em várias
// "partes" (subdocumentos), remontando ao ler.
const CHUNK_SIZE = 700_000; // ~700KB de base64 por chunk (margem segura)
const MAX_FILE_SIZE = 1_572_864; // ~1.5MB de arquivo original (após compressão, se aplicável)

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

async function cacheArquivoLocal(item) {
  const idb = await openDB();
  return new Promise((resolve, reject) => {
    const tx = idb.transaction(STORE_NAME, "readwrite");
    tx.objectStore(STORE_NAME).put(item);
    tx.oncomplete = () => resolve(item);
    tx.onerror = () => reject(tx.error);
  });
}

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

// Comprime uma imagem (reduz dimensão/qualidade) até caber no limite, mantendo
// formato JPEG. Retorna a data URL comprimida. Não-imagens retornam null (sem alteração).
async function comprimirImagem(file, maxBytes) {
  if (!file.type.startsWith("image/")) return null;

  const bitmap = await createImageBitmap(file);
  let { width, height } = bitmap;
  let quality = 0.85;
  let scale = 1;

  for (let tentativa = 0; tentativa < 8; tentativa++) {
    const w = Math.round(width * scale);
    const h = Math.round(height * scale);
    const canvas = document.createElement("canvas");
    canvas.width = w; canvas.height = h;
    const ctx = canvas.getContext("2d");
    ctx.drawImage(bitmap, 0, 0, w, h);
    const blob = await new Promise(res => canvas.toBlob(res, "image/jpeg", quality));
    if (blob.size <= maxBytes) {
      return await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
      });
    }
    // Reduz qualidade primeiro, depois dimensão
    if (quality > 0.4) quality -= 0.15;
    else scale *= 0.75;
  }
  return null; // não conseguiu comprimir o suficiente
}

// ── Firestore: metadados + conteúdo (em chunks se necessário) ─────────────────
function docsCollection(uid) {
  return collection(db, "usuarios_pwa", uid, "documentos");
}
function chunksCollection(uid, docId) {
  return collection(db, "usuarios_pwa", uid, "documentos", docId, "chunks");
}

// Escuta em tempo real a lista de documentos do usuário (sincroniza entre dispositivos)
export function ouvirDocumentos(uid, callback) {
  const q = query(docsCollection(uid), orderBy("criadoEm", "desc"));
  return onSnapshot(q, snap => {
    callback(snap.docs.map(d => ({ id: d.id, ...d.data() })));
  });
}

// Salva um documento (texto + arquivo opcional). Se o arquivo for grande,
// comprime (se for imagem) e/ou divide em chunks para caber no Firestore.
export async function salvarDocumento(uid, { texto, file }) {
  const id = String(Date.now());
  let meta = { texto: texto || "", criadoEm: Date.now() };

  if (file) {
    let base64 = await fileToBase64(file);
    let tamanhoBase64 = base64.length;

    // Se passar do limite, tenta comprimir (apenas imagens)
    if (tamanhoBase64 > MAX_FILE_SIZE) {
      const comprimida = await comprimirImagem(file, MAX_FILE_SIZE * 0.7);
      if (comprimida) {
        base64 = comprimida;
        tamanhoBase64 = base64.length;
      }
    }

    if (tamanhoBase64 > MAX_FILE_SIZE) {
      throw new Error(`Arquivo muito grande (${(tamanhoBase64/1024).toFixed(0)}KB). Máximo ~1.5MB (imagens são comprimidas automaticamente).`);
    }

    const numChunks = Math.ceil(base64.length / CHUNK_SIZE);
    meta = { ...meta, fileName: file.name, mimeType: file.type, numChunks, totalSize: base64.length };

    // Salva os chunks em paralelo
    await Promise.all(Array.from({length: numChunks}, (_, i) => {
      const parte = base64.slice(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
      return setDoc(doc(chunksCollection(uid, id), String(i)), { data: parte });
    }));

    // Cache local para abrir offline sem precisar baixar de novo
    await cacheArquivoLocal({ id, fileName: file.name, mimeType: file.type, data: base64 }).catch(()=>{});
  }

  await setDoc(doc(docsCollection(uid), id), meta);
  return { id, ...meta };
}

// Remove o documento, seus chunks e o cache local
export async function removerDocumento(uid, item) {
  if (item.numChunks) {
    const snap = await getDocs(chunksCollection(uid, item.id));
    await Promise.all(snap.docs.map(d => deleteDoc(d.ref)));
  }
  await deleteDoc(doc(docsCollection(uid), item.id));
  await removerArquivoLocal(item.id).catch(()=>{});
}

// Obtém o arquivo (base64) para visualização/download: usa o cache local se
// existir (funciona offline); senão remonta os chunks do Firestore.
export async function obterArquivo(uid, item) {
  if (!item.numChunks) return null;

  const local = await obterArquivoLocal(item.id).catch(()=>null);
  if (local?.data) return local.data;

  // Sem cópia local (ex: documento adicionado em outro dispositivo) — remonta do Firestore
  const colSnap = await getDocs(chunksCollection(uid, item.id));
  const map = {};
  colSnap.forEach(d => { map[d.id] = d.data().data; });
  const partes = [];
  for (let i = 0; i < item.numChunks; i++) partes.push(map[String(i)] || "");
  const base64 = partes.join("");

  await cacheArquivoLocal({ id: item.id, fileName: item.fileName, mimeType: item.mimeType, data: base64 }).catch(()=>{});
  return base64;
}
