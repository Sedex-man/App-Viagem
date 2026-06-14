// Cache em memória (por id de produto) para evitar recomputar URLs já resolvidas na sessão
export const imageCache = {};

// Nome do banco de armazenamento offline de imagens no celular
const IMAGE_CACHE_NAME = "app-imagens-produtos-v1";

// Lê uma imagem salva localmente no dispositivo (Cache Storage), pela URL original
async function getCachedImage(url) {
  try {
    if (!('caches' in window)) return null;
    const cache = await caches.open(IMAGE_CACHE_NAME);
    const response = await cache.match(url);
    if (response) {
      const blob = await response.blob();
      return URL.createObjectURL(blob);
    }
  } catch (e) {
    console.error("Erro ao ler cache offline de imagem:", e);
  }
  return null;
}

// Salva uma cópia da imagem no Cache Storage para uso offline
async function salvarImagemOffline(url) {
  try {
    if (!('caches' in window)) return;
    const cache = await caches.open(IMAGE_CACHE_NAME);
    await cache.add(new Request(url, { mode: 'cors' }));
  } catch (err) {
    // Alguns sites bloqueiam o download direto (CORS). Se falhar, segue usando a URL normal.
    console.warn("Não foi possível salvar esta imagem para o modo offline:", err?.message || err);
  }
}

/**
 * Resolve a imagem de um produto a partir da URL colada em produto.imagem.
 * 1. Tenta usar a cópia salva no dispositivo (funciona offline).
 * 2. Se não houver cópia local e o app estiver online, exibe a URL direta
 *    e salva uma cópia em segundo plano para a próxima vez.
 * 3. Sem imagem cadastrada, retorna null (cai no ícone padrão).
 */
export async function getProductImage(produto) {
  if (!produto || !produto.imagem) {
    return null;
  }

  const urlOriginal = produto.imagem;

  // 1. Tenta recuperar do cache offline do aparelho primeiro
  const imagemLocal = await getCachedImage(urlOriginal);
  if (imagemLocal) {
    return imagemLocal;
  }

  // 2. Se estiver online, salva uma cópia para a próxima vez (sem bloquear a exibição)
  if (typeof navigator !== 'undefined' && navigator.onLine) {
    salvarImagemOffline(urlOriginal);
  }

  // 3. Exibe a URL direta enquanto isso
  return urlOriginal;
}
