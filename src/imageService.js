// ─── IMAGE SERVICE ─────────────────────────────────────────────────────────
// Cache em memória para não refazer buscas na mesma sessão
export const imageCache = {};

// Proxy gratuito para contornar CORS
const PROXY = 'https://api.allorigins.win/get?url=';

// Extrai a melhor imagem de produto de um HTML cru
function extractBestImage(html, origin = '') {
  const candidates = [];

  // 1. og:image — geralmente a foto principal do produto
  const ogMatches = [
    ...html.matchAll(/<meta[^>]+property=["']og:image(?::secure_url)?["'][^>]+content=["']([^"']+)["']/gi),
    ...html.matchAll(/<meta[^>]+content=["']([^"']+)["'][^>]+property=["']og:image(?::secure_url)?["']/gi),
  ];
  ogMatches.forEach(m => m[1] && candidates.push({ url: m[1], score: 100 }));

  // 2. twitter:image
  const twMatches = [
    ...html.matchAll(/<meta[^>]+name=["']twitter:image(?::src)?["'][^>]+content=["']([^"']+)["']/gi),
    ...html.matchAll(/<meta[^>]+content=["']([^"']+)["'][^>]+name=["']twitter:image(?::src)?["']/gi),
  ];
  twMatches.forEach(m => m[1] && candidates.push({ url: m[1], score: 90 }));

  // 3. JSON-LD — Amazon, Walmart e outros colocam dados estruturados
  const jsonLdMatches = html.matchAll(/<script[^>]+type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi);
  for (const match of jsonLdMatches) {
    try {
      const data = JSON.parse(match[1]);
      const imgs = [data.image, data?.offers?.image, data?.mainEntity?.image].flat().filter(Boolean);
      imgs.forEach(img => {
        const url = typeof img === 'string' ? img : img?.url;
        if (url) candidates.push({ url, score: 95 });
      });
    } catch {}
  }

  // 4. <img> com atributos indicando produto (itemprop, data-zoom, class contendo "product")
  const imgMatches = html.matchAll(/<img[^>]+src=["']([^"']+)["'][^>]*>/gi);
  for (const m of imgMatches) {
    const tag = m[0].toLowerCase();
    const src = m[1];
    if (!src || src.startsWith('data:') || src.includes('logo') || src.includes('icon') || src.includes('sprite')) continue;
    let score = 10;
    if (tag.includes('itemprop="image"')) score = 85;
    else if (tag.includes('data-zoom') || tag.includes('data-large')) score = 80;
    else if (tag.includes('product')) score = 60;
    else if (src.includes('product') || src.includes('item')) score = 40;
    // Prefere imagens maiores (heurística pelo tamanho do nome)
    if (src.length > 60) score += 5;
    candidates.push({ url: src, score });
  }

  if (!candidates.length) return null;

  // Resolve URLs relativas
  candidates.forEach(c => {
    if (c.url.startsWith('//')) c.url = 'https:' + c.url;
    else if (c.url.startsWith('/') && origin) c.url = origin + c.url;
  });

  // Ordena pela melhor pontuação e retorna a vencedora
  candidates.sort((a, b) => b.score - a.score);
  return candidates[0].url;
}

// Busca via proxy — retorna HTML da página
async function fetchViaProxy(url, timeout = 9000) {
  const res = await fetch(PROXY + encodeURIComponent(url), {
    signal: AbortSignal.timeout(timeout),
  });
  const json = await res.json();
  return json.contents || '';
}

// DuckDuckGo image search — gratuito, sem chave
async function fetchDDGImage(query) {
  try {
    const searchUrl = `https://duckduckgo.com/?q=${encodeURIComponent(query)}&iax=images&ia=images`;
    const html = await fetchViaProxy(searchUrl, 10000);
    const vqd = html.match(/vqd=([\d-]+)/)?.[1];
    if (!vqd) return null;

    const apiUrl = `https://duckduckgo.com/i.js?q=${encodeURIComponent(query)}&vqd=${vqd}&p=1&o=json`;
    const apiHtml = await fetchViaProxy(apiUrl, 10000);
    const data = JSON.parse(apiHtml || '{}');
    const results = data.results || [];

    // Filtra resultados com tamanho razoável e sem logos
    const good = results.find(r =>
      r.image &&
      r.width > 200 &&
      r.height > 200 &&
      !r.image.includes('logo') &&
      !r.image.includes('icon')
    );
    return good?.image || results[0]?.image || null;
  } catch {
    return null;
  }
}

// Função principal exportada
export async function getProductImage(produto) {
  const key = String(produto.id);

  // Cache hit
  if (imageCache[key] !== undefined) return imageCache[key];

  // URL de imagem manual tem prioridade absoluta
  if (produto.imagem) {
    imageCache[key] = produto.imagem;
    return produto.imagem;
  }

  // 1. Tenta extrair imagem do link do anúncio
  if (produto.link) {
    try {
      const html = await fetchViaProxy(produto.link);
      if (html) {
        // Tenta descobrir a origem para resolver URLs relativas
        let origin = '';
        try { origin = new URL(produto.link).origin; } catch {}
        const img = extractBestImage(html, origin);
        if (img) {
          imageCache[key] = img;
          return img;
        }
      }
    } catch {}
  }

  // 2. Fallback: DuckDuckGo com nome + loja
  try {
    const query = `${produto.nome} ${produto.loja} product`;
    const ddg = await fetchDDGImage(query);
    if (ddg) {
      imageCache[key] = ddg;
      return ddg;
    }
  } catch {}

  // 3. Nada encontrado — usa emoji
  imageCache[key] = null;
  return null;
}
