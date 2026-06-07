// imageService.js
export const imageCache = {};

export async function getProductImage(produto) {
  const key = String(produto.id);
  if (imageCache[key]) return imageCache[key];

  if (produto.imagem && produto.imagem.trim() !== "") {
    imageCache[key] = produto.imagem.trim();
    return imageCache[key];
  }

  if (!produto.link) return null;

  try {
    // Usando o cors-anywhere ou seu próprio proxy configurado
    const proxyUrl = "https://cors-anywhere.herokuapp.com/"; 
    const response = await fetch(proxyUrl + produto.link, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      }
    });
    
    if (!response.ok) throw new Error("Falha na requisição");
    
    const html = await response.text();
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, "text/html");

    // Lista de seletores Open Graph e Twitter Cards (Mais prováveis de ter a imagem real)
    const metaSelectors = [
      'meta[property="og:image"]',
      'meta[property="og:image:secure_url"]',
      'meta[name="twitter:image"]',
      'meta[name="twitter:image:src"]',
      'link[rel="image_src"]'
    ];

    for (let selector of metaSelectors) {
      const el = doc.querySelector(selector);
      if (el && (el.getAttribute("content") || el.getAttribute("href"))) {
        let url = el.getAttribute("content") || el.getAttribute("href");
        if (url.startsWith("//")) url = "https:" + url;
        if (url.startsWith("/")) {
          const urlObj = new URL(produto.link);
          url = urlObj.origin + url;
        }
        imageCache[key] = url;
        return url;
      }
    }

    // Fallback: Se não achar meta tag, busca a primeira imagem relevante do corpo do texto
    const images = Array.from(doc.querySelectorAll("main img, #content img, article img, .product img"));
    for (let img of images) {
      let src = img.getAttribute("src") || img.getAttribute("data-src");
      if (src && src.match(/\.(jpeg|jpg|gif|png|webp)/i)) {
        if (src.startsWith("//")) src = "https:" + src;
        if (src.startsWith("/")) {
          const urlObj = new URL(produto.link);
          src = urlObj.origin + src;
        }
        imageCache[key] = src;
        return src;
      }
    }

    return null;
  } catch (error) {
    console.error("Erro ao buscar imagem para o produto:", produto.id, error);
    return null;
  }
}