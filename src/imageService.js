export const imageCache = {};

/**
 * Converte uma URL de imagem em uma String permanente de Texto (Base64)
 */
async function urlToBase64(url) {
  try {
    const response = await fetch(url, { mode: 'cors' });
    const blob = await response.blob();
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  } catch (e) {
    console.warn("Não foi possível converter a imagem para Base64 de forma offline:", e);
    return null;
  }
}

/**
 * Função principal chamada pelo App.jsx
 */
export async function getProductImage(produto) {
  if (!produto || !produto.imagem) {
    return null;
  }

  const urlOriginal = produto.imagem;
  const storageKey = `img_perm_${produto.id}`;

  // 1. Tenta pegar da memória RAM (Fast cache)
  if (imageCache[produto.id]) {
    return imageCache[produto.id];
  }

  // 2. Tenta recuperar do armazenamento permanente do Celular (LocalStorage)
  const imagemSalvaNoDisco = localStorage.getItem(storageKey);
  if (imagemSalvaNoDisco) {
    imageCache[produto.id] = imagemSalvaNoDisco; // Salva na RAM para as próximas leituras rápidas
    return imagemSalvaNoDisco;
  }

  // 3. Se não achou em nenhum lugar e está online, baixa o arquivo, converte em texto e salva para SEMPRE
  if (navigator.onLine) {
    const base64Str = await urlToBase64(urlOriginal);
    if (base64Str) {
      try {
        localStorage.setItem(storageKey, base64Str);
        imageCache[produto.id] = base64Str;
        return base64Str;
      } catch (storageError) {
        // Se o celular estiver sem espaço para o localStorage
        console.warn("Armazenamento local cheio, exibindo via link direto.");
      }
    }
  }

  // Fallback seguro se tudo falhar ou estiver estritamente offline no primeiro carregamento
  return urlOriginal;
}
