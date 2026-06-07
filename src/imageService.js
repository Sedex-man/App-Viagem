export const imageCache = {};

export async function getProductImage(produto) {

  if (produto.imagem) {
    return produto.imagem;
  }

  if (!produto.link) {
    return null;
  }

  try {

    const response = await fetch(
      `/api/buscar-imagem?url=${encodeURIComponent(produto.link)}`
    );

    const data = await response.json();

    if (data.image) {

      imageCache[String(produto.id)] = data.image;

      return data.image;
    }

  } catch (err) {
    console.error(err);
  }

  return null;
}