export default async function handler(req, res) {
  const url = req.query.url;

  if (!url) {
    return res.status(400).json({
      error: "URL não informada"
    });
  }

  try {
    const response = await fetch(url, {
      headers: {
        "User-Agent":
          "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/122 Safari/537.36"
      }
    });

    const html = await response.text();

    let image = null;

    // OG IMAGE
    const og = html.match(
      /<meta[^>]*property=["']og:image["'][^>]*content=["']([^"']+)["']/i
    );

    if (og) image = og[1];

    // TWITTER IMAGE
    if (!image) {
      const tw = html.match(
        /<meta[^>]*name=["']twitter:image["'][^>]*content=["']([^"']+)["']/i
      );

      if (tw) image = tw[1];
    }

    // JSON-LD
    if (!image) {
      const jsonLd = html.match(
        /"image"\s*:\s*"([^"]+)"/i
      );

      if (jsonLd) image = jsonLd[1];
    }

    // Walmart CDN
    if (!image) {
      const walmart = html.match(
        /https:\/\/i5\.walmartimages\.com\/[^"']+/i
      );

      if (walmart) image = walmart[0];
    }

    // Qualquer JPG
    if (!image) {
      const jpg = html.match(
        /https?:\/\/[^"']+\.(jpg|jpeg|png|webp)/i
      );

      if (jpg) image = jpg[0];
    }

    return res.status(200).json({
      image
    });

  } catch (err) {

    return res.status(500).json({
      error: err.message
    });

  }
}