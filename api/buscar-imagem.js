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
        "User-Agent": "Mozilla/5.0"
      }
    });

    const html = await response.text();

    let image = null;

    const ogImage = html.match(
      /<meta[^>]*property=["']og:image["'][^>]*content=["']([^"']+)["']/i
    );

    if (ogImage) {
      image = ogImage[1];
    }

    if (!image) {
      const twitterImage = html.match(
        /<meta[^>]*name=["']twitter:image["'][^>]*content=["']([^"']+)["']/i
      );

      if (twitterImage) {
        image = twitterImage[1];
      }
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