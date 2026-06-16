import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { VitePWA } from 'vite-plugin-pwa'
export default defineConfig({
  plugins: [
    react(),
    VitePWA({
      registerType: 'autoUpdate',
      injectRegister: 'auto',
      workbox: {
        clientsClaim: true,
        skipWaiting: true,
        globPatterns: ['**/*.{js,css,html,ico,png,svg,webmanifest}'],
        maximumFileSizeToCacheInBytes: 5 * 1024 * 1024,
        // Deixa passar sem interceptar qualquer requisição do Firebase/Google APIs
        // O Firestore usa WebSockets/streaming que o Workbox não consegue cachear
        navigateFallbackDenylist: [/^\/api/],
        runtimeCaching: [
          {
            urlPattern: /^https:\/\/fonts\.googleapis\.com\/.*/i,
            handler: 'CacheFirst',
            options: {
              cacheName: 'google-fonts-cache',
              expiration: { maxEntries: 10, maxAgeSeconds: 60*60*24*365 },
              cacheableResponse: { statuses: [0, 200] }
            }
          },
          {
            urlPattern: /^https:\/\/fonts\.gstatic\.com\/.*/i,
            handler: 'CacheFirst',
            options: {
              cacheName: 'gstatic-fonts-cache',
              expiration: { maxEntries: 10, maxAgeSeconds: 60*60*24*365 },
              cacheableResponse: { statuses: [0, 200] }
            }
          },
          // Firestore e Firebase Auth: NetworkOnly — o SDK do Firebase gerencia
          // o próprio cache offline via persistentLocalCache (IndexedDB).
          // Tentar cachear via Workbox quebra as conexões WebSocket/streaming.
          {
            urlPattern: /^https:\/\/firestore\.googleapis\.com\/.*/i,
            handler: 'NetworkOnly',
          },
          {
            urlPattern: /^https:\/\/firebase\.googleapis\.com\/.*/i,
            handler: 'NetworkOnly',
          },
          {
            urlPattern: /^https:\/\/identitytoolkit\.googleapis\.com\/.*/i,
            handler: 'NetworkOnly',
          },
          {
            urlPattern: /^https:\/\/securetoken\.googleapis\.com\/.*/i,
            handler: 'NetworkOnly',
          },
          // APIs de cotação — NetworkOnly (sem cache, evita erro offline)
          {
            urlPattern: /^https:\/\/economia\.awesomeapi\.com\.br\/.*/i,
            handler: 'NetworkOnly',
          },
          {
            urlPattern: /^https:\/\/olinda\.bcb\.gov\.br\/.*/i,
            handler: 'NetworkOnly',
          },
          {
            urlPattern: /^https:\/\/corsproxy\.io\/.*/i,
            handler: 'NetworkOnly',
          },
        ]
      },
      includeAssets: ['image_a375cf.png', 'favicon.ico'],
      manifest: {
        name: 'TravelShop Orlando',
        short_name: 'TravelShop',
        description: 'Gerenciador de compras internacionais',
        theme_color: '#2563EB',
        background_color: '#F8FAFC',
        display: 'standalone',
        orientation: 'portrait',
        scope: '/',
        start_url: '/',
        icons: [
          { src: 'image_a375cf.png', sizes: '192x192', type: 'image/png', purpose: 'any' },
          { src: 'image_a375cf.png', sizes: '512x512', type: 'image/png', purpose: 'any maskable' }
        ]
      }
    })
  ]
})
