import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { VitePWA } from 'vite-plugin-pwa'

export default defineConfig({
  plugins: [
    react(),
    VitePWA({
      registerType: 'autoUpdate',
      // Ativar SW imediatamente sem esperar reload
      injectRegister: 'auto',
      workbox: {
        clientsClaim: true,
        skipWaiting: true,
        globPatterns: ['**/*.{js,css,html,ico,png,svg,webmanifest}'],
        // Aumentar limite para bundles grandes do Firebase
        maximumFileSizeToCacheInBytes: 5 * 1024 * 1024,
        runtimeCaching: [
          // Fontes Google
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
          // Firebase Firestore — NetworkFirst com fallback de cache
          {
            urlPattern: /^https:\/\/firestore\.googleapis\.com\/.*/i,
            handler: 'NetworkFirst',
            options: {
              cacheName: 'firestore-cache',
              networkTimeoutSeconds: 5,
              expiration: { maxEntries: 20, maxAgeSeconds: 60*60*24 },
              cacheableResponse: { statuses: [0, 200] }
            }
          },
          // Firebase Auth
          {
            urlPattern: /^https:\/\/identitytoolkit\.googleapis\.com\/.*/i,
            handler: 'NetworkFirst',
            options: {
              cacheName: 'firebase-auth-cache',
              networkTimeoutSeconds: 5,
              expiration: { maxEntries: 5, maxAgeSeconds: 60*60 },
              cacheableResponse: { statuses: [0, 200] }
            }
          },
          // Cotação (AwesomeAPI / Yahoo proxy)
          {
            urlPattern: /^https:\/\/economia\.awesomeapi\.com\.br\/.*/i,
            handler: 'NetworkFirst',
            options: {
              cacheName: 'cotacao-cache',
              networkTimeoutSeconds: 4,
              expiration: { maxEntries: 3, maxAgeSeconds: 300 },
              cacheableResponse: { statuses: [0, 200] }
            }
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
