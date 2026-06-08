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
        // Nenhuma API externa em cache — só assets do próprio app
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
          // BCB e qualquer API externa: NetworkOnly (sem cache)
          {
            urlPattern: /^https:\/\/olinda\.bcb\.gov\.br\/.*/i,
            handler: 'NetworkOnly',
          },
          // Bloquear AwesomeAPI explicitamente (evitar que SW antigo interfira)
          {
            urlPattern: /^https:\/\/economia\.awesomeapi\.com\.br\/.*/i,
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
