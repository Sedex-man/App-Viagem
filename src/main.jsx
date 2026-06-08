import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'
import { registerSW } from 'virtual:pwa-register'

// Limpar caches antigos do Service Worker (remove caches de versões anteriores)
if ('caches' in window) {
  caches.keys().then(names => {
    names.forEach(name => {
      // Remover caches de APIs externas que não devem ser cacheadas
      if (name.includes('api-cache') || name.includes('cotacao-cache')) {
        caches.delete(name);
      }
    });
  });
}

// Registrar Service Worker com auto-reload silencioso
const updateSW = registerSW({
  onNeedRefresh() {
    updateSW(true);
  },
  onOfflineReady() {
    console.log('[TravelShop] App pronto para uso offline!');
  },
  onRegistered(r) {
    if (r) {
      setInterval(() => {
        if (navigator.onLine) r.update();
      }, 60_000);
    }
  },
  onRegisterError(error) {
    console.error('[TravelShop] Erro ao registrar SW:', error);
  }
});

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
)
