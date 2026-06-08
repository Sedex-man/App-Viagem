import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'
import { registerSW } from 'virtual:pwa-register'

// Registrar Service Worker com auto-reload silencioso
const updateSW = registerSW({
  // Quando uma nova versão estiver disponível e instalada, recarregar automaticamente
  onNeedRefresh() {
    updateSW(true); // força atualização silenciosa
  },
  onOfflineReady() {
    console.log('[TravelShop] App pronto para uso offline!');
  },
  onRegistered(r) {
    // Verificar atualizações a cada 60 segundos quando online
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
