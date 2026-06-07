# TravelShop PWA 🛍✈

Gerenciador de compras internacionais — funciona como app no iPhone e Android.

## Como publicar (5 minutos, tudo pelo celular)

### Passo 1 — GitHub (repositório)
1. Acesse **github.com** no celular e crie uma conta gratuita (se não tiver)
2. Clique em **"New repository"**
3. Nome: `travelshop` → clique **"Create repository"**
4. Na tela do repositório vazio, clique em **"uploading an existing file"**
5. Faça upload de **todos os arquivos desta pasta** (arraste e solte)
6. Clique **"Commit changes"**

### Passo 2 — Vercel (hospedagem grátis)
1. Acesse **vercel.com** e clique **"Sign up"** → entre com sua conta GitHub
2. Clique **"Add New Project"**
3. Selecione o repositório `travelshop`
4. Clique **"Deploy"** (Vercel detecta Vite automaticamente)
5. Aguarde ~2 minutos → seu app estará em `travelshop.vercel.app`

### Passo 3 — Instalar no iPhone
1. Abra o link no **Safari** (não Chrome!)
2. Toque em **Compartilhar** (ícone de caixa com seta ↑)
3. Role e toque **"Adicionar à Tela de Início"**
4. Confirme → ícone aparece na home screen ✅

### Compartilhar com amigos/família
- Mande o link do Vercel por WhatsApp
- Cada pessoa segue o Passo 3 no próprio celular
- **Dados são individuais** — cada dispositivo tem sua própria lista

## Estrutura
```
travelshop-pwa/
├── public/
│   ├── icon-192.png     (ícone do app)
│   └── icon-512.png     (ícone splash)
├── src/
│   ├── main.jsx         (entrada)
│   └── App.jsx          (app completo)
├── index.html
├── vite.config.js
├── package.json
└── vercel.json
```

## Funcionalidades
- ✅ Dashboard com cotação do dólar
- ✅ Lista de produtos com imagens automáticas
- ✅ Galeria visual
- ✅ Aba de gastos com divisão por pessoas
- ✅ Estatísticas de compras e gastos
- ✅ Calculadora USD→BRL + conversor de peso
- ✅ Importação de planilha Excel
- ✅ **Dados salvos localmente** (não some ao fechar)
- ✅ **Funciona offline** após primeiro acesso

## Dados salvos localmente
Os dados ficam no `localStorage` do navegador de cada pessoa.
Não são compartilhados entre dispositivos — cada um tem a sua lista.
