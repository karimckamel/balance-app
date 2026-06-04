# ₿alance — Finanças Pessoais

## Como publicar

### 1. Delete o repositório antigo no GitHub
- github.com/karimckamel/balance-app → Settings → Danger Zone → Delete repository

### 2. Crie um novo repositório privado
- Nome: balance-app
- Marque Private
- NÃO inicialize com README

### 3. Configure os Secrets no GitHub
No novo repositório → Settings → Secrets and variables → Actions → New repository secret:

- Nome: REACT_APP_SUPABASE_URL
  Valor: https://vnqbgqeqmvpshgsyyzff.supabase.co

- Nome: REACT_APP_SUPABASE_ANON_KEY
  Valor: (sua anon key do Supabase)

### 4. Ative o GitHub Pages
- Settings → Pages → Source: selecione "GitHub Actions"

### 5. Faça o upload dos arquivos
No Terminal:
  cd ~/Downloads/balance-v3
  git init
  git remote add origin https://TOKEN@github.com/karimckamel/balance-app.git
  git add .
  git commit -m "init"
  git push -u origin main

O GitHub Actions vai buildar e publicar automaticamente.
O app estará em: https://karimckamel.github.io/balance-app

### 6. Ativar 2FA no app
Após fazer login, clique em 🔐 e escaneie o QR code com Google Authenticator ou Authy.

## Privacidade
- Credenciais NUNCA no código — ficam nos Secrets do GitHub
- Dados criptografados no Supabase com Row Level Security
- Repositório privado
