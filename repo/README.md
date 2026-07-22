# Plano de Investimento — Macrorregião Sudeste

Dashboard estático (Previsão × Execução) publicado em
**https://planosudeste.netlify.app/**, gerado automaticamente a partir da
planilha de Monitoramento Fonte 95.

## Como funciona

Este repositório **não usa GitHub Pages** — ele só guarda o código e a
planilha, e roda o robô (GitHub Actions). O site publicado de verdade é
o Netlify.

Um workflow (`.github/workflows/atualizar.yml`) dispara sozinho **toda vez
que você sobe uma planilha nova na pasta `dados/`**:

1. Lê a planilha (`dados/*.xlsx`) — abas **Consolidado** e **Plano Sudeste**.
2. Roda `scripts/build_dashboard.py`, que gera o `index.html` completo.
3. Publica esse `index.html` direto no Netlify (via `netlify-cli`).
4. Guarda uma cópia do `index.html` gerado no próprio repositório, só como
   histórico.

Você não precisa rodar nada manualmente — só substituir a planilha.

## Configuração inicial (só uma vez)

### 1. Pegar o Site ID do Netlify

No painel do Netlify: entre no site `planosudeste` → `Site configuration`
→ `General` → `Site details` → copie o **Site ID** (parece um UUID, tipo
`a1b2c3d4-...`).

### 2. Gerar um Personal Access Token do Netlify

No Netlify: clique no seu avatar (canto superior direito) → `User settings`
→ `Applications` → `Personal access tokens` → `New access token` → dê um
nome (ex.: "GitHub Actions") → gere e **copie o token na hora** (ele só
aparece uma vez).

### 3. Cadastrar os dois segredos no GitHub

No repositório: `Settings` → `Secrets and variables` → `Actions` →
`New repository secret`, duas vezes:

- **Name**: `NETLIFY_SITE_ID` → **Value**: o Site ID do passo 1
- **Name**: `NETLIFY_AUTH_TOKEN` → **Value**: o token do passo 2

### 4. Subir a planilha pela primeira vez

Arraste o arquivo `.xlsx` mais atual para a pasta `dados/` do repositório
(pelo site do GitHub: entre na pasta `dados` → `Add file` → `Upload files`)
e confirme o commit. Isso já dispara o robô automaticamente.

## No dia a dia: como atualizar

Toda vez que a planilha do Monitoramento Fonte 95 mudar:

1. Baixe a versão mais recente do OneDrive/SharePoint.
2. No GitHub, entre na pasta `dados/` do repositório → `Add file` →
   `Upload files` → arraste o `.xlsx` novo (ele substitui o antigo, pode
   manter o mesmo nome ou não — o robô pega qualquer `.xlsx` que estiver
   na pasta).
3. Confirme o commit. Em menos de um minuto o Netlify já está atualizado.

Se quiser, também dá pra disparar manualmente sem subir arquivo novo:
`Actions` → `Atualizar Dashboard (Netlify)` → `Run workflow` (útil se você
só quer forçar uma nova publicação com a planilha que já está lá).

## Estrutura do repositório

```
index.html                       ← última versão gerada (guardada como histórico — não edite à mão)
dados/                            ← a planilha atual fica aqui (dispara o robô quando muda)
scripts/build_dashboard.py       ← TODA a lógica de negócio vive aqui
.github/workflows/atualizar.yml  ← o robô que gera e publica no Netlify
```

## Quero mudar alguma regra (cor, agrupamento, fórmula etc.)

Edite `scripts/build_dashboard.py` — é o único lugar onde a lógica de
negócio existe. Depois, dispare o workflow manualmente
(`Actions → Run workflow`) para republicar com a mudança, mesmo sem trocar
a planilha.

## Rodar localmente (para testar antes de subir)

```bash
pip install openpyxl
python3 scripts/build_dashboard.py caminho/para/planilha.xlsx index.html
```

Abra o `index.html` gerado no navegador para conferir antes de subir a
planilha para o repositório.
