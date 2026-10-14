# Sorteio Grupo Plasfran (formulário do evento)

Formulário mobile (HTML puro) → Google Apps Script → Google Sheets.

- `index.html` + `bg.webp`: a página.
- `apps-script/Code.gs`: script que grava na planilha (aba única **Leads**, com a coluna Marca), numera em sequência única (limite 1000) e devolve o mesmo número se telefone ou e-mail já existir.
- `qrcode-sorteio.png` / `.svg`: QR Code para https://grupoplasfransorteio.vercel.app/

## Configuração (uma vez)
1. Crie uma planilha no Google Sheets (nome sugerido: "Sorteio Plasfran").
2. Extensões > Apps Script. Apague o conteúdo e cole `apps-script/Code.gs`. Salve.
3. Implantar > Nova implantação > tipo **App da Web**: executar como **Eu**; quem tem acesso: **Qualquer pessoa**. Autorize quando pedir.
4. Copie a URL terminada em `/exec` e cole em `API_URL` no `index.html`.
5. Faça o deploy (projeto Vercel com *Root Directory* = `evento-plasfran`).
6. Teste enviando um cadastro; depois apague as linhas de teste da planilha (a numeração recomeça do maior número restante).
