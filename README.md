# Valei-me ✝ Sugestões

Site de sugestões do perfil [@valeime.catolico](https://instagram.com/valeime.catolico): <https://valeime.netlify.app>

Os seguidores podem sugerir:

- 🎵 **Música**: para virar edit
- ✦ **Santo da Semana**: santo que querem ver em destaque
- ✝ **Ideia de post**

As respostas vão para uma planilha do Google Sheets usando um Web App do Apps Script.

## Estrutura

| Arquivo | Descrição |
| --- | --- |
| `index.html` | Site estático (formulário) publicado no Netlify |
| `logo.png` | Logo Valei-me |
| `apps-script/Code.gs` | Backend: recebe os envios e organiza a planilha nas abas 📊 Resumo, 🎵 Músicas, ✦ Santos e ✝ Posts |
| `apps-script/appsscript.json` | Manifesto do Apps Script |

## Configuração

1. Na planilha, abra **Extensões → Apps Script** e cole o conteúdo de `apps-script/Code.gs` e `appsscript.json`.
2. Em `Code.gs`, troque o valor de `ADMIN_KEY` por uma chave secreta sua.
3. Publique em **Implantar → Nova implantação → App da Web**, com execução como você e acesso para qualquer pessoa.
4. Copie a URL `/exec` para a constante `SHEET_URL` em `index.html`.
5. Acesse `<URL>/exec?k=<ADMIN_KEY>&a=setup` uma vez para criar as abas.

Outras ações de admin disponíveis: `a=status` (contagens em JSON) e `a=limpar` (remove os envios de teste).

## Deploy do site

Arraste a pasta com `index.html` e `logo.png` para o Netlify ou use `netlify deploy --prod`.
