# Mapa de Fornecedores (página publicada)

Página do Mapa de Fornecedores do time de manutenção da Transmaroni, publicada pelo GitHub Pages.

- Acesso só com a conta Microsoft da empresa (Entra ID, locatário único). A senha é digitada na página da Microsoft; este site não a vê nem guarda.
- Nenhum dado de fornecedor fica neste repositório. Os dados ficam em uma pasta do SharePoint da empresa, com acesso restrito, e só são lidos depois do login.
- O código que gera o mapa fica em um repositório privado.

Biblioteca de login: `vendor/msal-browser.min.js` e `vendor/msal-redirect-bridge.min.js` (Microsoft Authentication Library, `@azure/msal-browser` 5.23.0, licença MIT em `vendor/LICENSE-msal.txt`), servidas por este mesmo site.
