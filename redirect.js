// Ponte de login do MSAL 5: recebe a resposta da Microsoft e devolve para a página que pediu o login.
// Fica neste mesmo site (nunca em CDN), como manda a documentação da Microsoft, porque recebe o código de autorização.
// Fica num arquivo à parte, e não dentro do redirect.html, porque o CSP dessa página só deixa rodar script do próprio site.
msalRedirectBridge.broadcastResponseToMainFrame().catch(e => {
  document.getElementById('msg').textContent = 'Não foi possível concluir o login. Volte à página do mapa e tente de novo.';
  console.error(e);
});
