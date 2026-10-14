// Versão publicada do Mapa de Fornecedores (GitHub Pages): login com a conta Microsoft da Transmaroni (Entra ID, locatário
// único) e leitura dos dados na pasta do SharePoint, com a permissão da própria pessoa. Nenhum dado fica neste site.
// O mapa local, aberto do disco, já vem com os dados embutidos e não usa este arquivo.
(() => {
  const CFG = {
    clientId: 'be92c8b6-d0fc-40e3-8e5c-f75de940e41b',   // registro "Mapa de Fornecedores - Manutenção" (IDs públicos, não são senha)
    tenant: '8babaf2b-248f-4651-98d6-1e5312fbe00a',     // Transmaroni
    site: 'grupotransmaroni.sharepoint.com:/sites/ProTech',   // SharePoint: ProTech - Manutenção Corporativa
    lista: '18419567-be85-4b62-8c5f-e1e8482686e2',           // biblioteca "PCM - FleetControl"
    pasta: 'Mapa de Fornecedores',                           // leitura: "Todos, exceto usuários externos"
    arquivo: 'mapa_dados.json'                               // gravado pelo passo 10 da atualização
  };
  const RAIZ = location.origin + location.pathname.replace(/[^/]*$/, '');
  const ESCOPOS = ['User.Read', 'Files.Read.All'];   // Files.Read.All: só lê o que a própria pessoa já pode abrir no SharePoint
  const pca = new msal.PublicClientApplication({
    auth: {
      clientId: CFG.clientId,
      authority: 'https://login.microsoftonline.com/' + CFG.tenant,
      redirectUri: RAIZ + 'redirect.html',
      postLogoutRedirectUri: RAIZ + 'redirect.html'
    },
    cache: { cacheLocation: 'localStorage' }   // continua conectado ao reabrir o site no celular
  });
  // o MSAL 5 devolve só o código do erro; aqui ele vira uma mensagem para a pessoa
  const ERROS = {
    consent_required: 'O acesso precisa ser aprovado por um administrador da Transmaroni no Azure.',
    access_denied: 'O login foi cancelado ou não foi autorizado.',
    sem_permissao: 'Sua conta entrou, mas não tem acesso à pasta de dados do mapa no SharePoint. Fale com a Manutenção Corporativa.',
    sem_dados: 'Os dados do mapa ainda não foram publicados na pasta do SharePoint.',
    rede: 'Não foi possível baixar os dados do mapa (sem internet?). Tente de novo.'
  };

  // ---------- tela de acesso, por cima do mapa, até o login e os dados chegarem ----------
  const estilo = document.createElement('style');
  estilo.textContent = `
    #acesso { position: fixed; inset: 0; z-index: 3000; display: flex; align-items: center; justify-content: center; padding: 16px;
      background: rgba(242, 244, 246, .94); }
    #acesso .cartao { width: 100%; max-width: 420px; background: var(--canvas); border-radius: 16px; box-shadow: var(--sh2); padding: 28px 24px; }
    #acesso h2 { font: 700 20px/28px var(--font); margin: 0; }
    #acesso .sub { margin: 4px 0 20px; }
    #acesso .btn { width: 100%; }
    #acesso .aviso { background: #fdecec; color: #b42318; border-radius: 12px; padding: 12px 14px; margin-bottom: 16px; font-size: 14px; line-height: 20px; }
    #acesso .m { color: var(--body); font-size: 13px; line-height: 19px; margin: 14px 0 0; }
    header .conta { margin-top: 4px; }
    header .conta a { color: var(--accent-ink); }`;
  document.head.appendChild(estilo);
  const tela = document.createElement('div');
  tela.id = 'acesso';
  tela.innerHTML = '<div class="cartao"><h2>Mapa de Fornecedores</h2><p class="sub">Acesso com a conta Microsoft da Transmaroni</p>' +
    '<div id="acessoMsg"><p class="m">Verificando acesso…</p></div></div>';
  document.body.appendChild(tela);
  const mostrar = html => { document.getElementById('acessoMsg').innerHTML = html; };

  function telaEntrar(aviso) {
    mostrar((aviso ? `<div class="aviso">${esc(aviso)}</div>` : '') +
      '<button class="btn btn-primary" id="entrar">Entrar com a conta Microsoft</button>' +
      '<p class="m">Use o mesmo e-mail e senha do Outlook. A senha é digitada na página da Microsoft: este site não a vê nem guarda.</p>');
    document.getElementById('entrar').onclick = () => pca.loginRedirect({ scopes: ESCOPOS });
  }
  function telaErro(texto) {
    mostrar(`<div class="aviso">${esc(texto)}</div><button class="btn btn-primary" id="denovo">Tentar de novo</button>`);
    document.getElementById('denovo').onclick = () => location.reload();
  }

  // ---------- dados no SharePoint (Microsoft Graph) ----------
  async function graph(tk, caminho) {
    const r = await fetch('https://graph.microsoft.com/v1.0' + caminho, { headers: { Authorization: 'Bearer ' + tk } });
    if (!r.ok) throw Object.assign(new Error('Graph ' + r.status), { status: r.status });
    return r.json();
  }
  // endereço de download do arquivo de dados. A biblioteca (drive) é procurada uma vez e fica guardada neste navegador;
  // o arquivo é lido pelo caminho dentro dela, com a permissão da própria pessoa.
  async function enderecoDados(tk) {
    let drive = null;
    try { drive = localStorage.getItem('mapaForn.drive'); } catch (e) { }
    if (!drive) {
      const site = await graph(tk, `/sites/${CFG.site}?$select=id`);
      drive = (await graph(tk, `/sites/${site.id}/lists/${CFG.lista}/drive?$select=id`)).id;
      try { localStorage.setItem('mapaForn.drive', drive); } catch (e) { }
    }
    const caminho = [CFG.pasta, CFG.arquivo].map(encodeURIComponent).join('/');
    try { return (await graph(tk, `/drives/${drive}/root:/${caminho}`))['@microsoft.graph.downloadUrl']; }
    catch (e) {
      try { localStorage.removeItem('mapaForn.drive'); } catch (x) { }   // biblioteca mudou: procura de novo na próxima vez
      throw e;
    }
  }

  async function iniciar() {
    try {
      await pca.initialize();
      const r = await pca.handleRedirectPromise();
      const conta = r?.account || pca.getActiveAccount() || pca.getAllAccounts()[0];
      if (!conta) return telaEntrar();
      pca.setActiveAccount(conta);
      mostrar('<p class="m">Carregando os dados do mapa…</p>');
      let tk;
      try { tk = (await pca.acquireTokenSilent({ scopes: ESCOPOS, account: conta })).accessToken; }
      catch (e) {
        if (e instanceof msal.InteractionRequiredAuthError) return pca.acquireTokenRedirect({ scopes: ESCOPOS, account: conta });
        throw e;
      }
      let dados;
      try {
        const resp = await fetch(await enderecoDados(tk));
        if (!resp.ok) throw Object.assign(new Error('download ' + resp.status), { status: resp.status });
        dados = await resp.json();
      } catch (e) {
        if (e.status === 401 || e.status === 403) return telaErro(ERROS.sem_permissao);
        if (e.status === 404) return telaErro(ERROS.sem_dados);
        if (e instanceof TypeError) return telaErro(ERROS.rede);   // sem conexão
        throw e;
      }
      iniciarMapa(dados);
      tela.remove();
      // quem está conectado; "Sair" desconecta só deste site, não do Outlook e dos outros aplicativos da Microsoft
      const linha = document.createElement('div');
      linha.className = 'sub conta';
      linha.innerHTML = `${esc(conta.name || conta.username)} · <a href="#" id="sair">Sair</a>`;
      document.querySelector('#side header').appendChild(linha);
      document.getElementById('sair').onclick = async ev => { ev.preventDefault(); await pca.clearCache({ account: conta }); location.reload(); };
    } catch (e) {
      telaEntrar((ERROS[e.errorCode] || 'Não foi possível entrar.') + (e.errorCode ? ` (código: ${e.errorCode})` : ''));
    }
  }
  iniciar();
})();
