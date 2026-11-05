// Versão publicada do Mapa de Fornecedores (GitHub Pages): login com a conta Microsoft da Transmaroni (Entra ID, locatário
// único) e leitura dos dados na pasta do SharePoint, com a permissão da própria pessoa. Nenhum dado fica neste site.
// O mapa local, aberto do disco, já vem com os dados embutidos e não usa este arquivo.
(() => {
  // trava contra moldura (decisão do gestor em 30/10/2026): dentro de um iframe de outra página o mapa não entra nem baixa
  // os dados, para ninguém cobrir a página com outra e induzir cliques ("clickjacking"). O GitHub Pages não deixa
  // configurar o cabeçalho X-Frame-Options, e o frame-ancestors do CSP não vale em <meta>.
  if (window.top !== window.self) {
    const a = document.createElement('a');
    a.href = location.href; a.target = '_top'; a.textContent = 'Abrir o Mapa de Fornecedores';
    document.body.replaceChildren(a);
    return;
  }
  const CFG = {
    clientId: 'be92c8b6-d0fc-40e3-8e5c-f75de940e41b',   // registro do aplicativo no Entra ID (IDs públicos, não são senha)
    tenant: '8babaf2b-248f-4651-98d6-1e5312fbe00a',     // Transmaroni
    // biblioteca e pasta dos dados no SharePoint da empresa (identificadores do Graph; sozinhos não dão acesso a nada).
    // A pasta tem leitura para "Todos, exceto usuários externos"; se for trocada, atualizar aqui.
    drive: 'b!Aze8ECmtmkSONo7jEHZ6-GtMAmSijYJMvjNi4-9e53LMVxOpPZr4TYlGlmf17edT',
    pasta: '014Y2MABWQGI36RFAB3FC33ISBX75WNI4E',   // identificador da pasta: continua valendo se ela for renomeada ou movida
    arquivo: 'mapa_dados.json',                     // gravado pelo passo 10
    arquivoGz: 'mapa_dados.json.gz'                 // o mesmo, comprimido (gzip): uns 6 vezes menor para baixar
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
  // o código técnico (ex.: "404 itemNotFound") permite diagnosticar o problema por um print da tela
  function telaErro(texto, detalhe) {
    mostrar(`<div class="aviso">${esc(texto)}</div><button class="btn btn-primary" id="denovo">Tentar de novo</button>` +
      (detalhe ? `<p class="m">Código técnico: ${esc(detalhe)}</p>` : ''));
    document.getElementById('denovo').onclick = () => location.reload();
  }

  // ---------- dados no SharePoint (Microsoft Graph) ----------
  async function graph(tk, caminho) {
    const r = await fetch('https://graph.microsoft.com/v1.0' + caminho, { headers: { Authorization: 'Bearer ' + tk } });
    if (!r.ok) {
      const j = await r.json().catch(() => null);
      throw Object.assign(new Error('Graph ' + r.status), { status: r.status, codigo: j?.error?.code });
    }
    return r.json();
  }
  // endereço de download do arquivo de dados: leitura direta pela pasta (identificador) e pelo nome do arquivo, com a
  // permissão da própria pessoa (Files.Read.All delegada). Funciona também para quem tem acesso só à pasta.
  async function itemDados(tk, nome) {
    try { return await graph(tk, `/drives/${CFG.drive}/items/${CFG.pasta}:/${encodeURIComponent(nome)}`); }
    catch (e) { if (e.status === 404) return null; throw e; }
  }
  async function baixar(item) {
    const r = await fetch(item['@microsoft.graph.downloadUrl']);
    if (!r.ok) throw Object.assign(new Error('download ' + r.status), { status: r.status });
    return r;
  }
  // Dados comprimidos (decisão do gestor em 03/11/2026): com os fornecedores de mercado o arquivo passa de 12 MB; a cópia
  // em gzip tem uns 2 MB e é aberta aqui mesmo (DecompressionStream). Vale só se não for mais velha que o .json; se faltar
  // ou der qualquer erro, o mapa usa o .json, como antes.
  async function lerDados(tk) {
    const [js, gz] = await Promise.all([itemDados(tk, CFG.arquivo), itemDados(tk, CFG.arquivoGz)]);
    if (gz && typeof DecompressionStream === 'function' && (!js || gz.lastModifiedDateTime >= js.lastModifiedDateTime)) {
      try {
        const b = new Uint8Array(await (await baixar(gz)).arrayBuffer());
        // gzip começa com 1f 8b; se o caminho já tiver descomprimido, o conteúdo é o próprio JSON
        const txt = b[0] === 0x1f && b[1] === 0x8b
          ? await new Response(new Blob([b]).stream().pipeThrough(new DecompressionStream('gzip'))).text()
          : new TextDecoder().decode(b);
        return JSON.parse(txt);
      } catch (e) {
        if (!js) throw e;
        console.warn('Dados comprimidos indisponíveis; usando o arquivo normal.', e);
      }
    }
    if (!js) throw Object.assign(new Error('Graph 404'), { status: 404, codigo: 'itemNotFound' });
    return (await baixar(js)).json();
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
        dados = await lerDados(tk);
      } catch (e) {
        const detalhe = [e.status, e.codigo].filter(Boolean).join(' ') || e.name;
        if (e.status === 401 || e.status === 403) return telaErro(ERROS.sem_permissao, detalhe);
        if (e.status === 404) return telaErro(ERROS.sem_dados, detalhe);
        if (e instanceof TypeError) return telaErro(ERROS.rede, detalhe);   // sem conexão
        return telaErro('Não foi possível carregar os dados do mapa.', detalhe);
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
