// Ativa ícones Lucide
if (typeof lucide !== 'undefined') {
  lucide.createIcons();
}

let dataStore = {
  reportSection: {},
  analiseCarteira: {}
};

// ---------------------------------------------------------------------
// Dados mensais FIXOS (Jan a Ago/2026): Vendas por Cliente, por Segmento e
// Top Produtos, mês a mês. Isso resolve o problema de a planilha ao vivo
// (Export_ReportSection / Export_...) só trazer o acumulado atual, sem
// detalhar mês a mês por cliente/segmento/produto.
//
// Você mantém essas 4 planilhas atualizadas direto no Google Drive, então
// o dashboard busca elas ao vivo (via o link de exportação do Drive) toda
// vez que você carrega uma das planilhas "Export" — exatamente como pedido:
// elas só valem como COMPLEMENTO de quando o Export já foi carregado.
//
// Se por qualquer motivo o navegador não conseguir buscar do Drive (link
// mudou de permissão, sem internet, bloqueio de CORS etc.), cai para o
// último snapshot salvo no localStorage e, se não houver nenhum, para o
// arquivo historico_fixo.json que já vai junto com o site (uma foto de
// Jan-Ago/2026 tirada em 19/09/2026) — assim o filtro de mês nunca fica
// totalmente quebrado, só desatualizado.
// ---------------------------------------------------------------------
let dadosFixosMensais = null;
let sincronizacaoDriveEmAndamento = false;

const DRIVE_SHEET_IDS = {
  cliente: '1AP_60koNw2moYQfoJbiXwepVl0EGgcb0',   // Vendas (R$) por Cliente MES A MES 2026
  segmento: '1z2Xt-nouxE5JapG-pGTZTkjjKca1mW7k',   // Vendas em reais por Segmento MES A MES 2026
  produto: '1lHcnnMJHKzlBt7El5GQ-qktMx897o0pR'     // 20 produtos mais vendidos mês a mês 2026
};
const MESES_FIXOS = [1, 2, 3, 4, 5, 6, 7, 8];

function encontrarAbaMes(workbook, mes) {
  const alvo = `MES ${mes}`;
  if (workbook.SheetNames.includes(alvo)) return alvo;
  return workbook.SheetNames.find(n => n.trim() === alvo) || null;
}

async function buscarWorkbookDrive(fileId) {
  // Exporta a planilha inteira (todas as abas) como .xlsx direto do Drive.
  // Só funciona se o arquivo estiver como "Qualquer pessoa com o link".
  const url = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`HTTP ${res.status} ao buscar planilha do Drive`);
  const buf = await res.arrayBuffer();
  return XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: true });
}

function linhaVazia(row) {
  return !row || row.every(c => c === null || c === undefined || c === '');
}

function extrairPorClienteDoWorkbook(workbook) {
  const resultado = {};
  MESES_FIXOS.forEach(mes => {
    const aba = encontrarAbaMes(workbook, mes);
    if (!aba) return;
    const linhasRaw = XLSX.utils.sheet_to_json(workbook.Sheets[aba], { header: 1, defval: null });
    let classeAtual = null;
    const linhas = [];
    for (let i = 1; i < linhasRaw.length; i++) {
      const row = linhasRaw[i];
      if (linhaVazia(row)) continue;
      const [classe, cod, nome] = row;
      const valor = row[4];
      const pct = row[5];
      if (classe) classeAtual = classe;
      if (classeAtual === 'Total') continue; // linha de total geral do mês
      if (cod === null || cod === undefined || cod === '' || !nome) continue; // subtotal de classe
      linhas.push({
        Classe: classeAtual,
        Cliente_Pai: `${Math.trunc(Number(cod))}-${String(nome).trim()}`,
        'Valor de venda (R$)': Number(valor) || 0,
        '% do Total': Number(pct) || 0
      });
    }
    resultado[String(mes)] = linhas;
  });
  return resultado;
}

function extrairPorSegmentoDoWorkbook(workbook) {
  const resultado = {};
  MESES_FIXOS.forEach(mes => {
    const aba = encontrarAbaMes(workbook, mes);
    if (!aba) return;
    const linhasRaw = XLSX.utils.sheet_to_json(workbook.Sheets[aba], { header: 1, defval: null });
    const linhas = [];
    for (let i = 1; i < linhasRaw.length; i++) {
      const row = linhasRaw[i];
      if (linhaVazia(row)) continue;
      const [sep, , , valor] = row;
      if (sep !== null && String(sep).trim().toLowerCase() === 'total') continue;
      linhas.push({
        Separador: sep ? String(sep).trim() : 'Outros',
        After_Tax_Amount: Number(valor) || 0
      });
    }
    resultado[String(mes)] = linhas;
  });
  return resultado;
}

function extrairPorProdutoDoWorkbook(workbook) {
  const resultado = {};
  MESES_FIXOS.forEach(mes => {
    const aba = encontrarAbaMes(workbook, mes);
    if (!aba) return;
    const linhasRaw = XLSX.utils.sheet_to_json(workbook.Sheets[aba], { header: 1, defval: null });
    const linhas = [];
    for (let i = 1; i < linhasRaw.length; i++) {
      const row = linhasRaw[i];
      if (linhaVazia(row)) continue;
      const [produto, , valor] = row;
      if (produto === null || produto === undefined || produto === '') continue;
      if (String(produto).trim().toLowerCase() === 'total') continue;
      linhas.push({ Produto: String(produto).trim(), 'Valor de Venda': Number(valor) || 0 });
    }
    linhas.sort((a, b) => b['Valor de Venda'] - a['Valor de Venda']);
    resultado[String(mes)] = linhas.slice(0, 20);
  });
  return resultado;
}

// Busca as 3 planilhas no Drive, monta o objeto no mesmo formato de sempre
// e atualiza dadosFixosMensais + o cache local. Chamada automaticamente
// sempre que uma planilha "Export" é carregada (ver readExcelFile).
async function sincronizarPlanilhasDoDrive() {
  if (sincronizacaoDriveEmAndamento) return;
  sincronizacaoDriveEmAndamento = true;
  try {
    const [wbCliente, wbSegmento, wbProduto] = await Promise.all([
      buscarWorkbookDrive(DRIVE_SHEET_IDS.cliente),
      buscarWorkbookDrive(DRIVE_SHEET_IDS.segmento),
      buscarWorkbookDrive(DRIVE_SHEET_IDS.produto)
    ]);

    const novosDados = {
      meses_disponiveis: MESES_FIXOS,
      ano: 2026,
      porCliente: extrairPorClienteDoWorkbook(wbCliente),
      porSegmento: extrairPorSegmentoDoWorkbook(wbSegmento),
      porProduto: extrairPorProdutoDoWorkbook(wbProduto)
    };

    dadosFixosMensais = novosDados;
    try {
      localStorage.setItem('rca61_mensalDrive', JSON.stringify({
        dados: novosDados,
        atualizadoEm: new Date().toISOString()
      }));
    } catch (e) { /* ignora falha de storage */ }

    console.info('[RCA 61] Planilhas mês a mês sincronizadas do Google Drive com sucesso.');
    popularSelectClientes();
    renderDashboard();
  } catch (err) {
    // Provável bloqueio de CORS, arquivo sem permissão pública, ou sem
    // internet. Mantém o que já estava carregado (cache local ou o
    // historico_fixo.json) e avisa só no console, sem travar a página.
    console.warn('[RCA 61] Não consegui sincronizar as planilhas mês a mês do Drive. Usando o último snapshot disponível.', err);
  } finally {
    sincronizacaoDriveEmAndamento = false;
  }
}

// Carrega, na ordem: 1) cache salvo no navegador (instantâneo) e, se não
// houver, 2) o snapshot fixo que já vai junto com o site.
function carregarDadosFixosMensais() {
  try {
    const cache = localStorage.getItem('rca61_mensalDrive');
    if (cache) {
      const parsed = JSON.parse(cache);
      if (parsed && parsed.dados) {
        dadosFixosMensais = parsed.dados;
        return Promise.resolve();
      }
    }
  } catch (e) { /* ignora cache corrompido */ }

  return fetch('historico_fixo.json')
    .then(r => { if (!r.ok) throw new Error('historico_fixo.json não encontrado'); return r.json(); })
    .then(data => { dadosFixosMensais = data; })
    .catch(err => {
      console.warn('Não foi possível carregar historico_fixo.json (dados mensais fixos).', err);
      dadosFixosMensais = null;
    });
}

// Devolve as linhas fixas de um mês (1-8) para 'porCliente' | 'porSegmento' | 'porProduto',
// ou null se aquele mês não estiver disponível.
function obterLinhasFixasMes(tipo, mes) {
  if (!dadosFixosMensais || !dadosFixosMensais[tipo]) return null;
  const linhas = dadosFixosMensais[tipo][String(mes)];
  return linhas && linhas.length > 0 ? linhas : null;
}

// Os dados mensais fixos (Jan-Ago) só devem "valer" depois que pelo menos
// uma das planilhas Export (Relatório Geral / Análise de Carteira) tiver
// sido carregada nesta sessão — como pedido: elas são complemento, não
// substituem o upload.
function usarDadosMensaisFixos() {
  return algumaPlanilhaCarregada();
}

let charts = {};

const MAPA_MESES_EN = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

// Elementos do DOM
const fileInput1 = document.getElementById('fileInput1');
const fileInput2 = document.getElementById('fileInput2');
const dropZone1 = document.getElementById('dropZone1');
const dropZone2 = document.getElementById('dropZone2');
const labelFile1 = document.getElementById('labelFile1');
const labelFile2 = document.getElementById('labelFile2');
const statusBadge = document.getElementById('statusBadge');
const badgeText = document.getElementById('badgeText');

const selectAno = document.getElementById('selectAno');
const selectMes = document.getElementById('selectMes');
const selectCliente = document.getElementById('selectCliente');
const searchInput = document.getElementById('searchClientInput');

// Listeners de Upload
if (fileInput1) fileInput1.addEventListener('change', (e) => e.target.files.length > 0 && readExcelFile(e.target.files[0], 1));
if (fileInput2) fileInput2.addEventListener('change', (e) => e.target.files.length > 0 && readExcelFile(e.target.files[0], 2));

// Guarda as planilhas já lidas na sessão do navegador (sessionStorage), para
// não precisar carregar de novo ao ir para a Base de Clientes e voltar.
// Some sozinho quando a guia/aba é fechada (sessionStorage é por aba).
function salvarSessaoAtual() {
  try {
    sessionStorage.setItem('rca61_sessaoPlanilhas', JSON.stringify({
      dataStore: dataStore,
      nomeArquivo1: labelFile1 ? labelFile1.textContent : '',
      nomeArquivo2: labelFile2 ? labelFile2.textContent : '',
      carregado1: dropZone1 ? dropZone1.classList.contains('loaded') : false,
      carregado2: dropZone2 ? dropZone2.classList.contains('loaded') : false
    }));
  } catch (err) {
    console.error('Não foi possível salvar a sessão das planilhas:', err);
  }
}

function restaurarSessaoAtual() {
  try {
    const raw = sessionStorage.getItem('rca61_sessaoPlanilhas');
    if (!raw) return false;
    const salvo = JSON.parse(raw);
    if (!salvo || !salvo.dataStore) return false;

    dataStore = salvo.dataStore;

    if (salvo.carregado1 && dropZone1) dropZone1.classList.add('loaded');
    if (salvo.carregado2 && dropZone2) dropZone2.classList.add('loaded');
    if (salvo.nomeArquivo1 && labelFile1) labelFile1.textContent = salvo.nomeArquivo1;
    if (salvo.nomeArquivo2 && labelFile2) labelFile2.textContent = salvo.nomeArquivo2;

    const temDados = (dataStore.reportSection && Object.keys(dataStore.reportSection).length > 0) ||
                      (dataStore.analiseCarteira && Object.keys(dataStore.analiseCarteira).length > 0);

    if (temDados) {
      if (statusBadge) statusBadge.classList.add('active');
      if (badgeText) badgeText.textContent = "Dados Sincronizados";
      popularSelectClientes();
      // Espera o navegador terminar de montar o layout da página antes de
      // desenhar os gráficos — evita que eles fiquem em branco quando o
      // restauro acontece muito cedo (canvas ainda sem tamanho definido).
      requestAnimationFrame(() => renderDashboard());
      // Já tinha Export carregado nesta aba: aproveita e busca a versão
      // mais recente das 4 planilhas mensais no Drive também.
      sincronizarPlanilhasDoDrive();
    }
    return temDados;
  } catch (err) {
    console.error('Não foi possível restaurar a sessão das planilhas:', err);
    return false;
  }
}

// Ao abrir/voltar para esta página, primeiro carrega os dados mensais fixos
// (Jan-Ago/2026) — eles já deixam o dashboard e o filtro de mês funcionando
// mesmo sem nenhum upload — e em seguida tenta restaurar as planilhas ao
// vivo já carregadas nesta mesma aba.
carregarDadosFixosMensais().then(() => {
  popularSelectClientes();
  renderDashboard();
  restaurarSessaoAtual();
});

// Alguns navegadores restauram a página do "cache de navegação" (bfcache) ao
// clicar em voltar, sem executar o script de novo. Nesse caso os dados em
// memória continuam certos, mas os gráficos (canvas) podem ficar em branco.
// Isso força o redesenho sempre que isso acontecer.
window.addEventListener('pageshow', (event) => {
  if (event.persisted) {
    requestAnimationFrame(() => renderDashboard());
  }
});

function readExcelFile(file, fileNum) {
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array', cellDates: true });
      
      let parsedSheets = {};
      workbook.SheetNames.forEach(sheetName => {
        parsedSheets[sheetName.trim()] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
      });

      if (fileNum === 1) {
        dataStore.reportSection = parsedSheets;
        if (dropZone1) dropZone1.classList.add('loaded');
        if (labelFile1) labelFile1.textContent = `✔ ${file.name}`;
      } else {
        dataStore.analiseCarteira = parsedSheets;
        if (dropZone2) dropZone2.classList.add('loaded');
        if (labelFile2) labelFile2.textContent = `✔ ${file.name}`;
      }

      if (statusBadge) statusBadge.classList.add('active');
      if (badgeText) badgeText.textContent = "Dados Sincronizados";

      popularSelectClientes();
      renderDashboard();
      atualizarDadosVivosLocalStorage();
      salvarSessaoAtual();

      // A partir do momento em que uma planilha Export é carregada, os
      // dados mensais fixos (Jan-Ago) "entram em vigor" — e busca a versão
      // mais recente delas direto do Google Drive em segundo plano (não
      // trava a tela; quando terminar, o dashboard se atualiza sozinho).
      sincronizarPlanilhasDoDrive();
    } catch (err) {
      console.error("Erro ao ler planilha:", err);
    }
  };
  reader.readAsArrayBuffer(file);
}

// Salva um resumo (faturamento e inatividade por cliente) no localStorage
// para que a página "Base de Clientes" (clientes.html) possa exibir essas
// informações somente quando as planilhas já tiverem sido carregadas aqui.
// Isso não altera nada na interface do index.html.
function normalizarNome(str) {
  return String(str || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/[^A-Z0-9 ]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function atualizarDadosVivosLocalStorage() {
  try {
    const mapa = {};

    const sheetVendas = getSheet(dataStore.reportSection, ['Vendas (R$) por Cliente', 'Cliente']);
    sheetVendas.forEach(r => {
      const nome = String(r['Cliente_Pai'] || r['Cliente'] || '').trim();
      if (!nome) return;
      const chave = normalizarNome(nome);
      if (!chave) return;
      if (!mapa[chave]) mapa[chave] = { nome };
      mapa[chave].faturamento = parseCurrency(r['Valor de venda (R$)'] || r['Valor']);
      if (r['Classe']) mapa[chave].classe = r['Classe'];
    });

    const sheetInativ = getSheet(dataStore.analiseCarteira, ['#Dias até primeira fatura', 'Clientes com ultima fatura', 'Ultima Fatura']);
    sheetInativ.forEach(r => {
      const nome = String(r['Cliente_Pai'] || r['Cliente'] || '').trim();
      if (!nome) return;
      const chave = normalizarNome(nome);
      if (!chave) return;
      if (!mapa[chave]) mapa[chave] = { nome };
      mapa[chave].diasInativo = parseInt(r['#dias desde a ultima fat'] || r['Dias Inativo'] || r['Dias'] || 0, 10);
      mapa[chave].ultimaFatura = formatDate(r['Ultima fat'] || r['Última Fatura']);
      if (r['Classe']) mapa[chave].classe = r['Classe'];
    });

    localStorage.setItem('rca61_dadosVivos', JSON.stringify({
      clientes: mapa,
      atualizadoEm: new Date().toISOString()
    }));
  } catch (err) {
    console.error('Erro ao salvar dados vivos para a Base de Clientes:', err);
  }
}

// Event Listeners
if (selectAno) selectAno.addEventListener('change', renderDashboard);
if (selectMes) selectMes.addEventListener('change', renderDashboard);
if (selectCliente) selectCliente.addEventListener('change', renderDashboard);
if (searchInput) searchInput.addEventListener('input', () => {
  // A busca é compartilhada: filtra as duas tabelas (Vendas por Cliente e
  // Inatividade) ao mesmo tempo, então achar o cliente em uma já "seleciona"
  // ele na outra automaticamente.
  renderVendasClienteTable();
  renderInatividadeTable();
});

function popularSelectClientes() {
  if (!selectCliente) return;
  const sheetCliente = getSheet(dataStore.reportSection, ['Vendas (R$) por Cliente', 'Cliente']);
  const clientesUnicos = new Set();

  sheetCliente.forEach(r => {
    const nome = r['Cliente_Pai'] || r['Cliente'];
    if (nome) clientesUnicos.add(String(nome).trim());
  });

  // Também inclui os clientes que só aparecem nos meses fixos (1-8), para
  // que o filtro de cliente + mês funcione mesmo antes de qualquer upload.
  if (dadosFixosMensais && dadosFixosMensais.porCliente) {
    Object.values(dadosFixosMensais.porCliente).forEach(linhasMes => {
      linhasMes.forEach(r => {
        if (r.Cliente_Pai) clientesUnicos.add(String(r.Cliente_Pai).trim());
      });
    });
  }

  const clienteAtual = selectCliente.value;
  selectCliente.innerHTML = '<option value="ALL">Todos os Clientes</option>';
  
  Array.from(clientesUnicos).sort().forEach(cliente => {
    const opt = document.createElement('option');
    opt.value = cliente;
    opt.textContent = cliente;
    selectCliente.appendChild(opt);
  });

  selectCliente.value = clienteAtual || 'ALL';
}

function getSheet(dataObj, keywords) {
  if (!dataObj) return [];
  const sheetNames = Object.keys(dataObj);
  for (let kw of keywords) {
    const found = sheetNames.find(s => s.toLowerCase().includes(kw.toLowerCase()));
    if (found) return dataObj[found];
  }
  return [];
}

// true assim que pelo menos uma das duas planilhas foi carregada. Usado para
// só exibir os números de demonstração (placeholder) ANTES do upload —
// depois que o usuário carrega a planilha, tudo deve refletir o real,
// inclusive quando o real for zero.
function algumaPlanilhaCarregada() {
  return (dataStore.reportSection && Object.keys(dataStore.reportSection).length > 0) ||
         (dataStore.analiseCarteira && Object.keys(dataStore.analiseCarteira).length > 0);
}

// Verifica se a planilha tem uma coluna de Cliente (linha a linha), ou seja,
// se dá pra saber qual linha pertence a qual cliente.
function sheetTemColunaCliente(sheet) {
  if (!sheet || sheet.length === 0) return false;
  const primeira = sheet[0];
  return Object.prototype.hasOwnProperty.call(primeira, 'Cliente_Pai') ||
         Object.prototype.hasOwnProperty.call(primeira, 'Cliente');
}

// Filtra uma planilha pelo cliente selecionado no topo (selectCliente).
// - "ALL": devolve tudo, sem restrição.
// - Cliente específico + planilha TEM coluna de cliente: filtra normalmente.
// - Cliente específico + planilha NÃO TEM coluna de cliente (é um dado
//   agregado da carteira toda, ex: histórico mensal): não tem como saber a
//   fatia desse cliente, então devolve vazio -> o gráfico/KPI fica zerado
//   em vez de continuar mostrando o total da carteira (o que pareceria bug).
function filtrarPorCliente(sheet, clienteSel) {
  if (!clienteSel || clienteSel === 'ALL') {
    return { linhas: sheet, semDetalhePorCliente: false };
  }
  if (!sheetTemColunaCliente(sheet)) {
    return { linhas: [], semDetalhePorCliente: true };
  }
  const linhas = sheet.filter(r => {
    const nome = String(r['Cliente_Pai'] || r['Cliente'] || '').trim();
    return nome === clienteSel;
  });
  return { linhas, semDetalhePorCliente: false };
}

// Conta quantos dias úteis (seg a sex, sem considerar feriados) faltam para
// acabar o mês atual, contando o dia de hoje se ainda for dia útil.
function diasUteisRestantesNoMes(dataRef = new Date()) {
  const ano = dataRef.getFullYear();
  const mes = dataRef.getMonth();
  const ultimoDia = new Date(ano, mes + 1, 0).getDate();
  let count = 0;
  for (let dia = dataRef.getDate(); dia <= ultimoDia; dia++) {
    const diaSemana = new Date(ano, mes, dia).getDay(); // 0 = domingo, 6 = sábado
    if (diaSemana !== 0 && diaSemana !== 6) count++;
  }
  return count;
}

function renderDashboard() {
  renderYTDBanner();
  renderKPIs();
  renderChartHistorico();
  renderChartBudget();
  renderChartTipoEncomenda();
  renderChartSegmentos();
  renderChartTopProdutos();
  renderVendasClienteTable();
  renderInatividadeTable();
}

function renderYTDBanner() {
  const sheet = getSheet(dataStore.reportSection, ['Venda Mensal em Reais', 'Venda Mensal', 'Image-6', 'Geral']);
  let lytd = 0, ytd = 0;

  if (sheet.length > 0) {
    sheet.forEach(r => {
      lytd += parseCurrency(r['2025'] || r['2025 (R$)']);
      ytd += parseCurrency(r['2026'] || r['2026 (R$)']);
    });
  }

  const elLytd = document.getElementById('kpiLytd');
  const elYtd = document.getElementById('kpiYtd');
  const elVar = document.getElementById('kpiVariacao');

  const variacao = ytd - lytd;
  const carregado = algumaPlanilhaCarregada();

  // Antes de carregar qualquer planilha, mostra números de demonstração.
  // Depois de carregar, mostra o valor real mesmo que seja zero.
  if (elLytd) elLytd.textContent = formatBRL(carregado ? lytd : 6386614.16);
  if (elYtd) elYtd.textContent = formatBRL(carregado ? ytd : 6636963.60);
  if (elVar) elVar.textContent = formatBRL(carregado ? variacao : 250349.43);
}

function renderKPIs() {
  const mesSel = selectMes ? selectMes.value : 'ALL';
  const anoSel = selectAno ? selectAno.value : '2026';
  const clienteSel = selectCliente ? selectCliente.value : 'ALL';
  const carregado = algumaPlanilhaCarregada();
  const clienteEspecifico = clienteSel !== 'ALL';

  const anoNum = Number(anoSel);
  const mesNum = mesSel !== 'ALL' ? parseInt(mesSel, 10) : null;

  // ---- Faturamento da Carteira (já é por cliente, então filtra normal) ----
  // Se um mês específico (1-8) foi selecionado, usa os dados FIXOS daquele
  // mês (que têm o detalhe por cliente); senão usa o acumulado da planilha
  // ao vivo carregada no Dashboard (comportamento de antes).
  let totalFat = 0;
  const linhasFixasMes = (mesNum !== null && usarDadosMensaisFixos()) ? obterLinhasFixasMes('porCliente', mesNum) : null;
  const usandoMesFixo = linhasFixasMes !== null;
  const sheetCliente = usandoMesFixo ? linhasFixasMes : getSheet(dataStore.reportSection, ['Vendas (R$) por Cliente', 'Cliente']);

  if (clienteEspecifico) {
    sheetCliente.forEach(r => {
      const nome = String(r['Cliente_Pai'] || r['Cliente'] || '').trim();
      if (nome === clienteSel) {
        totalFat += parseCurrency(r['Valor de venda (R$)'] || r['Valor']);
      }
    });
  } else {
    sheetCliente.forEach(r => totalFat += parseCurrency(r['Valor de venda (R$)'] || r['Valor']));
    if (totalFat === 0 && !usandoMesFixo) {
      const sheetVenda = getSheet(dataStore.reportSection, ['Venda Mensal em Reais', 'Venda Mensal', 'Image-6']);
      sheetVenda.forEach(r => totalFat += parseCurrency(r[anoSel] || r[`${anoSel} (R$)`] || r['Vendas (R$)']));
    }
  }

  const elValor = document.getElementById('kpiValorMensal');
  // Mostra o valor real assim que houver alguma fonte de dado (planilha ao
  // vivo OU mês fixo selecionado); só cai no número de demonstração se não
  // tiver nenhuma das duas ainda.
  if (elValor) elValor.textContent = formatBRL((carregado || usandoMesFixo) ? totalFat : 32766640.70);

  // ---- % Budget Atingido (indicador da carteira toda; planilha não traz
  // o budget quebrado por cliente, então zera quando um cliente é selecionado) ----
  const sheetBudgetBruto = getSheet(dataStore.reportSection, ['% do Budget atingida por', '% Budget Atingida', 'Budget']);
  const { linhas: sheetBudget, semDetalhePorCliente: budgetSemDetalhe } = filtrarPorCliente(sheetBudgetBruto, clienteSel);
  let avgBudget = 0;

  if (!budgetSemDetalhe && sheetBudget.length > 0) {
    if (mesNum !== null) {
      const row = sheetBudget.find(r => Number(r['Mês'] || r['Mes']) === mesNum);
      if (row) avgBudget = parsePct(row['% do Budget'] || row['Budget']);
    } else {
      let sum = 0, count = 0;
      sheetBudget.forEach(r => {
        const val = parsePct(r['% do Budget'] || r['Budget']);
        if (val > 0) { sum += val; count++; }
      });
      avgBudget = count > 0 ? (sum / count) : 0;
    }
  }

  const elBudget = document.getElementById('kpiBudgetAtingido');
  if (elBudget) elBudget.textContent = `${avgBudget.toFixed(1)}%`;

  // ---- Positivação da Carteira ----
  // É um indicador da carteira inteira (quantos clientes compraram no
  // período). Não faz sentido "por 1 cliente só", então quando um cliente
  // específico está selecionado, o card fica zerado com um aviso — em vez
  // de continuar mostrando o número da carteira toda (que parecia bug).
  const elPos = document.getElementById('kpiPositivacao');
  const elPosSub = document.getElementById('kpiPositivacaoSub');
  const kpiCardPositivacao = elPosSub ? elPosSub.closest('.kpi-card') : null;

  if (clienteEspecifico) {
    if (elPos) { elPos.textContent = '—'; elPos.style.color = ''; }
    if (elPosSub) elPosSub.innerHTML = `Indicador de carteira — selecione "Todos os Clientes" para ver a positivação.`;
    if (kpiCardPositivacao) { kpiCardPositivacao.style.borderColor = ''; kpiCardPositivacao.style.boxShadow = ''; }
  } else {
    const sheetUltimaFatura = getSheet(dataStore.analiseCarteira, ['Ultima fatura', 'Última fatura']);
    let positivados = carregado ? 0 : 82;
    let totalCarteira = carregado ? 0 : 271;

    if (sheetUltimaFatura && sheetUltimaFatura.length > 0) {
      const row = sheetUltimaFatura[0];
      const valPos = parseCurrency(row['Qtd_Positivados'] || row['Qtd_Positivado']);
      const valCart = parseCurrency(row['Carteira']);
      if (valPos > 0) positivados = valPos;
      if (valCart > 0) totalCarteira = valCart;
    }

    // A meta é sempre 60% da carteira atual da planilha — não usa mais um
    // valor de "Meta" fixo vindo da planilha, para não desalinhar do real.
    const metaQtd = Math.round(totalCarteira * 0.60);
    const metaPct = 60;
    const realPct = totalCarteira > 0 ? (positivados / totalCarteira) * 100 : 0;
    const faltaQtd = metaQtd - positivados;
    const faltaPct = metaPct - realPct;

    // Alerta "corra atrás": liga na reta final do mês — quando restam 10 dias
    // úteis ou menos para acabar e a meta ainda não foi batida.
    const diasUteisRestantes = diasUteisRestantesNoMes();
    const abaixoDaMeta = faltaQtd > 0;
    const alertaPositivacao = abaixoDaMeta && diasUteisRestantes <= 10;

    if (elPos) {
      elPos.textContent = `${positivados} / ${totalCarteira}`;
      elPos.style.color = alertaPositivacao ? '#ef4444' : '';
    }
    if (elPosSub) {
      if (faltaQtd <= 0) {
        elPosSub.innerHTML = `Real: <strong>${realPct.toFixed(2)}%</strong> | Meta: <strong>${metaQtd} clientes (${metaPct.toFixed(1)}%)</strong> | <span style="color:#10b981;font-weight:bold;">Meta Atingida!</span>`;
      } else {
        elPosSub.innerHTML = `Real: <strong>${realPct.toFixed(2)}%</strong> | Meta: <strong>${metaQtd} clientes (${metaPct.toFixed(1)}%)</strong> | Falta: <span style="color:#ef4444;font-weight:bold;">${faltaQtd} clientes (${faltaPct.toFixed(2)}%)</span>` +
          (alertaPositivacao ? `<br><span style="display:inline-block; margin-top:6px; padding:4px 8px; border-radius:6px; background:rgba(239,68,68,0.15); color:#ef4444; font-weight:700;">🚨 Faltam ${diasUteisRestantes} dias úteis para o fim do mês — corra atrás da meta!</span>` : '');
      }
    }

    // Destaca todo o card de Positivação em vermelho quando o alerta está ativo.
    if (kpiCardPositivacao) {
      if (alertaPositivacao) {
        kpiCardPositivacao.style.borderColor = '#ef4444';
        kpiCardPositivacao.style.boxShadow = '0 0 0 1px rgba(239,68,68,0.45)';
      } else {
        kpiCardPositivacao.style.borderColor = '';
        kpiCardPositivacao.style.boxShadow = '';
      }
    }
  }

  // ---- % Encomendas Gravadas ----
  // Mesma lógica: se a planilha não detalha por cliente, zera ao filtrar.
  const sheetGravadosBruto = getSheet(dataStore.analiseCarteira, ['% de encomendas gravadas', 'Encomendas Gravadas', 'Gravado']);
  const { linhas: sheetGravados, semDetalhePorCliente: gravadosSemDetalhe } = filtrarPorCliente(sheetGravadosBruto, clienteSel);
  let pctVal = carregado ? 0 : 44.72;
  const metaGravaçãoPct = 50.0;

  if (!gravadosSemDetalhe && sheetGravados.length > 0) {
    let row = null;
    if (mesNum !== null) {
      row = sheetGravados.find(r => Number(r.Ano) === anoNum && Number(r.Mes) === mesNum && String(r.Tipo).trim().toLowerCase() === 'gravado');
    }
    if (!row) row = sheetGravados.find(r => String(r['Ano'] || r['Tipo']).toLowerCase().includes('total')) || sheetGravados[0];

    const rawVal = row ? (row['% gravação'] ?? row['% Gravado'] ?? row['Total'] ?? row['Gravado']) : undefined;
    pctVal = rawVal !== undefined ? parsePct(rawVal) : 0;
  }

  const elGrav = document.getElementById('kpiPctGravadas');
  const elGravSub = document.getElementById('kpiPctGravadasSub');
  if (elGrav) elGrav.textContent = `${pctVal.toFixed(2)}%`;

  if (elGravSub) {
    if (clienteEspecifico && gravadosSemDetalhe) {
      elGravSub.innerHTML = `Indicador de carteira — selecione "Todos os Clientes" para ver este indicador.`;
    } else {
      const diffGrav = metaGravaçãoPct - pctVal;
      if (diffGrav <= 0) {
        elGravSub.innerHTML = `Meta: ${metaGravaçãoPct}% | <span style="color:#10b981;font-weight:bold;">Meta Atingida!</span>`;
      } else {
        elGravSub.innerHTML = `Meta: ${metaGravaçãoPct}% | Falta: <span style="color:#f59e0b;font-weight:bold;">${diffGrav.toFixed(2)}%</span> p/ a meta`;
      }
    }
  }
}

// Plugin nativo para desenhar os rótulos (números) diretamente nos gráficos
const pluginValoresNativos = {
  id: 'pluginValoresNativos',
  afterDatasetsDraw(chart, args, options) {
    const { ctx } = chart;
    
    chart.data.datasets.forEach((dataset, datasetIndex) => {
      const meta = chart.getDatasetMeta(datasetIndex);
      if (!meta.hidden) {
        meta.data.forEach((element, index) => {
          const value = dataset.data[index];
          if (value === null || value === undefined || value === 0) return;

          ctx.save();
          ctx.font = 'bold 10px sans-serif';
          ctx.fillStyle = '#ffffff';
          ctx.textAlign = 'center';

          let text = '';
          if (chart.config.type === 'doughnut') {
            const sum = dataset.data.reduce((a, b) => a + b, 0);
            const pct = sum > 0 ? ((value / sum) * 100).toFixed(1) + '%' : '';
            text = `${value} (${pct})`;
            const position = element.tooltipPosition();
            ctx.fillText(text, position.x, position.y);
          } else {
            if (chart.options.scales && chart.options.scales.y && chart.options.scales.y.type === 'linear') {
              if (String(dataset.label).includes('%') || (value <= 100 && dataset.label && dataset.label.includes('Budget'))) {
                text = `${Number(value).toFixed(1)}%`;
              } else if (value >= 1000) {
                text = `R$ ${(value / 1000).toFixed(0)}k`;
              } else {
                text = `${value}`;
              }
            } else {
              text = `${value}`;
            }

            const { x, y } = element.tooltipPosition ? element.tooltipPosition() : { x: element.x, y: element.y };
            
            const textWidth = ctx.measureText(text).width;
            ctx.fillStyle = 'rgba(15, 23, 42, 0.85)';
            ctx.fillRect(x - textWidth / 2 - 4, y - 16, textWidth + 8, 14);

            ctx.fillStyle = '#ffffff';
            ctx.fillText(text, x, y - 6);
          }
          ctx.restore();
        });
      }
    });
  }
};

function renderChartHistorico() {
  const ctx = document.getElementById('chartHistoricoFaturamento');
  if (!ctx) return;

  const clienteSel = selectCliente ? selectCliente.value : 'ALL';
  const meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
  let v2026 = new Array(12).fill(null);
  let semNenhumDado = true;

  // 1) Meses 1-8: sempre vêm dos dados FIXOS (Jan-Ago/2026), que já têm o
  // detalhe por cliente — então o histórico muda de verdade ao selecionar
  // um cliente específico.
  for (let m = 1; m <= 8; m++) {
    const linhasMes = usarDadosMensaisFixos() ? obterLinhasFixasMes('porCliente', m) : null;
    if (!linhasMes) continue;
    const { linhas } = filtrarPorCliente(linhasMes, clienteSel);
    const total = linhas.reduce((s, r) => s + parseCurrency(r['Valor de venda (R$)']), 0);
    v2026[m - 1] = total;
    semNenhumDado = false;
  }

  // 2) Meses 9+ (mês corrente, ainda sem planilha fixa): usa a planilha ao
  // vivo (Image-6), que só tem o total da carteira toda — por isso só
  // preenche quando "Todos os Clientes" está selecionado.
  const sheetBruto = getSheet(dataStore.reportSection, ['Image-6', 'Venda Mensal em Reais', 'Venda Mensal']);
  const { linhas: sheetLive, semDetalhePorCliente } = filtrarPorCliente(sheetBruto, clienteSel);
  sheetLive.forEach(r => {
    const idx = Number(r['Mês'] || r['Mes']) - 1;
    const anoRow = Number(r['Ano']);
    const val = parseCurrency(r['Vendas (R$)'] || r['Vendas'] || r['Valor']);
    if (idx >= 8 && idx < 12 && anoRow === 2026 && val > 0) {
      v2026[idx] = val;
      semNenhumDado = false;
    }
  });

  destroyChart('chartHistoricoFaturamento');
  charts['chartHistoricoFaturamento'] = new Chart(ctx.getContext('2d'), {
    type: 'line',
    data: {
      labels: meses,
      datasets: [
        { label: '2026', data: v2026, borderColor: '#10b981', backgroundColor: 'rgba(16,185,129,0.1)', borderWidth: 3, fill: true, spanGaps: true }
      ]
    },
    plugins: [pluginValoresNativos],
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        title: semNenhumDado
          ? { display: true, text: 'Sem dados para este cliente', color: '#94a3b8', font: { size: 11, weight: 'normal' } }
          : { display: false }
      },
      scales: {
        x: { grid: { color: '#1f293d' }, ticks: { color: '#94a3b8' } },
        y: { beginAtZero: false, grid: { color: '#1f293d' }, ticks: { color: '#94a3b8' } }
      }
    }
  });
}

function renderChartBudget() {
  const ctx = document.getElementById('chartBudget');
  if (!ctx) return;

  const clienteSel = selectCliente ? selectCliente.value : 'ALL';
  const sheetBruto = getSheet(dataStore.reportSection, ['% do Budget atingida por', '% Budget Atingida', 'Budget']);
  const { linhas: sheet, semDetalhePorCliente } = filtrarPorCliente(sheetBruto, clienteSel);
  const labels = mesesArray();
  let dataVals = new Array(12).fill(0);

  if (sheet.length > 0) {
    sheet.forEach(r => {
      const idx = Number(r['Mês'] || r['Mes']) - 1;
      if (idx >= 0 && idx < 12) {
        dataVals[idx] = parsePct(r['% do Budget'] || r['Budget']);
      }
    });
  }

  destroyChart('chartBudget');
  charts['chartBudget'] = new Chart(ctx.getContext('2d'), {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: '% Budget Atingido',
        data: dataVals,
        backgroundColor: '#6366f1',
        borderRadius: 4
      }]
    },
    plugins: [pluginValoresNativos],
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: true, labels: { color: '#94a3b8' } },
        title: semDetalhePorCliente
          ? { display: true, text: 'Planilha não detalha o budget por cliente', color: '#94a3b8', font: { size: 11, weight: 'normal' } }
          : { display: false }
      },
      scales: {
        x: { grid: { color: '#1f293d' }, ticks: { color: '#94a3b8' } },
        y: { beginAtZero: true, grid: { color: '#1f293d' }, ticks: { color: '#94a3b8' } }
      }
    }
  });
}

function renderChartTipoEncomenda() {
  const ctx = document.getElementById('chartTipoEncomenda');
  if (!ctx) return;

  const clienteSel = selectCliente ? selectCliente.value : 'ALL';
  const carregado = algumaPlanilhaCarregada();
  const sheetBruto = getSheet(dataStore.analiseCarteira, ['Clientes recentes que já', '% de encomendas gravadas', 'Encomenda por Tipo']);
  const { linhas: sheet, semDetalhePorCliente } = filtrarPorCliente(sheetBruto, clienteSel);

  // Antes de qualquer planilha carregada, mostra números de demonstração.
  // Depois de carregada, mostra o real (que pode ser 0/0).
  let gravado = carregado ? 0 : 89;
  let normal = carregado ? 0 : 110;

  if (sheet.length > 0) {
    const rowG = sheet.find(r => String(r.Tipo).toLowerCase().includes('gravado'));
    const rowN = sheet.find(r => String(r.Tipo).toLowerCase().includes('normal'));

    if (rowG) gravado = parseCurrency(rowG['Sum of Valor'] ?? rowG['% gravação']);
    if (rowN) normal = parseCurrency(rowN['Sum of Valor'] ?? rowN['% gravação']);
  }

  destroyChart('chartTipoEncomenda');
  charts['chartTipoEncomenda'] = new Chart(ctx.getContext('2d'), {
    type: 'doughnut',
    data: {
      labels: ['Gravado', 'Normal'],
      datasets: [{
        data: [gravado, normal],
        backgroundColor: ['#10b981', '#ef4444'],
        borderWidth: 0
      }]
    },
    // O plugin nativo já calcula o percentual em cima do que existir no
    // gráfico (gravado + normal), então mesmo com pouco volume (ex: só
    // 35,16 no total) o rótulo mostra o percentual correto daquilo que tem.
    plugins: [pluginValoresNativos],
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: true, labels: { color: '#94a3b8' } },
        title: semDetalhePorCliente
          ? { display: true, text: 'Planilha não detalha o tipo de encomenda por cliente', color: '#94a3b8', font: { size: 11, weight: 'normal' } }
          : { display: false }
      }
    }
  });
}

function renderChartSegmentos() {
  const ctx = document.getElementById('chartSegmentos');
  if (!ctx) return;

  const clienteSel = selectCliente ? selectCliente.value : 'ALL';
  const mesSel = selectMes ? selectMes.value : 'ALL';
  const mesNum = mesSel !== 'ALL' ? parseInt(mesSel, 10) : null;

  const linhasFixas = (mesNum !== null && usarDadosMensaisFixos()) ? obterLinhasFixasMes('porSegmento', mesNum) : null;
  const sheetBruto = linhasFixas || getSheet(dataStore.reportSection, ['Venda mensal em reais da', 'Separador Segmento', 'Segmento']);
  const { linhas: sheetFiltrado, semDetalhePorCliente } = filtrarPorCliente(sheetBruto, clienteSel);
  const sheet = [...sheetFiltrado].sort((a, b) => parseCurrency(b['After_Tax_Amount']) - parseCurrency(a['After_Tax_Amount']));
  const labels = sheet.map(r => r['Separador'] || r['Segmento'] || 'Outros').slice(0, 6);
  const dataVals = sheet.map(r => parseCurrency(r['After_Tax_Amount'] || r['Valor'])).slice(0, 6);

  destroyChart('chartSegmentos');
  charts['chartSegmentos'] = new Chart(ctx.getContext('2d'), {
    type: 'bar',
    data: {
      labels: labels.length > 0 ? labels : ['Sem Dados'],
      datasets: [{
        label: 'Vendas (R$)',
        data: dataVals.length > 0 ? dataVals : [0],
        backgroundColor: '#10b981',
        borderRadius: 4
      }]
    },
    plugins: [pluginValoresNativos],
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: true, labels: { color: '#94a3b8' } },
        title: semDetalhePorCliente
          ? { display: true, text: 'Planilha não detalha segmentos por cliente', color: '#94a3b8', font: { size: 11, weight: 'normal' } }
          : { display: false }
      },
      scales: {
        x: { grid: { color: '#1f293d' }, ticks: { color: '#94a3b8' } },
        y: { beginAtZero: true, grid: { color: '#1f293d' }, ticks: { color: '#94a3b8' } }
      }
    }
  });
}

function renderChartTopProdutos() {
  const ctx = document.getElementById('chartTopProdutos');
  if (!ctx) return;

  const clienteSel = selectCliente ? selectCliente.value : 'ALL';
  const mesSel = selectMes ? selectMes.value : 'ALL';
  const mesNum = mesSel !== 'ALL' ? parseInt(mesSel, 10) : null;

  const linhasFixas = (mesNum !== null && usarDadosMensaisFixos()) ? obterLinhasFixasMes('porProduto', mesNum) : null;
  const sheetCompleto = linhasFixas || getSheet(dataStore.reportSection, ['Image-7', 'Top 20 Produtos Mais Vendidos', 'Top 20']);
  const { linhas: sheetFiltrado, semDetalhePorCliente } = filtrarPorCliente(sheetCompleto, clienteSel);
  const sheet = [...sheetFiltrado]
    .sort((a, b) => parseCurrency(b['Valor de Venda'] || b['Valor de Venda (R$)']) - parseCurrency(a['Valor de Venda'] || a['Valor de Venda (R$)']))
    .slice(0, 5);
  const labels = sheet.map(r => String(r['Produto'] || r['Cod'] || ''));
  const dataVals = sheet.map(r => parseCurrency(r['Valor de Venda'] || r['Valor de Venda (R$)']));

  destroyChart('chartTopProdutos');
  charts['chartTopProdutos'] = new Chart(ctx.getContext('2d'), {
    type: 'bar',
    data: {
      labels: labels.length > 0 ? labels : ['Sem Dados'],
      datasets: [{
        label: 'Valor de Venda (R$)',
        data: dataVals.length > 0 ? dataVals : [0],
        backgroundColor: '#3b82f6',
        borderRadius: 4
      }]
    },
    plugins: [pluginValoresNativos],
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: true, labels: { color: '#94a3b8' } },
        title: semDetalhePorCliente
          ? { display: true, text: 'Planilha não detalha produtos por cliente', color: '#94a3b8', font: { size: 11, weight: 'normal' } }
          : { display: false }
      },
      scales: {
        x: { grid: { color: '#1f293d' }, ticks: { color: '#94a3b8' } },
        y: { beginAtZero: true, grid: { color: '#1f293d' }, ticks: { color: '#94a3b8' } }
      }
    }
  });
}

// Formata a célula "% do Total" da tabela: aceita tanto string pronta
// ("15.29%", vinda da planilha ao vivo) quanto número puro (0.1529,
// vindo dos dados fixos mensais).
function formatPctCell(val) {
  if (val === undefined || val === null || val === '') return '-';
  if (typeof val === 'string' && val.includes('%')) return val;
  return `${parsePct(val).toFixed(2)}%`;
}

function renderVendasClienteTable() {
  const tbody = document.getElementById('tbVendasCliente');
  if (!tbody) return;

  const mesSel = selectMes ? selectMes.value : 'ALL';
  const mesNum = mesSel !== 'ALL' ? parseInt(mesSel, 10) : null;
  const linhasFixas = (mesNum !== null && usarDadosMensaisFixos()) ? obterLinhasFixasMes('porCliente', mesNum) : null;
  const sheet = linhasFixas || getSheet(dataStore.reportSection, ['Vendas (R$) por Cliente', 'Cliente']);
  const query = searchInput ? searchInput.value.toLowerCase().trim() : '';
  const clienteSel = selectCliente ? selectCliente.value : 'ALL';

  tbody.innerHTML = '';
  const filtered = sheet.filter(r => {
    const nome = String(r['Cliente_Pai'] || r['Cliente'] || '').trim();
    const matchBusca = nome.toLowerCase().includes(query);
    const matchFiltroTopo = clienteSel === 'ALL' || nome === clienteSel;
    return matchBusca && matchFiltroTopo;
  }).sort((a, b) => parseCurrency(b['Valor de venda (R$)'] || b['Valor']) - parseCurrency(a['Valor de venda (R$)'] || a['Valor']));

  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4" class="empty-row">Nenhum registro localizado.</td></tr>';
    return;
  }

  filtered.forEach(r => {
    const clientName = r['Cliente_Pai'] || r['Cliente'] || '-';
    const valorVenda = parseCurrency(r['Valor de venda (R$)'] || r['Valor']);
    const pctTotal = formatPctCell(r['% do Total']);
    const classe = r['Classe'] || 'Fiel';

    const tr = document.createElement('tr');
    tr.className = 'clickable-row';
    tr.style.cursor = 'pointer';
    tr.innerHTML = `
      <td>${classe}</td>
      <td><strong>${clientName}</strong></td>
      <td>${formatBRL(valorVenda)}</td>
      <td>${pctTotal}</td>
    `;

    tr.addEventListener('click', () => abrirModalCliente({
      nome: clientName,
      classe: classe,
      faturamento: formatBRL(valorVenda),
      participacao: pctTotal
    }));

    tbody.appendChild(tr);
  });
}

function renderInatividadeTable() {
  const tbody = document.getElementById('tbInatividade');
  if (!tbody) return;

  const sheet = getSheet(dataStore.analiseCarteira, ['#Dias até primeira fatura', 'Clientes com ultima fatura', 'Ultima Fatura']);
  const clienteSel = selectCliente ? selectCliente.value : 'ALL';
  const query = searchInput ? searchInput.value.toLowerCase().trim() : '';
  tbody.innerHTML = '';

  const filtered = sheet.filter(r => {
    const nome = String(r['Cliente_Pai'] || r['Cliente'] || '').trim();
    const matchBusca = nome.toLowerCase().includes(query);
    const matchFiltroTopo = clienteSel === 'ALL' || nome === clienteSel;
    return matchBusca && matchFiltroTopo;
  });

  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4" class="empty-row">Nenhum registro localizado.</td></tr>';
    return;
  }

  filtered.forEach(r => {
    const clientName = r['Cliente_Pai'] || r['Cliente'] || '-';
    const dias = parseInt(r['#dias desde a ultima fat'] || r['Dias Inativo'] || r['Dias'] || 0, 10);
    const dataFat = formatDate(r['Ultima fat'] || r['Última Fatura']);
    const classe = r['Classe'] || 'Pontual';

    const tr = document.createElement('tr');
    tr.className = 'clickable-row';
    tr.style.cursor = 'pointer';
    tr.innerHTML = `
      <td>${classe}</td>
      <td><strong>${clientName}</strong></td>
      <td>${dataFat}</td>
      <td><span style="color: ${dias >= 60 ? '#ef4444' : '#10b981'}; font-weight: 700;">${dias} dias</span></td>
    `;

    tr.addEventListener('click', () => abrirModalCliente({
      nome: clientName,
      classe: classe,
      ultimaFatura: dataFat,
      inatividade: `${dias} dias`
    }));

    tbody.appendChild(tr);
  });
}

function abrirModalCliente(dados) {
  let modalContainer = document.getElementById('customClientModal');

  if (!modalContainer) {
    modalContainer = document.createElement('div');
    modalContainer.id = 'customClientModal';
    modalContainer.style.position = 'fixed';
    modalContainer.style.top = '0';
    modalContainer.style.left = '0';
    modalContainer.style.width = '100vw';
    modalContainer.style.height = '100vh';
    modalContainer.style.backgroundColor = 'rgba(11, 15, 25, 0.8)';
    modalContainer.style.backdropFilter = 'blur(6px)';
    modalContainer.style.zIndex = '99999';
    modalContainer.style.display = 'flex';
    modalContainer.style.alignItems = 'center';
    modalContainer.style.justifyContent = 'center';

    document.body.appendChild(modalContainer);
  }

  modalContainer.innerHTML = `
    <div style="background: #1e293b; border: 1px solid #334155; border-radius: 12px; width: 90%; max-width: 480px; padding: 24px; box-shadow: 0 20px 25px -5px rgba(0,0,0,0.5); font-family: sans-serif; color: #f8fafc;">
      <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #334155; padding-bottom: 12px; margin-bottom: 16px;">
        <h3 style="margin: 0; font-size: 1.1rem; color: #6366f1; display: flex; align-items: center; gap: 8px;">
          🏢 Detalhes do Cliente
        </h3>
        <button onclick="fecharModalCliente()" style="background: transparent; border: none; color: #94a3b8; font-size: 1.5rem; cursor: pointer;">&times;</button>
      </div>
      
      <div style="display: flex; flex-direction: column; gap: 12px; font-size: 0.95rem;">
        <div style="background: #0f172a; padding: 12px; border-radius: 8px; border-left: 4px solid #6366f1;">
          <span style="color: #94a3b8; font-size: 0.8rem; text-transform: uppercase; display: block;">Cliente</span>
          <strong style="font-size: 1.05rem; color: #ffffff;">${dados.nome}</strong>
        </div>

        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px;">
          <div style="background: #0f172a; padding: 10px; border-radius: 8px;">
            <span style="color: #94a3b8; font-size: 0.75rem; text-transform: uppercase; display: block;">Classe</span>
            <strong style="color: #38bdf8;">${dados.classe || '-'}</strong>
          </div>
          ${dados.faturamento ? `
          <div style="background: #0f172a; padding: 10px; border-radius: 8px;">
            <span style="color: #94a3b8; font-size: 0.75rem; text-transform: uppercase; display: block;">Faturamento</span>
            <strong style="color: #10b981;">${dados.faturamento}</strong>
          </div>` : ''}
          ${dados.participacao ? `
          <div style="background: #0f172a; padding: 10px; border-radius: 8px;">
            <span style="color: #94a3b8; font-size: 0.75rem; text-transform: uppercase; display: block;">% do Total</span>
            <strong style="color: #f59e0b;">${dados.participacao}</strong>
          </div>` : ''}
          ${dados.ultimaFatura ? `
          <div style="background: #0f172a; padding: 10px; border-radius: 8px;">
            <span style="color: #94a3b8; font-size: 0.75rem; text-transform: uppercase; display: block;">Última Fatura</span>
            <strong style="color: #e2e8f0;">${dados.ultimaFatura}</strong>
          </div>` : ''}
          ${dados.inatividade ? `
          <div style="background: #0f172a; padding: 10px; border-radius: 8px;">
            <span style="color: #94a3b8; font-size: 0.75rem; text-transform: uppercase; display: block;">Inatividade</span>
            <strong style="color: #ef4444;">${dados.inatividade}</strong>
          </div>` : ''}
        </div>
      </div>

      <div style="margin-top: 20px; text-align: right;">
        <button onclick="fecharModalCliente()" style="background: #6366f1; color: white; border: none; padding: 8px 18px; border-radius: 6px; font-weight: 600; cursor: pointer;">
          Fechar
        </button>
      </div>
    </div>
  `;

  modalContainer.style.display = 'flex';
}

function fecharModalCliente() {
  const modalContainer = document.getElementById('customClientModal');
  if (modalContainer) modalContainer.style.display = 'none';
}

function parseCurrency(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  let str = String(val).replace('R$', '').replace(/\s/g, '').trim();
  if (str.includes(',')) str = str.replace(/\./g, '').replace(',', '.');
  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

function parsePct(val) {
  if (val === undefined || val === null || val === '') return 0;
  if (typeof val === 'number') return val <= 1 ? val * 100 : val;
  let str = String(val).replace('%', '').replace(',', '.').trim();
  let num = parseFloat(str);
  if (isNaN(num)) return 0;
  return num <= 1 && str.indexOf('.') !== -1 ? num * 100 : num;
}

function formatDate(val) {
  if (!val) return '-';
  if (val instanceof Date) return val.toISOString().split('T')[0];
  return String(val).split('T')[0];
}

function formatBRL(val) {
  return val.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function mesesArray() {
  return ['Mês 1', 'Mês 2', 'Mês 3', 'Mês 4', 'Mês 5', 'Mês 6', 'Mês 7', 'Mês 8', 'Mês 9', 'Mês 10', 'Mês 11', 'Mês 12'];
}

function destroyChart(chartId) {
  if (charts[chartId]) {
    charts[chartId].destroy();
    charts[chartId] = null;
  }
  // Segurança extra: se por algum motivo sobrou um gráfico "grudado" nesse
  // canvas (ex: página restaurada pelo navegador), remove antes de desenhar
  // um novo, senão o Chart.js recusa e o gráfico fica em branco.
  const canvas = document.getElementById(chartId);
  if (canvas && typeof Chart !== 'undefined' && typeof Chart.getChart === 'function') {
    const existente = Chart.getChart(canvas);
    if (existente) existente.destroy();
  }
}
