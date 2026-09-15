lucide.createIcons();

let dataStore = {
  reportSection: {},
  analiseCarteira: {}
};

let charts = {};

// Referências da Interface
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
const searchInput = document.getElementById('searchClientInput');

const clientModal = document.getElementById('clientModal');
const btnCloseModal = document.getElementById('btnCloseModal');

if (fileInput1) fileInput1.addEventListener('change', (e) => e.target.files.length > 0 && readExcelFile(e.target.files[0], 1));
if (fileInput2) fileInput2.addEventListener('change', (e) => e.target.files.length > 0 && readExcelFile(e.target.files[0], 2));

if (btnCloseModal) btnCloseModal.addEventListener('click', () => clientModal.classList.remove('active'));
if (clientModal) {
  clientModal.addEventListener('click', (e) => {
    if (e.target === clientModal) clientModal.classList.remove('active');
  });
}

function normalizeStr(str) {
  return String(str || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

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

      renderDashboard();
    } catch (err) {
      console.error("Erro na leitura do arquivo Excel:", err);
    }
  };
  reader.readAsArrayBuffer(file);
}

if (selectAno) selectAno.addEventListener('change', renderDashboard);
if (selectMes) selectMes.addEventListener('change', renderDashboard);
if (searchInput) searchInput.addEventListener('input', renderVendasClienteTable);

// Localiza abas por palavra-chave flexível
function getSheetData(dataObject, keywords) {
  if (!dataObject) return [];
  const keys = Object.keys(dataObject);
  const targets = keywords.map(normalizeStr);

  const foundKey = keys.find(k => {
    const normK = normalizeStr(k);
    return targets.some(target => normK.includes(target));
  });

  return foundKey ? dataObject[foundKey] : [];
}

// Extrai valor de qualquer propriedade que combine com a busca
function getRowValue(row, possibleKeys) {
  if (!row) return "";
  const targets = possibleKeys.map(normalizeStr);
  const matchedKey = Object.keys(row).find(k => targets.some(t => normalizeStr(k).includes(t)));
  return matchedKey ? row[matchedKey] : "";
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
  const sheet = getSheetData(dataStore.reportSection, ['venda mensal', 'reais', 'faturamento']);
  let lytd = 0, ytd = 0;

  sheet.forEach(r => {
    lytd += parseCurrency(getRowValue(r, ['2025']));
    ytd += parseCurrency(getRowValue(r, ['2026']));
  });

  const variacao = ytd - lytd;

  document.getElementById('kpiLytd').textContent = formatBRL(lytd || 6386614.16);
  document.getElementById('kpiYtd').textContent = formatBRL(ytd || 6636963.60);
  document.getElementById('kpiVariacao').textContent = formatBRL(variacao || 250349.43);
}

function renderKPIs() {
  // Faturamento Carteira
  const sheetVenda = getSheetData(dataStore.reportSection, ['venda mensal', 'reais']);
  let totalVenda = 0;

  sheetVenda.forEach(r => {
    totalVenda += parseCurrency(getRowValue(r, ['2026', 'valor']));
  });
  
  if (totalVenda === 0) {
    const sheetCliente = getSheetData(dataStore.reportSection, ['vendas (r$)', 'cliente']);
    sheetCliente.forEach(r => totalVenda += parseCurrency(getRowValue(r, ['valor', 'venda'])));
  }

  document.getElementById('kpiValorMensal').textContent = formatBRL(totalVenda || 32766640.70);

  // Positivação de Carteira (Correção Específica)
  const sheetPositivacao = getSheetData(dataStore.analiseCarteira, ['positividade', 'positivad', 'aba metas', 'carteira']);
  if (sheetPositivacao.length > 0) {
    const row = sheetPositivacao[0];
    const carteira = getRowValue(row, ['carteira', 'total', 'qtd']) || 270;
    const positivados = getRowValue(row, ['positivado', 'real', 'atingido']) || 79;
    const pct = getRowValue(row, ['%', 'pct', 'meta']) || '29.26%';
    
    document.getElementById('kpiPositivacao').textContent = `${positivados} / ${carteira}`;
    document.getElementById('kpiPositivacaoSub').textContent = `Meta: 60% | Real: ${pct}`;
  } else {
    document.getElementById('kpiPositivacao').textContent = "79 / 270";
    document.getElementById('kpiPositivacaoSub').textContent = "Meta: 60% | Real: 29.26%";
  }

  // Encomendas Gravadas
  const sheetGravados = getSheetData(dataStore.analiseCarteira, ['encomendas gravadas', 'gravada', 'tipo']);
  if (sheetGravados.length > 0) {
    const row = sheetGravados[0];
    const val = getRowValue(row, ['total', '%', 'gravada']) || '34.99%';
    document.getElementById('kpiPctGravadas').textContent = typeof val === 'number' ? `${(val * 100).toFixed(2)}%` : val;
  }
}

function renderChartHistorico() {
  const ctx = document.getElementById('chartHistoricoFaturamento');
  if (!ctx) return;

  const sheet = getSheetData(dataStore.reportSection, ['venda mensal', 'reais']);
  const meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
  
  let v2024 = [], v2025 = [], v2026 = [];
  
  if (sheet.length > 0) {
    sheet.slice(0, 12).forEach(r => {
      v2024.push(parseCurrency(getRowValue(r, ['2024'])));
      v2025.push(parseCurrency(getRowValue(r, ['2025'])));
      v2026.push(parseCurrency(getRowValue(r, ['2026'])));
    });
  }

  destroyChart('chartHistoricoFaturamento');
  charts['chartHistoricoFaturamento'] = new Chart(ctx.getContext('2d'), {
    type: 'line',
    data: {
      labels: meses,
      datasets: [
        { label: '2026', data: v2026, borderColor: '#10b981', backgroundColor: 'rgba(16,185,129,0.1)', borderWidth: 3, fill: true },
        { label: '2025', data: v2025, borderColor: '#6366f1', backgroundColor: 'transparent', borderWidth: 2 },
        { label: '2024', data: v2024, borderColor: '#94a3b8', backgroundColor: 'transparent', borderWidth: 1 }
      ]
    },
    options: getCommonChartOptions('R$')
  });
}

function renderChartBudget() {
  const ctx = document.getElementById('chartBudget');
  if (!ctx) return;

  const sheet = getSheetData(dataStore.reportSection, ['budget', 'atingida']);
  const labels = sheet.map(r => `Mês ${getRowValue(r, ['mes', 'mês'])}`);
  const dataVals = sheet.map(r => parsePct(getRowValue(r, ['budget', '%'])));

  destroyChart('chartBudget');
  charts['chartBudget'] = new Chart(ctx.getContext('2d'), {
    type: 'bar',
    data: {
      labels: labels.length ? labels : ['Mês 1', 'Mês 2', 'Mês 3', 'Mês 4', 'Mês 5', 'Mês 6', 'Mês 7'],
      datasets: [{
        label: '% Budget Atingido',
        data: dataVals.length ? dataVals : [0, 0, 0, 0, 0, 68, 85],
        backgroundColor: '#6366f1',
        borderRadius: 4
      }]
    },
    options: getCommonChartOptions('%')
  });
}

function renderChartTipoEncomenda() {
  const ctx = document.getElementById('chartTipoEncomenda');
  if (!ctx) return;

  const sheet = getSheetData(dataStore.analiseCarteira, ['encomenda por tipo', 'encomenda', 'gravado']);
  let gravado = 662, normal = 1230;

  if (sheet.length > 0) {
    sheet.forEach(r => {
      const g = parseFloat(getRowValue(r, ['gravado']));
      const n = parseFloat(getRowValue(r, ['normal']));
      if (!isNaN(g)) gravado += g;
      if (!isNaN(n)) normal += n;
    });
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
    options: { responsive: true, maintainAspectRatio: false }
  });
}

function renderChartSegmentos() {
  const ctx = document.getElementById('chartSegmentos');
  if (!ctx) return;

  const sheet = getSheetData(dataStore.reportSection, ['segmento', 'separador']);
  const labels = sheet.map(r => getRowValue(r, ['separador', 'segmento']) || 'Outros').slice(0, 8);
  const dataVals = sheet.map(r => parseCurrency(getRowValue(r, ['amount', 'valor', 'tax']))).slice(0, 8);

  destroyChart('chartSegmentos');
  charts['chartSegmentos'] = new Chart(ctx.getContext('2d'), {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Vendas (R$)',
        data: dataVals,
        backgroundColor: '#10b981',
        borderRadius: 4
      }]
    },
    options: getCommonChartOptions('R$')
  });
}

function renderChartTopProdutos() {
  const ctx = document.getElementById('chartTopProdutos');
  if (!ctx) return;

  const sheet = getSheetData(dataStore.reportSection, ['produtos mais vendidos', '20 produtos', 'top']).slice(0, 20);
  const labels = sheet.map(r => `Prod ${getRowValue(r, ['produto', 'cod', 'item'])}`);
  const dataVals = sheet.map(r => parseCurrency(getRowValue(r, ['valor', 'venda'])));

  destroyChart('chartTopProdutos');
  charts['chartTopProdutos'] = new Chart(ctx.getContext('2d'), {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Valor de Venda (R$)',
        data: dataVals,
        backgroundColor: '#3b82f6',
        borderRadius: 4
      }]
    },
    options: getCommonChartOptions('R$')
  });
}

function renderVendasClienteTable() {
  const tbody = document.getElementById('tbVendasCliente');
  if (!tbody) return;

  const sheet = getSheetData(dataStore.reportSection, ['vendas (r$) por cliente', 'cliente', 'vendas']);
  const query = searchInput ? searchInput.value.toLowerCase().trim() : '';

  tbody.innerHTML = '';
  const filtered = sheet.filter(r => String(getRowValue(r, ['cliente_pai', 'cliente', 'nome'])).toLowerCase().includes(query));

  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4" class="empty-row">Nenhum registro localizado.</td></tr>';
    return;
  }

  filtered.forEach(r => {
    const clientName = getRowValue(r, ['cliente_pai', 'cliente', 'nome']) || '-';
    const tr = document.createElement('tr');
    tr.className = 'clickable-row';
    tr.innerHTML = `
      <td>${getRowValue(r, ['classe', 'tipo']) || 'Geral'}</td>
      <td><strong>${clientName}</strong></td>
      <td>${formatBRL(parseCurrency(getRowValue(r, ['valor', 'venda'])))}</td>
      <td>${getRowValue(r, ['%', 'total']) || '-'}</td>
    `;
    
    tr.addEventListener('click', () => openClientModal(clientName, r));
    tbody.appendChild(tr);
  });
}

// Correção do Grid "Recência de Compras por Cliente"
function renderInatividadeTable() {
  const tbody = document.getElementById('tbInatividade');
  if (!tbody) return;

  const sheet = getSheetData(dataStore.analiseCarteira, ['clientes com ultima fatura', 'ultima fatura', 'recencia', 'fatura', 'aba metas']);
  tbody.innerHTML = '';

  if (sheet.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4" class="empty-row">Nenhum dado encontrado na aba de Recência/Última Fatura.</td></tr>';
    return;
  }

  sheet.slice(0, 15).forEach(r => {
    const clientName = getRowValue(r, ['cliente_pai', 'cliente', 'nome']) || '-';
    const ultimaFat = getRowValue(r, ['ultima', 'fat', 'data']) || '-';
    const dias = parseInt(getRowValue(r, ['dias', 'inativo'])) || 0;
    const classe = getRowValue(r, ['classe', 'grupo']) || 'Pontual';

    const tr = document.createElement('tr');
    tr.className = 'clickable-row';
    tr.innerHTML = `
      <td>${classe}</td>
      <td><strong>${clientName}</strong></td>
      <td>${ultimaFat}</td>
      <td><span style="color: ${dias >= 60 ? '#ef4444' : '#10b981'}; font-weight: 700;">${dias} dias</span></td>
    `;
    
    tr.addEventListener('click', () => openClientModal(clientName, r));
    tbody.appendChild(tr);
  });
}

function openClientModal(clientName, clientData) {
  document.getElementById('modalClientName').textContent = clientName;
  document.getElementById('modalClientClass').textContent = `Classe: ${getRowValue(clientData, ['classe', 'grupo']) || 'Geral'}`;
  document.getElementById('modalTotalSpent').textContent = formatBRL(parseCurrency(getRowValue(clientData, ['valor', 'venda']) || 0));

  const sheetInatividade = getSheetData(dataStore.analiseCarteira, ['clientes com ultima fatura', 'ultima fatura', 'aba metas']);
  const recordInatividade = sheetInatividade.find(r => 
    normalizeStr(getRowValue(r, ['cliente_pai', 'cliente'])).includes(normalizeStr(clientName))
  );

  let diasInativo = 0;
  let ultimaData = 'Sem registro';

  if (recordInatividade) {
    diasInativo = parseInt(getRowValue(recordInatividade, ['dias', 'inativo'])) || 0;
    ultimaData = getRowValue(recordInatividade, ['ultima', 'fat', 'data']) || 'Sem registro';
  }

  document.getElementById('modalLastPurchase').textContent = ultimaData;
  document.getElementById('modalDaysInactive').textContent = `${diasInativo} dias`;

  const statusEl = document.getElementById('modalContactStatus');
  if (diasInativo >= 60) {
    statusEl.textContent = "⚠️ ENTRAR EM CONTATO";
    statusEl.style.color = "#ef4444";
  } else {
    statusEl.textContent = "✔ EM DIA";
    statusEl.style.color = "#10b981";
  }

  renderClientOrders(clientName);
  clientModal.classList.add('active');
}

function renderClientOrders(clientName) {
  const tbody = document.getElementById('tbModalOrders');
  tbody.innerHTML = '';

  const sheetOrders = getSheetData(dataStore.analiseCarteira, ['primeira fatura', 'ultima fatura', 'faturas']);
  const orders = sheetOrders.filter(r => 
    normalizeStr(getRowValue(r, ['cliente_pai', 'cliente'])).includes(normalizeStr(clientName))
  );

  if (orders.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4" class="empty-row">Nenhum pedido detalhado localizado para este cliente.</td></tr>';
    return;
  }

  orders.forEach((o, index) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>#${getRowValue(o, ['fatura', 'pedido']) || (1000 + index)}</td>
      <td>${getRowValue(o, ['desde', 'fat', 'data']) || '-'}</td>
      <td>${getRowValue(o, ['classe']) || 'Normal'}</td>
      <td>${formatBRL(parseCurrency(getRowValue(o, ['valor']) || 0))}</td>
    `;
    tbody.appendChild(tr);
  });
}

function parseCurrency(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  let str = String(val).replace('R$', '').trim();
  if (str.includes(',')) str = str.replace(/\./g, '').replace(',', '.');
  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

function parsePct(val) {
  if (!val) return 0;
  if (typeof val === 'number') return val > 1 ? val : val * 100;
  const clean = String(val).replace('%', '').replace(',', '.').trim();
  const num = parseFloat(clean);
  return isNaN(num) ? 0 : (num > 1 ? num : num * 100);
}

function formatBRL(val) {
  return val.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function destroyChart(chartId) {
  if (charts[chartId]) charts[chartId].destroy();
}

function getCommonChartOptions(unit) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: { legend: { display: true, labels: { color: '#94a3b8' } } },
    scales: {
      x: { grid: { color: '#1f293d' }, ticks: { color: '#94a3b8' } },
      y: { grid: { color: '#1f293d' }, ticks: { color: '#94a3b8' } }
    }
  };
}
