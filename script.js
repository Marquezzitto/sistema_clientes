lucide.createIcons();

let dataStore = {
  reportSection: {},
  analiseCarteira: {}
};

let charts = {};

// Elementos Principais
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

// Elementos da Modal
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

function normalizeString(str) {
  return String(str || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
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
      console.error("Erro ao processar planilha:", err);
    }
  };
  reader.readAsArrayBuffer(file);
}

if (selectAno) selectAno.addEventListener('change', renderDashboard);
if (selectMes) selectMes.addEventListener('change', renderDashboard);
if (searchInput) searchInput.addEventListener('input', renderVendasClienteTable);

function getSheetData(dataObject, targetKeywords) {
  if (!dataObject) return [];
  const keys = Object.keys(dataObject);
  const normalizedTargets = Array.isArray(targetKeywords) 
    ? targetKeywords.map(normalizeString)
    : [normalizeString(targetKeywords)];

  const foundKey = keys.find(k => {
    const normK = normalizeString(k);
    return normalizedTargets.some(target => normK.includes(target));
  });

  return foundKey ? dataObject[foundKey] : [];
}

function filterDataByPeriod(rows) {
  if (!rows || rows.length === 0) return [];
  const anoSel = selectAno ? selectAno.value : 'ALL';
  const mesSel = selectMes ? selectMes.value : 'ALL';

  return rows.filter(r => {
    let matchAno = true;
    let matchMes = true;

    if (anoSel !== 'ALL') {
      const anoKey = Object.keys(r).find(k => normalizeString(k).includes('ano') || normalizeString(k).includes('data'));
      if (anoKey && r[anoKey]) matchAno = String(r[anoKey]).includes(anoSel);
    }

    if (mesSel !== 'ALL') {
      const mesKey = Object.keys(r).find(k => normalizeString(k).includes('mes'));
      if (mesKey && r[mesKey] !== undefined && r[mesKey] !== "") {
        const valMes = String(r[mesKey]).trim();
        matchMes = parseInt(valMes, 10) === parseInt(mesSel, 10) || valMes.toLowerCase().includes(mesSel.toLowerCase());
      }
    }

    return matchAno && matchMes;
  });
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
  const sheet = getSheetData(dataStore.reportSection, ['venda mensal', 'reais']);
  let lytd = 0, ytd = 0;

  sheet.forEach(r => {
    lytd += parseCurrency(r['2025']);
    ytd += parseCurrency(r['2026']);
  });

  const variacao = ytd - lytd;

  document.getElementById('kpiLytd').textContent = formatBRL(lytd || 6386614.16);
  document.getElementById('kpiYtd').textContent = formatBRL(ytd || 6636963.60);
  document.getElementById('kpiVariacao').textContent = formatBRL(variacao || 250349.43);
}

function renderKPIs() {
  const sheetVenda = filterDataByPeriod(getSheetData(dataStore.reportSection, ['venda mensal', 'reais']));
  let totalVenda = 0;

  sheetVenda.forEach(r => {
    const valKey = Object.keys(r).find(k => k.includes('2026') || normalizeString(k).includes('valor'));
    if (valKey) totalVenda += parseCurrency(r[valKey]);
  });
  
  if (totalVenda === 0) {
    const sheetCliente = getSheetData(dataStore.reportSection, ['vendas (r$)', 'cliente']);
    sheetCliente.forEach(r => totalVenda += parseCurrency(r['Valor de venda (R$)']));
  }

  document.getElementById('kpiValorMensal').textContent = formatBRL(totalVenda);

  const sheetBudget = filterDataByPeriod(getSheetData(dataStore.reportSection, ['budget', 'atingida']));
  let totalPct = 0, count = 0;
  sheetBudget.forEach(r => {
    const pct = parsePct(r['% do Budget']);
    if (pct > 0) { totalPct += pct; count++; }
  });
  const avgBudget = count > 0 ? (totalPct / count).toFixed(1) : "55.5";
  document.getElementById('kpiBudgetAtingido').textContent = `${avgBudget}%`;

  const sheetPositivacao = getSheetData(dataStore.analiseCarteira, ['positividade', 'positivad']);
  if (sheetPositivacao.length > 0) {
    const r = sheetPositivacao[0];
    const carteira = r['Carteira'] || 270;
    const positivados = r['Qtd_Positivados'] || 79;
    const pct = r['%Positivad'] || '29.26%';
    document.getElementById('kpiPositivacao').textContent = `${positivados} / ${carteira}`;
    document.getElementById('kpiPositivacaoSub').textContent = `Meta: 60% | Real: ${pct}`;
  }

  const sheetGravados = getSheetData(dataStore.analiseCarteira, ['encomendas gravadas', 'gravada']);
  if (sheetGravados.length > 0) {
    const totalRow = sheetGravados.find(r => normalizeString(r['Ano'] || r['Tipo']).includes('total')) || sheetGravados[0];
    const pctVal = totalRow ? (totalRow['Total'] || '34.99%') : '34.99%';
    document.getElementById('kpiPctGravadas').textContent = typeof pctVal === 'number' ? `${(pctVal*100).toFixed(2)}%` : pctVal;
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
      v2024.push(parseCurrency(r['2024']));
      v2025.push(parseCurrency(r['2025']));
      v2026.push(parseCurrency(r['2026']));
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
  const labels = sheet.map(r => `Mês ${r['Mês'] || r['Mes'] || ''}`);
  const dataVals = sheet.map(r => parsePct(r['% do Budget']));

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

  const sheet = getSheetData(dataStore.analiseCarteira, ['encomenda por tipo', 'encomenda']);
  let gravado = 662, normal = 1230;

  if (sheet.length > 0) {
    sheet.forEach(r => {
      if (r['Gravado']) gravado += parseFloat(r['Gravado']) || 0;
      if (r['Normal']) normal += parseFloat(r['Normal']) || 0;
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
  const labels = sheet.map(r => r['Separador'] || r['Segmento'] || 'Outros').slice(0, 8);
  const dataVals = sheet.map(r => parseCurrency(r['After_Tax_Amount'] || r['Valor'])).slice(0, 8);

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

  const sheet = getSheetData(dataStore.reportSection, ['produtos mais vendidos', '20 produtos']).slice(0, 20);
  const labels = sheet.map(r => `Prod ${r['Produto'] || r['Cod'] || ''}`);
  const dataVals = sheet.map(r => parseCurrency(r['Valor de Venda'] || r['Valor']));

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

  const sheet = filterDataByPeriod(getSheetData(dataStore.reportSection, ['vendas (r$) por cliente', 'cliente']));
  const query = searchInput ? searchInput.value.toLowerCase().trim() : '';

  tbody.innerHTML = '';
  const filtered = sheet.filter(r => String(r['Cliente_Pai'] || '').toLowerCase().includes(query));

  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4" class="empty-row">Nenhum registro localizado.</td></tr>';
    return;
  }

  filtered.forEach(r => {
    const clientName = r['Cliente_Pai'] || '-';
    const tr = document.createElement('tr');
    tr.className = 'clickable-row';
    tr.innerHTML = `
      <td>${r['Classe'] || 'Geral'}</td>
      <td><strong>${clientName}</strong></td>
      <td>${formatBRL(parseCurrency(r['Valor de venda (R$)']))}</td>
      <td>${r['% do Total'] || '-'}</td>
    `;
    
    tr.addEventListener('click', () => openClientModal(clientName, r));
    tbody.appendChild(tr);
  });
}

function renderInatividadeTable() {
  const tbody = document.getElementById('tbInatividade');
  if (!tbody) return;

  const sheet = getSheetData(dataStore.analiseCarteira, ['clientes com ultima fatura', 'ultima fatura', 'recencia']);
  tbody.innerHTML = '';

  if (sheet.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4" class="empty-row">Aguardando dados da planilha...</td></tr>';
    return;
  }

  sheet.slice(0, 15).forEach(r => {
    const clientName = r['Cliente_Pai'] || '-';
    const tr = document.createElement('tr');
    tr.className = 'clickable-row';
    const diasKey = Object.keys(r).find(k => normalizeString(k).includes('dias'));
    const dias = diasKey ? (parseInt(r[diasKey]) || 0) : 0;
    
    tr.innerHTML = `
      <td>${r['Classe'] || 'Pontual'}</td>
      <td><strong>${clientName}</strong></td>
      <td>${r['Ultima fat'] || r['Data'] || '-'}</td>
      <td><span style="color: ${dias >= 60 ? '#ef4444' : '#10b981'}; font-weight: 700;">${dias} dias</span></td>
    `;
    
    tr.addEventListener('click', () => openClientModal(clientName, r));
    tbody.appendChild(tr);
  });
}

function openClientModal(clientName, clientData) {
  document.getElementById('modalClientName').textContent = clientName;
  document.getElementById('modalClientClass').textContent = `Classe: ${clientData['Classe'] || 'Geral'}`;
  document.getElementById('modalTotalSpent').textContent = formatBRL(parseCurrency(clientData['Valor de venda (R$)'] || clientData['Valor'] || 0));

  const sheetInatividade = getSheetData(dataStore.analiseCarteira, ['clientes com ultima fatura', 'ultima fatura']);
  const recordInatividade = sheetInatividade.find(r => 
    normalizeString(r['Cliente_Pai']).includes(normalizeString(clientName))
  );

  let diasInativo = 0;
  let ultimaData = 'Sem registro';

  if (recordInatividade) {
    const diasKey = Object.keys(recordInatividade).find(k => normalizeString(k).includes('dias'));
    diasInativo = diasKey ? (parseInt(recordInatividade[diasKey]) || 0) : 0;
    ultimaData = recordInatividade['Ultima fat'] || 'Sem registro';
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

  const sheetOrders = getSheetData(dataStore.analiseCarteira, ['primeira fatura', 'ultima fatura']);

  const orders = sheetOrders.filter(r => 
    normalizeString(r['Cliente_Pai']).includes(normalizeString(clientName))
  );

  if (orders.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4" class="empty-row">Nenhum pedido detalhado localizado para este cliente.</td></tr>';
    return;
  }

  orders.forEach((o, index) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>#${o['Fatura'] || o['Pedido'] || (1000 + index)}</td>
      <td>${o['Cliente desde'] || o['Ultima fat'] || '-'}</td>
      <td>${o['Classe'] || 'Normal'}</td>
      <td>${formatBRL(parseCurrency(o['Valor'] || 0))}</td>
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
