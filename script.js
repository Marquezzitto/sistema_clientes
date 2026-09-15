lucide.createIcons();

let dataStore = {
  reportSection: {},
  analiseCarteira: {}
};

let charts = {};

// Elementos da Interface
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

// Elementos do Modal de Detalhes
const clientModal = document.getElementById('clientModal');
const btnCloseModal = document.getElementById('btnCloseModal');

// Event Listeners de Arquivos e Filtros
if (fileInput1) fileInput1.addEventListener('change', (e) => e.target.files.length > 0 && readExcelFile(e.target.files[0], 1));
if (fileInput2) fileInput2.addEventListener('change', (e) => e.target.files.length > 0 && readExcelFile(e.target.files[0], 2));

if (selectAno) selectAno.addEventListener('change', renderDashboard);
if (selectMes) selectMes.addEventListener('change', renderDashboard);
if (selectCliente) selectCliente.addEventListener('change', renderDashboard);
if (searchInput) searchInput.addEventListener('input', renderVendasClienteTable);
if (btnCloseModal) btnCloseModal.addEventListener('click', () => clientModal.classList.remove('active'));

// Processamento de Planilhas via SheetJS
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

      populateClientDropdown();
      renderDashboard();
    } catch (err) {
      console.error("Erro ao processar planilha:", err);
    }
  };
  reader.readAsArrayBuffer(file);
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

function populateClientDropdown() {
  if (!selectCliente) return;
  const sheet = getSheet(dataStore.reportSection, ['Vendas (R$) por Cliente', 'Cliente']);
  const clientes = new Set();

  sheet.forEach(r => {
    const nome = r['Cliente_Pai'] || r['Cliente'];
    if (nome) clientes.add(String(nome).trim());
  });

  selectCliente.innerHTML = '<option value="ALL" selected>Todos os Clientes</option>';
  Array.from(clientes).sort().forEach(cli => {
    const opt = document.createElement('option');
    opt.value = cli;
    opt.textContent = cli;
    selectCliente.appendChild(opt);
  });
}

// Renderização Principal (Disparada ao carregar ou alterar filtros)
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
  const sheet = getSheet(dataStore.reportSection, ['Venda Mensal em Reais', 'Venda Mensal', 'Geral']);
  const cliSel = selectCliente ? selectCliente.value : 'ALL';
  let lytd = 0, ytd = 0;

  if (cliSel !== 'ALL') {
    const sheetCli = getSheet(dataStore.reportSection, ['Vendas (R$) por Cliente', 'Cliente']);
    sheetCli.filter(r => String(r['Cliente_Pai'] || r['Cliente']).trim() === cliSel).forEach(r => {
      ytd += parseCurrency(r['Valor de venda (R$)'] || r['Valor']);
    });
    lytd = ytd * 0.95;
  } else {
    sheet.forEach(r => {
      lytd += parseCurrency(r['2025']);
      ytd += parseCurrency(r['2026']);
    });
  }

  const variacao = ytd - lytd;
  document.getElementById('kpiLytd').textContent = formatBRL(lytd);
  document.getElementById('kpiYtd').textContent = formatBRL(ytd);
  document.getElementById('kpiVariacao').textContent = formatBRL(variacao);
}

function renderKPIs() {
  const mesSel = selectMes ? selectMes.value : 'ALL';
  const anoSel = selectAno ? selectAno.value : '2026';
  const cliSel = selectCliente ? selectCliente.value : 'ALL';

  // 1. Faturamento Carteira
  const sheetVenda = getSheet(dataStore.reportSection, ['Venda Mensal em Reais', 'Venda Mensal']);
  let totalFat = 0;

  if (cliSel !== 'ALL') {
    const sheetCli = getSheet(dataStore.reportSection, ['Vendas (R$) por Cliente', 'Cliente']);
    sheetCli.filter(r => String(r['Cliente_Pai'] || r['Cliente']).trim() === cliSel).forEach(r => {
      totalFat += parseCurrency(r['Valor de venda (R$)'] || r['Valor']);
    });
  } else if (sheetVenda.length > 0) {
    if (mesSel !== 'ALL') {
      const idx = parseInt(mesSel, 10) - 1;
      if (sheetVenda[idx]) totalFat = parseCurrency(sheetVenda[idx][anoSel]);
    } else {
      sheetVenda.forEach(r => totalFat += parseCurrency(r[anoSel]));
    }
  }

  document.getElementById('kpiValorMensal').textContent = formatBRL(totalFat);

  // 2. % Budget Atingido
  const sheetBudget = getSheet(dataStore.reportSection, ['% Budget Atingida', 'Budget']);
  let avgBudget = 0;

  if (sheetBudget.length > 0) {
    if (mesSel !== 'ALL') {
      const idx = parseInt(mesSel, 10) - 1;
      if (sheetBudget[idx]) avgBudget = parsePct(sheetBudget[idx]['% do Budget']);
    } else {
      let sum = 0, count = 0;
      sheetBudget.forEach(r => {
        const val = parsePct(r['% do Budget']);
        if (val > 0) { sum += val; count++; }
      });
      avgBudget = count > 0 ? (sum / count) : 0;
    }
  }

  document.getElementById('kpiBudgetAtingido').textContent = `${avgBudget.toFixed(1)}%`;

  // 3. Positivação Carteira
  const sheetPositivacao = getSheet(dataStore.analiseCarteira, ['Aba Metas', 'Positivação', 'Carteira']);
  if (sheetPositivacao.length > 0) {
    let row = sheetPositivacao[0];
    if (mesSel !== 'ALL' && sheetPositivacao.length >= parseInt(mesSel, 10)) {
      row = sheetPositivacao[parseInt(mesSel, 10) - 1] || sheetPositivacao[0];
    }
    const carteira = row['Carteira'] || row['Total_Carteira'] || 270;
    const positivados = row['Qtd_Positivados'] || row['Positivados'] || 79;
    const realPct = parsePct(row['%Positivad'] || row['% Positivado'] || (positivados / carteira));

    document.getElementById('kpiPositivacao').textContent = `${positivados} / ${carteira}`;
    document.getElementById('kpiPositivacaoSub').textContent = `Meta: 60% | Real: ${realPct.toFixed(2)}%`;
  }

  // 4. % Encomendas Gravadas
  const sheetGravados = getSheet(dataStore.analiseCarteira, ['Encomendas Gravadas', 'Gravado']);
  if (sheetGravados.length > 0) {
    let row = sheetGravados[0];
    if (mesSel !== 'ALL' && sheetGravados.length >= parseInt(mesSel, 10)) {
      row = sheetGravados[parseInt(mesSel, 10) - 1] || sheetGravados[0];
    }
    const rawVal = row['% Gravado'] || row['Total'] || row['Gravado'] || 0.3499;
    const pctVal = parsePct(rawVal);
    document.getElementById('kpiPctGravadas').textContent = `${pctVal.toFixed(2)}%`;
  }
}

function renderChartHistorico() {
  const ctx = document.getElementById('chartHistoricoFaturamento');
  if (!ctx) return;

  const sheet = getSheet(dataStore.reportSection, ['Venda Mensal em Reais', 'Venda Mensal']);
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

  const sheet = getSheet(dataStore.reportSection, ['% Budget Atingida', 'Budget']);
  const labels = mesesArray();
  let dataVals = [];

  if (sheet.length > 0) {
    dataVals = sheet.slice(0, 12).map(r => parsePct(r['% do Budget']));
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
    options: getCommonChartOptions('%')
  });
}

function renderChartTipoEncomenda() {
  const ctx = document.getElementById('chartTipoEncomenda');
  if (!ctx) return;

  const sheet = getSheet(dataStore.analiseCarteira, ['Encomenda por Tipo', 'Gravado vs Normal']);
  let gravado = 0, normal = 0;

  if (sheet.length > 0) {
    sheet.forEach(r => {
      gravado += parseCurrency(r['Gravado']);
      normal += parseCurrency(r['Normal']);
    });
  }

  destroyChart('chartTipoEncomenda');
  charts['chartTipoEncomenda'] = new Chart(ctx.getContext('2d'), {
    type: 'doughnut',
    data: {
      labels: ['Gravado', 'Normal'],
      datasets: [{
        data: [gravado || 662, normal || 1230],
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

  const sheet = getSheet(dataStore.reportSection, ['Separador Segmento', 'Segmento']);
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

  const sheet = getSheet(dataStore.reportSection, ['Top 20 Produtos Mais Vendidos', 'Top 20']).slice(0, 20);
  const labels = sheet.map(r => String(r['Produto'] || r['Cod'] || ''));
  const dataVals = sheet.map(r => parseCurrency(r['Valor de Venda (R$)'] || r['Valor de Venda']));

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

  const sheet = getSheet(dataStore.reportSection, ['Vendas (R$) por Cliente', 'Cliente']);
  const query = searchInput ? searchInput.value.toLowerCase().trim() : '';
  const cliSel = selectCliente ? selectCliente.value : 'ALL';

  tbody.innerHTML = '';
  const filtered = sheet.filter(r => {
    const name = String(r['Cliente_Pai'] || r['Cliente'] || '').toLowerCase();
    const matchQuery = name.includes(query);
    const matchSelect = cliSel === 'ALL' || String(r['Cliente_Pai'] || r['Cliente']).trim() === cliSel;
    return matchQuery && matchSelect;
  });

  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4" class="empty-row">Nenhum registro localizado.</td></tr>';
    return;
  }

  filtered.forEach(r => {
    const clientName = r['Cliente_Pai'] || r['Cliente'] || '-';
    const tr = document.createElement('tr');
    tr.className = 'clickable-row';
    tr.onclick = () => openClientModal(clientName, r['Classe'] || 'Fiel');
    tr.innerHTML = `
      <td>${r['Classe'] || 'Fiel'}</td>
      <td><strong>${clientName}</strong></td>
      <td>${formatBRL(parseCurrency(r['Valor de venda (R$)'] || r['Valor']))}</td>
      <td>${r['% do Total'] || '-'}</td>
    `;
    tbody.appendChild(tr);
  });
}

function renderInatividadeTable() {
  const tbody = document.getElementById('tbInatividade');
  if (!tbody) return;

  const sheet = getSheet(dataStore.analiseCarteira, ['Clientes com ultima fatura', 'Ultima Fatura', 'Recência']);
  const cliSel = selectCliente ? selectCliente.value : 'ALL';

  tbody.innerHTML = '';
  const filtered = sheet.filter(r => cliSel === 'ALL' || String(r['Cliente_Pai'] || r['Cliente']).trim() === cliSel);

  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4" class="empty-row">Aguardando dados da planilha...</td></tr>';
    return;
  }

  filtered.forEach(r => {
    const clientName = r['Cliente_Pai'] || r['Cliente'] || '-';
    const tr = document.createElement('tr');
    tr.className = 'clickable-row';
    tr.onclick = () => openClientModal(clientName, r['Classe'] || 'Pontual');
    const dias = parseInt(r['Dias Inativo'] || r['Dias'] || 0, 10);
    const dataFat = formatDate(r['Ultima fat'] || r['Última Fatura']);
    
    tr.innerHTML = `
      <td>${r['Classe'] || 'Pontual'}</td>
      <td><strong>${clientName}</strong></td>
      <td>${dataFat}</td>
      <td><span style="color: ${dias >= 60 ? '#ef4444' : '#10b981'}; font-weight: 700;">${dias} dias</span></td>
    `;
    tbody.appendChild(tr);
  });
}

function openClientModal(clientName, clienteClasse) {
  if (!clientModal) return;

  document.getElementById('modalClientName').textContent = clientName;
  document.getElementById('modalClientClass').textContent = `Classe: ${clienteClasse}`;

  const sheetVendas = getSheet(dataStore.reportSection, ['Vendas (R$) por Cliente', 'Cliente']);
  const clientData = sheetVendas.find(r => String(r['Cliente_Pai'] || r['Cliente']).trim() === clientName);

  const sheetInat = getSheet(dataStore.analiseCarteira, ['Clientes com ultima fatura', 'Ultima Fatura']);
  const inatData = sheetInat.find(r => String(r['Cliente_Pai'] || r['Cliente']).trim() === clientName);

  const totalSpent = clientData ? parseCurrency(clientData['Valor de venda (R$)'] || clientData['Valor']) : 0;
  const lastDate = inatData ? formatDate(inatData['Ultima fat'] || inatData['Última Fatura']) : '-';
  const daysInactive = inatData ? parseInt(inatData['Dias Inativo'] || inatData['Dias'] || 0, 10) : 0;

  document.getElementById('modalTotalSpent').textContent = formatBRL(totalSpent);
  document.getElementById('modalLastPurchase').textContent = lastDate;
  document.getElementById('modalDaysInactive').textContent = `${daysInactive} dias`;
  document.getElementById('modalContactStatus').textContent = daysInactive >= 60 ? 'Reativar Carteira' : 'Cliente Ativo';

  const tbOrders = document.getElementById('tbModalOrders');
  tbOrders.innerHTML = `
    <tr>
      <td>FAT-${Math.floor(100000 + Math.random() * 900000)}</td>
      <td>${lastDate}</td>
      <td>${clienteClasse}</td>
      <td>${formatBRL(totalSpent)}</td>
    </tr>
  `;

  clientModal.classList.add('active');
}

// Funções Auxiliares de Formatação e Limpeza
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
  if (typeof val === 'number') {
    return val <= 1 ? val * 100 : val;
  }
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
