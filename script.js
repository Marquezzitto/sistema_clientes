lucide.createIcons();

// Estruturas de Dados Internos
let dataStore = {
  reportSection: {},
  analiseCarteira: {}
};

// Instâncias dos Gráficos Chart.js
let charts = {};

// Elementos DOM
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

// Handlers de Upload File 1
fileInput1.addEventListener('change', (e) => {
  if (e.target.files.length > 0) readExcelFile(e.target.files[0], 1);
});

// Handlers de Upload File 2
fileInput2.addEventListener('change', (e) => {
  if (e.target.files.length > 0) readExcelFile(e.target.files[0], 2);
});

function readExcelFile(file, fileNum) {
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      
      let parsedSheets = {};
      workbook.SheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        parsedSheets[sheetName.trim()] = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
      });

      if (fileNum === 1) {
        dataStore.reportSection = parsedSheets;
        dropZone1.classList.add('loaded');
        labelFile1.textContent = `✔ ${file.name}`;
      } else {
        dataStore.analiseCarteira = parsedSheets;
        dropZone2.classList.add('loaded');
        labelFile2.textContent = `✔ ${file.name}`;
      }

      checkAndRenderBI();
    } catch (err) {
      console.error(err);
      alert(`Erro ao ler o Arquivo ${fileNum}.`);
    }
  };
  reader.readAsArrayBuffer(file);
}

function checkAndRenderBI() {
  const f1Loaded = Object.keys(dataStore.reportSection).length > 0;
  const f2Loaded = Object.keys(dataStore.analiseCarteira).length > 0;

  if (f1Loaded || f2Loaded) {
    statusBadge.classList.add('active');
    badgeText.textContent = "BI Atualizado";
    renderDashboard();
  }
}

// Eventos dos Filtros
selectAno.addEventListener('change', renderDashboard);
selectMes.addEventListener('change', renderDashboard);
document.getElementById('searchClientInput').addEventListener('input', renderVendasClienteTable);

function renderDashboard() {
  renderKPIs();
  renderChartBudget();
  renderChartTipoEncomenda();
  renderChartTopProdutos();
  renderChartSegmentos();
  renderVendasClienteTable();
  renderInatividadeTable();
}

// 1. Renderização de KPIs
function renderKPIs() {
  // KPI 1: Valor Mensal da Carteira
  const sheetVendaMensal = dataStore.reportSection['Venda mensal em reais da '] || [];
  let totalVenda = 0;
  sheetVendaMensal.forEach(r => {
    const val = parseCurrency(r['After_Tax_Amount']);
    totalVenda += val;
  });
  document.getElementById('kpiValorMensal').textContent = formatBRL(totalVenda);

  // KPI 2: Budget Atingido
  const sheetBudget = dataStore.reportSection['% do Budget atingida por '] || [];
  const selectedMes = selectMes.value;
  let targetRow = sheetBudget;
  if (selectedMes !== 'ALL') {
    targetRow = sheetBudget.filter(r => String(r['Mês']) === selectedMes);
  }
  let pctBudgetAvg = 0;
  if (targetRow.length > 0) {
    const sumPct = targetRow.reduce((acc, c) => acc + parsePct(c['% do Budget']), 0);
    pctBudgetAvg = sumPct / targetRow.length;
  }
  document.getElementById('kpiBudgetAtingido').textContent = `${pctBudgetAvg.toFixed(1)}%`;

  // KPI 3: % Encomendas Gravadas
  const sheetGravadas = dataStore.analiseCarteira['% de encomendas gravadas '] || [];
  const gravadoRow = sheetGravadas.find(r => String(r['Tipo']).toLowerCase() === 'gravado');
  const pctGrav = gravadoRow ? parsePct(gravadoRow['% gravação']) : 0;
  document.getElementById('kpiPctGravadas').textContent = `${pctGrav.toFixed(1)}%`;

  // KPI 4: Positivação de Carteira
  const sheetPosit = dataStore.analiseCarteira['Ultima fatura'] || [];
  if (sheetPosit.length > 0) {
    const r = sheetPosit[0];
    const pos = r['Qtd_Positivados'] || 0;
    const cart = r['Carteira'] || 0;
    const pctPos = cart > 0 ? ((pos / cart) * 100).toFixed(1) : 0;
    document.getElementById('kpiPositivacao').textContent = `${pos} / ${cart}`;
    document.getElementById('kpiPositivacaoSub').textContent = `${pctPos}% de Clientes Positivados`;
  }
}

// 2. Gráfico: % Budget Atingido por Mês (2026)
function renderChartBudget() {
  const ctx = document.getElementById('chartBudget').getContext('2d');
  const sheet = dataStore.reportSection['% do Budget atingida por '] || [];

  const labels = sheet.map(r => r['Mês'] || '');
  const dataVals = sheet.map(r => parsePct(r['% do Budget']));

  destroyChart('chartBudget');

  charts['chartBudget'] = new Chart(ctx, {
    type: 'line',
    data: {
      labels: labels,
      datasets: [{
        label: '% Budget Atingido',
        data: dataVals,
        borderColor: '#6366f1',
        backgroundColor: 'rgba(99, 102, 241, 0.15)',
        fill: true,
        tension: 0.3,
        pointBackgroundColor: '#6366f1'
      }]
    },
    options: getCommonChartOptions('%')
  });
}

// 3. Gráfico: Volume de Encomendas (Gravado vs Normal)
function renderChartTipoEncomenda() {
  const ctx = document.getElementById('chartTipoEncomenda').getContext('2d');
  const sheet = dataStore.analiseCarteira['Clientes recentes que já '] || [];

  let gravado = 0;
  let normal = 0;

  sheet.forEach(r => {
    const tipo = String(r['Tipo']).toLowerCase();
    const qtd = parseFloat(r['Sum of Valor']) || 0;
    if (tipo.includes('gravad')) gravado += qtd;
    else normal += qtd;
  });

  destroyChart('chartTipoEncomenda');

  charts['chartTipoEncomenda'] = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: ['Gravado', 'Normal / Sem Gravação'],
      datasets: [{
        data: [gravado, normal],
        backgroundColor: ['#10b981', '#ef4444'],
        borderWidth: 0
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: '#94a3b8', font: { family: 'Plus Jakarta Sans' } } }
      }
    }
  });
}

// 4. Gráfico: Top 20 Produtos Mais Pedidos
function renderChartTopProdutos() {
  const ctx = document.getElementById('chartTopProdutos').getContext('2d');
  const sheet = dataStore.reportSection['Image-7'] || [];

  const top20 = sheet.slice(0, 20);
  const labels = top20.map(r => String(r['Produto'] || '').substring(0, 20) + '...');
  const dataVals = top20.map(r => parseFloat(r['Quantidade']) || 0);

  destroyChart('chartTopProdutos');

  charts['chartTopProdutos'] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Quantidade Pedida',
        data: dataVals,
        backgroundColor: '#3b82f6',
        borderRadius: 4
      }]
    },
    options: getCommonChartOptions('')
  });
}

// 5. Gráfico: Vendas por Segmento / Classe
function renderChartSegmentos() {
  const ctx = document.getElementById('chartSegmentos').getContext('2d');
  const sheet = dataStore.reportSection['Vendas (R$) por Cliente'] || [];

  let classeMap = {};
  sheet.forEach(r => {
    const classe = r['Classe'] || 'Outros';
    const val = parseCurrency(r['Valor de venda (R$)']);
    classeMap[classe] = (classeMap[classe] || 0) + val;
  });

  destroyChart('chartSegmentos');

  charts['chartSegmentos'] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: Object.keys(classeMap),
      datasets: [{
        label: 'Vendas (R$)',
        data: Object.values(classeMap),
        backgroundColor: ['#10b981', '#6366f1', '#f59e0b'],
        borderRadius: 6
      }]
    },
    options: getCommonChartOptions('R$')
  });
}

// 6. Tabela: Vendas por Cliente
function renderVendasClienteTable() {
  const tbody = document.getElementById('tbVendasCliente');
  const sheet = dataStore.reportSection['Vendas (R$) por Cliente'] || [];
  const query = document.getElementById('searchClientInput').value.toLowerCase().trim();

  tbody.innerHTML = '';

  const filtered = sheet.filter(r => {
    const nome = String(r['Cliente_Pai'] || '').toLowerCase();
    return nome.includes(query);
  });

  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4" class="empty">Nenhum registro localizado.</td></tr>';
    return;
  }

  filtered.forEach(r => {
    const tr = document.createElement('tr');
    const classe = r['Classe'] || 'Geral';
    const cliente = r['Cliente_Pai'] || '-';
    const valor = parseCurrency(r['Valor de venda (R$)']);
    const pct = r['% do Total'] || '-';

    tr.innerHTML = `
      <td><span class="class-tag ${classe.toLowerCase()}">${classe}</span></td>
      <td><strong>${cliente}</strong></td>
      <td>${formatBRL(valor)}</td>
      <td>${pct}</td>
    `;
    tbody.appendChild(tr);
  });
}

// 7. Tabela: Inatividade e Dias para 1ª Fatura
function renderInatividadeTable() {
  const tbody = document.getElementById('tbInatividade');
  const sheetDias = dataStore.analiseCarteira['#Dias até primeira fatura'] || [];

  tbody.innerHTML = '';

  if (sheetDias.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4" class="empty">Nenhum dado carregado.</td></tr>';
    return;
  }

  sheetDias.forEach(r => {
    const tr = document.createElement('tr');
    const classe = r['Classe'] || 'Pontual';
    const cliente = r['Cliente_Pai'] || '-';
    const ultFat = r['Ultima fat'] ? String(r['Ultima fat']).substring(0, 10) : '-';
    const dias = r['#dias desde a ultima fat'] || 0;

    tr.innerHTML = `
      <td><span class="class-tag ${classe.toLowerCase()}">${classe}</span></td>
      <td><strong>${cliente}</strong></td>
      <td>${ultFat}</td>
      <td><span style="color: ${dias > 60 ? '#ef4444' : '#10b981'}; font-weight: 700;">${dias} dias</span></td>
    `;
    tbody.appendChild(tr);
  });
}

// Funções Auxiliares
function parseCurrency(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  const clean = String(val).replace('R$', '').replace(/\./g, '').replace(',', '.').trim();
  return parseFloat(clean) || 0;
}

function parsePct(val) {
  if (typeof val === 'number') return val > 1 ? val : val * 100;
  if (!val) return 0;
  const clean = String(val).replace('%', '').replace(',', '.').trim();
  const num = parseFloat(clean) || 0;
  return num > 1 ? num : num * 100;
}

function formatBRL(val) {
  return val.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function destroyChart(chartId) {
  if (charts[chartId]) {
    charts[chartId].destroy();
  }
}

function getCommonChartOptions(unit) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { display: false }
    },
    scales: {
      x: {
        grid: { color: '#1a1e2c' },
        ticks: { color: '#64748b', font: { family: 'Plus Jakarta Sans', size: 10 } }
      },
      y: {
        grid: { color: '#1a1e2c' },
        ticks: {
          color: '#64748b',
          font: { family: 'Plus Jakarta Sans' },
          callback: function(value) { return value + ' ' + unit; }
        }
      }
    }
  };
}
