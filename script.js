// Inicializa ícones do Lucide
lucide.createIcons();

// Repositório global das planilhas importadas
let dataStore = {
  reportSection: {},
  analiseCarteira: {}
};

// Guarda as instâncias dos gráficos Chart.js para destruição/redesenho
let charts = {};

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

// Event Listeners de Carga de Arquivos
fileInput1.addEventListener('change', (e) => {
  if (e.target.files.length > 0) readExcelFile(e.target.files[0], 1);
});

fileInput2.addEventListener('change', (e) => {
  if (e.target.files.length > 0) readExcelFile(e.target.files[0], 2);
});

// Leitura de Arquivo via SheetJS
function readExcelFile(file, fileNum) {
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array', cellDates: true });
      
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
      console.error("Erro ao ler o Excel:", err);
      alert(`Erro ao processar o Arquivo ${fileNum}. Verifique o formato do arquivo.`);
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

// Event Listeners de Filtros
selectAno.addEventListener('change', renderDashboard);
selectMes.addEventListener('change', renderDashboard);
document.getElementById('searchClientInput').addEventListener('input', renderVendasClienteTable);

// Função Mestra de Renderização
function renderDashboard() {
  renderKPIs();
  renderChartBudget();
  renderChartTipoEncomenda();
  renderChartTopProdutos();
  renderChartSegmentos();
  renderVendasClienteTable();
  renderInatividadeTable();
}

// 1. Renderização dos Cards KPIs
function renderKPIs() {
  // KPI 1: Faturamento Mensal Carteira
  const sheetVendaMensal = dataStore.reportSection['Venda mensal em reais da '] || [];
  let totalVenda = 0;
  sheetVendaMensal.forEach(r => {
    totalVenda += parseCurrency(r['After_Tax_Amount']);
  });
  document.getElementById('kpiValorMensal').textContent = formatBRL(totalVenda);

  // KPI 2: Atingimento do Budget
  const sheetBudget = dataStore.reportSection['% do Budget atingida por '] || [];
  const selectedMes = selectMes.value;
  let rows = sheetBudget;
  if (selectedMes !== 'ALL') {
    rows = sheetBudget.filter(r => String(r['Mês']) === selectedMes);
  }
  let pctBudgetAvg = 0;
  if (rows.length > 0) {
    const validPcts = rows.map(r => parsePct(r['% do Budget'])).filter(v => !isNaN(v) && v > 0);
    if (validPcts.length > 0) {
      pctBudgetAvg = validPcts.reduce((a, b) => a + b, 0) / validPcts.length;
    }
  }
  document.getElementById('kpiBudgetAtingido').textContent = `${pctBudgetAvg.toFixed(1)}%`;

  // KPI 3: % Encomendas Gravadas
  const sheetGravadas = dataStore.analiseCarteira['% de encomendas gravadas '] || [];
  const gravadoRow = sheetGravadas.find(r => String(r['Tipo']).toLowerCase() === 'gravado');
  const pctGrav = gravadoRow ? parsePct(gravadoRow['% gravação']) : 0;
  document.getElementById('kpiPctGravadas').textContent = `${pctGrav.toFixed(1)}%`;

  // KPI 4: Positivação da Carteira
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

// 2. Gráfico: % Budget Atingido por Mês
function renderChartBudget() {
  const ctx = document.getElementById('chartBudget').getContext('2d');
  const sheet = dataStore.reportSection['% do Budget atingida por '] || [];

  const labels = sheet.map(r => `Mês ${r['Mês'] || ''}`);
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
        pointBackgroundColor: '#6366f1',
        pointRadius: 4
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
  const labels = top20.map(r => String(r['Produto'] || '').substring(0, 18) + '...');
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

// 5. Gráfico: Vendas (R$) por Classe / Segmento
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
        backgroundColor: ['#10b981', '#6366f1', '#f59e0b', '#3b82f6'],
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

// 7. Tabela: Recência de Compra e Dias até 1ª Fatura
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
    
    // Tratamento seguro de objeto de Data do Excel
    let ultFat = '-';
    if (r['Ultima fat']) {
      const d = new Date(r['Ultima fat']);
      ultFat = !isNaN(d.getTime()) ? d.toLocaleDateString('pt-BR') : String(r['Ultima fat']);
    }

    const dias = parseInt(r['#dias desde a ultima fat']) || 0;

    tr.innerHTML = `
      <td><span class="class-tag ${classe.toLowerCase()}">${classe}</span></td>
      <td><strong>${cliente}</strong></td>
      <td>${ultFat}</td>
      <td><span style="color: ${dias > 60 ? '#ef4444' : '#10b981'}; font-weight: 700;">${dias} dias</span></td>
    `;
    tbody.appendChild(tr);
  });
}

// Functions Auxiliares Truncadas e Seguras

function parseCurrency(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  let str = String(val).replace('R$', '').trim();
  if (str.includes(',')) {
    str = str.replace(/\./g, '').replace(',', '.');
  }
  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

function parsePct(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val > 1 ? val : val * 100;
  const clean = String(val).replace('%', '').replace(',', '.').trim();
  const num = parseFloat(clean);
  if (isNaN(num)) return 0;
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
          callback: function(value) { return unit === 'R$' ? 'R$ ' + value.toLocaleString('pt-BR') : value + ' ' + unit; }
        }
      }
    }
  };
}
