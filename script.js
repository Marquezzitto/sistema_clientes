// Inicializa ícones do Lucide
lucide.createIcons();

// Oculta a tela de carregamento do Power BI quando o Iframe carrega
const iframe = document.getElementById('powerbiFrame');
const overlay = document.getElementById('loadingOverlay');

if (iframe && overlay) {
  iframe.addEventListener('load', () => {
    overlay.classList.add('hidden');
  });
}

// Repositório global de planilhas locais (se houver upload)
let dataStore = {
  reportSection: {},
  analiseCarteira: {}
};

let charts = {};

// Elementos do DOM
const fileInput1 = document.getElementById('fileInput1');
const fileInput2 = document.getElementById('fileInput2');
const dropZone1 = document.getElementById('dropZone1');
const dropZone2 = document.getElementById('dropZone2');
const labelFile1 = document.getElementById('labelFile1');
const labelFile2 = document.getElementById('labelFile2');

const selectAno = document.getElementById('selectAno');
const selectMes = document.getElementById('selectMes');

if (fileInput1) fileInput1.addEventListener('change', (e) => e.target.files.length > 0 && readExcelFile(e.target.files[0], 1));
if (fileInput2) fileInput2.addEventListener('change', (e) => e.target.files.length > 0 && readExcelFile(e.target.files[0], 2));

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

      renderDashboard();
    } catch (err) {
      console.error("Erro ao ler Excel:", err);
    }
  };
  reader.readAsArrayBuffer(file);
}

if (selectAno) selectAno.addEventListener('change', renderDashboard);
if (selectMes) selectMes.addEventListener('change', renderDashboard);

function getSheetData(dataObject, targetName) {
  if (!dataObject) return [];
  const normalizedTarget = targetName.trim().toLowerCase();
  const key = Object.keys(dataObject).find(
    k => k.trim().toLowerCase() === normalizedTarget || k.trim().toLowerCase().includes(normalizedTarget)
  );
  return key ? dataObject[key] : [];
}

function filterDataByPeriod(rows) {
  if (!rows || rows.length === 0) return [];
  const anoSel = selectAno ? selectAno.value : 'ALL';
  const mesSel = selectMes ? selectMes.value : 'ALL';

  return rows.filter(r => {
    let matchAno = true;
    let matchMes = true;

    if (anoSel !== 'ALL') {
      const anoKey = Object.keys(r).find(k => k.toLowerCase().includes('ano') || k.toLowerCase().includes('data'));
      if (anoKey && r[anoKey]) matchAno = String(r[anoKey]).includes(anoSel);
    }

    if (mesSel !== 'ALL') {
      const mesKey = Object.keys(r).find(k => k.toLowerCase().includes('mês') || k.toLowerCase().includes('mes'));
      if (mesKey && r[mesKey] !== undefined && r[mesKey] !== "") {
        const valMes = String(r[mesKey]).trim();
        matchMes = parseInt(valMes, 10) === parseInt(mesSel, 10) || valMes.toLowerCase().includes(mesSel.toLowerCase());
      }
    }

    return matchAno && matchMes;
  });
}

function renderDashboard() {
  renderKPIs();
  renderChartBudget();
  renderChartTipoEncomenda();
  renderChartTopProdutos();
  renderChartSegmentos();
  renderVendasClienteTable();
  renderInatividadeTable();
}

function renderKPIs() {
  const sheetVenda = filterDataByPeriod(getSheetData(dataStore.reportSection, 'Venda mensal em reais da'));
  let totalVenda = 0;
  sheetVenda.forEach(r => {
    const valKey = Object.keys(r).find(k => k.toLowerCase().includes('after_tax') || k.toLowerCase().includes('valor'));
    if (valKey) totalVenda += parseCurrency(r[valKey]);
  });
  const elVenda = document.getElementById('kpiValorMensal');
  if (elVenda) elVenda.textContent = formatBRL(totalVenda);
}

function renderChartBudget() {
  const ctx = document.getElementById('chartBudget');
  if (!ctx) return;

  const sheet = filterDataByPeriod(getSheetData(dataStore.reportSection, '% do Budget atingida por'));
  const labels = sheet.map(r => `Mês ${r['Mês'] || ''}`);
  const dataVals = sheet.map(r => parsePct(r['% do Budget']));

  destroyChart('chartBudget');

  charts['chartBudget'] = new Chart(ctx.getContext('2d'), {
    type: 'line',
    data: {
      labels: labels,
      datasets: [{
        label: '% Budget Atingido',
        data: dataVals,
        borderColor: '#6366f1',
        backgroundColor: 'rgba(99, 102, 241, 0.15)',
        fill: true,
        tension: 0.3
      }]
    },
    options: getCommonChartOptions('%')
  });
}

function renderChartTipoEncomenda() {
  const ctx = document.getElementById('chartTipoEncomenda');
  if (!ctx) return;

  const sheet = filterDataByPeriod(getSheetData(dataStore.analiseCarteira, 'Clientes recentes que já'));
  let gravado = 0, normal = 0;

  sheet.forEach(r => {
    const strRow = JSON.stringify(r).toLowerCase();
    const valKey = Object.keys(r).find(k => k.toLowerCase().includes('valor') || k.toLowerCase().includes('qtd'));
    const qtd = valKey ? parseFloat(r[valKey]) || 0 : 0;
    if (strRow.includes('gravad')) gravado += qtd;
    else normal += qtd;
  });

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

function renderChartTopProdutos() {
  const ctx = document.getElementById('chartTopProdutos');
  if (!ctx) return;

  const sheet = filterDataByPeriod(getSheetData(dataStore.reportSection, 'Image-7')).slice(0, 20);
  const labels = sheet.map(r => String(r['Produto'] || '').substring(0, 15) + '...');
  const dataVals = sheet.map(r => parseFloat(r['Quantidade']) || 0);

  destroyChart('chartTopProdutos');

  charts['chartTopProdutos'] = new Chart(ctx.getContext('2d'), {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Qtd Pedida',
        data: dataVals,
        backgroundColor: '#3b82f6',
        borderRadius: 4
      }]
    },
    options: getCommonChartOptions('')
  });
}

function renderChartSegmentos() {
  const ctx = document.getElementById('chartSegmentos');
  if (!ctx) return;

  const sheet = filterDataByPeriod(getSheetData(dataStore.reportSection, 'Vendas (R$) por Cliente'));
  let classeMap = {};

  sheet.forEach(r => {
    const classe = r['Classe'] || 'Outros';
    const val = parseCurrency(r['Valor de venda (R$)']);
    classeMap[classe] = (classeMap[classe] || 0) + val;
  });

  destroyChart('chartSegmentos');

  charts['chartSegmentos'] = new Chart(ctx.getContext('2d'), {
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

function renderVendasClienteTable() {
  const tbody = document.getElementById('tbVendasCliente');
  if (!tbody) return;

  const sheet = filterDataByPeriod(getSheetData(dataStore.reportSection, 'Vendas (R$) por Cliente'));
  const query = document.getElementById('searchClientInput') ? document.getElementById('searchClientInput').value.toLowerCase().trim() : '';

  tbody.innerHTML = '';
  const filtered = sheet.filter(r => String(r['Cliente_Pai'] || '').toLowerCase().includes(query));

  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4" class="empty">Nenhum registro localizado.</td></tr>';
    return;
  }

  filtered.forEach(r => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><span class="class-tag">${r['Classe'] || 'Geral'}</span></td>
      <td><strong>${r['Cliente_Pai'] || '-'}</strong></td>
      <td>${formatBRL(parseCurrency(r['Valor de venda (R$)']))}</td>
      <td>${r['% do Total'] || '-'}</td>
    `;
    tbody.appendChild(tr);
  });
}

function renderInatividadeTable() {
  const tbody = document.getElementById('tbInatividade');
  if (!tbody) return;

  const sheet = filterDataByPeriod(getSheetData(dataStore.analiseCarteira, '#Dias até primeira fatura'));
  tbody.innerHTML = '';

  if (sheet.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4" class="empty">Nenhum dado carregado.</td></tr>';
    return;
  }

  sheet.forEach(r => {
    const tr = document.createElement('tr');
    const dias = parseInt(r['#dias desde a ultima fat']) || 0;
    tr.innerHTML = `
      <td><span class="class-tag">${r['Classe'] || 'Pontual'}</span></td>
      <td><strong>${r['Cliente_Pai'] || '-'}</strong></td>
      <td>${r['Ultima fat'] || '-'}</td>
      <td><span style="color: ${dias > 60 ? '#ef4444' : '#10b981'}; font-weight: 700;">${dias} dias</span></td>
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
    plugins: { legend: { display: false } },
    scales: {
      x: { grid: { color: '#1a1e2c' }, ticks: { color: '#64748b' } },
      y: { grid: { color: '#1a1e2c' }, ticks: { color: '#64748b' } }
    }
  };
}
