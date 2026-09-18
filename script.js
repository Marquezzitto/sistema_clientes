// Ativa ícones Lucide
if (typeof lucide !== 'undefined') {
  lucide.createIcons();
}

let dataStore = {
  reportSection: {},
  analiseCarteira: {}
};

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
      requestAnimationFrame(() => renderDashboard());
    }
    return temDados;
  } catch (err) {
    console.error('Não foi possível restaurar a sessão das planilhas:', err);
    return false;
  }
}

restaurarSessaoAtual();

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
    } catch (err) {
      console.error("Erro ao ler planilha:", err);
    }
  };
  reader.readAsArrayBuffer(file);
}

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

function diasUteisRestantesNoMes(dataRef = new Date()) {
  const ano = dataRef.getFullYear();
  const mes = dataRef.getMonth();
  const ultimoDia = new Date(ano, mes + 1, 0).getDate();
  let count = 0;
  for (let dia = dataRef.getDate(); dia <= ultimoDia; dia++) {
    const diaSemana = new Date(ano, mes, dia).getDay(); 
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

  if (elLytd) elLytd.textContent = formatBRL(lytd || 6386614.16);
  if (elYtd) elYtd.textContent = formatBRL(ytd || 6636963.60);
  if (elVar) elVar.textContent = formatBRL(variacao || 250349.43);
}

function renderKPIs() {
  const mesSel = selectMes ? selectMes.value : 'ALL';
  const anoSel = selectAno ? selectAno.value : '2026';
  const clienteSel = selectCliente ? selectCliente.value : 'ALL';
  
  const anoNum = Number(anoSel);
  const mesNum = mesSel !== 'ALL' ? parseInt(mesSel, 10) : null;

  let totalFat = 0;
  const sheetCliente = getSheet(dataStore.reportSection, ['Vendas (R$) por Cliente', 'Cliente']);

  if (clienteSel !== 'ALL') {
    sheetCliente.forEach(r => {
      const nome = String(r['Cliente_Pai'] || r['Cliente'] || '').trim();
      if (nome.toUpperCase() === clienteSel.toUpperCase()) {
        totalFat += parseCurrency(r['Valor de venda (R$)'] || r['Valor']);
      }
    });
  } else {
    sheetCliente.forEach(r => totalFat += parseCurrency(r['Valor de venda (R$)'] || r['Valor']));
    if (totalFat === 0) {
      const sheetVenda = getSheet(dataStore.reportSection, ['Venda Mensal em Reais', 'Venda Mensal', 'Image-6']);
      sheetVenda.forEach(r => totalFat += parseCurrency(r[anoSel] || r[`${anoSel} (R$)`] || r['Vendas (R$)']));
    }
  }

  const elValor = document.getElementById('kpiValorMensal');
  if (elValor) elValor.textContent = formatBRL(totalFat || 32766640.70);

  const sheetBudget = getSheet(dataStore.reportSection, ['% do Budget atingida por', '% Budget Atingida', 'Budget']);
  let avgBudget = 0;

  if (sheetBudget.length > 0) {
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

  const sheetUltimaFatura = getSheet(dataStore.analiseCarteira, ['Ultima fatura', 'Última fatura']);
  let positivados = 82;
  let totalCarteira = 271;
  let metaQtdPlanilha = null;

  if (sheetUltimaFatura && sheetUltimaFatura.length > 0) {
    const row = sheetUltimaFatura[0];
    const valPos = parseCurrency(row['Qtd_Positivados'] || row['Qtd_Positivado']);
    const valCart = parseCurrency(row['Carteira']);
    const valMeta = parseCurrency(row['Meta']);
    if (valPos > 0) positivados = valPos;
    if (valCart > 0) totalCarteira = valCart;
    if (valMeta > 0) metaQtdPlanilha = valMeta;
  }

  const metaQtd = metaQtdPlanilha !== null ? metaQtdPlanilha : Math.round(totalCarteira * 0.60);
  const metaPct = totalCarteira > 0 ? (metaQtd / totalCarteira) * 100 : 60;
  const realPct = (positivados / totalCarteira) * 100;
  const faltaQtd = metaQtd - positivados;
  const faltaPct = metaPct - realPct;

  const elPos = document.getElementById('kpiPositivacao');
  const elPosSub = document.getElementById('kpiPositivacaoSub');

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

  const kpiCardPositivacao = elPosSub ? elPosSub.closest('.kpi-card') : null;
  if (kpiCardPositivacao) {
    if (alertaPositivacao) {
      kpiCardPositivacao.style.borderColor = '#ef4444';
      kpiCardPositivacao.style.boxShadow = '0 0 0 1px rgba(239,68,68,0.45)';
    } else {
      kpiCardPositivacao.style.borderColor = '';
      kpiCardPositivacao.style.boxShadow = '';
    }
  }

  const sheetGravados = getSheet(dataStore.analiseCarteira, ['% de encomendas gravadas', 'Encomendas Gravadas', 'Gravado']);
  let pctVal = 44.72;
  const metaGravaçãoPct = 60.0;

  if (sheetGravados.length > 0) {
    let row = null;
    if (mesNum !== null) {
      row = sheetGravados.find(r => Number(r.Ano) === anoNum && Number(r.Mes) === mesNum && String(r.Tipo).trim().toLowerCase() === 'gravado');
    }
    if (!row) row = sheetGravados.find(r => String(r['Ano'] || r['Tipo']).toLowerCase().includes('total')) || sheetGravados[0];

    const rawVal = row ? (row['% gravação'] || row['% Gravado'] || row['Total'] || row['Gravado'] || 0.4472) : 0.4472;
    pctVal = parsePct(rawVal);
  }

  const elGrav = document.getElementById('kpiPctGravadas');
  const elGravSub = document.getElementById('kpiPctGravadasSub');
  if (elGrav) elGrav.textContent = `${pctVal.toFixed(2)}%`;
  
  if (elGravSub) {
    const diffGrav = metaGravaçãoPct - pctVal;
    if (diffGrav <= 0) {
      elGravSub.innerHTML = `Meta: ${metaGravaçãoPct}% | <span style="color:#10b981;font-weight:bold;">Meta Atingida!</span>`;
    } else {
      elGravSub.innerHTML = `Meta: ${metaGravaçãoPct}% | Falta: <span style="color:#f59e0b;font-weight:bold;">${diffGrav.toFixed(2)}%</span> p/ a meta`;
    }
  }
}

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

  const sheet = getSheet(dataStore.reportSection, ['Image-6', 'Venda Mensal em Reais', 'Venda Mensal']);
  const meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
  const clienteSel = selectCliente ? selectCliente.value : 'ALL';
  
  let v2024 = new Array(12).fill(null);
  let v2025 = new Array(12).fill(null);
  let v2026 = new Array(12).fill(null);

  if (sheet.length > 0) {
    sheet.forEach(r => {
      // Filtra pelo cliente (ignora maiúsculas/minúsculas)
      if (clienteSel !== 'ALL') {
        const nome = String(r['Cliente_Pai'] || r['Cliente'] || '').trim().toUpperCase();
        if (!nome || nome !== clienteSel.toUpperCase()) return;
      }

      const idx = Number(r['Mês'] || r['Mes']) - 1;
      if (idx >= 0 && idx < 12) {
        
        // CORREÇÃO AQUI: Em vez de substituir (=), ele soma (+=) caso o cliente tenha mais de um registro no mesmo mês.
        // E também valida o formato da planilha (se colunas são os anos, ou se tem uma coluna 'Ano')
        if (r['2026'] !== undefined || r['2025'] !== undefined || r['2026 (R$)'] !== undefined) {
          const val24 = parseCurrency(r['2024'] || r['2024 (R$)']);
          const val25 = parseCurrency(r['2025'] || r['2025 (R$)']);
          const val26 = parseCurrency(r['2026'] || r['2026 (R$)']);
          
          if (val24 > 0) v2024[idx] = (v2024[idx] || 0) + val24;
          if (val25 > 0) v2025[idx] = (v2025[idx] || 0) + val25;
          if (val26 > 0) v2026[idx] = (v2026[idx] || 0) + val26;
        } else {
          const anoRow = Number(r['Ano']);
          const val = parseCurrency(r['Vendas (R$)'] || r['Vendas'] || r['Valor']);
          
          if (val > 0) {
            if (anoRow === 2024) v2024[idx] = (v2024[idx] || 0) + val;
            if (anoRow === 2025) v2025[idx] = (v2025[idx] || 0) + val;
            if (anoRow === 2026) v2026[idx] = (v2026[idx] || 0) + val;
          }
        }
      }
    });
  }

  destroyChart('chartHistoricoFaturamento');
  charts['chartHistoricoFaturamento'] = new Chart(ctx.getContext('2d'), {
    type: 'line',
    data: {
      labels: meses,
      datasets: [
        { label: '2026', data: v2026, borderColor: '#10b981', backgroundColor: 'rgba(16,185,129,0.1)', borderWidth: 3, fill: true, spanGaps: true },
        { label: '2025', data: v2025, borderColor: '#6366f1', backgroundColor: 'transparent', borderWidth: 2, spanGaps: true },
        { label: '2024', data: v2024, borderColor: '#94a3b8', backgroundColor: 'transparent', borderWidth: 1, spanGaps: true }
      ]
    },
    plugins: [pluginValoresNativos],
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: true, labels: { color: '#94a3b8' } } },
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

  const sheet = getSheet(dataStore.reportSection, ['% do Budget atingida por', '% Budget Atingida', 'Budget']);
  const labels = mesesArray();
  const clienteSel = selectCliente ? selectCliente.value : 'ALL';
  let dataVals = new Array(12).fill(0);

  if (sheet.length > 0) {
    sheet.forEach(r => {
      if (clienteSel !== 'ALL') {
        const nome = String(r['Cliente_Pai'] || r['Cliente'] || '').trim().toUpperCase();
        if (!nome || nome !== clienteSel.toUpperCase()) return;
      }

      const idx = Number(r['Mês'] || r['Mes']) - 1;
      if (idx >= 0 && idx < 12) {
        const val = parsePct(r['% do Budget'] || r['Budget']);
        if (val > 0) dataVals[idx] = val; // Consideramos o último Budget lançado ou o único
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
      plugins: { legend: { display: true, labels: { color: '#94a3b8' } } },
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

  const sheet = getSheet(dataStore.analiseCarteira, ['Clientes recentes que já', '% de encomendas gravadas', 'Encomenda por Tipo']);
  const clienteSel = selectCliente ? selectCliente.value : 'ALL';
  
  let gravado = 35.16; 
  let normal = 64.84;

  if (sheet.length > 0) {
    let rowG = sheet.find(r => String(r.Tipo).toLowerCase().includes('gravado'));
    let rowN = sheet.find(r => String(r.Tipo).toLowerCase().includes('normal'));

    if (clienteSel !== 'ALL') {
      const cRowG = sheet.find(r => String(r.Tipo).toLowerCase().includes('gravado') && String(r['Cliente_Pai'] || r['Cliente']).trim().toUpperCase() === clienteSel.toUpperCase());
      const cRowN = sheet.find(r => String(r.Tipo).toLowerCase().includes('normal') && String(r['Cliente_Pai'] || r['Cliente']).trim().toUpperCase() === clienteSel.toUpperCase());
      if (cRowG || cRowN) {
        rowG = cRowG;
        rowN = cRowN;
      } else {
        rowG = null; rowN = null; gravado = 0; normal = 0;
      }
    }

    if (rowG) gravado = parsePct(rowG['Sum of Valor'] || rowG['% gravação'] || rowG['Gravado']);
    if (rowN) normal = parsePct(rowN['Sum of Valor'] || rowN['% gravação'] || rowN['Normal']);
  }

  destroyChart('chartTipoEncomenda');
  charts['chartTipoEncomenda'] = new Chart(ctx.getContext('2d'), {
    type: 'doughnut',
    data: {
      labels: ['Gravado (Meta: 60%)', 'Normal'],
      datasets: [{
        data: [gravado, normal],
        backgroundColor: ['#10b981', '#ef4444'],
        borderWidth: 0
      }]
    },
    plugins: [pluginValoresNativos],
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: true, labels: { color: '#94a3b8' } } }
    }
  });
}

function renderChartSegmentos() {
  const ctx = document.getElementById('chartSegmentos');
  if (!ctx) return;

  const sheet = getSheet(dataStore.reportSection, ['Venda mensal em reais da', 'Separador Segmento', 'Segmento']);
  const clienteSel = selectCliente ? selectCliente.value : 'ALL';
  
  let filteredSheet = sheet;
  if (clienteSel !== 'ALL') {
    filteredSheet = sheet.filter(r => {
      const nome = String(r['Cliente_Pai'] || r['Cliente'] || '').trim().toUpperCase();
      return nome === clienteSel.toUpperCase();
    });
  }

  // CORREÇÃO: Agrupa e soma os valores caso haja mais de um registro do mesmo segmento
  let grupos = {};
  filteredSheet.forEach(r => {
    const seg = String(r['Separador'] || r['Segmento'] || 'Outros').trim();
    const val = parseCurrency(r['After_Tax_Amount'] || r['Valor'] || r['Vendas (R$)']);
    if (val > 0) {
      grupos[seg] = (grupos[seg] || 0) + val;
    }
  });

  const sorted = Object.entries(grupos).sort((a,b) => b[1] - a[1]).slice(0, 6);
  const labels = sorted.map(x => x[0]);
  const dataVals = sorted.map(x => x[1]);

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
      plugins: { legend: { display: true, labels: { color: '#94a3b8' } } },
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

  const sheet = getSheet(dataStore.reportSection, ['Image-7', 'Top 20 Produtos Mais Vendidos', 'Top 20']);
  const clienteSel = selectCliente ? selectCliente.value : 'ALL';
  
  let filteredSheet = sheet;
  if (clienteSel !== 'ALL') {
    filteredSheet = sheet.filter(r => {
      const nome = String(r['Cliente_Pai'] || r['Cliente'] || '').trim().toUpperCase();
      return nome === clienteSel.toUpperCase();
    });
  }

  // CORREÇÃO: Agrupa e soma os valores caso haja mais de um registro do mesmo produto
  let grupos = {};
  filteredSheet.forEach(r => {
    const prod = String(r['Produto'] || r['Cod'] || 'Desconhecido').trim();
    const val = parseCurrency(r['Valor de Venda'] || r['Valor de Venda (R$)'] || r['Valor']);
    if (val > 0) {
      grupos[prod] = (grupos[prod] || 0) + val;
    }
  });

  const sorted = Object.entries(grupos).sort((a,b) => b[1] - a[1]).slice(0, 5);
  const labels = sorted.map(x => x[0]);
  const dataVals = sorted.map(x => x[1]);

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
      plugins: { legend: { display: true, labels: { color: '#94a3b8' } } },
      scales: {
        x: { grid: { color: '#1f293d' }, ticks: { color: '#94a3b8' } },
        y: { beginAtZero: true, grid: { color: '#1f293d' }, ticks: { color: '#94a3b8' } }
      }
    }
  });
}

function renderVendasClienteTable() {
  const tbody = document.getElementById('tbVendasCliente');
  if (!tbody) return;

  const sheet = getSheet(dataStore.reportSection, ['Vendas (R$) por Cliente', 'Cliente']);
  const query = searchInput ? searchInput.value.toLowerCase().trim() : '';
  const clienteSel = selectCliente ? selectCliente.value : 'ALL';

  tbody.innerHTML = '';
  const filtered = sheet.filter(r => {
    const nome = String(r['Cliente_Pai'] || r['Cliente'] || '').trim();
    const matchBusca = nome.toLowerCase().includes(query);
    const matchFiltroTopo = clienteSel === 'ALL' || nome.toUpperCase() === clienteSel.toUpperCase();
    return matchBusca && matchFiltroTopo;
  });

  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4" class="empty-row">Nenhum registro localizado.</td></tr>';
    return;
  }

  filtered.forEach(r => {
    const clientName = r['Cliente_Pai'] || r['Cliente'] || '-';
    const valorVenda = parseCurrency(r['Valor de venda (R$)'] || r['Valor']);
    const pctTotal = r['% do Total'] || '-';
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
    const matchFiltroTopo = clienteSel === 'ALL' || nome.toUpperCase() === clienteSel.toUpperCase();
    const matchBusca = nome.toLowerCase().includes(query);
    return matchFiltroTopo && matchBusca;
  });

  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4" class="empty-row">Aguardando dados da planilha...</td></tr>';
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
  const canvas = document.getElementById(chartId);
  if (canvas && typeof Chart !== 'undefined' && typeof Chart.getChart === 'function') {
    const existente = Chart.getChart(canvas);
    if (existente) existente.destroy();
  }
}
