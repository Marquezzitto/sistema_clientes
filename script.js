// ===== RCA 61 - script.js completo corrigido =====

if (typeof lucide !== 'undefined') lucide.createIcons();

let dataStore = {
  reportSection: {},
  analiseCarteira: {}
};

let dadosFixosMensais = null;
let sincronizacaoDriveEmAndamento = false;
let charts = {};

const DRIVE_SHEET_IDS = {
  cliente: '1AP_60koNw2moYQfoJbiXwepVl0EGgcb0',
  segmento: '1z2Xt-nouxE5JapG-pGTZTkjjKca1mW7k',
  produto: '1lHcnnMJHKzlBt7El5GQ-qktMx897o0pR'
};

const MESES_FIXOS = [1, 2, 3, 4, 5, 6, 7, 8];

const NOMES_MESES = [
  'Janeiro',
  'Fevereiro',
  'Março',
  'Abril',
  'Maio',
  'Junho',
  'Julho',
  'Agosto',
  'Setembro',
  'Outubro',
  'Novembro',
  'Dezembro'
];

const LABELS_MESES = [
  'Jan',
  'Fev',
  'Mar',
  'Abr',
  'Mai',
  'Jun',
  'Jul',
  'Ago',
  'Set',
  'Out',
  'Nov',
  'Dez'
];

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

function parseCurrency(val) {
  if (typeof val === 'number') {
    return Number.isFinite(val) ? val : 0;
  }

  if (val === null || val === undefined || val === '') {
    return 0;
  }

  let s = String(val)
    .replace(/R\$/g, '')
    .replace(/\s/g, '')
    .trim();

  if (s.includes(',')) {
    s = s
      .replace(/\./g, '')
      .replace(',', '.');
  }

  const n = Number.parseFloat(s);

  return Number.isFinite(n) ? n : 0;
}

function parsePct(val) {
  if (
    val === null ||
    val === undefined ||
    val === ''
  ) {
    return 0;
  }

  if (typeof val === 'number') {
    return val <= 1 ? val * 100 : val;
  }

  const s = String(val)
    .replace('%', '')
    .replace(',', '.')
    .trim();

  const n = Number.parseFloat(s);

  if (!Number.isFinite(n)) {
    return 0;
  }

  return n <= 1 && s.includes('.')
    ? n * 100
    : n;
}

function formatBRL(val) {
  return Number(val || 0).toLocaleString(
    'pt-BR',
    {
      style: 'currency',
      currency: 'BRL'
    }
  );
}

function formatCompactBRL(val) {
  const n = Number(val || 0);
  const abs = Math.abs(n);
  const sign = n < 0 ? '-' : '';

  if (abs >= 1000000) {
    return `${sign}R$ ${(abs / 1000000)
      .toFixed(1)
      .replace('.', ',')} mi`;
  }

  if (abs >= 1000) {
    return `${sign}R$ ${(abs / 1000)
      .toFixed(1)
      .replace('.', ',')}k`;
  }

  return `${sign}R$ ${abs.toFixed(0)}`;
}

function formatDate(val) {
  if (!val) return '-';

  if (val instanceof Date) {
    return val.toISOString().split('T')[0];
  }

  return String(val).split('T')[0];
}

function getSheet(dataObj, keywords) {
  if (!dataObj) return [];

  const names = Object.keys(dataObj);

  for (const kw of keywords) {
    const found = names.find(
      n =>
        n
          .toLowerCase()
          .includes(kw.toLowerCase())
    );

    if (found) {
      return Array.isArray(dataObj[found])
        ? dataObj[found]
        : [];
    }
  }

  return [];
}

function algumaPlanilhaCarregada() {
  return (
    (
      dataStore.reportSection &&
      Object.keys(dataStore.reportSection).length > 0
    ) ||
    (
      dataStore.analiseCarteira &&
      Object.keys(dataStore.analiseCarteira).length > 0
    )
  );
}

function sheetTemColunaCliente(sheet) {
  return !!(
    sheet &&
    sheet.length &&
    (
      Object.prototype.hasOwnProperty.call(
        sheet[0],
        'Cliente_Pai'
      ) ||
      Object.prototype.hasOwnProperty.call(
        sheet[0],
        'Cliente'
      )
    )
  );
}

function filtrarPorCliente(
  sheet,
  cliente
) {
  const rows = Array.isArray(sheet)
    ? sheet
    : [];

  if (!cliente || cliente === 'ALL') {
    return {
      linhas: rows,
      semDetalhePorCliente: false
    };
  }

  if (!sheetTemColunaCliente(rows)) {
    return {
      linhas: [],
      semDetalhePorCliente: true
    };
  }

  const alvo = String(cliente).trim();

  return {
    linhas: rows.filter(r =>
      String(
        r['Cliente_Pai'] ||
        r['Cliente'] ||
        ''
      ).trim() === alvo
    ),
    semDetalhePorCliente: false
  };
}

function diasUteisRestantesNoMes(
  dataRef = new Date()
) {
  const ano = dataRef.getFullYear();
  const mes = dataRef.getMonth();

  const ultimo =
    new Date(
      ano,
      mes + 1,
      0
    ).getDate();

  let total = 0;

  for (
    let d = dataRef.getDate();
    d <= ultimo;
    d++
  ) {
    const dow =
      new Date(
        ano,
        mes,
        d
      ).getDay();

    if (
      dow !== 0 &&
      dow !== 6
    ) {
      total++;
    }
  }

  return total;
}

function normalizarNome(str) {
  return String(str || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/[^A-Z0-9 ]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizarChave(str) {
  const s =
    String(str || '')
      .trim();

  const cod =
    s.match(/^(\d+)\s*-/);

  if (cod) {
    return cod[1];
  }

  return normalizarNome(s);
}

function destruirGrafico(id) {
  if (charts[id]) {
    charts[id].destroy();
    charts[id] = null;
  }

  const canvas =
    document.getElementById(id);

  if (
    canvas &&
    typeof Chart !== 'undefined' &&
    Chart.getChart
  ) {
    const chart =
      Chart.getChart(canvas);

    if (chart) {
      chart.destroy();
    }
  }
}

function linhaVazia(row) {
  return (
    !row ||
    row.every(
      v =>
        v === null ||
        v === undefined ||
        v === ''
    )
  );
}

function encontrarAbaMes(
  workbook,
  mes
) {
  const alvo =
    `MES ${mes}`;

  return workbook.SheetNames.includes(
    alvo
  )
    ? alvo
    : workbook.SheetNames.find(
        n => n.trim() === alvo
      ) || null;
}

async function buscarWorkbookDrive(
  fileId
) {
  const url =
    `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`;

  const res =
    await fetch(url);

  if (!res.ok) {
    throw new Error(
      `HTTP ${res.status}`
    );
  }

  const buf =
    await res.arrayBuffer();

  return XLSX.read(
    new Uint8Array(buf),
    {
      type: 'array',
      cellDates: true
    }
  );
}

function extrairPorCliente(
  workbook
) {
  const out = {};

  MESES_FIXOS.forEach(
    mes => {
      const aba =
        encontrarAbaMes(
          workbook,
          mes
        );

      if (!aba) return;

      const raw =
        XLSX.utils.sheet_to_json(
          workbook.Sheets[aba],
          {
            header: 1,
            defval: null
          }
        );

      let classeAtual = null;

      const linhas = [];

      for (
        let i = 1;
        i < raw.length;
        i++
      ) {
        const row =
          raw[i];

        if (
          linhaVazia(row)
        ) {
          continue;
        }

        const [
          classe,
          cod,
          nome
        ] = row;

        const valor = row[4];
        const pct = row[5];

        if (classe) {
          classeAtual =
            classe;
        }

        if (
          classeAtual ===
          'Total'
        ) {
          continue;
        }

        if (
          cod === null ||
          cod === undefined ||
          cod === '' ||
          !nome
        ) {
          continue;
        }

        linhas.push({
          Classe:
            classeAtual,

          Cliente_Pai:
            `${Math.trunc(
              Number(cod)
            )}-${String(nome).trim()}`,

          'Valor de venda (R$)':
            Number(valor) || 0,

          '% do Total':
            Number(pct) || 0
        });
      }

      out[String(mes)] =
        linhas;
    }
  );

  return out;
}

function extrairPorSegmento(
  workbook
) {
  const out = {};

  MESES_FIXOS.forEach(
    mes => {
      const aba =
        encontrarAbaMes(
          workbook,
          mes
        );

      if (!aba) return;

      const raw =
        XLSX.utils.sheet_to_json(
          workbook.Sheets[aba],
          {
            header: 1,
            defval: null
          }
        );

      const linhas = [];

      for (
        let i = 1;
        i < raw.length;
        i++
      ) {
        const row =
          raw[i];

        if (
          linhaVazia(row)
        ) {
          continue;
        }

        const [
          segmento,
          ,
          ,
          valor
        ] = row;

        if (
          String(
            segmento || ''
          )
            .trim()
            .toLowerCase() ===
          'total'
        ) {
          continue;
        }

        linhas.push({
          Separador:
            segmento
              ? String(segmento).trim()
              : 'Outros',

          After_Tax_Amount:
            Number(valor) || 0
        });
      }

      out[String(mes)] =
        linhas;
    }
  );

  return out;
}

function extrairPorProduto(
  workbook
) {
  const out = {};

  MESES_FIXOS.forEach(
    mes => {
      const aba =
        encontrarAbaMes(
          workbook,
          mes
        );

      if (!aba) return;

      const raw =
        XLSX.utils.sheet_to_json(
          workbook.Sheets[aba],
          {
            header: 1,
            defval: null
          }
        );

      const linhas = [];

      for (
        let i = 1;
        i < raw.length;
        i++
      ) {
        const row =
          raw[i];

        if (
          linhaVazia(row)
        ) {
          continue;
        }

        const produto =
          row[0];

        const valor =
          row[2];

        if (
          produto === null ||
          produto === undefined ||
          produto === ''
        ) {
          continue;
        }

        if (
          String(produto)
            .trim()
            .toLowerCase() ===
          'total'
        ) {
          continue;
        }

        linhas.push({
          Produto:
            String(
              produto
            ).trim(),

          'Valor de Venda':
            Number(valor) || 0
        });
      }

      linhas.sort(
        (a, b) =>
          b[
            'Valor de Venda'
          ] -
          a[
            'Valor de Venda'
          ]
      );

      out[String(mes)] =
        linhas.slice(0, 20);
    }
  );

  return out;
}

async function sincronizarPlanilhasDoDrive() {
  if (
    sincronizacaoDriveEmAndamento
  ) {
    return;
  }

  sincronizacaoDriveEmAndamento =
    true;

  try {
    const [
      wbCliente,
      wbSegmento,
      wbProduto
    ] =
      await Promise.all([
        buscarWorkbookDrive(
          DRIVE_SHEET_IDS.cliente
        ),

        buscarWorkbookDrive(
          DRIVE_SHEET_IDS.segmento
        ),

        buscarWorkbookDrive(
          DRIVE_SHEET_IDS.produto
        )
      ]);

    dadosFixosMensais = {
      ano: 2026,

      meses_disponiveis:
        MESES_FIXOS,

      porCliente:
        extrairPorCliente(
          wbCliente
        ),

      porSegmento:
        extrairPorSegmento(
          wbSegmento
        ),

      porProduto:
        extrairPorProduto(
          wbProduto
        )
    };

    try {
      localStorage.setItem(
        'rca61_mensalDrive',
        JSON.stringify({
          dados:
            dadosFixosMensais,

          atualizadoEm:
            new Date()
              .toISOString()
        })
      );
    } catch (e) {}

    popularSelectClientes();
    renderDashboard();

  } catch (e) {
    console.warn(
      '[RCA 61] Drive mensal indisponível:',
      e
    );
  } finally {
    sincronizacaoDriveEmAndamento =
      false;
  }
}

async function carregarDadosFixosMensais() {
  try {
    const cache =
      localStorage.getItem(
        'rca61_mensalDrive'
      );

    if (cache) {
      const obj =
        JSON.parse(cache);

      if (
        obj &&
        obj.dados
      ) {
        dadosFixosMensais =
          obj.dados;

        return;
      }
    }
  } catch (e) {}

  try {
    const res =
      await fetch(
        'historico_fixo.json'
      );

    if (!res.ok) {
      throw new Error(
        'historico_fixo.json não encontrado'
      );
    }

    dadosFixosMensais =
      await res.json();

  } catch (e) {
    dadosFixosMensais =
      null;

    console.warn(
      '[RCA 61] Não foi possível carregar historico_fixo.json',
      e
    );
  }
}

function obterLinhasFixas(
  tipo,
  mes
) {
  if (
    !dadosFixosMensais ||
    !dadosFixosMensais[tipo]
  ) {
    return null;
  }

  const rows =
    dadosFixosMensais[
      tipo
    ][String(mes)];

  return (
    Array.isArray(rows) &&
    rows.length
  )
    ? rows
    : null;
}

function montarFatiaMesCorrente(
  tipo,
  mes
) {
  if (mes !== 9) {
    return null;
  }

  const configs = {

    porCliente: {
      keys: [
        'Vendas (R$) por Cliente',
        'Cliente'
      ],

      key: r =>
        String(
          r['Cliente_Pai'] ||
          r['Cliente'] ||
          ''
        ).trim(),

      value: r =>
        parseCurrency(
          r[
            'Valor de venda (R$)'
          ] ||
          r['Valor']
        ),

      make:
        (
          key,
          value,
          classe
        ) => ({
          Classe:
            classe ||
            'Fiel',

          Cliente_Pai:
            key,

          'Valor de venda (R$)':
            value
        })
    },

    porSegmento: {
      keys: [
        'Venda mensal em reais da',
        'Separador Segmento',
        'Segmento'
      ],

      key: r =>
        String(
          r['Separador'] ||
          r['Segmento'] ||
          ''
        ).trim(),

      value: r =>
        parseCurrency(
          r[
            'After_Tax_Amount'
          ] ||
          r['Valor']
        ),

      make:
        (
          key,
          value
        ) => ({
          Separador:
            key,

          After_Tax_Amount:
            value
        })
    },

    porProduto: {
      keys: [
        'Imagem-7',
        'Image-7',
        'Top 20 Produtos Mais Vendidos',
        'Top 20'
      ],

      key: r =>
        String(
          r['Produto'] ||
          r['Cod'] ||
          ''
        ).trim(),

      value: r =>
        parseCurrency(
          r['Valor de Venda'] ||
          r[
            'Valor de Venda (R$)'
          ] ||
          r['Valor']
        ),

      make:
        (
          key,
          value
        ) => ({
          Produto:
            key,

          'Valor de Venda':
            value
        })
    }
  };

  const cfg =
    configs[tipo];

  if (!cfg) {
    return null;
  }

  const live =
    getSheet(
      dataStore.reportSection,
      cfg.keys
    );

  if (!live.length) {
    return null;
  }

  const somaFixa = {};

  for (
    const m of MESES_FIXOS
  ) {
    const rows =
      obterLinhasFixas(
        tipo,
        m
      );

    if (!rows) {
      continue;
    }

    rows.forEach(
      r => {
        const k =
          normalizarChave(
            cfg.key(r)
          );

        if (!k) return;

        somaFixa[k] =
          (
            somaFixa[k] ||
            0
          ) +
          cfg.value(r);
      }
    );
  }

  let reconhecido = 0;
  let naoReconhecido = 0;

  const out = [];

  live.forEach(
    r => {
      const original =
        cfg.key(r);

      const k =
        normalizarChave(
          original
        );

      if (!k) return;

      const acumulado =
        cfg.value(r);

      if (
        Object.prototype.hasOwnProperty.call(
          somaFixa,
          k
        )
      ) {
        reconhecido +=
          acumulado;
      } else {
        naoReconhecido +=
          acumulado;
      }

      const mesValor =
        acumulado -
        (
          somaFixa[k] ||
          0
        );

      if (
        Math.abs(mesValor) >
        0.005
      ) {
        out.push(
          cfg.make(
            original,
            mesValor,
            r['Classe']
          )
        );
      }
    }
  );

  const base =
    reconhecido +
    naoReconhecido;

  if (
    base > 0 &&
    naoReconhecido >
      reconhecido
  ) {
    return null;
  }

  return out.length
    ? out
    : null;
}

function obterLinhasDoMes(
  tipo,
  mes
) {
  return (
    obterLinhasFixas(
      tipo,
      mes
    ) ||
    montarFatiaMesCorrente(
      tipo,
      mes
    )
  );
}

function usarDadosMensaisFixos(
  ano
) {
  const anoFixo =
    String(
      (
        dadosFixosMensais &&
        dadosFixosMensais.ano
      ) ||
      2026
    );

  return !!dadosFixosMensais &&
    (
      !ano ||
      ano === 'ALL' ||
      String(ano) === anoFixo
    );
}

function salvarSessaoAtual() {
  try {
    sessionStorage.setItem(
      'rca61_sessaoPlanilhas',
      JSON.stringify({
        dataStore,

        nomeArquivo1:
          labelFile1
            ? labelFile1.textContent
            : '',

        nomeArquivo2:
          labelFile2
            ? labelFile2.textContent
            : '',

        carregado1:
          !!(
            dropZone1 &&
            dropZone1.classList.contains(
              'loaded'
            )
          ),

        carregado2:
          !!(
            dropZone2 &&
            dropZone2.classList.contains(
              'loaded'
            )
          )
      })
    );
  } catch (e) {}
}

function restaurarSessaoAtual() {
  try {
    const raw =
      sessionStorage.getItem(
        'rca61_sessaoPlanilhas'
      );

    if (!raw) {
      return false;
    }

    const saved =
      JSON.parse(raw);

    if (
      !saved ||
      !saved.dataStore
    ) {
      return false;
    }

    dataStore =
      saved.dataStore;

    if (
      saved.carregado1 &&
      dropZone1
    ) {
      dropZone1.classList.add(
        'loaded'
      );
    }

    if (
      saved.carregado2 &&
      dropZone2
    ) {
      dropZone2.classList.add(
        'loaded'
      );
    }

    if (
      saved.nomeArquivo1 &&
      labelFile1
    ) {
      labelFile1.textContent =
        saved.nomeArquivo1;
    }

    if (
      saved.nomeArquivo2 &&
      labelFile2
    ) {
      labelFile2.textContent =
        saved.nomeArquivo2;
    }

    if (
      algumaPlanilhaCarregada()
    ) {
      if (statusBadge) {
        statusBadge.classList.add(
          'active'
        );
      }

      if (badgeText) {
        badgeText.textContent =
          'Dados Sincronizados';
      }

      popularSelectClientes();

      requestAnimationFrame(
        renderDashboard
      );

      sincronizarPlanilhasDoDrive();
    }

    return true;

  } catch (e) {
    console.warn(
      '[RCA 61] Sessão não restaurada',
      e
    );

    return false;
  }
}

function readExcelFile(
  file,
  fileNum
) {
  const reader =
    new FileReader();

  reader.onload =
    e => {
      try {
        const workbook =
          XLSX.read(
            new Uint8Array(
              e.target.result
            ),
            {
              type: 'array',
              cellDates: true
            }
          );

        const parsed = {};

        workbook.SheetNames.forEach(
          name => {
            parsed[name.trim()] =
              XLSX.utils.sheet_to_json(
                workbook.Sheets[name],
                {
                  defval: ''
                }
              );
          }
        );

        if (
          fileNum === 1
        ) {
          dataStore.reportSection =
            parsed;

          if (dropZone1) {
            dropZone1.classList.add(
              'loaded'
            );
          }

          if (labelFile1) {
            labelFile1.textContent =
              `✔ ${file.name}`;
          }

        } else {
          dataStore.analiseCarteira =
            parsed;

          if (dropZone2) {
            dropZone2.classList.add(
              'loaded'
            );
          }

          if (labelFile2) {
            labelFile2.textContent =
              `✔ ${file.name}`;
          }
        }

        if (statusBadge) {
          statusBadge.classList.add(
            'active'
          );
        }

        if (badgeText) {
          badgeText.textContent =
            'Dados Sincronizados';
        }

        popularSelectClientes();

        renderDashboard();

        atualizarDadosVivosLocalStorage();

        salvarSessaoAtual();

        sincronizarPlanilhasDoDrive();

      } catch (err) {
        console.error(
          '[RCA 61] Erro ao ler planilha:',
          err
        );
      }
    };

  reader.readAsArrayBuffer(
    file
  );
}

function popularSelectClientes() {
  if (!selectCliente) {
    return;
  }

  const nomes =
    new Set();

  const live =
    getSheet(
      dataStore.reportSection,
      [
        'Vendas (R$) por Cliente',
        'Cliente'
      ]
    );

  live.forEach(
    r => {
      const nome =
        r['Cliente_Pai'] ||
        r['Cliente'];

      if (nome) {
        nomes.add(
          String(nome).trim()
        );
      }
    }
  );

  if (
    dadosFixosMensais &&
    dadosFixosMensais.porCliente
  ) {
    Object.values(
      dadosFixosMensais.porCliente
    ).forEach(
      rows => {
        rows.forEach(
          r => {
            if (
              r.Cliente_Pai
            ) {
              nomes.add(
                String(
                  r.Cliente_Pai
                ).trim()
              );
            }
          }
        );
      }
    );
  }

  const atual =
    selectCliente.value ||
    'ALL';

  selectCliente.innerHTML =
    '<option value="ALL">Todos os Clientes</option>';

  [
    ...nomes
  ]
    .sort()
    .forEach(
      nome => {
        const opt =
          document.createElement(
            'option'
          );

        opt.value =
          nome;

        opt.textContent =
          nome;

        selectCliente.appendChild(
          opt
        );
      }
    );

  selectCliente.value =
    [
      ...selectCliente.options
    ].some(
      o =>
        o.value === atual
    )
      ? atual
      : 'ALL';
}

function atualizarDadosVivosLocalStorage() {
  try {
    const mapa = {};

    const vendas =
      getSheet(
        dataStore.reportSection,
        [
          'Vendas (R$) por Cliente',
          'Cliente'
        ]
      );

    vendas.forEach(
      r => {
        const nome =
          String(
            r['Cliente_Pai'] ||
            r['Cliente'] ||
            ''
          ).trim();

        if (!nome) return;

        const chave =
          normalizarNome(
            nome
          );

        if (!mapa[chave]) {
          mapa[chave] = {
            nome
          };
        }

        mapa[chave].faturamento =
          parseCurrency(
            r[
              'Valor de venda (R$)'
            ] ||
            r['Valor']
          );

        if (r['Classe']) {
          mapa[chave].classe =
            r['Classe'];
        }
      }
    );

    const inativ =
      getSheet(
        dataStore.analiseCarteira,
        [
          '#Dias até primeira fatura',
          'Clientes com ultima fatura',
          'Ultima Fatura'
        ]
      );

    inativ.forEach(
      r => {
        const nome =
          String(
            r['Cliente_Pai'] ||
            r['Cliente'] ||
            ''
          ).trim();

        if (!nome) return;

        const chave =
          normalizarNome(
            nome
          );

        if (!mapa[chave]) {
          mapa[chave] = {
            nome
          };
        }

        mapa[chave].diasInativo =
          parseInt(
            r[
              '#dias desde a ultima fat'
            ] ||
            r['Dias Inativo'] ||
            r['Dias'] ||
            0,
            10
          );

        mapa[chave].ultimaFatura =
          formatDate(
            r['Ultima fat'] ||
            r['Última Fatura']
          );

        if (r['Classe']) {
          mapa[chave].classe =
            r['Classe'];
        }
      }
    );

    localStorage.setItem(
      'rca61_dadosVivos',
      JSON.stringify({
        clientes: mapa,
        atualizadoEm:
          new Date().toISOString()
      })
    );

  } catch (e) {
    console.warn(
      '[RCA 61] Não foi possível salvar dados vivos',
      e
    );
  }
}

function calcularTotalAno2026() {
  let total = 0;

  const mesesFixos =
    new Set();

  if (
    usarDadosMensaisFixos(
      '2026'
    )
  ) {
    for (
      let m = 1;
      m <= 8;
      m++
    ) {
      const rows =
        obterLinhasFixas(
          'porCliente',
          m
        );

      if (!rows) {
        continue;
      }

      total +=
        rows.reduce(
          (
            s,
            r
          ) =>
            s +
            parseCurrency(
              r[
                'Valor de venda (R$)'
              ] ||
              r['Valor']
            ),
          0
        );

      mesesFixos.add(m);
    }
  }

  const fat =
    getSheet(
      dataStore.reportSection,
      [
        'Imagem-6',
        'Image-6',
        'Venda Mensal em Reais',
        'Venda Mensal'
      ]
    );

  fat.forEach(
    r => {
      const mes =
        Number(
          r['Mês'] ||
          r['Mes']
        );

      const ano =
        Number(
          r['Ano']
        );

      if (
        ano === 2026 &&
        mes &&
        !mesesFixos.has(
          mes
        )
      ) {
        total +=
          parseCurrency(
            r[
              'Vendas (R$)'
            ] ||
            r['Vendas'] ||
            r['Valor']
          );
      }
    }
  );

  return total;
}

function renderYTDBanner() {
  const elLytd =
    document.getElementById(
      'kpiLytd'
    );

  const elYtd =
    document.getElementById(
      'kpiYtd'
    );

  const elVar =
    document.getElementById(
      'kpiVariacao'
    );

  if (
    !algumaPlanilhaCarregada()
  ) {
    if (elLytd) {
      elLytd.textContent =
        formatBRL(
          6386614.16
        );
    }

    if (elYtd) {
      elYtd.textContent =
        formatBRL(
          6636963.60
        );
    }

    if (elVar) {
      elVar.textContent =
        formatBRL(
          250349.43
        );
    }

    return;
  }

  let ytd =
    calcularTotalAno2026();

  let lytd =
    null;

  const sheet =
    getSheet(
      dataStore.analiseCarteira,
      [
        'Positivação de carteira'
      ]
    );

  if (sheet.length) {
    const row =
      sheet[0];

    const a =
      parseCurrency(
        row['Vendas LYTD']
      );

    const b =
      parseCurrency(
        row['Vendas YTD']
      );

    if (a > 0) {
      lytd = a;
    }

    if (b > 0) {
      ytd = b;
    }
  }

  if (elYtd) {
    elYtd.textContent =
      formatBRL(ytd);
  }

  if (
    lytd !== null
  ) {
    const variacao =
      ytd - lytd;

    const pct =
      lytd > 0
        ? (variacao / lytd) *
          100
        : 0;

    if (elLytd) {
      elLytd.textContent =
        formatBRL(
          lytd
        );
    }

    if (elVar) {
      elVar.innerHTML =
        `<span style="color:${
          variacao >= 0
            ? '#10b981'
            : '#ef4444'
        };">${
          variacao >= 0
            ? '+'
            : ''
        }${formatBRL(
          variacao
        )} (${
          pct >= 0
            ? '+'
            : ''
        }${pct.toFixed(
          1
        )}%)</span>`;
    }

  } else {

    if (elLytd) {
      elLytd.innerHTML =
        'R$ 0,00<br>' +
        '<span style="font-size:.7rem;font-weight:400;color:#64748b;">' +
        'Ainda sem base de 2025 na planilha' +
        '</span>';
    }

    if (elVar) {
      elVar.innerHTML =
        '<span style="font-size:.9rem;color:#64748b;">N/A</span>';
    }
  }
}

function rw(
  row,
  keys
) {
  if (!row) {
    return undefined;
  }

  for (
    const key of keys
  ) {
    if (
      row[key] !== undefined &&
      row[key] !== null &&
      row[key] !== ''
    ) {
      return row[key];
    }
  }

  return undefined;
}

function renderKPIs() {
  const mesSel =
    selectMes
      ? selectMes.value
      : 'ALL';

  const anoSel =
    selectAno
      ? selectAno.value
      : '2026';

  const clienteSel =
    selectCliente
      ? selectCliente.value
      : 'ALL';

  const clienteEspecifico =
    clienteSel !== 'ALL';

  const carregado =
    algumaPlanilhaCarregada();

  const mesNum =
    mesSel !== 'ALL'
      ? parseInt(
          mesSel,
          10
        )
      : null;

  const anoTemDados =
    anoSel === '2026' ||
    anoSel === 'ALL';

  // ----------------------------------
  // FATURAMENTO
  // ----------------------------------

  let totalFat = 0;

  const sheetFaturamento =
    getSheet(
      dataStore.reportSection,
      [
        'Imagem-6',
        'Image-6',
        'Venda Mensal em Reais',
        'Venda Mensal'
      ]
    );

  if (
    anoTemDados &&
    mesNum !== null
  ) {

    if (
      !clienteEspecifico
    ) {

      const rowMes =
        sheetFaturamento.find(
          r =>
            Number(
              r['Ano']
            ) ===
              Number(
                anoSel
              ) &&
            Number(
              r['Mês'] ||
              r['Mes']
            ) ===
              mesNum
        );

      if (rowMes) {
        totalFat =
          parseCurrency(
            rowMes[
              'Vendas (R$)'
            ] ||
            rowMes[
              'Vendas'
            ] ||
            rowMes[
              'Valor'
            ]
          );
      }

      if (
        totalFat === 0 &&
        usarDadosMensaisFixos(
          anoSel
        )
      ) {
        const rows =
          obterLinhasFixas(
            'porCliente',
            mesNum
          );

        if (rows) {
          totalFat =
            rows.reduce(
              (
                s,
                r
              ) =>
                s +
                parseCurrency(
                  r[
                    'Valor de venda (R$)'
                  ] ||
                  r['Valor']
                ),
              0
            );
        }
      }

    } else {

      const rows =
        obterLinhasDoMes(
          'porCliente',
          mesNum
        ) || [];

      totalFat =
        rows
          .filter(
            r =>
              String(
                r[
                  'Cliente_Pai'
                ] ||
                r[
                  'Cliente'
                ] ||
                ''
              ).trim() ===
              String(
                clienteSel
              ).trim()
          )
          .reduce(
            (
              s,
              r
            ) =>
              s +
              parseCurrency(
                r[
                  'Valor de venda (R$)'
                ] ||
                r['Valor']
              ),
            0
          );
    }

  } else if (
    anoTemDados
  ) {

    totalFat =
      calcularTotalAno2026();
  }

  const elFat =
    document.getElementById(
      'kpiValorMensal'
    );

  const elFatSub =
    elFat
      ? elFat.parentElement.querySelector(
          '.kpi-sub'
        )
      : null;

  if (elFat) {
    elFat.textContent =
      (
        carregado ||
        totalFat > 0
      ) &&
      anoTemDados
        ? formatBRL(
            totalFat
          )
        : 'R$ 0,00';
  }

  if (elFatSub) {
    elFatSub.textContent =
      mesNum !== null &&
      anoTemDados
        ? `Faturamento de ${
            NOMES_MESES[
              mesNum - 1
            ]
          }/${anoSel}`
        : 'Total acumulado 2026';
  }

  // ----------------------------------
  // BUDGET
  // ----------------------------------

  const budget =
    getSheet(
      dataStore.reportSection,
      [
        '% do Budget atingida por',
        '% Budget Atingida',
        'Budget'
      ]
    );

  const budgetFiltrado =
    filtrarPorCliente(
      budget,
      clienteSel
    );

  let budgetPct = 0;

  if (
    !budgetFiltrado.semDetalhePorCliente &&
    budgetFiltrado.linhas.length
  ) {

    if (
      mesNum !== null
    ) {

      const row =
        budgetFiltrado.linhas.find(
          r =>
            Number(
              r['Mês'] ||
              r['Mes']
            ) ===
              mesNum
        );

      if (row) {
        budgetPct =
          parsePct(
            rw(
              row,
              [
                '% do Budget',
                'Budget'
              ]
            )
          );
      }

    } else {

      const vals =
        budgetFiltrado.linhas
          .map(
            r =>
              parsePct(
                rw(
                  r,
                  [
                    '% do Budget',
                    'Budget'
                  ]
                )
              )
          )
          .filter(
            v =>
              v > 0
          );

      if (vals.length) {
        budgetPct =
          vals.reduce(
            (
              a,
              b
            ) =>
              a + b,
            0
          ) /
          vals.length;
      }
    }
  }

  const elBudget =
    document.getElementById(
      'kpiBudgetAtingido'
    );

  if (elBudget) {
    elBudget.textContent =
      `${budgetPct.toFixed(
        1
      )}%`;
  }

  // ----------------------------------
  // POSITIVAÇÃO
  // ----------------------------------

  const elPos =
    document.getElementById(
      'kpiPositivacao'
    );

  const elPosSub =
    document.getElementById(
      'kpiPositivacaoSub'
    );

  const posCard =
    elPosSub
      ? elPosSub.closest(
          '.kpi-card'
        )
      : null;

  if (
    clienteEspecifico
  ) {

    if (elPos) {
      elPos.textContent =
        '—';
    }

    if (elPosSub) {
      elPosSub.innerHTML =
        'Indicador de carteira — ' +
        'selecione "Todos os Clientes" ' +
        'para ver a positivação.';
    }

    if (posCard) {
      posCard.style.borderColor =
        '';

      posCard.style.boxShadow =
        '';
    }

  } else {

    const ultima =
      getSheet(
        dataStore.analiseCarteira,
        [
          'Ultima fatura',
          'Última fatura'
        ]
      );

    const rowPos =
      ultima[0] ||
      {};

    const positivados =
      parseCurrency(
        rw(
          rowPos,
          [
            'Qtd_Positivados',
            'Qtd_Positivado'
          ]
        )
      ) ||
      (
        carregado
          ? 0
          : 82
      );

    const carteira =
      parseCurrency(
        rowPos[
          'Carteira'
        ]
      ) ||
      (
        carregado
          ? 0
          : 271
      );

    const metaOficial =
      parseCurrency(
        rowPos[
          'Meta'
        ]
      );

    const faltaOficial =
      parseCurrency(
        rowPos[
          'Qtd_Falta'
        ]
      );

    const pctOficial =
      parsePct(
        rowPos[
          '%Positivados'
        ]
      );

    const meta =
      metaOficial > 0
        ? metaOficial
        : Math.round(
            carteira *
              0.60
          );

    const metaPct =
      carteira > 0
        ? (
            meta /
            carteira
          ) *
          100
        : 60;

    const realPct =
      pctOficial > 0
        ? pctOficial
        : (
            carteira > 0
              ? (
                  positivados /
                  carteira
                ) *
                100
              : 0
          );

    const falta =
      faltaOficial >= 0
        ? faltaOficial
        : Math.max(
            meta -
              positivados,
            0
          );

    const faltaPct =
      Math.max(
        metaPct -
          realPct,
        0
      );

    const alerta =
      falta > 0 &&
      diasUteisRestantesNoMes() <=
        10;

    if (elPos) {
      elPos.textContent =
        `${positivados} / ${carteira}`;

      elPos.style.color =
        alerta
          ? '#ef4444'
          : '';
    }

    if (elPosSub) {

      if (falta <= 0) {

        elPosSub.innerHTML =
          `Real: <strong>${realPct.toFixed(
            2
          )}%</strong> | ` +
          `Meta: <strong>${meta} clientes (${metaPct.toFixed(
            1
          )}%)</strong> | ` +
          `<span style="color:#10b981;font-weight:bold;">Meta Atingida!</span>`;

      } else {

        elPosSub.innerHTML =
          `Real: <strong>${realPct.toFixed(
            2
          )}%</strong> | ` +
          `Meta: <strong>${meta} clientes (${metaPct.toFixed(
            1
          )}%)</strong> | ` +
          `Falta: <span style="color:#ef4444;font-weight:bold;">${falta} clientes (${faltaPct.toFixed(
            2
          )}%)</span>` +
          (
            alerta
              ? `<br><span style="display:inline-block;margin-top:6px;padding:4px 8px;border-radius:6px;background:rgba(239,68,68,.15);color:#ef4444;font-weight:700;">🚨 Faltam ${diasUteisRestantesNoMes()} dias úteis para o fim do mês</span>`
              : ''
          );
      }
    }

    if (posCard) {
      posCard.style.borderColor =
        alerta
          ? '#ef4444'
          : '';

      posCard.style.boxShadow =
        alerta
          ? '0 0 0 1px rgba(239,68,68,.45)'
          : '';
    }
  }

  // ----------------------------------
  // ENCOMENDAS GRAVADAS
  // ----------------------------------

  const gravSheet =
    getSheet(
      dataStore.analiseCarteira,
      [
        '% de encomendas gravadas',
        'Encomendas Gravadas',
        'Gravado'
      ]
    );

  const grav =
    filtrarPorCliente(
      gravSheet,
      clienteSel
    );

  let pctGrav =
    carregado
      ? 0
      : 44.72;

  if (
    !grav.semDetalhePorCliente &&
    grav.linhas.length
  ) {

    const row =
      mesNum !== null

        ? grav.linhas.find(
            r =>
              Number(
                r['Ano']
              ) ===
                Number(
                  anoSel
                ) &&
              Number(
                r['Mes']
              ) ===
                mesNum &&
              String(
                r['Tipo'] ||
                  ''
              )
                .toLowerCase() ===
                'gravado'
          )

        : null;

    const alvo =
      row ||
      grav.linhas.find(
        r =>
          String(
            r['Ano'] ||
              r['Tipo']
          )
            .toLowerCase()
            .includes(
              'total'
            )
      ) ||
      grav.linhas[0];

    pctGrav =
      parsePct(
        rw(
          alvo,
          [
            '% gravação',
            '% Gravado',
            'Total',
            'Gravado'
          ]
        )
      );
  }

  const elGrav =
    document.getElementById(
      'kpiPctGravadas'
    );

  const elGravSub =
    document.getElementById(
      'kpiPctGravadasSub'
    );

  if (elGrav) {
    elGrav.textContent =
      `${pctGrav.toFixed(
        2
      )}%`;
  }

  if (elGravSub) {

    const diff =
      40 -
      pctGrav;

    elGravSub.innerHTML =
      grav.semDetalhePorCliente &&
      clienteEspecifico
        ? 'Indicador de carteira — selecione "Todos os Clientes".'

        : diff <= 0
          ? 'Meta: 40% | ' +
            '<span style="color:#10b981;font-weight:bold;">' +
            'Meta Atingida!' +
            '</span>'

          : `Meta: 40% | ` +
            `Falta: <span style="color:#f59e0b;font-weight:bold;">${diff.toFixed(
              2
            )}%</span> p/ a meta`;
  }
}

const pluginValoresNativos = {
  id: 'pluginValoresNativos',

  afterDatasetsDraw(
    chart
  ) {
    if (
      !chart ||
      !chart.data ||
      !chart.data.datasets
    ) {
      return;
    }

    chart.data.datasets.forEach(
      (
        dataset,
        dsIndex
      ) => {

        const meta =
          chart.getDatasetMeta(
            dsIndex
          );

        if (meta.hidden) {
          return;
        }

        if (
          chart.config.type !==
            'doughnut' &&
          meta.data.length >
            24
        ) {
          return;
        }

        meta.data.forEach(
          (
            element,
            index
          ) => {

            const value =
              dataset.data[
                index
              ];

            if (
              value === null ||
              value === undefined ||
              value === 0
            ) {
              return;
            }

            const ctx =
              chart.ctx;

            ctx.save();

            const horizontal =
              chart.options
                .indexAxis ===
              'y';

            if (
              chart.config.type ===
              'doughnut'
            ) {

              const nums =
                dataset.data.filter(
                  v =>
                    typeof v ===
                    'number'
                );

              const total =
                nums.reduce(
                  (
                    a,
                    b
                  ) =>
                    a + b,
                  0
                );

              const pct =
                total > 0
                  ? `${(
                      (
                        value /
                        total
                      ) *
                      100
                    ).toFixed(
                      1
                    )}%`
                  : '';

              const p =
                element.tooltipPosition();

              ctx.font =
                'bold 10px sans-serif';

              ctx.fillStyle =
                '#fff';

              ctx.textAlign =
                'center';

              ctx.fillText(
                `${value} (${pct})`,
                p.x,
                p.y
              );

            } else {

              const isPct =
                String(
                  dataset.label ||
                    ''
                )
                  .includes('%') ||
                String(
                  dataset.label ||
                    ''
                )
                  .toLowerCase()
                  .includes(
                    'budget'
                  );

              const text =
                isPct
                  ? `${Number(
                      value
                    ).toFixed(
                      1
                    )}%`
                  : formatCompactBRL(
                      value
                    );

              ctx.font =
                `bold ${
                  meta.data
                    .length >
                    12
                    ? 8
                    : 10
                }px sans-serif`;

              ctx.fillStyle =
                '#fff';

              const p =
                element.tooltipPosition
                  ? element.tooltipPosition()
                  : {
                      x: element.x,
                      y: element.y
                    };

              const w =
                ctx.measureText(
                  text
                ).width;

              ctx.fillStyle =
                'rgba(15,23,42,.85)';

              if (horizontal) {

                ctx.fillRect(
                  p.x + 2,
                  p.y - 7,
                  w + 8,
                  14
                );

                ctx.fillStyle =
                  '#fff';

                ctx.textAlign =
                  'left';

                ctx.fillText(
                  text,
                  p.x + 6,
                  p.y
                );

              } else {

                ctx.fillRect(
                  p.x -
                    w / 2 -
                    4,
                  p.y -
                    16,
                  w + 8,
                  14
                );

                ctx.fillStyle =
                  '#fff';

                ctx.textAlign =
                  'center';

                ctx.fillText(
                  text,
                  p.x,
                  p.y - 6
                );
              }
            }

            ctx.restore();
          }
        );
      }
    );
  }
};

function renderChartHistorico() {
  const ctx =
    document.getElementById(
      'chartHistoricoFaturamento'
    );

  if (!ctx) {
    return;
  }

  const cliente =
    selectCliente
      ? selectCliente.value
      : 'ALL';

  const fat =
    getSheet(
      dataStore.reportSection,
      [
        'Imagem-6',
        'Image-6',
        'Venda Mensal em Reais',
        'Venda Mensal'
      ]
    );

  const serie2026 =
    new Array(
      12
    ).fill(null);

  for (
    let m = 1;
    m <= 8;
    m++
  ) {

    const rows =
      usarDadosMensaisFixos(
        '2026'
      )
        ? obterLinhasFixas(
            'porCliente',
            m
          )
        : null;

    if (!rows) {
      continue;
    }

    const flt =
      filtrarPorCliente(
        rows,
        cliente
      ).linhas;

    const total =
      flt.reduce(
        (
          s,
          r
        ) =>
          s +
          parseCurrency(
            r[
              'Valor de venda (R$)'
            ] ||
            r['Valor']
          ),
        0
      );

    if (total > 0) {
      serie2026[m - 1] =
        total;
    }
  }

  fat.forEach(
    r => {
      const m =
        Number(
          r['Mês'] ||
          r['Mes']
        );

      const a =
        Number(
          r['Ano']
        );

      const v =
        parseCurrency(
          r[
            'Vendas (R$)'
          ] ||
          r['Vendas'] ||
          r['Valor']
        );

      if (
        a === 2026 &&
        m >= 9 &&
        m <= 12 &&
        v > 0 &&
        serie2026[
          m - 1
        ] === null
      ) {
        serie2026[
          m - 1
        ] = v;
      }
    }
  );

  if (
    cliente !== 'ALL'
  ) {

    for (
      let m = 9;
      m <= 12;
      m++
    ) {

      if (
        serie2026[
          m - 1
        ] !== null
      ) {
        continue;
      }

      const rows =
        obterLinhasDoMes(
          'porCliente',
          m
        );

      if (!rows) {
        continue;
      }

      const total =
        filtrarPorCliente(
          rows,
          cliente
        ).linhas.reduce(
          (
            s,
            r
          ) =>
            s +
            parseCurrency(
              r[
                'Valor de venda (R$)'
              ] ||
              r['Valor']
            ),
          0
        );

      if (total > 0) {
        serie2026[
          m - 1
        ] = total;
      }
    }
  }

  const series = {
    2023:
      new Array(
        12
      ).fill(null),

    2024:
      new Array(
        12
      ).fill(null),

    2025:
      new Array(
        12
      ).fill(null),

    2026:
      serie2026
  };

  if (
    cliente === 'ALL'
  ) {

    fat.forEach(
      r => {

        const m =
          Number(
            r['Mês'] ||
            r['Mes']
          );

        const a =
          Number(
            r['Ano']
          );

        const v =
          parseCurrency(
            r[
              'Vendas (R$)'
            ] ||
            r['Vendas'] ||
            r['Valor']
          );

        if (
          m < 1 ||
          m > 12 ||
          !series[a] ||
          v < 0
        ) {
          return;
        }

        if (
          a === 2026 &&
          series[2026][
            m - 1
          ] !== null
        ) {
          return;
        }

        series[a][
          m - 1
        ] = v;
      }
    );
  }

  destruirGrafico(
    'chartHistoricoFaturamento'
  );

  charts.chartHistoricoFaturamento =
    new Chart(
      ctx.getContext(
        '2d'
      ),
      {
        type:
          'line',

        data: {
          labels:
            LABELS_MESES,

          datasets:
            Object.keys(
              series
            ).map(
              ano => ({
                label:
                  ano,

                data:
                  series[
                    ano
                  ],

                borderColor:
                  ano === '2026'
                    ? '#facc15'
                    : ano === '2025'
                      ? '#f59e0b'
                      : ano === '2024'
                        ? '#ec4899'
                        : '#64748b',

                backgroundColor:
                  'transparent',

                borderWidth:
                  ano ===
                  '2026'
                    ? 3.5
                    : 2,

                pointRadius:
                  ano ===
                  '2026'
                    ? 3
                    : 2,

                tension:
                  0.25,

                spanGaps:
                  true
              })
            )
        },

        options: {
          responsive:
            true,

          maintainAspectRatio:
            false,

          interaction: {
            mode:
              'index',

            intersect:
              false
          },

          plugins: {

            legend: {
              position:
                'top',

              labels: {
                color:
                  '#cbd5e1',

                usePointStyle:
                  true,

                padding:
                  16
              }
            },

            title: {
              display:
                true,

              text:
                'Passe o mouse sobre cada mês para ver os valores',

              color:
                '#94a3b8',

              font: {
                size:
                  10,

                weight:
                  'normal'
              }
            },

            tooltip: {

              callbacks: {

                label:
                  c =>
                    `${c.dataset.label}: ${formatBRL(
                      c.parsed.y
                    )}`
              }
            }
          },

          scales: {

            x: {

              grid: {
                color:
                  'rgba(148,163,184,.08)'
              },

              ticks: {
                color:
                  '#94a3b8'
              }
            },

            y: {

              beginAtZero:
                false,

              grid: {
                color:
                  'rgba(148,163,184,.08)'
              },

              ticks: {
                color:
                  '#94a3b8',

                callback:
                  v =>
                    formatCompactBRL(
                      v
                    )
              }
            }
          }
        }
      }
    );
}

function renderChartBudget() {
  const ctx =
    document.getElementById(
      'chartBudget'
    );

  if (!ctx) {
    return;
  }

  const cliente =
    selectCliente
      ? selectCliente.value
      : 'ALL';

  const sheet =
    filtrarPorCliente(
      getSheet(
        dataStore.reportSection,
        [
          '% do Budget atingida por',
          '% Budget Atingida',
          'Budget'
        ]
      ),
      cliente
    );

  const vals =
    new Array(
      12
    ).fill(0);

  sheet.linhas.forEach(
    r => {

      const idx =
        Number(
          r['Mês'] ||
          r['Mes']
        ) - 1;

      if (
        idx >= 0 &&
        idx < 12
      ) {
        vals[idx] =
          parsePct(
            rw(
              r,
              [
                '% do Budget',
                'Budget'
              ]
            )
          );
      }
    }
  );

  destruirGrafico(
    'chartBudget'
  );

  ctx.parentElement.style.height =
    '280px';

  charts.chartBudget =
    new Chart(
      ctx.getContext('2d'),
      {
        type:
          'bar',

        data: {
          labels:
            LABELS_MESES,

          datasets: [
            {
              label:
                '% Budget Atingido',

              data:
                vals,

              backgroundColor:
                '#6366f1',

              borderRadius:
                4
            }
          ]
        },

        plugins: [
          pluginValoresNativos
        ],

        options: {
          responsive:
            true,

          maintainAspectRatio:
            false,

          indexAxis:
            'y',

          plugins: {

            legend: {
              display:
                true,

              labels: {
                color:
                  '#94a3b8'
              }
            }
          },

          scales: {

            x: {

              beginAtZero:
                true,

              grid: {
                color:
                  '#1f293d'
              },

              ticks: {

                color:
                  '#94a3b8',

                callback:
                  v =>
                    `${v}%`
              }
            },

            y: {

              grid: {
                display:
                  false
              },

              ticks: {
                color:
                  '#94a3b8'
              }
            }
          }
        }
      }
    );
}

function renderChartTipoEncomenda() {
  const ctx =
    document.getElementById(
      'chartTipoEncomenda'
    );

  if (!ctx) {
    return;
  }

  const cliente =
    selectCliente
      ? selectCliente.value
      : 'ALL';

  const sheet =
    filtrarPorCliente(
      getSheet(
        dataStore.analiseCarteira,
        [
          'Clientes recentes que já',
          '% de encomendas gravadas',
          'Encomenda por Tipo'
        ]
      ),
      cliente
    ).linhas;

  let gravado =
    algumaPlanilhaCarregada()
      ? 0
      : 89;

  let normal =
    algumaPlanilhaCarregada()
      ? 0
      : 110;

  const rg =
    sheet.find(
      r =>
        String(
          r.Tipo ||
            ''
        )
          .toLowerCase()
          .includes(
            'gravado'
          )
    );

  const rn =
    sheet.find(
      r =>
        String(
          r.Tipo ||
            ''
        )
          .toLowerCase()
          .includes(
            'normal'
          )
    );

  if (rg) {
    gravado =
      parseCurrency(
        rg[
          'Sum of Valor'
        ] ??
        rg[
          'Valor'
        ] ??
        rg[
          '% gravação'
        ]
      );
  }

  if (rn) {
    normal =
      parseCurrency(
        rn[
          'Sum of Valor'
        ] ??
        rn[
          'Valor'
        ] ??
        rn[
          '% gravação'
        ]
      );
  }

  destruirGrafico(
    'chartTipoEncomenda'
  );

  charts.chartTipoEncomenda =
    new Chart(
      ctx.getContext(
        '2d'
      ),
      {
        type:
          'doughnut',

        data: {
          labels: [
            'Gravado',
            'Normal'
          ],

          datasets: [
            {
              data: [
                gravado,
                normal
              ],

              backgroundColor: [
                '#10b981',
                '#ef4444'
              ],

              borderWidth:
                0
            }
          ]
        },

        plugins: [
          pluginValoresNativos
        ],

        options: {

          responsive:
            true,

          maintainAspectRatio:
            false,

          plugins: {

            legend: {
              display:
                true,

              labels: {
                color:
                  '#94a3b8'
              }
            }
          }
        }
      }
    );
}

function renderChartSegmentos() {
  const ctx =
    document.getElementById(
      'chartSegmentos'
    );

  if (!ctx) {
    return;
  }

  const cliente =
    selectCliente
      ? selectCliente.value
      : 'ALL';

  const mes =
    selectMes &&
    selectMes.value !==
      'ALL'
      ? parseInt(
          selectMes.value,
          10
        )
      : null;

  const ano =
    selectAno
      ? selectAno.value
      : '2026';

  const rowsFixos =
    mes !== null &&
    usarDadosMensaisFixos(
      ano
    )
      ? obterLinhasDoMes(
          'porSegmento',
          mes
        )
      : null;

  const sheet =
    filtrarPorCliente(
      rowsFixos ||
      getSheet(
        dataStore.reportSection,
        [
          'Venda mensal em reais da',
          'Separador Segmento',
          'Segmento'
        ]
      ),
      cliente
    ).linhas;

  const data =
    [
      ...sheet
    ]
      .map(
        r => ({
          label:
            String(
              r[
                'Separador'
              ] ||
              r[
                'Segmento'
              ] ||
              'Outros'
            ),

          value:
            parseCurrency(
              r[
                'After_Tax_Amount'
              ] ||
              r['Valor']
            )
        })
      )
      .sort(
        (
          a,
          b
        ) =>
          b.value -
          a.value
      );

  destruirGrafico(
    'chartSegmentos'
  );

  ctx.parentElement.style.height =
    Math.max(
      300,
      data.length * 26
    ) + 'px';

  charts.chartSegmentos =
    new Chart(
      ctx.getContext(
        '2d'
      ),
      {

        type:
          'bar',

        data: {

          labels:
            data.length
              ? data.map(
                  x =>
                    x.label
                )
              : [
                  'Sem Dados'
                ],

          datasets: [
            {
              label:
                'Vendas (R$)',

              data:
                data.length
                  ? data.map(
                      x =>
                        x.value
                    )
                  : [0],

              backgroundColor:
                '#10b981',

              borderRadius:
                4
            }
          ]
        },

        plugins: [
          pluginValoresNativos
        ],

        options: {

          responsive:
            true,

          maintainAspectRatio:
            false,

          indexAxis:
            'y',

          plugins: {

            legend: {
              display:
                true,

              labels: {
                color:
                  '#94a3b8'
              }
            }
          },

          scales: {

            x: {

              beginAtZero:
                true,

              grid: {
                color:
                  '#1f293d'
              },

              ticks: {

                color:
                  '#94a3b8',

                callback:
                  v =>
                    formatCompactBRL(
                      v
                    )
              }
            },

            y: {

              grid: {
                display:
                  false
              },

              ticks: {

                color:
                  '#94a3b8',

                autoSkip:
                  false,

                font: {
                  size:
                    9
                }
              }
            }
          }
        }
      }
    );
}

function renderChartTopProdutos() {
  const ctx =
    document.getElementById(
      'chartTopProdutos'
    );

  if (!ctx) {
    return;
  }

  const cliente =
    selectCliente
      ? selectCliente.value
      : 'ALL';

  const mes =
    selectMes &&
    selectMes.value !==
      'ALL'
      ? parseInt(
          selectMes.value,
          10
        )
      : null;

  const ano =
    selectAno
      ? selectAno.value
      : '2026';

  const rowsFixos =
    mes !== null &&
    mes <= 8 &&
    usarDadosMensaisFixos(
      ano
    )
      ? obterLinhasFixas(
          'porProduto',
          mes
        )
      : null;

  const raw =
    rowsFixos ||
    getSheet(
      dataStore.reportSection,
      [
        'Imagem-7',
        'Image-7',
        'Top 20 Produtos Mais Vendidos',
        'Top 20'
      ]
    );

  const filtrado =
    filtrarPorCliente(
      raw,
      cliente
    );

  const data =
    filtrado.linhas
      .map(
        r => ({
          label:
            String(
              r[
                'Produto'
              ] ||
              r[
                'Descrição'
              ] ||
              r[
                'Descricao'
              ] ||
              r[
                'Cod'
              ] ||
              ''
            ).trim(),

          value:
            parseCurrency(
              r[
                'Valor de Venda'
              ] ||
              r[
                'Valor de Venda (R$)'
              ] ||
              r[
                'Valor'
              ] ||
              r[
                'Vendas (R$)'
              ]
            )
        })
      )
      .filter(
        x =>
          x.label &&
          Number.isFinite(
            x.value
          )
      )
      .sort(
        (
          a,
          b
        ) =>
          b.value -
          a.value
      )
      .slice(
        0,
        20
      );

  destruirGrafico(
    'chartTopProdutos'
  );

  ctx.parentElement.style.height =
    '430px';

  charts.chartTopProdutos =
    new Chart(
      ctx.getContext(
        '2d'
      ),
      {

        type:
          'bar',

        data: {

          labels:
            data.length
              ? data.map(
                  x =>
                    x.label
                )
              : [
                  'Sem Dados'
                ],

          datasets: [
            {
              label:
                'Valor de Venda (R$)',

              data:
                data.length
                  ? data.map(
                      x =>
                        x.value
                    )
                  : [0],

              backgroundColor:
                '#3b82f6',

              borderRadius:
                4,

              barPercentage:
                0.72,

              categoryPercentage:
                0.82
            }
          ]
        },

        plugins: [
          pluginValoresNativos
        ],

        options: {

          responsive:
            true,

          maintainAspectRatio:
            false,

          indexAxis:
            'y',

          plugins: {

            legend: {
              display:
                true,

              labels: {
                color:
                  '#94a3b8'
              }
            },

            tooltip: {

              callbacks: {

                label:
                  c =>
                    formatBRL(
                      c.parsed.x
                    )
              }
            }
          },

          scales: {

            x: {

              beginAtZero:
                true,

              grid: {
                color:
                  'rgba(148,163,184,.08)'
              },

              ticks: {

                color:
                  '#94a3b8',

                callback:
                  v =>
                    formatCompactBRL(
                      v
                    )
              }
            },

            y: {

              grid: {
                display:
                  false
              },

              ticks: {

                color:
                  '#cbd5e1',

                autoSkip:
                  false,

                font: {
                  size:
                    8.5
                }
              }
            }
          }
        }
      }
    );
}

function formatPctCell(
  val
) {
  if (
    val === undefined ||
    val === null ||
    val === ''
  ) {
    return '-';
  }

  if (
    typeof val ===
      'string' &&
    val.includes('%')
  ) {
    return val;
  }

  return `${parsePct(
    val
  ).toFixed(
    2
  )}%`;
}

function renderVendasClienteTable() {
  const tbody =
    document.getElementById(
      'tbVendasCliente'
    );

  if (!tbody) {
    return;
  }

  const mes =
    selectMes &&
    selectMes.value !==
      'ALL'
      ? parseInt(
          selectMes.value,
          10
        )
      : null;

  const ano =
    selectAno
      ? selectAno.value
      : '2026';

  const cliente =
    selectCliente
      ? selectCliente.value
      : 'ALL';

  const query =
    searchInput
      ? searchInput.value
          .toLowerCase()
          .trim()
      : '';

  const rows =
    mes !== null &&
    usarDadosMensaisFixos(
      ano
    )
      ? obterLinhasDoMes(
          'porCliente',
          mes
        )
      : getSheet(
          dataStore.reportSection,
          [
            'Vendas (R$) por Cliente',
            'Cliente'
          ]
        );

  const filtered =
    rows
      .filter(
        r => {

          const nome =
            String(
              r[
                'Cliente_Pai'
              ] ||
              r[
                'Cliente'
              ] ||
              ''
            ).trim();

          return (
            nome
              .toLowerCase()
              .includes(
                query
              ) &&
            (
              cliente ===
                'ALL' ||
              nome ===
                cliente
            )
          );
        }
      )
      .sort(
        (
          a,
          b
        ) =>
          parseCurrency(
            b[
              'Valor de venda (R$)'
            ] ||
            b['Valor']
          ) -
          parseCurrency(
            a[
              'Valor de venda (R$)'
            ] ||
            a['Valor']
          )
      );

  tbody.innerHTML =
    '';

  if (!filtered.length) {

    tbody.innerHTML =
      '<tr><td colspan="4" class="empty-row">Nenhum registro localizado.</td></tr>';

    return;
  }

  filtered.forEach(
    r => {

      const nome =
        r[
          'Cliente_Pai'
        ] ||
        r[
          'Cliente'
        ] ||
        '-';

      const valor =
        parseCurrency(
          r[
            'Valor de venda (R$)'
          ] ||
          r['Valor']
        );

      const tr =
        document.createElement(
          'tr'
        );

      tr.className =
        'clickable-row';

      tr.style.cursor =
        'pointer';

      tr.innerHTML =
        `<td>${
          r['Classe'] ||
          'Fiel'
        }</td>` +

        `<td><strong>${nome}</strong></td>` +

        `<td>${formatBRL(
          valor
        )}</td>` +

        `<td>${formatPctCell(
          r[
            '% do Total'
          ]
        )}</td>`;

      tbody.appendChild(
        tr
      );
    }
  );
}

function renderInatividadeTable() {
  const tbody =
    document.getElementById(
      'tbInatividade'
    );

  if (!tbody) {
    return;
  }

  const cliente =
    selectCliente
      ? selectCliente.value
      : 'ALL';

  const query =
    searchInput
      ? searchInput.value
          .toLowerCase()
          .trim()
      : '';

  const rows =
    getSheet(
      dataStore.analiseCarteira,
      [
        '#Dias até primeira fatura',
        'Clientes com ultima fatura',
        'Ultima Fatura'
      ]
    );

  const filtered =
    rows.filter(
      r => {

        const nome =
          String(
            r[
              'Cliente_Pai'
            ] ||
            r[
              'Cliente'
            ] ||
            ''
          ).trim();

        return (
          nome
            .toLowerCase()
            .includes(
              query
            ) &&
          (
            cliente ===
              'ALL' ||
            nome ===
              cliente
          )
        );
      }
    );

  tbody.innerHTML =
    '';

  if (!filtered.length) {

    tbody.innerHTML =
      '<tr><td colspan="4" class="empty-row">Nenhum registro localizado.</td></tr>';

    return;
  }

  filtered.forEach(
    r => {

      const nome =
        r[
          'Cliente_Pai'
        ] ||
        r[
          'Cliente'
        ] ||
        '-';

      const dias =
        parseInt(
          r[
            '#dias desde a ultima fat'
          ] ||
          r[
            'Dias Inativo'
          ] ||
          r[
            'Dias'
          ] ||
          0,
          10
        );

      const data =
        formatDate(
          r[
            'Ultima fat'
          ] ||
          r[
            'Última Fatura'
          ]
        );

      const tr =
        document.createElement(
          'tr'
        );

      tr.className =
        'clickable-row';

      tr.style.cursor =
        'pointer';

      tr.innerHTML =
        `<td>${
          r['Classe'] ||
          'Pontual'
        }</td>` +

        `<td><strong>${nome}</strong></td>` +

        `<td>${data}</td>` +

        `<td><span style="color:${
          dias >= 60
            ? '#ef4444'
            : '#10b981'
        };font-weight:700;">${dias} dias</span></td>`;

      tbody.appendChild(
        tr
      );
    }
  );
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

if (fileInput1) {
  fileInput1.addEventListener(
    'change',
    e =>
      e.target.files.length &&
      readExcelFile(
        e.target.files[0],
        1
      )
  );
}

if (fileInput2) {
  fileInput2.addEventListener(
    'change',
    e =>
      e.target.files.length &&
      readExcelFile(
        e.target.files[0],
        2
      )
  );
}

if (selectAno) {
  selectAno.addEventListener(
    'change',
    renderDashboard
  );
}

if (selectMes) {
  selectMes.addEventListener(
    'change',
    renderDashboard
  );
}

if (selectCliente) {
  selectCliente.addEventListener(
    'change',
    renderDashboard
  );
}

if (searchInput) {
  searchInput.addEventListener(
    'input',
    () => {
      renderVendasClienteTable();
      renderInatividadeTable();
    }
  );
}

carregarDadosFixosMensais()
  .then(
    () => {
      popularSelectClientes();
      renderDashboard();
      restaurarSessaoAtual();
    }
  );

window.addEventListener(
  'pageshow',
  e => {
    if (e.persisted) {
      requestAnimationFrame(
        renderDashboard
      );
    }
  }
);

if (
  typeof Chart !==
  'undefined'
) {

  Chart.defaults.plugins =
    Chart.defaults.plugins ||
    {};

  Chart.defaults.plugins.tooltip =
    Chart.defaults.plugins.tooltip ||
    {};

  Chart.defaults.plugins.tooltip.callbacks =
    Chart.defaults.plugins.tooltip.callbacks ||
    {};

  Chart.defaults.plugins.tooltip.callbacks.label =
    function (
      context
    ) {

      const label =
        context.dataset?.label ||
        context.label ||
        '';

      let val =
        context.parsed;

      if (
        val &&
        typeof val ===
          'object'
      ) {
        val =
          val.y ??
          val.x ??
          val.r;
      }

      if (
        typeof val !==
          'number'
      ) {
        return label;
      }

      return label
        ? `${label}: ${formatBRL(
            val
          )}`
        : formatBRL(
            val
          );
    };
}

function abrirModalCliente(
  dados
) {
  let modal =
    document.getElementById(
      'customClientModal'
    );

  if (!modal) {

    modal =
      document.createElement(
        'div'
      );

    modal.id =
      'customClientModal';

    modal.style.cssText =
      'position:fixed;inset:0;background:rgba(11,15,25,.8);backdrop-filter:blur(6px);z-index:99999;display:flex;align-items:center;justify-content:center;';

    document.body.appendChild(
      modal
    );
  }

  modal.innerHTML =
    `<div style="background:#1e293b;border:1px solid #334155;border-radius:12px;width:90%;max-width:480px;padding:24px;color:#f8fafc;font-family:sans-serif;">` +

    `<div style="display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #334155;padding-bottom:12px;margin-bottom:16px;">` +

    `<h3 style="margin:0;color:#6366f1;">🏢 Detalhes do Cliente</h3>` +

    `<button onclick="fecharModalCliente()" style="background:transparent;border:none;color:#94a3b8;font-size:1.5rem;cursor:pointer;">&times;</button>` +

    `</div>` +

    `<div style="display:grid;gap:10px;">` +

    `<div><span style="color:#94a3b8;font-size:.75rem;display:block;">Cliente</span><strong>${dados.nome || '-'}</strong></div>` +

    (
      dados.faturamento
        ? `<div><span style="color:#94a3b8;font-size:.75rem;display:block;">Faturamento</span><strong>${dados.faturamento}</strong></div>`
        : ''
    ) +

    (
      dados.ultimaFatura
        ? `<div><span style="color:#94a3b8;font-size:.75rem;display:block;">Última Fatura</span><strong>${dados.ultimaFatura}</strong></div>`
        : ''
    ) +

    (
      dados.inatividade
        ? `<div><span style="color:#94a3b8;font-size:.75rem;display:block;">Inatividade</span><strong>${dados.inatividade}</strong></div>`
        : ''
    ) +

    `</div>` +

    `<div style="margin-top:20px;text-align:right;">` +

    `<button onclick="fecharModalCliente()" style="background:#6366f1;color:#fff;border:none;padding:8px 18px;border-radius:6px;font-weight:600;cursor:pointer;">Fechar</button>` +

    `</div>` +

    `</div>`;

  modal.style.display =
    'flex';
}

function fecharModalCliente() {
  const modal =
    document.getElementById(
      'customClientModal'
    );

  if (modal) {
    modal.style.display =
      'none';
  }
}
