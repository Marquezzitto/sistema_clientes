let chartPosObj = null;
let chartGravObj = null;

// Base estática de segurança da carteira caso a planilha venha apenas com os pedidos do mês
// O sistema cruza e calcula tudo dinamicamente.
document.getElementById('excelFile').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        const sheetNames = workbook.SheetNames;

        // Identifica as abas corretamente
        let sheetPedidosName = sheetNames.find(s => s.toUpperCase().includes('PEDIDO')) || sheetNames[0];
        let sheetCarteiraName = sheetNames.find(s => s.toUpperCase().includes('CARTEIRA') || s.toUpperCase().includes('CLIENTE')) || sheetNames[1];

        const jsonPedidos = XLSX.utils.sheet_to_json(workbook.Sheets[sheetPedidosName]);
        const jsonCarteira = sheetCarteiraName ? XLSX.utils.sheet_to_json(workbook.Sheets[sheetCarteiraName]) : [];

        processarDadosMes(jsonPedidos, jsonCarteira);
    };
    reader.readAsArrayBuffer(file);
});

function processarDadosMes(pedidos, carteiraBase) {
    // Definimos a meta da base total (ex: 270 ou o tamanho total da carteira base carregada)
    const totalClientesBase = carteiraBase.length > 0 ? carteiraBase.length : 270;
    const metaPositivacaoQtd = Math.round(totalClientesBase * 0.60);

    let clientesComPrimeiroPedido = new Set();
    let totalPedidos = pedidos.length;
    let totalGravados = 0;

    // Processa os pedidos do mês
    pedidos.forEach(p => {
        let clienteNome = String(p['CLIENTE'] || p['Cliente'] || '').trim().toUpperCase();
        let gravado = String(p['GRAVADO'] || p['Gravado'] || '').trim().toUpperCase();
        let primeiroPedido = String(p['1º PDEIDO?'] || p['1º PEDIDO?'] || '').trim().toUpperCase();

        if (gravado.includes('SIM')) {
            totalGravados++;
        }
        if (primeiroPedido.includes('SIM') && clienteNome) {
            clientesComPrimeiroPedido.add(clienteNome);
        }
    });

    const qtdPositivados = clientesComPrimeiroPedido.size;
    const pctPositivacao = (qtdPositivados / totalClientesBase) * 100;
    const pctGravacao = totalPedidos > 0 ? (totalGravados / totalPedidos) * 100 : 0;

    // Atualiza os KPIs na tela
    document.getElementById('kpiPositivados').innerText = `${qtdPositivados} / ${metaPositivacaoQtd}`;
    document.getElementById('kpiPctPositivacao').innerText = `${pctPositivacao.toFixed(1)}%`;
    document.getElementById('kpiGravados').innerText = `${totalGravados} / ${totalPedidos}`;
    document.getElementById('kpiPctGravacao').innerText = `${pctGravacao.toFixed(1)}%`;

    // Mapeia a carteira base completa
    let carteiraMapeada = [];
    if (carteiraBase.length > 0) {
        carteiraBase.forEach(c => {
            let nomeCliente = String(c['Cliente'] || c['CLIENTE'] || '').trim();
            let nomeClean = nomeCliente.toUpperCase();
            if (nomeClean) {
                carteiraMapeada.push({
                    codigo: c['Código Cliente'] || c['COD. CLIENTE'] || '-',
                    nome: nomeCliente,
                    cidade: c['Cidade'] || '-',
                    uf: c['UF'] || '-',
                    status: c['Status'] || c['STATUS'] || 'BASE',
                    clean: nomeClean
                });
            }
        });
    }

    // Cruza a base de clientes com os pedidos do mês para achar quem ainda não comprou
    let clientesFaltantes = [];
    if (carteiraMapeada.length > 0) {
        clientesFaltantes = carteiraMapeada.filter(c => 
            !clientesComPrimeiroPedido.has(c.clean) && 
            !pedidos.some(p => String(p['CLIENTE']||p['Cliente']||'').trim().toUpperCase() === c.clean)
        );
    }

    document.getElementById('totalFaltantes').innerText = `${clientesFaltantes.length} pendentes`;

    // Preenche a tabela de clientes da base que ainda não compraram no mês
    let tbody = document.getElementById('tabelaFaltantesBody');
    tbody.innerHTML = '';
    
    if (clientesFaltantes.length === 0) {
        tbody.innerHTML = `<tr><td colspan="4" class="px-6 py-4 text-center text-gray-500">Nenhum cliente pendente ou aba de carteira não identificada na planilha.</td></tr>`;
    } else {
        clientesFaltantes.forEach(cli => {
            let tr = document.createElement('tr');
            tr.innerHTML = `
                <td class="px-6 py-4 whitespace-nowrap text-gray-500">${cli.codigo}</td>
                <td class="px-6 py-4 whitespace-nowrap font-medium text-gray-900">${cli.nome}</td>
                <td class="px-6 py-4 whitespace-nowrap text-gray-500">${cli.cidade} / ${cli.uf}</td>
                <td class="px-6 py-4 whitespace-nowrap"><span class="px-2 inline-flex text-xs leading-5 font-semibold rounded-full bg-yellow-100 text-yellow-800">${cli.status}</span></td>
            `;
            tbody.appendChild(tr);
        });
    }

    // Atualiza os gráficos visuais
    atualizarGraficos(qtdPositivados, Math.max(0, metaPositivacaoQtd - qtdPositivados), totalGravados, totalPedidos - totalGravados);
}

function atualizarGraficos(posiAtv, posiFalta, gravSim, gravNao) {
    const ctxPos = document.getElementById('chartPositivacao').getContext('2d');
    if (chartPosObj) chartPosObj.destroy();
    chartPosObj = new Chart(ctxPos, {
        type: 'doughnut',
        data: {
            labels: ['Positivados', 'Faltam para Meta'],
            datasets: [{ data: [posiAtv, posiFalta], backgroundColor: ['#3b82f6', '#e5e7eb'] }]
        },
        options: { responsive: true, maintainAspectRatio: false }
    });

    const ctxGrav = document.getElementById('chartGravacao').getContext('2d');
    if (chartGravObj) chartGravObj.destroy();
    chartGravObj = new Chart(ctxGrav, {
        type: 'pie',
        data: {
            labels: ['Com Gravação', 'Sem Gravação'],
            datasets: [{ data: [gravSim, gravNao], backgroundColor: ['#10b981', '#f3f4f6'] }]
        },
        options: { responsive: true, maintainAspectRatio: false }
    });
}