const CORES = {
    roxo: { fundo: "rgba(153, 102, 255, 0.5)", borda: "rgba(153, 102, 255, 1)" },
    vermelho: { fundo: "rgba(255, 99, 132, 0.5)", borda: "rgba(255, 99, 132, 1)" },
    azul: { fundo: "rgba(54, 162, 235, 0.5)", borda: "rgba(54, 162, 235, 1)" },
    verde: { fundo: "rgba(75, 192, 192, 0.5)", borda: "rgba(75, 192, 192, 1)" },
    laranja: { fundo: "rgba(255, 159, 64, 0.5)", borda: "rgba(255, 159, 64, 1)" }
};

// 📦 Importa a URL do arquivo url.js
const excelUrl = "https://fundacaooswaldoaranha-my.sharepoint.com/personal/wesley_balbino_foa_org_br/_layouts/15/download.aspx?share=EdsT2JkTPstFhYTAoyB0kWwB0T83o-R9AR4Wu2Yex8hxBw";

let dadosExcel = [];

const charts = {}; // Armazena instâncias de gráficos

// 🆕 Função para verificar se deve atualizar os dados baseado no horário
function deveAtualizarDados() {
    const agora = new Date();
    const horas = agora.getHours();
    const minutos = agora.getMinutes();
    const horaAtual = horas + (minutos / 60);

    // Converte 08:15 e 16:15 para formato decimal
    const horarioManha = 8 + (15 / 60); // 8.25
    const horarioTarde = 16 + (15 / 60); // 16.25

    // Verifica se passou dos horários de atualização
    return horaAtual >= horarioManha || horaAtual >= horarioTarde;
}

// 🆕 Função para verificar se os dados foram atualizados hoje nos horários corretos
function dadosAtualizadosHoje() {
    const ultimaAtualizacao = localStorage.getItem("ultimaAtualizacaoExcel");
    if (!ultimaAtualizacao) return false;

    const dataUltimaAtualizacao = new Date(ultimaAtualizacao);
    const hoje = new Date();

    // Verifica se foi atualizado hoje
    const mesmodia = dataUltimaAtualizacao.toDateString() === hoje.toDateString();

    if (!mesmodia) return false;

    const horasUltimaAtualizacao = dataUltimaAtualizacao.getHours();
    const minutosUltimaAtualizacao = dataUltimaAtualizacao.getMinutes();
    const horaUltimaAtualizacao = horasUltimaAtualizacao + (minutosUltimaAtualizacao / 60);

    // Verifica se a última atualização foi após um dos horários de corte
    const horarioManha = 8 + (15 / 60);
    const horarioTarde = 16 + (15 / 60);

    return horaUltimaAtualizacao >= horarioManha || horaUltimaAtualizacao >= horarioTarde;
}

function exibirUltimaAtualizacao() {
    const el = document.getElementById("infoAtualizacao");
    if (!el) return;

    const raw = localStorage.getItem("ultimaAtualizacaoExcel");

    if (!raw) {
        el.innerHTML = `<i class="material-icons">update</i> <em>Sem dados</em>`;
        return;
    }

    try {
        const data = new Date(raw);
        const dataFormatada = data.toLocaleDateString("pt-BR", {
            day: "2-digit",
            month: "long",
            year: "numeric"
        });

        const hora = data.toLocaleTimeString("pt-BR", {
            hour: "2-digit",
            minute: "2-digit"
        });

        el.innerHTML = `
        <span title="${raw}">
            <i class="material-icons">update</i>
            Atualizado em <strong>${dataFormatada}</strong> às <strong>${hora}</strong>
        </span>
        `;
    } catch (e) {
        console.warn("Erro ao formatar data:", e);
        el.innerHTML = `<i class="material-icons">error</i> <em>Erro na data</em>`;
    }
}




function converterDataExcelParaPtBR(valorData) {
    if (!valorData || valorData === "-" || valorData === "") return valorData;

    // Se já estiver em formato de string pt-BR, retorna como está
    if (typeof valorData === 'string' && valorData.includes('/')) {
        return valorData;
    }

    // Se for número serial do Excel, converte
    if (typeof valorData === 'number' && valorData > 0) {
        try {
            // Excel conta dias desde 01/01/1900, mas com bug do ano 1900
            // JavaScript conta milissegundos desde 01/01/1970
            const diasDesde1900 = valorData - 25569; // Ajuste para JavaScript
            const data = new Date(diasDesde1900 * 86400 * 1000);

            // Verifica se a data é válida
            if (isNaN(data.getTime())) {
                console.warn(`Data inválida do Excel: ${valorData}`);
                return valorData;
            }

            // Formatar para DD/MM/AAAA HH:MM
            const dia = String(data.getDate()).padStart(2, '0');
            const mes = String(data.getMonth() + 1).padStart(2, '0');
            const ano = data.getFullYear();
            const horas = String(data.getHours()).padStart(2, '0');
            const minutos = String(data.getMinutes()).padStart(2, '0');

            // Se não tem horário específico (00:00), retorna só a data
            if (horas === '00' && minutos === '00') {
                return `${dia}/${mes}/${ano}`;
            } else {
                return `${dia}/${mes}/${ano} ${horas}:${minutos}`;
            }
        } catch (e) {
            console.warn(`Erro ao converter data do Excel: ${valorData}`, e);
            return valorData;
        }
    }

    // Tenta converter outros formatos de data
    if (valorData instanceof Date) {
        const dia = String(valorData.getDate()).padStart(2, '0');
        const mes = String(valorData.getMonth() + 1).padStart(2, '0');
        const ano = valorData.getFullYear();
        const horas = String(valorData.getHours()).padStart(2, '0');
        const minutos = String(valorData.getMinutes()).padStart(2, '0');

        if (horas === '00' && minutos === '00') {
            return `${dia}/${mes}/${ano}`;
        } else {
            return `${dia}/${mes}/${ano} ${horas}:${minutos}`;
        }
    }

    // Se for string, tenta converter
    if (typeof valorData === 'string') {
        try {
            const data = new Date(valorData);
            if (!isNaN(data.getTime())) {
                const dia = String(data.getDate()).padStart(2, '0');
                const mes = String(data.getMonth() + 1).padStart(2, '0');
                const ano = data.getFullYear();
                const horas = String(data.getHours()).padStart(2, '0');
                const minutos = String(data.getMinutes()).padStart(2, '0');

                if (horas === '00' && minutos === '00') {
                    return `${dia}/${mes}/${ano}`;
                } else {
                    return `${dia}/${mes}/${ano} ${horas}:${minutos}`;
                }
            }
        } catch (e) {
            // Se não conseguir converter, retorna o valor original
        }
    }

    return valorData;
}

async function carregarExcel() {
    if (dadosExcel.length && dadosAtualizadosHoje()) {
        return dadosExcel;
    }

    // 🆕 Verifica se deve forçar atualização baseado no horário
    const forcarAtualizacao = deveAtualizarDados() && !dadosAtualizadosHoje();

    const dadosSalvos = localStorage.getItem("dadosExcel");
    if (dadosSalvos && !forcarAtualizacao) {
        dadosExcel = JSON.parse(dadosSalvos);
        console.log(`✅ Dados carregados do localStorage: ${dadosExcel.length} registros`);
        return dadosExcel;
    }

    console.log("🔄 Carregando dados do Excel...");

    try {
        const response = await fetch(excelUrl);
        const blob = await response.blob();
        const buffer = await blob.arrayBuffer();

        const workbook = XLSX.read(buffer, { type: "array" });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        let dadosBrutos = XLSX.utils.sheet_to_json(worksheet, {
            range: 5,
            defval: ""
        });

        console.log(`📊 Dados brutos carregados: ${dadosBrutos.length} registros`);

        // 🆕 Converte todas as datas para formato pt-BR antes de salvar
        dadosExcel = dadosBrutos.map((registro, index) => {
            const registroConvertido = { ...registro };

            // Lista de campos que podem conter datas
            const camposData = [
                "Data do agendamento",
                "Data de Criação",
                "Data de Conclusão",
                "Data de Vencimento",
                "Data da Audiência",
                "Data do Protocolo"
            ];

            // Converte cada campo de data encontrado
            camposData.forEach(campo => {
                if (registroConvertido[campo]) {
                    const valorOriginal = registroConvertido[campo];
                    const valorConvertido = converterDataExcelParaPtBR(valorOriginal);

                    registroConvertido[campo] = valorConvertido;

                    // Log das primeiras 5 conversões para debug
                    if (index < 5 && valorOriginal !== valorConvertido) {
                        console.log(`🔄 [${index + 1}] ${campo}: "${valorOriginal}" -> "${valorConvertido}"`);
                    }
                }
            });

            return registroConvertido;
        });

        // 🆕 Salva os dados já convertidos e marca o timestamp da atualização
        localStorage.setItem("dadosExcel", JSON.stringify(dadosExcel));
        localStorage.setItem("ultimaAtualizacaoExcel", new Date().toISOString());

        console.log("✅ Dados atualizados e convertidos com sucesso!");
        console.log(`📊 Total de registros processados: ${dadosExcel.length}`);

        // Debug das primeiras datas convertidas
        console.log("🔍 Primeiras 5 datas convertidas:");
        dadosExcel.slice(0, 5).forEach((item, i) => {
            const data = item["Data do agendamento"];
            console.log(`  ${i + 1}. "${data}"`);
        });

    } catch (error) {
        console.error("❌ Erro ao carregar dados do Excel:", error);
        throw error;
    }

    return dadosExcel;
}

// 🆕 Função simplificada para filtrar dados até hoje (datas já estão em formato pt-BR)
function filtrarDadosAteHoje(dados) {
    const hoje = new Date();
    hoje.setHours(23, 59, 59, 999); // Final do dia

    console.log(`🔍 Filtrando dados até: ${hoje.toLocaleDateString('pt-BR')}`);

    const dadosFiltrados = dados.filter(item => {
        const dataStr = item["Data do agendamento"];

        // Ignora registros sem data ou com data vazia
        if (!dataStr || dataStr === "" || dataStr === "-") return false;

        let dataAgendamento = null;

        // Como as datas já estão em formato pt-BR (DD/MM/YYYY), processa diretamente
        if (typeof dataStr === 'string' && dataStr.includes('/')) {
            // Remove a parte do horário se existir (DD/MM/YYYY HH:MM)
            const parteData = dataStr.split(' ')[0];
            const partes = parteData.split('/');

            if (partes.length === 3) {
                const dia = parseInt(partes[0]);
                const mes = parseInt(partes[1]) - 1; // Mês em JS é 0-11
                const ano = parseInt(partes[2]);

                // Validação básica dos valores
                if (dia >= 1 && dia <= 31 && mes >= 0 && mes <= 11 && ano >= 2000) {
                    dataAgendamento = new Date(ano, mes, dia);
                }
            }
        }

        // Se não conseguiu parsear, tenta outros formatos (fallback)
        if (!dataAgendamento && dataStr) {
            try {
                dataAgendamento = new Date(dataStr);
            } catch (e) {
                console.warn(`Erro ao parsear data: ${dataStr}`);
                return false;
            }
        }

        // Verifica se a data é válida
        if (!dataAgendamento || isNaN(dataAgendamento.getTime())) {
            console.warn(`Data inválida encontrada: "${dataStr}"`);
            return false;
        }

        // Verifica se a data é até hoje (inclusive)
        const dataValida = dataAgendamento <= hoje;

        return dataValida;
    });

    console.log(`📊 Total de registros originais: ${dados.length}`);
    console.log(`📊 Registros filtrados até hoje: ${dadosFiltrados.length}`);

    // Debug: mostra as 5 primeiras datas filtradas
    console.log('✅ Primeiras 5 datas que passaram no filtro:');
    dadosFiltrados.slice(0, 5).forEach((item, i) => {
        console.log(`  ${i + 1}. "${item["Data do agendamento"]}"`);
    });

    // Debug: mostra 5 datas que foram rejeitadas (futuras)
    const rejeitados = dados.filter(item => {
        const dataStr = item["Data do agendamento"];
        if (!dataStr || dataStr === "" || dataStr === "-") return false;

        if (typeof dataStr === 'string' && dataStr.includes('/')) {
            const parteData = dataStr.split(' ')[0];
            const partes = parteData.split('/');
            if (partes.length === 3) {
                const dia = parseInt(partes[0]);
                const mes = parseInt(partes[1]) - 1;
                const ano = parseInt(partes[2]);
                const data = new Date(ano, mes, dia);
                return data > hoje;
            }
        }
        return false;
    });

    if (rejeitados.length > 0) {
        console.log('❌ Primeiras 5 datas futuras rejeitadas:');
        rejeitados.slice(0, 5).forEach((item, i) => {
            console.log(`  ${i + 1}. "${item["Data do agendamento"]}"`);
        });
    }

    return dadosFiltrados;
}
function gerarGraficoPorColuna(coluna, dados, canvasId, cor = CORES.roxo, filtrarAteHoje = false) {
    // 🆕 Aplica filtro de data se solicitado
    let dadosFiltrados = dados;
    if (filtrarAteHoje) {
        dadosFiltrados = filtrarDadosAteHoje(dados);
        console.log(`🎯 Gráfico "${coluna}" - Total original: ${dados.length}`);
        console.log(`🎯 Gráfico "${coluna}" - Filtrado até hoje: ${dadosFiltrados.length} registros`);
    }

    const contagem = {};
    dadosFiltrados.forEach(item => {
        const chave = item[coluna];
        if (chave && chave.trim() && chave.trim() !== "-") {
            const chaveLimpa = chave.trim();
            contagem[chaveLimpa] = (contagem[chaveLimpa] || 0) + 1;
        }
    });

    const nomesCompletos = Object.keys(contagem);
    const valores = Object.values(contagem);

    // Para responsável, mostra apenas o primeiro nome no gráfico
    const labels = nomesCompletos.map(nome => {
        return coluna.toLowerCase() === "responsável" ? nome.split(" ")[0] : nome;
    });

    const ctx = document.getElementById(canvasId).getContext("2d");

    const chart = new Chart(ctx, {
        type: "bar",
        data: {
            labels: labels,
            datasets: [{
                label: `Quantidade por ${coluna}${filtrarAteHoje ? ' (até 01/08/2025)' : ''}`,
                data: valores,
                backgroundColor: cor.fundo,
                borderColor: cor.borda,
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            onClick: (e, elements) => {
                if (elements.length === 0) return;

                const index = elements[0].index;
                const valorClicado = chart.data.labels[index];

                let resultados;

                if (coluna.toLowerCase() === "responsável") {
                    // Para responsável, encontra nomes completos que batem com o primeiro nome
                    const nomesCorrespondentes = nomesCompletos.filter(nome => nome.startsWith(valorClicado));
                    resultados = dadosFiltrados.filter(item =>
                        nomesCorrespondentes.includes(item["Responsável"]?.trim())
                    );
                } else {
                    // Para outras colunas, filtra diretamente pelo valor
                    resultados = dadosFiltrados.filter(item => {
                        const valorItem = item[coluna]?.trim();
                        return valorItem === nomesCompletos.find(nome =>
                            (coluna.toLowerCase() === "responsável" ? nome.split(" ")[0] : nome) === valorClicado
                        );
                    });
                }

                console.log(`🔍 Clique no gráfico "${coluna}": ${valorClicado} - ${resultados.length} resultados`);

                // Debug: mostra algumas datas dos resultados
                console.log('📅 Primeiras 5 datas dos resultados:');
                resultados.slice(0, 5).forEach((item, i) => {
                    console.log(`  ${i + 1}. "${item["Data do agendamento"]}"`);
                });

                exibirTabela(canvasId, resultados);
            },
            plugins: {
                datalabels: {
                    anchor: 'end',
                    align: 'top',
                    color: '#000',
                    font: { weight: 'bold', size: 12 },
                    formatter: Math.round
                },
                tooltip: {
                    callbacks: {
                        title: function (context) {
                            const index = context[0].dataIndex;
                            // Mostra o nome completo no tooltip
                            return nomesCompletos[index];
                        },
                        label: function (context) {
                            return `${context.dataset.label}: ${context.parsed.y}`;
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: `Quantidade${filtrarAteHoje ? ' (até hoje)' : ''}`
                    }
                },
                x: {
                    title: { display: true, text: coluna }
                }
            }
        },
        plugins: [ChartDataLabels]
    });

    return chart;
}

function gerarGraficoEvolucaoMensal(dados, canvasId) {
    const contagemPorMes = {};

    dados.forEach(item => {
        const dataRaw = item["Data do agendamento"];
        const data = new Date(dataRaw);
        if (!isNaN(data)) {
            const chave = `${String(data.getMonth() + 1).padStart(2, '0')}/${data.getFullYear()}`;
            contagemPorMes[chave] = (contagemPorMes[chave] || 0) + 1;
        }
    });

    const labels = Object.keys(contagemPorMes).sort((a, b) => {
        const [mA, yA] = a.split("/").map(Number);
        const [mB, yB] = b.split("/").map(Number);
        return new Date(yA, mA - 1) - new Date(yB, mB - 1);
    });

    const valores = labels.map(label => contagemPorMes[label]);

    const ctx = document.getElementById(canvasId).getContext("2d");

    const chart = new Chart(ctx, {
        type: "line",
        data: {
            labels,
            datasets: [{
                label: "Processos por Mês",
                data: valores,
                fill: true,
                borderColor: CORES.azul.borda,
                backgroundColor: CORES.azul.fundo,
                tension: 0.2,
                pointRadius: 5,
                pointHoverRadius: 7
            }]
        },
        options: {
            responsive: true,
            plugins: {
                datalabels: {
                    anchor: 'end',
                    align: 'top',
                    color: '#000',
                    font: {
                        weight: 'bold',
                        size: 12
                    },
                    formatter: Math.round
                },
                tooltip: {
                    callbacks: {
                        label: context => `${context.parsed.y} processo(s)`
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    title: { display: true, text: 'Quantidade' }
                },
                x: {
                    title: { display: true, text: 'Mês/Ano' }
                }
            }
        },
        plugins: [ChartDataLabels]
    });

    return chart;
}

// 🧹 Função para limpar o cache e forçar recarregamento dos dados
function limparCacheERecarregar() {
    console.log("🧹 Limpando cache e recarregando dados...");

    // Remove dados do localStorage
    localStorage.removeItem("dadosExcel");
    localStorage.removeItem("ultimaAtualizacaoExcel");

    // Limpa variável global
    dadosExcel = [];

    console.log("✅ Cache limpo! Recarregue a página para baixar dados atualizados.");

    // Opcionalmente, pode recarregar a página automaticamente:
    // window.location.reload();
}

// 🆕 Função simplificada para formatar data (dados já estão em pt-BR)
function formatarDataExcel(valorData) {
    if (!valorData || valorData === "-" || valorData === "") return "-";

    // Se já estiver em formato pt-BR, retorna como está
    if (typeof valorData === 'string' && valorData.includes('/')) {
        return valorData;
    }

    // Fallback: se por algum motivo ainda vier como número do Excel, converte
    if (typeof valorData === 'number' && valorData > 0) {
        return converterDataExcelParaPtBR(valorData);
    }

    // Fallback: se vier como Date object, converte
    if (valorData instanceof Date && !isNaN(valorData.getTime())) {
        const dia = String(valorData.getDate()).padStart(2, '0');
        const mes = String(valorData.getMonth() + 1).padStart(2, '0');
        const ano = valorData.getFullYear();
        const horas = String(valorData.getHours()).padStart(2, '0');
        const minutos = String(valorData.getMinutes()).padStart(2, '0');

        if (horas === '00' && minutos === '00') {
            return `${dia}/${mes}/${ano}`;
        } else {
            return `${dia}/${mes}/${ano} ${horas}:${minutos}`;
        }
    }

    // Se não conseguir converter, retorna o valor original
    return valorData;
}

// 🆕 Função para obter a descrição preenchida (ordem de prioridade)
function obterDescricao(registro) {
    // Função auxiliar para verificar se um campo está realmente preenchido
    function campoPreenchido(valor) {
        return valor &&
            typeof valor === 'string' &&
            valor.trim() !== '' &&
            valor.trim() !== '-' &&
            valor.trim().toLowerCase() !== 'null' &&
            valor.trim().toLowerCase() !== 'undefined';
    }

    // Ordem de prioridade - retorna o primeiro que estiver preenchido

    // 1ª prioridade: Processo - Descrição
    if (campoPreenchido(registro["Processo - Descrição"])) {
        return registro["Processo - Descrição"].trim();
    }

    // 2ª prioridade: Solicitação - Descrição da Solicitação
    if (campoPreenchido(registro["Solicitação - Descrição da Solicitação"])) {
        return registro["Solicitação - Descrição da Solicitação"].trim();
    }

    // 3ª prioridade: Observação
    if (campoPreenchido(registro["Observação"])) {
        return registro["Observação"].trim();
    }

    // 4ª prioridade: Sub Tipo
    if (campoPreenchido(registro["Sub Tipo"])) {
        return registro["Sub Tipo"].trim();
    }

    // Se nenhum estiver preenchido
    return "-";
}

function exibirTabela(canvasId, registros) {
    // Remove modal anterior se existir
    const modalAnterior = document.getElementById('modal-tabela');
    if (modalAnterior) {
        modalAnterior.remove();
    }

    // Cria o modal
    const modal = document.createElement('div');
    modal.id = 'modal-tabela';
    modal.className = 'modal modal-fixed-footer';
    modal.style.maxHeight = '80%';

    modal.innerHTML = `
        <div class="modal-content">
            <h4>Detalhes dos Processos</h4>
            <div style="max-height: calc(80vh - 150px); overflow-y: auto;">
                <table class="striped highlight responsive-table">
                    <thead>
                        <tr>
                            <th>ID do Processo</th>
                            <th>Data do Agendamento</th>
                            <th>Tipo</th>
                            <th>Status da Tarefa</th>
                            <th>Descrição</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${registros.map(reg => `
                            <tr>
                                <td>${reg["Processo - ID"] || "-"}</td>
                                <td>${formatarDataExcel(reg["Data do agendamento"])}</td>
                                <td>${reg["Tipo"] || "-"}</td>
                                <td>${reg["Status da tarefa"] || "-"}</td>
                                <td>${obterDescricao(reg)}</td>
                            </tr>
                        `).join("")}
                    </tbody>
                </table>
            </div>
        </div>
        <div class="modal-footer">
            <a href="#!" class="modal-close waves-effect waves-green btn-flat">Fechar</a>
        </div>
    `;

    // Adiciona o modal ao body
    document.body.appendChild(modal);

    // Inicializa e abre o modal
    const modalInstance = M.Modal.init(modal, {
        dismissible: true,
        opacity: 0.5,
        inDuration: 300,
        outDuration: 200
    });

    modalInstance.open();
}

async function adicionarGrafico(coluna, filtrarAteHoje = false) {
    const dados = await carregarExcel();

    const idUnico = `grafico_${Math.random().toString(36).substr(2, 9)}`;
    const container = document.createElement("div");
    container.className = "grafico-container";

    const tituloExtra = filtrarAteHoje ? " (até hoje)" : "";

    container.innerHTML = `
        <div class="grafico-header">
            <strong>Gráfico por ${coluna}${tituloExtra}</strong>
            <select onchange="trocarCor('${idUnico}', '${coluna}', this.value)">
                ${Object.keys(CORES).map(cor => `<option value="${cor}">${cor[0].toUpperCase() + cor.slice(1)}</option>`).join('')}
            </select>
        </div>
        <canvas id="${idUnico}"></canvas>
    `;

    document.getElementById("graficos").appendChild(container);
    const chart = gerarGraficoPorColuna(coluna, dados, idUnico, CORES.roxo, filtrarAteHoje);
    charts[idUnico] = { chart, coluna };
}

function trocarCor(canvasId, coluna, corSelecionada) {
    const { chart } = charts[canvasId];
    const novaCor = CORES[corSelecionada];

    chart.data.datasets[0].backgroundColor = novaCor.fundo;
    chart.data.datasets[0].borderColor = novaCor.borda;
    chart.update();
}

// 🍕 Função para gerar gráfico de pizza do Status das Tarefas
function gerarGraficoPizza(dados, canvasId) {
    // Conta a ocorrência de cada status
    const contagem = {};
    dados.forEach(item => {
        const status = item["Status da tarefa"];
        if (status && status.trim() && status.trim() !== "-") {
            const statusLimpo = status.trim();
            contagem[statusLimpo] = (contagem[statusLimpo] || 0) + 1;
        }
    });

    const labels = Object.keys(contagem);
    const valores = Object.values(contagem);

    // Cores para o gráfico de pizza
    const coresPizza = [
        '#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0', '#9966FF',
        '#FF9F40', '#FF6384', '#C9CBCF', '#4BC0C0', '#FF6384'
    ];

    const ctx = document.getElementById(canvasId).getContext("2d");

    const chart = new Chart(ctx, {
        type: "pie",
        data: {
            labels: labels,
            datasets: [{
                label: 'Status das Tarefas',
                data: valores,
                backgroundColor: coresPizza.slice(0, labels.length),
                borderColor: '#fff',
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            plugins: {
                legend: {
                    position: 'right',
                    labels: {
                        padding: 20,
                        usePointStyle: true,
                        font: { size: 12 }
                    }
                },
                datalabels: {
                    color: '#fff',
                    font: { weight: 'bold', size: 14 },
                    formatter: (value, context) => {
                        const total = context.dataset.data.reduce((a, b) => a + b, 0);
                        const porcentagem = ((value / total) * 100).toFixed(1);
                        return `${value}\n(${porcentagem}%)`;
                    },
                    textAlign: 'center'
                },
                tooltip: {
                    callbacks: {
                        label: function (context) {
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const porcentagem = ((context.parsed / total) * 100).toFixed(1);
                            return `${context.label}: ${context.parsed} (${porcentagem}%)`;
                        }
                    }
                }
            },
            onClick: (e, elements) => {
                if (elements.length === 0) return;

                const index = elements[0].index;
                const statusClicado = chart.data.labels[index];

                // Filtra registros pelo status clicado
                const resultados = dados.filter(item =>
                    item["Status da tarefa"] &&
                    item["Status da tarefa"].trim() === statusClicado
                );

                exibirTabela(canvasId, resultados);
            }
        },
        plugins: [ChartDataLabels]
    });

    return chart;
}

async function adicionarGraficoPizza() {
    const dados = await carregarExcel();

    const idUnico = `pizza_status_${Math.random().toString(36).substr(2, 9)}`;
    const container = document.createElement("div");
    container.className = "grafico-container";
    container.innerHTML = `
        <div class="grafico-header">
            <strong>📊 Gráfico de Pizza - Status das Tarefas</strong>
        </div>
        <div style="width: 100%; height: 400px; display: flex; justify-content: center; align-items: center;">
            <canvas id="${idUnico}" style="max-width: 600px; max-height: 400px;"></canvas>
        </div>
    `;

    document.getElementById("graficos").appendChild(container);
    const chart = gerarGraficoPizza(dados, idUnico);
    charts[idUnico] = { chart, coluna: "Status da tarefa" };
}

function gerarTabelaAudiencias(dados) {
    const audiencias = dados.filter(item => {
        const tipo = item["Tipo"];
        return tipo && (
            tipo.toLowerCase().includes("audiência") ||
            tipo.toLowerCase().includes("audiencia") ||
            tipo.toLowerCase().includes("hearing")
        );
    });

    audiencias.sort((a, b) => {
        const dataA = new Date(a["Data do agendamento"]);
        const dataB = new Date(b["Data do agendamento"]);
        return dataA - dataB;
    });

    const container = document.createElement("div");
    container.className = "row";
    container.id = "audiencias";
    container.innerHTML = `
        <div class="col s12">
            <div class="card">
                <div class="card-content">
                    <span class="card-title">📅 Audiências Agendadas</span>
                    <div style="max-height: 400px; overflow-y: auto;">
                        <table class="striped highlight responsive-table">
                            <thead>
                                <tr>
                                    <th>Processo ID</th>
                                    <th>Data do Agendamento</th>
                                    <th>Empresa</th>
                                    <th>Responsável</th>
                                    <th>Tipo</th>
                                    <th>Status</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${audiencias.map(reg => `
                                    <tr>
                                        <td><strong>${reg["Processo - ID"] || "-"}</strong></td>
                                        <td><strong>${formatarDataExcel(reg["Data do agendamento"])}</strong></td>
                                        <td>${reg["Empresa"] || "-"}</td>
                                        <td>${reg["Responsável"] || "-"}</td>
                                        <td><span class="chip blue white-text">${reg["Tipo"] || "-"}</span></td>
                                        <td><span class="chip ${obterCorStatus(reg["Status da tarefa"])}">${reg["Status da tarefa"] || "-"}</span></td>
                                    </tr>
                                `).join("")}
                            </tbody>
                        </table>
                    </div>
                    <div class="card-action">
                        <span><strong>Total de audiências:</strong> ${audiencias.length}</span>
                    </div>
                </div>
            </div>
        </div>
    `;

    return container;
}

function gerarTabelaPrazosFatais(dados) {
    const hoje = new Date();
    hoje.setHours(23, 59, 59, 999); // Final do dia de hoje

    const seteDiasDepois = new Date(hoje);
    seteDiasDepois.setDate(hoje.getDate() + 7); // 7 dias a partir de hoje

    console.log(`🚨 Analisando prazos fatais:`);
    console.log(`   📅 Hoje: ${hoje.toLocaleDateString('pt-BR')}`);
    console.log(`   📅 Limite (7 dias): ${seteDiasDepois.toLocaleDateString('pt-BR')}`);

    // Função auxiliar para verificar se um campo indica "Sim" para prazo fatal
    const temPrazoFatal = (valor) => {
        if (!valor) return false;
        const valorLower = valor.toString().toLowerCase().trim();
        return valorLower === 'sim' ||
            valorLower === 's' ||
            valorLower === 'yes' ||
            valorLower === 'y' ||
            valorLower === 'true' ||
            valorLower === '1';
    };

    // Função auxiliar para converter data pt-BR para objeto Date
    const converterDataPtBR = (dataStr) => {
        if (!dataStr || dataStr === "-" || dataStr === "") return null;

        if (typeof dataStr === 'string' && dataStr.includes('/')) {
            const parteData = dataStr.split(' ')[0]; // Remove horário se existir
            const partes = parteData.split('/');

            if (partes.length === 3) {
                const dia = parseInt(partes[0]);
                const mes = parseInt(partes[1]) - 1; // Mês em JS é 0-11
                const ano = parseInt(partes[2]);

                if (dia >= 1 && dia <= 31 && mes >= 0 && mes <= 11 && ano >= 2000) {
                    return new Date(ano, mes, dia);
                }
            }
        }

        return null;
    };

    // Filtra registros com prazo fatal
    const prazosFatais = dados.filter(item => {
        // Verifica campos que podem indicar prazo fatal
        const campos = [
            "Solicitação - Há Prazo Fatal",
            "Há Prazo Fatal",
            "Prazo Fatal",
            "Prazo Crítico",
            "Urgente"
        ];

        const temPrazo = campos.some(campo => temPrazoFatal(item[campo]));

        if (!temPrazo) return false;

        // Verifica a data de vencimento
        const camposData = [
            "Data do agendamento",
            "Data de Vencimento",
            "Data Limite",
            "Prazo"
        ];

        let dataVencimento = null;

        // Procura a primeira data válida nos campos
        for (const campo of camposData) {
            if (item[campo]) {
                dataVencimento = converterDataPtBR(item[campo]);
                if (dataVencimento) break;
            }
        }

        if (!dataVencimento) return false;

        // Inclui se está atrasado (data já passou) ou vence nos próximos 7 dias
        const estaAtrasado = dataVencimento < hoje;
        const venceEm7Dias = dataVencimento >= hoje && dataVencimento <= seteDiasDepois;

        return estaAtrasado || venceEm7Dias;
    });

    // Separa atrasados e próximos do vencimento
    const atrasados = prazosFatais.filter(item => {
        const camposData = ["Data do agendamento", "Data de Vencimento", "Data Limite", "Prazo"];
        let dataVencimento = null;

        for (const campo of camposData) {
            if (item[campo]) {
                dataVencimento = converterDataPtBR(item[campo]);
                if (dataVencimento) break;
            }
        }

        return dataVencimento && dataVencimento < hoje;
    });

    const proximosVencimento = prazosFatais.filter(item => {
        const camposData = ["Data do agendamento", "Data de Vencimento", "Data Limite", "Prazo"];
        let dataVencimento = null;

        for (const campo of camposData) {
            if (item[campo]) {
                dataVencimento = converterDataPtBR(item[campo]);
                if (dataVencimento) break;
            }
        }

        return dataVencimento && dataVencimento >= hoje && dataVencimento <= seteDiasDepois;
    });

    // Ordena por data (mais urgentes primeiro)
    prazosFatais.sort((a, b) => {
        const getDataVencimento = (item) => {
            const camposData = ["Data do agendamento", "Data de Vencimento", "Data Limite", "Prazo"];
            for (const campo of camposData) {
                if (item[campo]) {
                    const data = converterDataPtBR(item[campo]);
                    if (data) return data;
                }
            }
            return new Date(0); // Data muito antiga se não encontrar
        };

        return getDataVencimento(a) - getDataVencimento(b);
    });

    console.log(`🚨 Prazos fatais encontrados: ${prazosFatais.length}`);
    console.log(`   ❌ Atrasados: ${atrasados.length}`);
    console.log(`   ⚠️ Vencem em 7 dias: ${proximosVencimento.length}`);

    // Função para obter a classe CSS baseada no status do prazo
    const getClassePrazo = (item) => {
        const camposData = ["Data do agendamento", "Data de Vencimento", "Data Limite", "Prazo"];
        let dataVencimento = null;

        for (const campo of camposData) {
            if (item[campo]) {
                dataVencimento = converterDataPtBR(item[campo]);
                if (dataVencimento) break;
            }
        }

        if (!dataVencimento) return "yellow lighten-4";

        if (dataVencimento < hoje) {
            return "red lighten-4"; // Atrasado
        } else if (dataVencimento <= seteDiasDepois) {
            return "orange lighten-4"; // Vence em breve
        } else {
            return "green lighten-4"; // OK
        }
    };

    // Função para obter o texto do status
    const getStatusPrazo = (item) => {
        const camposData = ["Data do agendamento", "Data de Vencimento", "Data Limite", "Prazo"];
        let dataVencimento = null;

        for (const campo of camposData) {
            if (item[campo]) {
                dataVencimento = converterDataPtBR(item[campo]);
                if (dataVencimento) break;
            }
        }

        if (!dataVencimento) return "⚠️ Sem data";

        const diffDias = Math.ceil((dataVencimento - hoje) / (1000 * 60 * 60 * 24));

        if (diffDias < 0) {
            return `🔴 ${Math.abs(diffDias)} dia(s) atrasado`;
        } else if (diffDias === 0) {
            return "🟡 Vence hoje";
        } else if (diffDias <= 7) {
            return `🟠 Vence em ${diffDias} dia(s)`;
        } else {
            return `🟢 Vence em ${diffDias} dia(s)`;
        }
    };

    const container = document.createElement("div");
    container.className = "row";
    container.id = "prazos-fatais";
    container.innerHTML = `
        <div class="col s12">
            <div class="card">
                <div class="card-content">
                    <span class="card-title red-text">🚨 Prazos Fatais</span>
                    <p class="grey-text">
                        <strong>Critérios:</strong> Processos com prazo fatal = "Sim" que estão atrasados ou vencem nos próximos 7 dias
                    </p>
                    <div style="max-height: 500px; overflow-y: auto;">
                        <table class="striped highlight responsive-table">
                            <thead>
                                <tr>
                                    <th>Status do Prazo</th>
                                    <th>Processo ID</th>
                                    <th>Data de Vencimento</th>
                                    <th>Responsável</th>
                                    <th>Empresa</th>
                                    <th>Tipo</th>
                                    <th>Status da Tarefa</th>
                                    <th>Descrição</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${prazosFatais.map(reg => {
        const camposData = ["Data do agendamento", "Data de Vencimento", "Data Limite", "Prazo"];
        let dataVencimento = "Não informada";

        for (const campo of camposData) {
            if (reg[campo]) {
                dataVencimento = reg[campo];
                break;
            }
        }

        return `
                                        <tr class="${getClassePrazo(reg)}">
                                            <td><strong>${getStatusPrazo(reg)}</strong></td>
                                            <td><strong>${reg["Processo - ID"] || "-"}</strong></td>
                                            <td><strong>${dataVencimento}</strong></td>
                                            <td>${reg["Responsável"] || "-"}</td>
                                            <td>${reg["Empresa"] || "-"}</td>
                                            <td><span class="chip blue white-text">${reg["Tipo"] || "-"}</span></td>
                                            <td><span class="chip ${obterCorStatus(reg["Status da tarefa"])}">${reg["Status da tarefa"] || "-"}</span></td>
                                            <td>${obterDescricao(reg)}</td>
                                        </tr>
                                    `;
    }).join("")}
                            </tbody>
                        </table>
                    </div>
                    <div class="card-action">
                        <div class="row">
                            <div class="col s12 m4">
                                <span><strong class="red-text">🔴 Atrasados:</strong> ${atrasados.length}</span>
                            </div>
                            <div class="col s12 m4">
                                <span><strong class="orange-text">🟠 Vencem em 7 dias:</strong> ${proximosVencimento.length}</span>
                            </div>
                            <div class="col s12 m4">
                                <span><strong>📊 Total:</strong> ${prazosFatais.length}</span>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    `;

    return container;
}

function gerarRadarResponsaveis(dados, canvasId) {
    const contagem = {};

    dados.forEach(item => {
        const nome = item["Responsável"];
        if (!nome || nome.trim() === "") return;
        contagem[nome.trim()] = (contagem[nome.trim()] || 0) + 1;
    });

    const labels = Object.keys(contagem);
    const valores = Object.values(contagem);

    const ctx = document.getElementById(canvasId).getContext("2d");

    const chart = new Chart(ctx, {
        type: 'radar',
        data: {
            labels,
            datasets: [{
                label: "Carga por Responsável",
                data: valores,
                backgroundColor: CORES.laranja.fundo,
                borderColor: CORES.laranja.borda,
                borderWidth: 2,
                pointBackgroundColor: "#fff"
            }]
        },
        options: {
            responsive: true,
            plugins: {
                tooltip: {
                    callbacks: {
                        label: context => `${context.label}: ${context.formattedValue} processo(s)`
                    }
                }
            },
            scales: {
                r: {
                    beginAtZero: true,
                    pointLabels: {
                        font: { size: 12 }
                    }
                }
            }
        }
    });

    return chart;
}

async function adicionarGraficoRadarResponsaveis() {
    const dados = await carregarExcel();
    const id = "graficoRadarResponsaveis";

    if (document.getElementById(id)) return;

    const container = document.createElement("div");
    container.className = "grafico-container";
    container.innerHTML = `
        <div class="grafico-header">
            <strong>📊 Distribuição da Carga por Responsável</strong>
        </div>
        <canvas id="${id}" style="max-height: 450px;"></canvas>
    `;

    document.getElementById("graficos").appendChild(container);
    gerarRadarResponsaveis(dados, id);
}


// 🆕 Função corrigida para gerar gráfico de pendências com filtro correto
function gerarGraficoPendenciasStacked(dados, canvasId) {
    console.log(`🎯 Gráfico Pendências - Iniciando filtro`);

    const responsaveis = new Set();
    const areas = new Set();
    const mapa = {};
    const registrosPorAreaEResponsavel = {}; // Para armazenar registros para o modal

    // Usa a função filtrarDadosAteHoje para garantir consistência
    const dadosFiltradosPorData = filtrarDadosAteHoje(dados);

    // Filtra apenas os com status pendente
    const dadosFiltrados = dadosFiltradosPorData.filter(item => {
        const status = item["Status da tarefa"];
        const statusPendente = status && (
            status.toLowerCase().includes("pendente") ||
            status.toLowerCase().includes("ativo") ||
            status.toLowerCase().includes("em andamento") ||
            status.toLowerCase().includes("a vencer")
        );

        return statusPendente;
    });

    console.log(`📊 Pendências - Total original: ${dados.length}`);
    console.log(`📊 Pendências - Filtrado por data: ${dadosFiltradosPorData.length}`);
    console.log(`📊 Pendências - Final (data + status): ${dadosFiltrados.length}`);

    dadosFiltrados.forEach(item => {
        const responsavel = item["Responsável"]?.trim();
        const area = item["Área do Direito"]?.trim();

        if (responsavel && area) {
            responsaveis.add(responsavel);
            areas.add(area);

            if (!mapa[area]) mapa[area] = {};
            mapa[area][responsavel] = (mapa[area][responsavel] || 0) + 1;

            // Armazena registros para o modal
            const chave = `${area}_${responsavel}`;
            if (!registrosPorAreaEResponsavel[chave]) {
                registrosPorAreaEResponsavel[chave] = [];
            }
            registrosPorAreaEResponsavel[chave].push(item);
        }
    });

    const responsaveisArray = Array.from(responsaveis).sort();
    const cores = Object.values(CORES);
    const areasArray = Array.from(areas);

    const datasets = areasArray.map((area, idx) => ({
        label: area,
        data: responsaveisArray.map(responsavel => mapa[area]?.[responsavel] || 0),
        backgroundColor: cores[idx % cores.length].fundo,
        borderColor: cores[idx % cores.length].borda,
        borderWidth: 1
    }));

    const ctx = document.getElementById(canvasId).getContext("2d");

    const chart = new Chart(ctx, {
        type: "bar",
        data: {
            labels: responsaveisArray,
            datasets: datasets
        },
        options: {
            responsive: true,
            plugins: {
                tooltip: {
                    mode: "index",
                    intersect: false
                },
                legend: {
                    position: "top"
                }
            },
            interaction: {
                mode: "nearest",
                axis: "x",
                intersect: false
            },
            scales: {
                x: {
                    stacked: true,
                    title: { display: true, text: 'Responsável' }
                },
                y: {
                    stacked: true,
                    beginAtZero: true,
                    title: { display: true, text: 'Quantidade de Pendências (até 01/08/2025)' }
                }
            },
            onClick: (e, elements) => {
                if (elements.length === 0) return;

                const element = elements[0];
                const datasetIndex = element.datasetIndex;
                const index = element.index;

                const area = chart.data.datasets[datasetIndex].label;
                const responsavel = chart.data.labels[index];

                const chave = `${area}_${responsavel}`;
                const registros = registrosPorAreaEResponsavel[chave] || [];

                if (registros.length > 0) {
                    exibirTabela(canvasId, registros);
                }
            }
        }
    });

    return chart;
}

async function adicionarGraficoPendenciasPorAreaEUsuario() {
    const dados = await carregarExcel();
    const id = "graficoPendenciasAreaResponsavel";

    if (document.getElementById(id)) return;

    const container = document.createElement("div");
    container.className = "grafico-container";
    container.innerHTML = `
        <div class="grafico-header">
            <strong>📌 Pendências por Responsável e Área do Direito (até 01/08/2025)</strong>
        </div>
        <canvas id="${id}" style="max-height: 500px;"></canvas>
    `;

    document.getElementById("graficos").appendChild(container);
    gerarGraficoPendenciasStacked(dados, id);
}

// 🎨 Função auxiliar para definir cores dos status
function obterCorStatus(status) {
    if (!status) return "grey";

    const statusLower = status.toLowerCase();

    if (statusLower.includes("ativo") || statusLower.includes("pendente")) {
        return "green white-text";
    } else if (statusLower.includes("concluído") || statusLower.includes("finalizado")) {
        return "blue white-text";
    } else if (statusLower.includes("atrasado") || statusLower.includes("vencido")) {
        return "red white-text";
    } else if (statusLower.includes("aguardando")) {
        return "orange white-text";
    } else {
        return "grey white-text";
    }
}

// 🔍 Função de debug para verificar como as datas estão sendo processadas
function debugDatas(dados) {
    console.log("🔍 === DEBUG DE DATAS ===");

    const hoje = new Date();
    hoje.setHours(23, 59, 59, 999);
    console.log(`📅 Data de referência (hoje): ${hoje.toLocaleDateString('pt-BR')}`);

    // Analisa os primeiros 20 registros
    const amostra = dados.slice(0, 20);

    console.log("\n📊 Análise das primeiras 20 datas:");
    amostra.forEach((item, i) => {
        const dataStr = item["Data do agendamento"];
        let status = "❌ Inválida";
        let dataParsed = null;
        let incluiNoFiltro = false;

        if (dataStr && typeof dataStr === 'string' && dataStr.includes('/')) {
            const parteData = dataStr.split(' ')[0];
            const partes = parteData.split('/');

            if (partes.length === 3) {
                const dia = parseInt(partes[0]);
                const mes = parseInt(partes[1]) - 1;
                const ano = parseInt(partes[2]);

                if (dia >= 1 && dia <= 31 && mes >= 0 && mes <= 11 && ano >= 2000) {
                    dataParsed = new Date(ano, mes, dia);
                    incluiNoFiltro = dataParsed <= hoje;

                    if (incluiNoFiltro) {
                        status = "✅ Incluída";
                    } else {
                        status = "🔮 Futura";
                    }
                }
            }
        }

        console.log(`${i + 1}. "${dataStr}" -> ${status} ${dataParsed ? `(${dataParsed.toLocaleDateString('pt-BR')})` : ''}`);
    });

    // Conta quantas são futuras vs passadas
    const futuras = dados.filter(item => {
        const dataStr = item["Data do agendamento"];
        if (!dataStr || typeof dataStr !== 'string' || !dataStr.includes('/')) return false;

        const parteData = dataStr.split(' ')[0];
        const partes = parteData.split('/');

        if (partes.length === 3) {
            const dia = parseInt(partes[0]);
            const mes = parseInt(partes[1]) - 1;
            const ano = parseInt(partes[2]);

            if (dia >= 1 && dia <= 31 && mes >= 0 && mes <= 11 && ano >= 2000) {
                const data = new Date(ano, mes, dia);
                return data > hoje;
            }
        }
        return false;
    });

    console.log(`\n📈 Resumo:`);
    console.log(`   Total de registros: ${dados.length}`);
    console.log(`   Registros com datas futuras: ${futuras.length}`);
    console.log(`   Registros que devem passar no filtro: ${dados.length - futuras.length}`);

    console.log("🔍 === FIM DEBUG ===\n");
}

// Para usar o debug, chame: debugDatas(dadosExcel) no console do navegador



// 🆕 Função corrigida para atualizar estatísticas
async function atualizarEstatisticas() {
    const dados = await carregarExcel();

    // Total de processos
    document.getElementById('totalProcessos').textContent = dados.length;

    // Total de audiências
    const audiencias = dados.filter(item => {
        const tipo = item["Tipo"];
        return tipo && (
            tipo.toLowerCase().includes("audiência") ||
            tipo.toLowerCase().includes("audiencia")
        );
    });
    document.getElementById('totalAudiencias').textContent = audiencias.length;

    // Tarefas pendentes (até hoje) - usando a função filtrarDadosAteHoje
    const dadosFiltradosPorData = filtrarDadosAteHoje(dados);

    const pendentes = dadosFiltradosPorData.filter(item => {
        const status = item["Status da tarefa"];
        const statusPendente = status && (
            status.toLowerCase().includes("pendente") ||
            status.toLowerCase().includes("ativo") ||
            status.toLowerCase().includes("em andamento") ||
            status.toLowerCase().includes("a vencer")
        );

        return statusPendente;
    });

    console.log(`📊 Estatísticas - Total de dados: ${dados.length}`);
    console.log(`📊 Estatísticas - Filtrados por data (até hoje): ${dadosFiltradosPorData.length}`);
    console.log(`📊 Estatísticas - Pendentes até hoje: ${pendentes.length}`);

    document.getElementById('totalPendentes').textContent = pendentes.length;

    // Prazos fatais (atrasados ou que vencem em 7 dias)
    const hoje = new Date();
    hoje.setHours(23, 59, 59, 999);
    const seteDiasDepois = new Date(hoje);
    seteDiasDepois.setDate(hoje.getDate() + 7);

    // Função auxiliar para verificar se um campo indica "Sim" para prazo fatal
    const temPrazoFatal = (valor) => {
        if (!valor) return false;
        const valorLower = valor.toString().toLowerCase().trim();
        return valorLower === 'sim' ||
            valorLower === 's' ||
            valorLower === 'yes' ||
            valorLower === 'y' ||
            valorLower === 'true' ||
            valorLower === '1';
    };

    // Função auxiliar para converter data pt-BR para objeto Date
    const converterDataPtBR = (dataStr) => {
        if (!dataStr || dataStr === "-" || dataStr === "") return null;

        if (typeof dataStr === 'string' && dataStr.includes('/')) {
            const parteData = dataStr.split(' ')[0];
            const partes = parteData.split('/');

            if (partes.length === 3) {
                const dia = parseInt(partes[0]);
                const mes = parseInt(partes[1]) - 1;
                const ano = parseInt(partes[2]);

                if (dia >= 1 && dia <= 31 && mes >= 0 && mes <= 11 && ano >= 2000) {
                    return new Date(ano, mes, dia);
                }
            }
        }

        return null;
    };

    const prazosFataisCriticos = dados.filter(item => {
        // Verifica se tem prazo fatal marcado como "Sim"
        const campos = [
            "Solicitação - Há Prazo Fatal",
            "Há Prazo Fatal",
            "Prazo Fatal",
            "Prazo Crítico",
            "Urgente"
        ];

        const temPrazo = campos.some(campo => temPrazoFatal(item[campo]));
        if (!temPrazo) return false;

        // Verifica a data de vencimento
        const camposData = [
            "Data do agendamento",
            "Data de Vencimento",
            "Data Limite",
            "Prazo"
        ];

        let dataVencimento = null;

        for (const campo of camposData) {
            if (item[campo]) {
                dataVencimento = converterDataPtBR(item[campo]);
                if (dataVencimento) break;
            }
        }

        if (!dataVencimento) return false;

        // Inclui se está atrasado ou vence nos próximos 7 dias
        return dataVencimento <= seteDiasDepois;
    });

    document.getElementById('totalAtrasados').textContent = prazosFataisCriticos.length;
}

// 🔍 Função de debug para analisar prazos fatais
function debugPrazosFatais(dados) {
    console.log("🚨 === DEBUG PRAZOS FATAIS ===");

    const hoje = new Date();
    hoje.setHours(23, 59, 59, 999);
    const seteDiasDepois = new Date(hoje);
    seteDiasDepois.setDate(hoje.getDate() + 7);

    console.log(`📅 Período de análise: ${hoje.toLocaleDateString('pt-BR')} até ${seteDiasDepois.toLocaleDateString('pt-BR')}`);

    // Campos que podem indicar prazo fatal
    const camposPrazoFatal = [
        "Solicitação - Há Prazo Fatal",
        "Há Prazo Fatal",
        "Prazo Fatal",
        "Prazo Crítico",
        "Urgente"
    ];

    // Campos que podem conter datas de vencimento
    const camposData = [
        "Data do agendamento",
        "Data de Vencimento",
        "Data Limite",
        "Prazo"
    ];

    console.log("\n🔍 Analisando campos de prazo fatal disponíveis:");
    camposPrazoFatal.forEach(campo => {
        const valores = dados
            .map(item => item[campo])
            .filter(valor => valor && valor !== "" && valor !== "-")
            .slice(0, 10); // Primeiros 10 valores únicos

        if (valores.length > 0) {
            console.log(`   ${campo}: ${[...new Set(valores)].join(', ')}`);
        }
    });

    console.log("\n🔍 Analisando campos de data disponíveis:");
    camposData.forEach(campo => {
        const count = dados.filter(item => item[campo] && item[campo] !== "" && item[campo] !== "-").length;
        if (count > 0) {
            console.log(`   ${campo}: ${count} registros com data`);
        }
    });

    // Analisa registros com "Sim" nos campos de prazo fatal
    const comPrazoFatal = dados.filter(item => {
        return camposPrazoFatal.some(campo => {
            const valor = item[campo];
            if (!valor) return false;
            const valorLower = valor.toString().toLowerCase().trim();
            return valorLower === 'sim' ||
                valorLower === 's' ||
                valorLower === 'yes' ||
                valorLower === 'y' ||
                valorLower === 'true' ||
                valorLower === '1';
        });
    });

    console.log(`\n📊 Registros com prazo fatal = "Sim": ${comPrazoFatal.length}`);

    // Mostra alguns exemplos
    console.log("\n🔍 Primeiros 5 exemplos de prazos fatais:");
    comPrazoFatal.slice(0, 5).forEach((item, i) => {
        const processoId = item["Processo - ID"] || "N/A";

        // Encontra qual campo tem o prazo fatal
        const campoComPrazo = camposPrazoFatal.find(campo => {
            const valor = item[campo];
            if (!valor) return false;
            const valorLower = valor.toString().toLowerCase().trim();
            return valorLower === 'sim' || valorLower === 's' || valorLower === 'yes' || valorLower === 'y' || valorLower === 'true' || valorLower === '1';
        });

        // Encontra a data
        const campoComData = camposData.find(campo => item[campo] && item[campo] !== "" && item[campo] !== "-");
        const data = campoComData ? item[campoComData] : "Sem data";

        console.log(`   ${i + 1}. Processo ${processoId}: ${campoComPrazo} = "${item[campoComPrazo]}" | Data: ${data}`);
    });

    console.log("🚨 === FIM DEBUG ===\n");

    return {
        totalComPrazoFatal: comPrazoFatal.length,
        camposEncontrados: {
            prazoFatal: camposPrazoFatal.filter(campo =>
                dados.some(item => item[campo] && item[campo] !== "" && item[campo] !== "-")
            ),
            datas: camposData.filter(campo =>
                dados.some(item => item[campo] && item[campo] !== "" && item[campo] !== "-")
            )
        }
    };
}


// 🆕 Função para filtrar dados dos próximos 7 dias
function filtrarDadosProximos7Dias(dados) {
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0); // Início do dia de hoje

    const seteDiasDepois = new Date(hoje);
    seteDiasDepois.setDate(hoje.getDate() + 7);
    seteDiasDepois.setHours(23, 59, 59, 999); // Final do 7º dia

    console.log(`🔍 Filtrando próximos 7 dias:`);
    console.log(`   📅 A partir de: ${hoje.toLocaleDateString('pt-BR')}`);
    console.log(`   📅 Até: ${seteDiasDepois.toLocaleDateString('pt-BR')}`);

    const dadosFiltrados = dados.filter(item => {
        const dataStr = item["Data do agendamento"];

        // Ignora registros sem data ou com data vazia
        if (!dataStr || dataStr === "" || dataStr === "-") return false;

        let dataAgendamento = null;

        // Como as datas já estão em formato pt-BR (DD/MM/YYYY), processa diretamente
        if (typeof dataStr === 'string' && dataStr.includes('/')) {
            // Remove a parte do horário se existir (DD/MM/YYYY HH:MM)
            const parteData = dataStr.split(' ')[0];
            const partes = parteData.split('/');

            if (partes.length === 3) {
                const dia = parseInt(partes[0]);
                const mes = parseInt(partes[1]) - 1; // Mês em JS é 0-11
                const ano = parseInt(partes[2]);

                // Validação básica dos valores
                if (dia >= 1 && dia <= 31 && mes >= 0 && mes <= 11 && ano >= 2000) {
                    dataAgendamento = new Date(ano, mes, dia);
                }
            }
        }

        // Se não conseguiu parsear, tenta outros formatos (fallback)
        if (!dataAgendamento && dataStr) {
            try {
                dataAgendamento = new Date(dataStr);
            } catch (e) {
                console.warn(`Erro ao parsear data: ${dataStr}`);
                return false;
            }
        }

        // Verifica se a data é válida
        if (!dataAgendamento || isNaN(dataAgendamento.getTime())) {
            console.warn(`Data inválida encontrada: "${dataStr}"`);
            return false;
        }

        // Verifica se a data está nos próximos 7 dias (depois de hoje e dentro do período)
        const dataValida = dataAgendamento > hoje && dataAgendamento <= seteDiasDepois;

        return dataValida;
    });

    console.log(`📊 Total de registros originais: ${dados.length}`);
    console.log(`📊 Registros dos próximos 7 dias: ${dadosFiltrados.length}`);

    // Debug: mostra as 5 primeiras datas filtradas
    console.log('✅ Primeiras 5 datas dos próximos 7 dias:');
    dadosFiltrados.slice(0, 5).forEach((item, i) => {
        console.log(`  ${i + 1}. "${item["Data do agendamento"]}"`);
    });

    return dadosFiltrados;
}

// 🆕 Função para filtrar todas as atividades futuras
function filtrarDadosFuturos(dados) {
    const hoje = new Date();
    hoje.setHours(23, 59, 59, 999); // Final do dia de hoje

    console.log(`🔍 Filtrando todas as atividades futuras a partir de: ${hoje.toLocaleDateString('pt-BR')}`);

    const dadosFiltrados = dados.filter(item => {
        const dataStr = item["Data do agendamento"];

        // Ignora registros sem data ou com data vazia
        if (!dataStr || dataStr === "" || dataStr === "-") return false;

        let dataAgendamento = null;

        // Como as datas já estão em formato pt-BR (DD/MM/YYYY), processa diretamente
        if (typeof dataStr === 'string' && dataStr.includes('/')) {
            // Remove a parte do horário se existir (DD/MM/YYYY HH:MM)
            const parteData = dataStr.split(' ')[0];
            const partes = parteData.split('/');

            if (partes.length === 3) {
                const dia = parseInt(partes[0]);
                const mes = parseInt(partes[1]) - 1; // Mês em JS é 0-11
                const ano = parseInt(partes[2]);

                // Validação básica dos valores
                if (dia >= 1 && dia <= 31 && mes >= 0 && mes <= 11 && ano >= 2000) {
                    dataAgendamento = new Date(ano, mes, dia);
                }
            }
        }

        // Se não conseguiu parsear, tenta outros formatos (fallback)
        if (!dataAgendamento && dataStr) {
            try {
                dataAgendamento = new Date(dataStr);
            } catch (e) {
                console.warn(`Erro ao parsear data: ${dataStr}`);
                return false;
            }
        }

        // Verifica se a data é válida
        if (!dataAgendamento || isNaN(dataAgendamento.getTime())) {
            console.warn(`Data inválida encontrada: "${dataStr}"`);
            return false;
        }

        // Verifica se a data é futura (depois de hoje)
        const dataValida = dataAgendamento > hoje;

        return dataValida;
    });

    console.log(`📊 Total de registros originais: ${dados.length}`);
    console.log(`📊 Registros futuros: ${dadosFiltrados.length}`);

    // Debug: mostra as 5 primeiras datas filtradas
    console.log('✅ Primeiras 5 datas futuras:');
    dadosFiltrados.slice(0, 5).forEach((item, i) => {
        console.log(`  ${i + 1}. "${item["Data do agendamento"]}"`);
    });

    return dadosFiltrados;
}

// 🆕 Função modificada para aceitar filtros personalizados
function gerarGraficoPorColunaComFiltro(coluna, dados, canvasId, cor = CORES.roxo, filtroCallback = null, tituloExtra = "") {
    // Aplica filtro personalizado se fornecido
    let dadosFiltrados = dados;
    if (filtroCallback && typeof filtroCallback === 'function') {
        dadosFiltrados = filtroCallback(dados);
        console.log(`🎯 Gráfico "${coluna}" - Total original: ${dados.length}`);
        console.log(`🎯 Gráfico "${coluna}" - Filtrado: ${dadosFiltrados.length} registros`);
    }

    const contagem = {};
    dadosFiltrados.forEach(item => {
        const chave = item[coluna];
        if (chave && chave.trim() && chave.trim() !== "-") {
            const chaveLimpa = chave.trim();
            contagem[chaveLimpa] = (contagem[chaveLimpa] || 0) + 1;
        }
    });

    const nomesCompletos = Object.keys(contagem);
    const valores = Object.values(contagem);

    // Para responsável, mostra apenas o primeiro nome no gráfico
    const labels = nomesCompletos.map(nome => {
        return coluna.toLowerCase() === "responsável" ? nome.split(" ")[0] : nome;
    });

    const ctx = document.getElementById(canvasId).getContext("2d");

    const chart = new Chart(ctx, {
        type: "bar",
        data: {
            labels: labels,
            datasets: [{
                label: `Quantidade por ${coluna}${tituloExtra}`,
                data: valores,
                backgroundColor: cor.fundo,
                borderColor: cor.borda,
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            onClick: (e, elements) => {
                if (elements.length === 0) return;

                const index = elements[0].index;
                const valorClicado = chart.data.labels[index];

                let resultados;

                if (coluna.toLowerCase() === "responsável") {
                    // Para responsável, encontra nomes completos que batem com o primeiro nome
                    const nomesCorrespondentes = nomesCompletos.filter(nome => nome.startsWith(valorClicado));
                    resultados = dadosFiltrados.filter(item =>
                        nomesCorrespondentes.includes(item["Responsável"]?.trim())
                    );
                } else {
                    // Para outras colunas, filtra diretamente pelo valor
                    resultados = dadosFiltrados.filter(item => {
                        const valorItem = item[coluna]?.trim();
                        return valorItem === nomesCompletos.find(nome =>
                            (coluna.toLowerCase() === "responsável" ? nome.split(" ")[0] : nome) === valorClicado
                        );
                    });
                }

                console.log(`🔍 Clique no gráfico "${coluna}": ${valorClicado} - ${resultados.length} resultados`);

                exibirTabela(canvasId, resultados);
            },
            plugins: {
                datalabels: {
                    anchor: 'end',
                    align: 'top',
                    color: '#000',
                    font: { weight: 'bold', size: 12 },
                    formatter: Math.round
                },
                tooltip: {
                    callbacks: {
                        title: function (context) {
                            const index = context[0].dataIndex;
                            // Mostra o nome completo no tooltip
                            return nomesCompletos[index];
                        },
                        label: function (context) {
                            return `${context.dataset.label}: ${context.parsed.y}`;
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: `Quantidade${tituloExtra}`
                    }
                },
                x: {
                    title: { display: true, text: coluna }
                }
            }
        },
        plugins: [ChartDataLabels]
    });

    return chart;
}

// 🆕 Função para gerar gráfico dos próximos 7 dias
async function adicionarGraficoProximos7Dias(coluna) {
    const dados = await carregarExcel();

    const idUnico = `grafico_7dias_${Math.random().toString(36).substr(2, 9)}`;
    const container = document.createElement("div");
    container.className = "grafico-container";

    const hoje = new Date();
    const seteDiasDepois = new Date(hoje);
    seteDiasDepois.setDate(hoje.getDate() + 7);

    container.innerHTML = `
        <div class="grafico-header">
            <strong>📅 ${coluna} - Próximos 7 Dias (${hoje.toLocaleDateString('pt-BR')} a ${seteDiasDepois.toLocaleDateString('pt-BR')})</strong>
            <select onchange="trocarCor('${idUnico}', '${coluna}', this.value)">
                ${Object.keys(CORES).map(cor => `<option value="${cor}">${cor[0].toUpperCase() + cor.slice(1)}</option>`).join('')}
            </select>
        </div>
        <canvas id="${idUnico}"></canvas>
    `;

    document.getElementById("graficos").appendChild(container);
    const chart = gerarGraficoPorColunaComFiltro(coluna, dados, idUnico, CORES.verde, filtrarDadosProximos7Dias, " (próximos 7 dias)");
    charts[idUnico] = { chart, coluna };
}

// 🆕 Função para gerar gráfico de todas as atividades futuras
async function adicionarGraficoAtividadesFuturas(coluna) {
    const dados = await carregarExcel();

    const idUnico = `grafico_futuras_${Math.random().toString(36).substr(2, 9)}`;
    const container = document.createElement("div");
    container.className = "grafico-container";

    container.innerHTML = `
        <div class="grafico-header">
            <strong>🔮 ${coluna} - Todas as Atividades Futuras</strong>
            <select onchange="trocarCor('${idUnico}', '${coluna}', this.value)">
                ${Object.keys(CORES).map(cor => `<option value="${cor}">${cor[0].toUpperCase() + cor.slice(1)}</option>`).join('')}
            </select>
        </div>
        <canvas id="${idUnico}"></canvas>
    `;

    document.getElementById("graficos").appendChild(container);
    const chart = gerarGraficoPorColunaComFiltro(coluna, dados, idUnico, CORES.azul, filtrarDadosFuturos, " (futuras)");
    charts[idUnico] = { chart, coluna };
}

// 🆕 Função para gerar tabela das próximas atividades (7 dias)
function gerarTabelaProximasAtividades(dados) {
    const proximasAtividades = filtrarDadosProximos7Dias(dados);

    // Ordena por data (mais próximas primeiro)
    proximasAtividades.sort((a, b) => {
        const dataA = new Date(a["Data do agendamento"].split('/').reverse().join('-'));
        const dataB = new Date(b["Data do agendamento"].split('/').reverse().join('-'));
        return dataA - dataB;
    });

    const hoje = new Date();
    const seteDiasDepois = new Date(hoje);
    seteDiasDepois.setDate(hoje.getDate() + 7);

    const container = document.createElement("div");
    container.className = "row";
    container.id = "proximas-atividades";
    container.innerHTML = `
        <div class="col s12">
            <div class="card">
                <div class="card-content">
                    <span class="card-title green-text">📅 Próximas Atividades (7 dias)</span>
                    <p class="grey-text">
                        <strong>Período:</strong> ${hoje.toLocaleDateString('pt-BR')} a ${seteDiasDepois.toLocaleDateString('pt-BR')}
                    </p>
                    <div style="max-height: 400px; overflow-y: auto;">
                        <table class="striped highlight responsive-table">
                            <thead>
                                <tr>
                                    <th>Data</th>
                                    <th>Processo ID</th>
                                    <th>Responsável</th>
                                    <th>Empresa</th>
                                    <th>Tipo</th>
                                    <th>Status</th>
                                    <th>Descrição</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${proximasAtividades.map(reg => `
                                    <tr>
                                        <td><strong class="green-text">${formatarDataExcel(reg["Data do agendamento"])}</strong></td>
                                        <td><strong>${reg["Processo - ID"] || "-"}</strong></td>
                                        <td>${reg["Responsável"] || "-"}</td>
                                        <td>${reg["Empresa"] || "-"}</td>
                                        <td><span class="chip blue white-text">${reg["Tipo"] || "-"}</span></td>
                                        <td><span class="chip ${obterCorStatus(reg["Status da tarefa"])}">${reg["Status da tarefa"] || "-"}</span></td>
                                        <td>${obterDescricao(reg)}</td>
                                    </tr>
                                `).join("")}
                            </tbody>
                        </table>
                    </div>
                    <div class="card-action">
                        <span><strong>Total de atividades nos próximos 7 dias:</strong> ${proximasAtividades.length}</span>
                    </div>
                </div>
            </div>
        </div>
    `;

    return container;
}

// 🆕 Função para gerar tabela de todas as atividades futuras
function gerarTabelaAtividadesFuturas(dados) {
    const atividadesFuturas = filtrarDadosFuturos(dados);

    // Ordena por data (mais próximas primeiro)
    atividadesFuturas.sort((a, b) => {
        const dataA = new Date(a["Data do agendamento"].split('/').reverse().join('-'));
        const dataB = new Date(b["Data do agendamento"].split('/').reverse().join('-'));
        return dataA - dataB;
    });

    const container = document.createElement("div");
    container.className = "row";
    container.id = "atividades-futuras";
    container.innerHTML = `
        <div class="col s12">
            <div class="card">
                <div class="card-content">
                    <span class="card-title blue-text">🔮 Todas as Atividades Futuras</span>
                    <p class="grey-text">
                        <strong>Critério:</strong> Todas as atividades com data posterior a hoje
                    </p>
                    <div style="max-height: 500px; overflow-y: auto;">
                        <table class="striped highlight responsive-table">
                            <thead>
                                <tr>
                                    <th>Data</th>
                                    <th>Processo ID</th>
                                    <th>Responsável</th>
                                    <th>Empresa</th>
                                    <th>Área do Direito</th>
                                    <th>Tipo</th>
                                    <th>Status</th>
                                    <th>Descrição</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${atividadesFuturas.map(reg => `
                                    <tr>
                                        <td><strong class="blue-text">${formatarDataExcel(reg["Data do agendamento"])}</strong></td>
                                        <td><strong>${reg["Processo - ID"] || "-"}</strong></td>
                                        <td>${reg["Responsável"] || "-"}</td>
                                        <td>${reg["Empresa"] || "-"}</td>
                                        <td>${reg["Área do Direito"] || "-"}</td>
                                        <td><span class="chip blue white-text">${reg["Tipo"] || "-"}</span></td>
                                        <td><span class="chip ${obterCorStatus(reg["Status da tarefa"])}">${reg["Status da tarefa"] || "-"}</span></td>
                                        <td>${obterDescricao(reg)}</td>
                                    </tr>
                                `).join("")}
                            </tbody>
                        </table>
                    </div>
                    <div class="card-action">
                        <span><strong>Total de atividades futuras:</strong> ${atividadesFuturas.length}</span>
                    </div>
                </div>
            </div>
        </div>
    `;

    return container;
}

// 🆕 Função atualizada para incluir as novas tabelas
async function adicionarTabelasEspeciaisCompletas() {
    const dados = await carregarExcel();

    const containerPrincipal = document.createElement("div");
    containerPrincipal.className = "container";
    containerPrincipal.style.marginTop = "20px";

    const titulo = document.createElement("h4");
    titulo.textContent = "📊 Relatórios Especiais";
    titulo.className = "center-align";
    containerPrincipal.appendChild(titulo);

    // Tabelas existentes
    const tabelaAudiencias = gerarTabelaAudiencias(dados);
    containerPrincipal.appendChild(tabelaAudiencias);

    const tabelaPrazos = gerarTabelaPrazosFatais(dados);
    containerPrincipal.appendChild(tabelaPrazos);

    // 🆕 Novas tabelas de atividades futuras
    const tabelaProximas = gerarTabelaProximasAtividades(dados);
    containerPrincipal.appendChild(tabelaProximas);

    const tabelaFuturas = gerarTabelaAtividadesFuturas(dados);
    containerPrincipal.appendChild(tabelaFuturas);

    document.body.appendChild(containerPrincipal);
}




// Para usar: debugPrazosFatais(dadosExcel) no console do navegador

async function adicionarGraficoEvolucaoMensal() {
    const dados = await carregarExcel();
    const id = "graficoEvolucaoMensal";

    if (document.getElementById(id)) return;

    const container = document.createElement("div");
    container.className = "grafico-container";
    container.innerHTML = `
        <div class="grafico-header">
            <strong>📅 Evolução Mensal de Processos</strong>
        </div>
        <canvas id="${id}" style="max-height: 400px;"></canvas>
    `;

    document.getElementById("graficos").appendChild(container);
    gerarGraficoEvolucaoMensal(dados, id);
}

function rolarPara(id) {
    const elemento = document.getElementById(id);
    if (elemento) {
        elemento.scrollIntoView({
            behavior: "smooth",
            block: "start"
        });
    } else {
        console.warn(`Elemento com id '${id}' não encontrado.`);
    }
}

// 🚀 Inicialização principal
window.onload = async () => {
    try {
        await carregarExcel();

        // Atualiza estatísticas
        await atualizarEstatisticas();

        exibirUltimaAtualizacao(); // ⬅️ Adicionada aqui


        // 📊 Gráficos principais com filtro "até hoje"
        const colunas = ["Responsável", "Área do Direito"];
        for (const coluna of colunas) {
            await adicionarGrafico(coluna, true); // true = filtrar até hoje
        }

        // 📌 Gráfico de pendências por área e responsável
        await adicionarGraficoPendenciasPorAreaEUsuario();

        // 🍕 Gráfico de pizza
        await adicionarGraficoPizza();

        // 📈 Gráfico de evolução mensal
        await adicionarGraficoEvolucaoMensal();

        // 🧑‍💼 Gráfico radar de responsáveis
        await adicionarGraficoRadarResponsaveis();


        // 🗓️ Gráfico de próximos 7 dias
        await adicionarGraficoProximos7Dias("Responsável");

        // 📅 Gráfico atividades futuras
        await adicionarGraficoAtividadesFuturas("Responsável");

        // 📋 Tabelas especiais

        await adicionarTabelasEspeciaisCompletas();

        // Adiciona animações aos elementos
        document.querySelectorAll('.grafico-container').forEach((el, index) => {
            el.classList.add('fade-in');
            el.style.animationDelay = `${index * 0.1}s`;
        });

        // Remove loading overlay
        const loadingOverlay = document.getElementById('loadingOverlay');
        if (loadingOverlay) {
            loadingOverlay.style.display = 'none';
        }

        console.log('✅ Dashboard carregado com sucesso!');

    } catch (error) {
        console.error('❌ Erro ao carregar dashboard:', error);
        const loadingOverlay = document.getElementById('loadingOverlay');
        if (loadingOverlay) {
            loadingOverlay.innerHTML = `
                <div class="loading-content">
                    <i class="material-icons" style="font-size: 4rem; color: #f44336;">error</i>
                    <h5 style="color: #f44336;">Erro ao carregar dados</h5>
                    <p>Verifique a conexão e tente novamente</p>
                </div>
            `;
        }
    }
};