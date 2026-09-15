const CORES = {
    roxo: { fundo: "rgba(153, 102, 255, 0.5)", borda: "rgba(153, 102, 255, 1)" },
    vermelho: { fundo: "rgba(255, 99, 132, 0.5)", borda: "rgba(255, 99, 132, 1)" },
    azul: { fundo: "rgba(54, 162, 235, 0.5)", borda: "rgba(54, 162, 235, 1)" },
    verde: { fundo: "rgba(75, 192, 192, 0.5)", borda: "rgba(75, 192, 192, 1)" },
    laranja: { fundo: "rgba(255, 159, 64, 0.5)", borda: "rgba(255, 159, 64, 1)" }
};

// 📦 Importa a URL do arquivo url.js
const excelUrl = "https://fundacaooswaldoaranha-my.sharepoint.com/personal/wesley_balbino_foa_org_br/_layouts/15/download.aspx?share=EdsT2JkTPstFhYTAoyB0kWwB0T83o-R9AR4Wu2Yex8hxBw";
// 🗂️ Variável global para armazenar os dados do Excel
let dadosExcel = [];

let dataB3 = null; // 🆕 Nova variável para armazenar a data da célula B3
let dataB3Formatada = null; // 🆕 Data B3 em formato pt-BR

// 📊 Armazena instâncias de gráficos
const charts = {}; 



function limparLocalStorage() {
    localStorage.removeItem("dadosExcel");
    localStorage.removeItem("ultimaAtualizacaoExcel");
    localStorage.removeItem("dataB3");           // 🆕
    localStorage.removeItem("dataB3Formatada");  // 🆕
}

limparLocalStorage(); // Chama a função para limpar o localStorage
document.querySelector("#logo").addEventListener("click", () => location.reload()); // Adiciona evento de clique no logo


// 🆕 Função para exibir a data B3 na interface
function exibirDataB3() {
    const { formatada } = obterDataB3();
    
    // Cria ou atualiza elemento para mostrar a data B3
    let elementoDataB3 = document.getElementById('dataB3Info');
    
    if (!elementoDataB3) {
        // Cria o elemento se não existir
        elementoDataB3 = document.createElement('div');
        elementoDataB3.id = 'dataB3Info';
        elementoDataB3.className = 'card-panel blue lighten-5';
        elementoDataB3.style.marginTop = '20px';
        elementoDataB3.style.textAlign = 'center';
        
        // Adiciona ao container principal (você pode mudar o local)
        const container = document.querySelector('.container') || document.body;
        container.insertBefore(elementoDataB3, container.firstChild);
    }
    
    // Atualiza o conteúdo
    if (formatada) {
        elementoDataB3.innerHTML = `
            <h6><i class="material-icons">schedule</i> Data da Célula B3</h6>
            <p><strong>${formatada}</strong></p>
        `;
    } else {
        elementoDataB3.innerHTML = `
            <h6><i class="material-icons">error</i> Data B3 não encontrada</h6>
            <p><em>Célula B3 vazia ou não localizada</em></p>
        `;
    }
}


// ⏱️ Verifica se já passou do horário limite (08:15)
function deveAtualizarDados() {
    const agora = new Date();
    const horaAtual = agora.getHours() + agora.getMinutes() / 60;
    const horarioManha = 8 + 15 / 60;  // 8:15
    return horaAtual >= horarioManha;
}


function exibirUltimaAtualizacao() {
    const el = document.getElementById("infoAtualizacao");
    if (!el) return;

    const raw = localStorage.getItem("ultimaAtualizacaoExcel");

    if (!raw) {
        el.innerHTML = `<span class="dot" style="background:var(--ink-faint);"></span><em>Sem dados</em>`;
        return;
    }

    try {
        el.innerHTML = `
        <span class="dot"></span>
        <span title="${raw}">Pauta atualizada ${dataB3Formatada}</span>
        `;
    } catch (e) {
        console.warn("Erro ao formatar data:", e);
        el.innerHTML = `<span class="dot" style="background:var(--crit);"></span><em>Erro na data</em>`;
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

// 🆕 Função para converter data do Excel para pt-BR com hora
function converterDataExcelParaPtBRComHora(valorData) {
    if (!valorData || valorData === "-" || valorData === "") return valorData;

    // Se já estiver em formato de string pt-BR, retorna como está
    if (typeof valorData === 'string' && valorData.includes('/')) {
        return valorData;
    }

    // Se for número serial do Excel, converte
    if (typeof valorData === 'number' && valorData > 0) {
        try {
            // Excel conta dias desde 01/01/1900, mas com bug do ano 1900
            // Ajuste para JavaScript (que conta desde 01/01/1970)
            const diasDesde1900 = valorData - 25569;
            
            // 🔧 CORREÇÃO: Cria a data em UTC primeiro para evitar problemas de fuso horário
            const dataUTC = new Date(diasDesde1900 * 86400 * 1000);
            
            // 🔧 Ajusta o fuso horário para o Brasil (UTC-3)
            // Como o Excel não considera fuso horário, precisamos ajustar manualmente
            const offsetBrasil = +6 * 60; // -3 horas em minutos
            const offsetLocal = dataUTC.getTimezoneOffset(); // Offset local em minutos
            const diferencaOffset = offsetBrasil - offsetLocal;
            
            // Aplica o ajuste
            const dataAjustada = new Date(dataUTC.getTime() + (diferencaOffset * 60 * 1000));

            // Verifica se a data é válida
            if (isNaN(dataAjustada.getTime())) {
                console.warn(`Data inválida do Excel: ${valorData}`);
                return valorData;
            }

            // Formatar para DD/MM/AAAA HH:MM:SS
            const dia = String(dataAjustada.getDate()).padStart(2, '0');
            const mes = String(dataAjustada.getMonth() + 1).padStart(2, '0');
            const ano = dataAjustada.getFullYear();
            const horas = String(dataAjustada.getHours()).padStart(2, '0');
            const minutos = String(dataAjustada.getMinutes()).padStart(2, '0');
            const segundos = String(dataAjustada.getSeconds()).padStart(2, '0');

            // Sempre inclui hora, minuto e segundo para B3
            return `${dia}/${mes}/${ano} ${horas}:${minutos}:${segundos}`;

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
        const segundos = String(valorData.getSeconds()).padStart(2, '0');

        return `${dia}/${mes}/${ano} ${horas}:${minutos}:${segundos}`;
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
                const segundos = String(data.getSeconds()).padStart(2, '0');

                return `${dia}/${mes}/${ano} ${horas}:${minutos}:${segundos}`;
            }
        } catch (e) {
            // Se não conseguir converter, retorna o valor original
        }
    }

    return valorData;
}


// 🆕 Função para capturar dados específicos de células
function capturarDadosCelulas(worksheet) {
    try {
        // Captura a célula B3
        const celulaB3 = worksheet['B3'];
        
        if (celulaB3 && celulaB3.v !== undefined) {
            dataB3 = celulaB3.v; // Valor bruto da célula
            dataB3Formatada = converterDataExcelParaPtBRComHora(celulaB3.v);
            
            console.log(`📅 Célula B3 capturada:`);
            console.log(`   Valor bruto: ${dataB3}`);
            console.log(`   Formatado: ${dataB3Formatada}`);
        } else {
            console.warn('⚠️ Célula B3 não encontrada ou vazia');
            dataB3 = null;
            dataB3Formatada = null;
        }

        // 🆕 Você pode capturar outras células aqui se necessário
        // Exemplo: const celulaC3 = worksheet['C3'];
        
    } catch (error) {
        console.error('❌ Erro ao capturar dados das células:', error);
        dataB3 = null;
        dataB3Formatada = null;
    }
}


// 🔄 Função carregarExcel MODIFICADA para incluir captura da célula B3
async function carregarExcel() {
    if (dadosExcel.length) {
        return dadosExcel;
    }

    const dadosSalvos = localStorage.getItem("dadosExcel");
    const dataB3Salva = localStorage.getItem("dataB3");
    const dataB3FormatadaSalva = localStorage.getItem("dataB3Formatada");
    
    if (dadosSalvos) {
        dadosExcel = JSON.parse(dadosSalvos);
        dataB3 = dataB3Salva ? JSON.parse(dataB3Salva) : null;
        dataB3Formatada = dataB3FormatadaSalva || null;
        
        console.log(`✅ Dados carregados do localStorage: ${dadosExcel.length} registros`);
        console.log(`📅 Data B3 recuperada: ${dataB3Formatada}`);
        return dadosExcel;
    }

    console.log("🔄 Carregando dados do Excel...");

    try {
        const response = await fetch(excelUrl);
        const blob = await response.blob();
        const buffer = await blob.arrayBuffer();

        const workbook = XLSX.read(buffer, { type: "array" });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        
        // 🆕 NOVA LINHA: Captura a célula B3 ANTES de processar os dados da tabela
        capturarDadosCelulas(worksheet);
        
        let dadosBrutos = XLSX.utils.sheet_to_json(worksheet, {
            range: 5,
            defval: ""
        });

        console.log(`📊 Dados brutos carregados: ${dadosBrutos.length} registros`);

        // Converte todas as datas para formato pt-BR antes de salvar
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
                        // console.log(`🔄 [${index + 1}] ${campo}: "${valorOriginal}" -> "${valorConvertido}"`);
                    }
                }
            });

            return registroConvertido;
        });

        // Salva os dados já convertidos e marca o timestamp da atualização
        localStorage.setItem("dadosExcel", JSON.stringify(dadosExcel));
        localStorage.setItem("ultimaAtualizacaoExcel", new Date().toISOString());
        
        // 🆕 Salva também os dados da célula B3
        localStorage.setItem("dataB3", JSON.stringify(dataB3));
        localStorage.setItem("dataB3Formatada", dataB3Formatada);

        console.log("✅ Dados atualizados e convertidos com sucesso!");
        console.log(`📊 Total de registros processados: ${dadosExcel.length}`);
        console.log(`📅 Data da célula B3: ${dataB3Formatada}`);

    } catch (error) {
        console.error("❌ Erro ao carregar dados do Excel:", error);
        throw error;
    }

    return dadosExcel;
}


// 🆕 Função para obter a data B3 formatada (para usar em outros lugares)
function obterDataB3() {
    return {
        valorBruto: dataB3,
        formatada: dataB3Formatada
    };
}



// 🆕 Função simplificada para filtrar dados até hoje (datas já estão em formato pt-BR)
function filtrarDadosAteHoje(dados) {
    const hoje = new Date();
    hoje.setHours(23, 59, 59, 999); // Final do dia

    
    // console.log(`🔍 Filtrando dados até: ${hoje.toLocaleDateString('pt-BR')}`);

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
        // console.log('❌ Primeiras 5 datas futuras rejeitadas:');
        rejeitados.slice(0, 5).forEach((item, i) => {
            // console.log(`  ${i + 1}. "${item["Data do agendamento"]}"`);
        });
    }
    
    return dadosFiltrados;
}




function gerarGraficoPorColuna(coluna, dados, canvasId, cor = CORES.roxo, filtrarAteHoje = false) {
    // 🆕 Aplica filtro de data se solicitado
    let dadosFiltrados = dados;
    if (filtrarAteHoje) {
        dadosFiltrados = filtrarDadosAteHoje(dados);
        // console.log(`🎯 Gráfico "${coluna}" - Total original: ${dados.length}`);
        // console.log(`🎯 Gráfico "${coluna}" - Filtrado até hoje: ${dadosFiltrados.length} registros`);
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
                label: `Quantidade por ${coluna}${filtrarAteHoje ? ' (até hoje)' : ''}`,
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

                // console.log(`🔍 Clique no gráfico "${coluna}": ${valorClicado} - ${resultados.length} resultados`);

                // Debug: mostra algumas datas dos resultados
                // console.log('📅 Primeiras 5 datas dos resultados:');
                resultados.slice(0, 5).forEach((item, i) => {
                    // console.log(`  ${i + 1}. "${item["Data do agendamento"]}"`);
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
    // console.log("🧹 Limpando cache e recarregando dados...");

    // Remove dados do localStorage
    localStorage.removeItem("dadosExcel");
    localStorage.removeItem("ultimaAtualizacaoExcel");

    // Limpa variável global
    dadosExcel = [];

    // console.log("✅ Cache limpo! Recarregue a página para baixar dados atualizados.");

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
        '#63c6ffff', '#eb367bff', '#FFCE56', '#4BC0C0', '#9966FF',
        '#FF9F40', '#63c1ffff', '#C9CBCF', '#4BC0C0', '#FF6384'
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

    // console.log(`🚨 Analisando prazos fatais:`);
    // console.log(`   📅 Hoje: ${hoje.toLocaleDateString('pt-BR')}`);
    // console.log(`   📅 Limite (7 dias): ${seteDiasDepois.toLocaleDateString('pt-BR')}`);

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

    // console.log(`🚨 Prazos fatais encontrados: ${prazosFatais.length}`);
    // console.log(`   ❌ Atrasados: ${atrasados.length}`);
    // console.log(`   ⚠️ Vencem em 7 dias: ${proximosVencimento.length}`);

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
    // console.log(`🎯 Gráfico Pendências - Iniciando filtro`);

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

    // console.log(`📊 Pendências - Total original: ${dados.length}`);
    // console.log(`📊 Pendências - Filtrado por data: ${dadosFiltradosPorData.length}`);
    // console.log(`📊 Pendências - Final (data + status): ${dadosFiltrados.length}`);

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
                    title: { display: true, text: 'Quantidade de Pendências (até hoje)' }
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
            <strong>📌 Pendências por Responsável e Área do Direito (até hoje)</strong>
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
    // console.log("🔍 === DEBUG DE DATAS ===");

    const hoje = new Date();
    hoje.setHours(23, 59, 59, 999);
    // console.log(`📅 Data de referência (hoje): ${hoje.toLocaleDateString('pt-BR')}`);

    // Analisa os primeiros 20 registros
    const amostra = dados.slice(0, 20);

    // console.log("\n📊 Análise das primeiras 20 datas:");
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

        // console.log(`${i + 1}. "${dataStr}" -> ${status} ${dataParsed ? `(${dataParsed.toLocaleDateString('pt-BR')})` : ''}`);
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

    // console.log(`\n📈 Resumo:`);
    // console.log(`   Total de registros: ${dados.length}`);
    // console.log(`   Registros com datas futuras: ${futuras.length}`);
    // console.log(`   Registros que devem passar no filtro: ${dados.length - futuras.length}`);

    // console.log("🔍 === FIM DEBUG ===\n");
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

    // console.log(`📊 Estatísticas - Total de dados: ${dados.length}`);
    // console.log(`📊 Estatísticas - Filtrados por data (até hoje): ${dadosFiltradosPorData.length}`);
    // console.log(`📊 Estatísticas - Pendentes até hoje: ${pendentes.length}`);

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
    // console.log("🚨 === DEBUG PRAZOS FATAIS ===");

    const hoje = new Date();
    hoje.setHours(23, 59, 59, 999);
    const seteDiasDepois = new Date(hoje);
    seteDiasDepois.setDate(hoje.getDate() + 7);

    // console.log(`📅 Período de análise: ${hoje.toLocaleDateString('pt-BR')} até ${seteDiasDepois.toLocaleDateString('pt-BR')}`);

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

    // console.log("\n🔍 Analisando campos de prazo fatal disponíveis:");
    camposPrazoFatal.forEach(campo => {
        const valores = dados
            .map(item => item[campo])
            .filter(valor => valor && valor !== "" && valor !== "-")
            .slice(0, 10); // Primeiros 10 valores únicos

        if (valores.length > 0) {
            // console.log(`   ${campo}: ${[...new Set(valores)].join(', ')}`);
        }
    });

    // console.log("\n🔍 Analisando campos de data disponíveis:");
    camposData.forEach(campo => {
        const count = dados.filter(item => item[campo] && item[campo] !== "" && item[campo] !== "-").length;
        if (count > 0) {
            // console.log(`   ${campo}: ${count} registros com data`);
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

    // console.log(`\n📊 Registros com prazo fatal = "Sim": ${comPrazoFatal.length}`);

    // Mostra alguns exemplos
    // console.log("\n🔍 Primeiros 5 exemplos de prazos fatais:");
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

        // console.log(`   ${i + 1}. Processo ${processoId}: ${campoComPrazo} = "${item[campoComPrazo]}" | Data: ${data}`);
    });

    // console.log("🚨 === FIM DEBUG ===\n");

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

    // console.log(`🔍 Filtrando próximos 7 dias:`);
    // console.log(`   📅 A partir de: ${hoje.toLocaleDateString('pt-BR')}`);
    // console.log(`   📅 Até: ${seteDiasDepois.toLocaleDateString('pt-BR')}`);

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

    // console.log(`📊 Total de registros originais: ${dados.length}`);
    // console.log(`📊 Registros dos próximos 7 dias: ${dadosFiltrados.length}`);

    // Debug: mostra as 5 primeiras datas filtradas
    // console.log('✅ Primeiras 5 datas dos próximos 7 dias:');
    dadosFiltrados.slice(0, 5).forEach((item, i) => {
        // console.log(`  ${i + 1}. "${item["Data do agendamento"]}"`);
    });

    return dadosFiltrados;
}

// 🆕 Função para filtrar todas as atividades futuras
function filtrarDadosFuturos(dados) {
    const hoje = new Date();
    hoje.setHours(23, 59, 59, 999); // Final do dia de hoje

    // console.log(`🔍 Filtrando todas as atividades futuras a partir de: ${hoje.toLocaleDateString('pt-BR')}`);

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

    // console.log(`📊 Total de registros originais: ${dados.length}`);
    // console.log(`📊 Registros futuros: ${dadosFiltrados.length}`);

    // Debug: mostra as 5 primeiras datas filtradas
    // console.log('✅ Primeiras 5 datas futuras:');
    dadosFiltrados.slice(0, 5).forEach((item, i) => {
        // console.log(`  ${i + 1}. "${item["Data do agendamento"]}"`);
    });

    return dadosFiltrados;
}

// 🆕 Função modificada para aceitar filtros personalizados
function gerarGraficoPorColunaComFiltro(coluna, dados, canvasId, cor = CORES.roxo, filtroCallback = null, tituloExtra = "") {
    // Aplica filtro personalizado se fornecido
    let dadosFiltrados = dados;
    if (filtroCallback && typeof filtroCallback === 'function') {
        dadosFiltrados = filtroCallback(dados);
        // console.log(`🎯 Gráfico "${coluna}" - Total original: ${dados.length}`);
        // console.log(`🎯 Gráfico "${coluna}" - Filtrado: ${dadosFiltrados.length} registros`);
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

                // console.log(`🔍 Clique no gráfico "${coluna}": ${valorClicado} - ${resultados.length} resultados`);

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

// 📅 Função para gerar gráfico de linha do volume de prazos por dia
function gerarGraficoVolumePrazosDiario(dados, canvasId) {
    const contagemPorDia = {};
    const registrosPorDia = {}; // Para armazenar registros para o modal

    // Processa todos os dados com datas válidas
    dados.forEach(item => {
        const dataStr = item["Data do agendamento"];

        // Ignora registros sem data ou com data vazia
        if (!dataStr || dataStr === "" || dataStr === "-") return;

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
                return;
            }
        }

        // Verifica se a data é válida
        if (!dataAgendamento || isNaN(dataAgendamento.getTime())) {
            console.warn(`Data inválida encontrada: "${dataStr}"`);
            return;
        }

        // Cria chave no formato DD/MM/YYYY para agrupamento
        const chaveData = `${String(dataAgendamento.getDate()).padStart(2, '0')}/${String(dataAgendamento.getMonth() + 1).padStart(2, '0')}/${dataAgendamento.getFullYear()}`;
        
        // Conta ocorrências por dia
        contagemPorDia[chaveData] = (contagemPorDia[chaveData] || 0) + 1;

        // Armazena registros para o modal
        if (!registrosPorDia[chaveData]) {
            registrosPorDia[chaveData] = [];
        }
        registrosPorDia[chaveData].push(item);
    });

    // Ordena as datas cronologicamente
    const datasOrdenadas = Object.keys(contagemPorDia).sort((a, b) => {
        const [diaA, mesA, anoA] = a.split('/').map(Number);
        const [diaB, mesB, anoB] = b.split('/').map(Number);
        const dataA = new Date(anoA, mesA - 1, diaA);
        const dataB = new Date(anoB, mesB - 1, diaB);
        return dataA - dataB;
    });

    const valores = datasOrdenadas.map(data => contagemPorDia[data]);

    // Identifica o dia de hoje para destacar no gráfico
    const hoje = new Date();
    const hojeStr = `${String(hoje.getDate()).padStart(2, '0')}/${String(hoje.getMonth() + 1).padStart(2, '0')}/${hoje.getFullYear()}`;

    // Cria cores diferenciadas: passado (azul), hoje (laranja), futuro (verde)
    const coresPontos = datasOrdenadas.map(data => {
        const [dia, mes, ano] = data.split('/').map(Number);
        const dataAtual = new Date(ano, mes - 1, dia);
        const hoje = new Date();
        hoje.setHours(0, 0, 0, 0);

        if (data === hojeStr) {
            return '#FF9F40'; // Laranja para hoje
        } else if (dataAtual < hoje) {
            return '#36A2EB'; // Azul para passado
        } else {
            return '#4BC0C0'; // Verde para futuro
        }
    });

    const ctx = document.getElementById(canvasId).getContext("2d");

    const chart = new Chart(ctx, {
        type: "line",
        data: {
            labels: datasOrdenadas,
            datasets: [{
                label: "Volume de Prazos por Dia",
                data: valores,
                fill: true,
                borderColor: CORES.azul.borda,
                backgroundColor: CORES.azul.fundo,
                tension: 0.3,
                pointRadius: 4,
                pointHoverRadius: 8,
                pointBackgroundColor: coresPontos,
                pointBorderColor: '#fff',
                pointBorderWidth: 2
            }]
        },
        options: {
            responsive: true,
            onClick: (e, elements) => {
                if (elements.length === 0) return;

                const index = elements[0].index;
                const dataClicada = chart.data.labels[index];
                const registros = registrosPorDia[dataClicada] || [];

                if (registros.length > 0) {
                    console.log(`🔍 Clique no gráfico de volume diário: ${dataClicada} - ${registros.length} atividades`);
                    exibirTabela(canvasId, registros);
                }
            },
            plugins: {
                datalabels: {
                    display: false // Desabilita labels nos pontos para não poluir
                },
                tooltip: {
                    callbacks: {
                        title: function(context) {
                            const data = context[0].label;
                            const [dia, mes, ano] = data.split('/').map(Number);
                            const dataObj = new Date(ano, mes - 1, dia);
                            const diasSemana = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];
                            const diaSemana = diasSemana[dataObj.getDay()];
                            return `${diaSemana}, ${data}`;
                        },
                        label: function(context) {
                            const quantidade = context.parsed.y;
                            return `${quantidade} ${quantidade === 1 ? 'atividade' : 'atividades'}`;
                        },
                        afterLabel: function(context) {
                            const data = context.label;
                            const hoje = new Date();
                            const hojeStr = `${String(hoje.getDate()).padStart(2, '0')}/${String(hoje.getMonth() + 1).padStart(2, '0')}/${hoje.getFullYear()}`;
                            
                            if (data === hojeStr) {
                                return '📅 Hoje';
                            }
                            
                            const [dia, mes, ano] = data.split('/').map(Number);
                            const dataAtual = new Date(ano, mes - 1, dia);
                            hoje.setHours(0, 0, 0, 0);
                            
                            if (dataAtual < hoje) {
                                return '⏪ Passado';
                            } else {
                                return '⏩ Futuro';
                            }
                        }
                    }
                },
                legend: {
                    display: true,
                    position: 'top'
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Número de Atividades'
                    },
                    ticks: {
                        stepSize: 1 // Força números inteiros no eixo Y
                    }
                },
                x: {
                    title: {
                        display: true,
                        text: 'Data (DD/MM/YYYY)'
                    },
                    ticks: {
                        maxTicksLimit: 15, // Limita o número de labels no eixo X
                        callback: function(value, index) {
                            // Mostra apenas algumas datas para não poluir o eixo
                            const totalDatas = this.chart.data.labels.length;
                            const intervalo = Math.ceil(totalDatas / 10);
                            return index % intervalo === 0 ? this.chart.data.labels[index] : '';
                        }
                    }
                }
            },
            elements: {
                point: {
                    hoverRadius: 8
                }
            }
        },
        plugins: [ChartDataLabels]
    });

    return chart;
}

// 📊 Função para adicionar o gráfico de volume diário
async function adicionarGraficoVolumePrazosDiario() {
    const dados = await carregarExcel();
    const id = "graficoVolumePrazosDiario";

    // Remove gráfico anterior se existir
    const graficoExistente = document.getElementById(id);
    if (graficoExistente) {
        graficoExistente.closest('.grafico-container').remove();
    }

    const container = document.createElement("div");
    container.className = "grafico-container";
    container.innerHTML = `
        <div class="grafico-header">
            <strong>📈 Volume de Prazos por Dia</strong>
            <div style="font-size: 0.9em; color: #666; margin-top: 5px;">
                🔵 Passado | 🟠 Hoje | 🟢 Futuro | Clique nos pontos para ver detalhes
            </div>
        </div>
        <canvas id="${id}" style="max-height: 450px;"></canvas>
    `;

    document.getElementById("graficos").appendChild(container);
    
    const chart = gerarGraficoVolumePrazosDiario(dados, id);
    charts[id] = { chart, coluna: "Volume Diário" };

    // Estatísticas do gráfico
    const totalDias = Object.keys(chart.data.labels).length;
    const totalAtividades = chart.data.datasets[0].data.reduce((a, b) => a + b, 0);
    const mediaDiaria = (totalAtividades / totalDias).toFixed(1);
    const maiorVolume = Math.max(...chart.data.datasets[0].data);

    console.log(`📊 Gráfico de Volume Diário criado:`);
    console.log(`   📅 Total de dias: ${totalDias}`);
    console.log(`   📋 Total de atividades: ${totalAtividades}`);
    console.log(`   📊 Média diária: ${mediaDiaria} atividades`);
    console.log(`   🔥 Maior volume em um dia: ${maiorVolume} atividades`);

    return chart;
}

// 🚀 Para adicionar o gráfico, chame:
// adicionarGraficoVolumePrazosDiario();

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

    // 🆕 Audiências Agendadas agora vive na própria aba do header
    const wrapAudiencias = document.getElementById("audienciasTabWrap");
    if (wrapAudiencias) {
        wrapAudiencias.appendChild(gerarTabelaAudiencias(dados));
    }

    const containerPrincipal = document.createElement("div");
    containerPrincipal.className = "container";
    containerPrincipal.style.marginTop = "20px";

    const titulo = document.createElement("h4");
    titulo.textContent = "📊 Relatórios Especiais";
    titulo.className = "center-align";
    containerPrincipal.appendChild(titulo);

    const tabelaPrazos = gerarTabelaPrazosFatais(dados);
    containerPrincipal.appendChild(tabelaPrazos);

    // 🆕 Novas tabelas de atividades futuras
    const tabelaProximas = gerarTabelaProximasAtividades(dados);
    containerPrincipal.appendChild(tabelaProximas);

    const tabelaFuturas = gerarTabelaAtividadesFuturas(dados);
    containerPrincipal.appendChild(tabelaFuturas);

    const wrapRelatorios = document.getElementById("relatoriosEspeciaisWrap");
    (wrapRelatorios || document.body).appendChild(containerPrincipal);
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
        const details = elemento.closest('details');
        if (details && !details.open) details.open = true;

        elemento.scrollIntoView({
            behavior: "smooth",
            block: "start"
        });
    } else {
        console.warn(`Elemento com id '${id}' não encontrado.`);
    }
}

// 🆕 Alterna as abas do header (Produtividade / Atividades / Prazos / Audiências)
function ativarAba(nome) {
    document.querySelectorAll('.topbar-tab').forEach(botao => {
        botao.classList.toggle('active', botao.dataset.aba === nome);
    });

    document.querySelectorAll('.tab-panel').forEach(secao => {
        secao.hidden = secao.dataset.aba !== nome;
    });

    window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ============================================================
// 🆕 Painel de Produtividade da Equipe (redesign)
// ============================================================

// Considera "ativa" toda tarefa cujo status não indique conclusão/cancelamento
function statusIndicaConcluida(status) {
    if (!status) return false;
    const s = status.toString().toLowerCase();
    return s.includes("concluíd") || s.includes("concluid") ||
        s.includes("finalizad") || s.includes("cancelad");
}

// Converte "Data do agendamento" (pt-BR, DD/MM/YYYY) em Date, sem hora
function parseDataAgendamento(dataStr) {
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
}

function mesmoDia(a, b) {
    return a.getFullYear() === b.getFullYear() &&
        a.getMonth() === b.getMonth() &&
        a.getDate() === b.getDate();
}

function temPrazoFatalSim(valor) {
    if (!valor) return false;
    return valor.toString().trim().toLowerCase() === "sim";
}

// Lê um token de cor do :root (resolve o tema light/dark atual) para uso no Chart.js,
// que não entende `var(--x)` diretamente no canvas.
function cssVar(nome) {
    return getComputedStyle(document.documentElement).getPropertyValue(nome).trim();
}

// 📊 Resumo da fila para os tiles do topo
function calcularResumoFila(dados) {
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    const amanha = new Date(hoje);
    amanha.setDate(hoje.getDate() + 1);

    const ativas = dados.filter(item => !statusIndicaConcluida(item["Status da tarefa"]));

    let contHoje = 0, contAmanha = 0, contAPrazo = 0;
    ativas.forEach(item => {
        const data = parseDataAgendamento(item["Data do agendamento"]);
        if (data && mesmoDia(data, hoje)) {
            contHoje++;
        } else if (data && mesmoDia(data, amanha)) {
            contAmanha++;
        } else {
            contAPrazo++;
        }
    });

    const totalAtivas = ativas.length;
    const pctHoje = totalAtivas ? Math.round((contHoje / totalAtivas) * 100) : 0;
    const pctAmanha = totalAtivas ? Math.round((contAmanha / totalAtivas) * 100) : 0;

    const prazoFatalAberto = ativas.filter(item => temPrazoFatalSim(item["Solicitação - Há Prazo Fatal"]));

    const porResponsavelPF = {};
    prazoFatalAberto.forEach(item => {
        const resp = (item["Responsável"] || "").trim() || "Sem responsável";
        porResponsavelPF[resp] = (porResponsavelPF[resp] || 0) + 1;
    });
    const topResponsaveisPF = Object.entries(porResponsavelPF)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 2);

    // Reaproveita o ranking por responsável para os insights de "hoje"/"amanhã"
    const ranking = calcularRankingResponsaveis(dados);
    const respComHoje = ranking.filter(r => r.hoje > 0).length;
    const topAmanha = ranking
        .filter(r => r.amanha > 0)
        .sort((a, b) => b.amanha - a.amanha)
        .slice(0, 2);

    return {
        totalAtivas,
        contHoje,
        contAmanha,
        contAPrazo,
        pctHoje,
        pctAmanha,
        prazoFatalAberto: prazoFatalAberto.length,
        topResponsaveisPF,
        respComHoje,
        topAmanha
    };
}

// 🧑‍💼 Ranking de tarefas ativas por responsável
function calcularRankingResponsaveis(dados) {
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    const amanha = new Date(hoje);
    amanha.setDate(hoje.getDate() + 1);

    const ativas = dados.filter(item => !statusIndicaConcluida(item["Status da tarefa"]));

    const porResponsavel = {};
    ativas.forEach(item => {
        const resp = (item["Responsável"] || "").trim();
        if (!resp) return;

        if (!porResponsavel[resp]) {
            porResponsavel[resp] = { nome: resp, hoje: 0, amanha: 0, total: 0 };
        }
        porResponsavel[resp].total++;

        const data = parseDataAgendamento(item["Data do agendamento"]);
        if (data && mesmoDia(data, hoje)) {
            porResponsavel[resp].hoje++;
        } else if (data && mesmoDia(data, amanha)) {
            porResponsavel[resp].amanha++;
        }
    });

    return Object.values(porResponsavel).sort((a, b) => b.total - a.total);
}

// 🚨 Prazos fatais ainda pendentes, agrupados por responsável + área
function calcularPrazosFatais(dados) {
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    const amanha = new Date(hoje);
    amanha.setDate(hoje.getDate() + 1);

    const pendentesComPrazo = dados.filter(item =>
        temPrazoFatalSim(item["Solicitação - Há Prazo Fatal"]) &&
        !statusIndicaConcluida(item["Status da tarefa"])
    );

    const grupos = {};
    pendentesComPrazo.forEach(item => {
        const resp = (item["Responsável"] || "").trim() || "Sem responsável";
        const area = (item["Área do Direito"] || "").trim() || "Não informada";
        const chave = `${resp}||${area}`;

        if (!grupos[chave]) {
            grupos[chave] = { responsavel: resp, area, quantidade: 0, hoje: 0, amanha: 0 };
        }
        grupos[chave].quantidade++;

        const data = parseDataAgendamento(item["Data do agendamento"]);
        if (data && mesmoDia(data, hoje)) {
            grupos[chave].hoje++;
        } else if (data && mesmoDia(data, amanha)) {
            grupos[chave].amanha++;
        }
    });

    return Object.values(grupos).sort((a, b) => b.quantidade - a.quantidade);
}

// 🎨 Renderiza os tiles de resumo do topo
function renderizarTilesResumo(dados) {
    const resumo = calcularResumoFila(dados);

    const elTotal = document.getElementById("tileAtivasTotal");
    if (elTotal) elTotal.textContent = resumo.totalAtivas;

    const elSub = document.getElementById("tileAtivasSub");
    if (elSub) {
        elSub.innerHTML = `<span><b>${resumo.contAPrazo}</b> a prazo</span><span><b>${resumo.contHoje}</b> hoje</span><span><b>${resumo.contAmanha}</b> amanhã</span>`;
    }

    const total = resumo.totalAtivas || 1;
    const pctPrazo = Math.round((resumo.contAPrazo / total) * 100);
    const pctHojeBarra = Math.round((resumo.contHoje / total) * 100);
    const pctAmanhaBarra = Math.max(0, 100 - pctPrazo - pctHojeBarra);

    const segPrazo = document.getElementById("segPrazo");
    const segHoje = document.getElementById("segHoje");
    const segAmanha = document.getElementById("segAmanha");
    if (segPrazo) segPrazo.style.width = `${pctPrazo}%`;
    if (segHoje) segHoje.style.width = `${pctHojeBarra}%`;
    if (segAmanha) segAmanha.style.width = `${pctAmanhaBarra}%`;

    const elHojeCount = document.getElementById("tileHojeCount");
    if (elHojeCount) elHojeCount.textContent = resumo.contHoje;
    const elHojeBadge = document.getElementById("tileHojeBadge");
    if (elHojeBadge) elHojeBadge.textContent = `${resumo.pctHoje}% da fila`;
    const elHojeResp = document.getElementById("tileHojeResp");
    if (elHojeResp) {
        elHojeResp.innerHTML = resumo.respComHoje
            ? `Concentradas em <b>${resumo.respComHoje}</b> responsável${resumo.respComHoje > 1 ? "eis" : ""}`
            : "Nenhuma tarefa vence hoje";
    }

    const elAmanhaCount = document.getElementById("tileAmanhaCount");
    if (elAmanhaCount) elAmanhaCount.textContent = resumo.contAmanha;
    const elAmanhaBadge = document.getElementById("tileAmanhaBadge");
    if (elAmanhaBadge) elAmanhaBadge.textContent = `${resumo.pctAmanha}% da fila`;
    const elAmanhaResp = document.getElementById("tileAmanhaResp");
    if (elAmanhaResp) {
        if (resumo.topAmanha.length) {
            const nomes = resumo.topAmanha.map(r => `<b>${r.nome.split(" ")[0]}</b>`).join(" e ");
            elAmanhaResp.innerHTML = `Pico em ${nomes}`;
        } else {
            elAmanhaResp.textContent = "Nenhuma tarefa amanhã";
        }
    }

    const elPrazoFatalCount = document.getElementById("tilePrazoFatalCount");
    if (elPrazoFatalCount) elPrazoFatalCount.textContent = resumo.prazoFatalAberto;

    const elPrazoFatalResp = document.getElementById("tilePrazoFatalResp");
    if (elPrazoFatalResp) {
        if (resumo.topResponsaveisPF.length) {
            elPrazoFatalResp.innerHTML = resumo.topResponsaveisPF
                .map(([nome, qtd]) => `${qtd} com <b>${nome.split(" ")[0]}</b>`)
                .join(" · ");
        } else {
            elPrazoFatalResp.textContent = "Nenhum prazo fatal em aberto";
        }
    }
}

// 🎨 Renderiza a tabela de ranking de tarefas por responsável
function renderizarRankingResponsaveis(dados) {
    const ranking = calcularRankingResponsaveis(dados);
    const corpo = document.getElementById("tabelaRankingBody");
    if (!corpo) return;

    const maiorTotal = ranking.reduce((max, r) => Math.max(max, r.total), 0) || 1;

    corpo.innerHTML = ranking.map(r => {
        const iniciais = r.nome.split(" ")
            .filter(Boolean)
            .slice(0, 2)
            .map(p => p[0].toUpperCase())
            .join("") || "?";
        const largura = Math.round((r.total / maiorTotal) * 100);

        return `
            <tr>
                <td>
                    <div class="ranking-row-name">
                        <span class="avatar">${iniciais}</span>
                        <span>${r.nome}</span>
                    </div>
                </td>
                <td class="num">${r.hoje || "—"}</td>
                <td class="num">${r.amanha || "—"}</td>
                <td class="num">${r.total}</td>
                <td>
                    <div class="load-bar-track">
                        <div class="load-bar-fill" style="width:${largura}%"></div>
                    </div>
                </td>
            </tr>
        `;
    }).join("");
}

// 🎨 Renderiza a tabela de prazos com risco fatal
function renderizarPainelPrazosFataisNovo(dados) {
    const lista = calcularPrazosFatais(dados);
    const corpo = document.getElementById("tabelaPrazosFataisNovaBody");
    if (!corpo) return;

    const meta = document.getElementById("prazosFataisMeta");
    if (meta) {
        const totalItens = lista.reduce((acc, item) => acc + item.quantidade, 0);
        meta.textContent = `${totalItens} prazo(s) fatal(is) em aberto no momento do carregamento.`;
    }

    if (!lista.length) {
        corpo.innerHTML = `<tr><td colspan="4" style="text-align:center; color:var(--ink-faint);">Nenhum prazo fatal em aberto</td></tr>`;
        return;
    }

    corpo.innerHTML = lista.map(item => {
        let pill = `<span class="pill pill-neutral">sem vencimento próximo</span>`;
        if (item.hoje > 0) {
            pill = `<span class="pill pill-crit">${item.hoje} vence${item.hoje > 1 ? "m" : ""} hoje</span>`;
        } else if (item.amanha > 0) {
            pill = `<span class="pill pill-warn">${item.amanha} vence${item.amanha > 1 ? "m" : ""} amanhã</span>`;
        }

        const iniciais = item.responsavel.split(" ")
            .filter(Boolean)
            .slice(0, 2)
            .map(p => p[0].toUpperCase())
            .join("") || "?";

        return `
            <tr>
                <td>
                    <div class="ranking-row-name">
                        <span class="avatar">${iniciais}</span>
                        <span>${item.responsavel}</span>
                    </div>
                </td>
                <td>${item.area}</td>
                <td class="num">${item.quantidade}</td>
                <td>${pill}</td>
            </tr>
        `;
    }).join("");
}

// 🧑‍💼 Lista de responsáveis distintos presentes na planilha (ordenada)
function obterListaResponsaveis(dados) {
    const nomes = new Set();
    dados.forEach(item => {
        const nome = (item["Responsável"] || "").trim();
        if (nome) nomes.add(nome);
    });
    return Array.from(nomes).sort((a, b) => a.localeCompare(b, 'pt-BR'));
}

// 🎨 Preenche o <select> da aba "Atividades por Responsável" e renderiza a primeira opção
function popularSelectResponsaveis(dados) {
    const select = document.getElementById("selectResponsavelAtividades");
    if (!select) return;

    const nomes = obterListaResponsaveis(dados);
    select.innerHTML = nomes.map(nome => `<option value="${nome}">${nome}</option>`).join("");

    select.onchange = () => renderizarAtividadesPorResponsavel(dados, select.value);

    if (nomes.length) {
        select.value = nomes[0];
        renderizarAtividadesPorResponsavel(dados, nomes[0]);
    }
}

// 🎨 Renderiza a tabela de prazos/tarefas do responsável selecionado
function renderizarAtividadesPorResponsavel(dados, nome) {
    const corpo = document.getElementById("tabelaAtividadesResponsavelBody");
    if (!corpo) return;

    const atividades = dados
        .filter(item => (item["Responsável"] || "").trim() === nome)
        .slice()
        .sort((a, b) => {
            const dataA = parseDataAgendamento(a["Data do agendamento"]);
            const dataB = parseDataAgendamento(b["Data do agendamento"]);
            if (!dataA && !dataB) return 0;
            if (!dataA) return 1;
            if (!dataB) return -1;
            return dataA - dataB;
        });

    const meta = document.getElementById("tabelaAtividadesResponsavelMeta");
    if (meta) meta.textContent = `${atividades.length} tarefa(s) de ${nome}`;

    if (!atividades.length) {
        corpo.innerHTML = `<tr><td colspan="6" style="text-align:center; color:var(--ink-faint);">Nenhuma atividade encontrada para este responsável</td></tr>`;
        return;
    }

    corpo.innerHTML = atividades.map(item => {
        const prazoFatal = temPrazoFatalSim(item["Solicitação - Há Prazo Fatal"])
            ? `<span class="pill pill-crit">Sim</span>`
            : `<span class="pill pill-neutral">Não</span>`;

        return `
            <tr>
                <td>${formatarDataExcel(item["Data do agendamento"])}</td>
                <td>${item["Processo - ID"] || "-"}</td>
                <td>${item["Tipo"] || "-"}</td>
                <td>${item["Status da tarefa"] || "-"}</td>
                <td>${prazoFatal}</td>
                <td>${obterDescricao(item)}</td>
            </tr>
        `;
    }).join("");
}

// 📊 Gráfico de barras horizontal: tarefas ativas por Área do Direito
function gerarGraficoAreaDireito(dados, canvasId) {
    const ativas = dados.filter(item => !statusIndicaConcluida(item["Status da tarefa"]));

    const contagem = {};
    ativas.forEach(item => {
        const area = (item["Área do Direito"] || "").trim();
        if (area && area !== "-") {
            contagem[area] = (contagem[area] || 0) + 1;
        }
    });

    const entradas = Object.entries(contagem).sort((a, b) => b[1] - a[1]);
    const labels = entradas.map(e => e[0]);
    const valores = entradas.map(e => e[1]);

    const canvas = document.getElementById(canvasId);
    if (!canvas) return null;

    if (charts[canvasId]) {
        charts[canvasId].chart.destroy();
    }

    const corAccent = cssVar('--accent') || '#0058A3';
    const corBorda = cssVar('--border') || '#E2E6EC';
    const corTexto = cssVar('--ink-muted') || '#5B6472';

    const ctx = canvas.getContext("2d");
    const chart = new Chart(ctx, {
        type: "bar",
        data: {
            labels,
            datasets: [{
                label: "Tarefas ativas",
                data: valores,
                backgroundColor: corAccent,
                borderRadius: 4,
                maxBarThickness: 28
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                datalabels: { display: false }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    ticks: { precision: 0, color: corTexto },
                    grid: { color: corBorda }
                },
                y: {
                    ticks: { autoSkip: false, color: corTexto },
                    grid: { display: false }
                }
            }
        },
        plugins: [ChartDataLabels]
    });

    charts[canvasId] = { chart, coluna: "Área do Direito" };
    return chart;
}

// 📊 Gráfico de barras: contagem por Status da tarefa
function gerarGraficoStatusTarefa(dados, canvasId) {
    const contagem = {};
    dados.forEach(item => {
        const status = (item["Status da tarefa"] || "").trim();
        if (status && status !== "-") {
            contagem[status] = (contagem[status] || 0) + 1;
        }
    });

    const entradas = Object.entries(contagem).sort((a, b) => b[1] - a[1]);
    const labels = entradas.map(e => e[0]);
    const valores = entradas.map(e => e[1]);

    const coresBarra = [
        cssVar('--accent') || '#0058A3',
        cssVar('--good') || '#1E8E5A',
        cssVar('--warn') || '#B76E00',
        cssVar('--crit') || '#C4362E',
        cssVar('--ink-muted') || '#5B6472',
        cssVar('--ink-faint') || '#8992A0'
    ];
    const corBorda = cssVar('--border') || '#E2E6EC';
    const corTexto = cssVar('--ink-muted') || '#5B6472';

    const canvas = document.getElementById(canvasId);
    if (!canvas) return null;

    if (charts[canvasId]) {
        charts[canvasId].chart.destroy();
    }

    const ctx = canvas.getContext("2d");
    const chart = new Chart(ctx, {
        type: "bar",
        data: {
            labels,
            datasets: [{
                label: "Tarefas",
                data: valores,
                backgroundColor: labels.map((_, i) => coresBarra[i % coresBarra.length]),
                borderRadius: 4,
                maxBarThickness: 40
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                datalabels: { display: false }
            },
            scales: {
                x: {
                    ticks: { color: corTexto },
                    grid: { display: false }
                },
                y: {
                    beginAtZero: true,
                    ticks: { precision: 0, color: corTexto },
                    grid: { color: corBorda }
                }
            }
        },
        plugins: [ChartDataLabels]
    });

    charts[canvasId] = { chart, coluna: "Status da tarefa" };
    return chart;
}

// 🚀 Inicialização principal
window.onload = async () => {
    try {
        const dados = await carregarExcel();

        // Atualiza estatísticas
        await atualizarEstatisticas();

        // exibirDataB3();

        exibirUltimaAtualizacao(); // ⬅️ Adicionada aqui

        // 🆕 Painel de Produtividade da Equipe
        renderizarTilesResumo(dados);
        renderizarRankingResponsaveis(dados);
        renderizarPainelPrazosFataisNovo(dados);
        gerarGraficoAreaDireito(dados, "graficoAreaDireito");
        gerarGraficoStatusTarefa(dados, "graficoStatusTarefa");

        // 🆕 Aba "Atividades por Responsável"
        popularSelectResponsaveis(dados);

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
        // await adicionarGraficoEvolucaoMensal();

        // 🧑‍💼 Gráfico radar de responsáveis
        await adicionarGraficoRadarResponsaveis();

        // 📊 Gráfico de volume de prazos por dia
        await adicionarGraficoVolumePrazosDiario();


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

        // console.log('✅ Dashboard carregado com sucesso!');

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