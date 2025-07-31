const CORES = {
      roxo:    { fundo: "rgba(153, 102, 255, 0.5)", borda: "rgba(153, 102, 255, 1)" },
      vermelho:{ fundo: "rgba(255, 99, 132, 0.5)",  borda: "rgba(255, 99, 132, 1)" },
      azul:    { fundo: "rgba(54, 162, 235, 0.5)",  borda: "rgba(54, 162, 235, 1)" },
      verde:   { fundo: "rgba(75, 192, 192, 0.5)",  borda: "rgba(75, 192, 192, 1)" },
      laranja: { fundo: "rgba(255, 159, 64, 0.5)",  borda: "rgba(255, 159, 64, 1)" }
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

    async function carregarExcel() {
      if (dadosExcel.length && dadosAtualizadosHoje()) {
        return dadosExcel;
      }

      // 🆕 Verifica se deve forçar atualização baseado no horário
      const forcarAtualizacao = deveAtualizarDados() && !dadosAtualizadosHoje();

      const dadosSalvos = localStorage.getItem("dadosExcel");
      if (dadosSalvos && !forcarAtualizacao) {
        dadosExcel = JSON.parse(dadosSalvos);
        return dadosExcel;
      }

      console.log("🔄 Carregando dados do Excel...");
      
      const response = await fetch(excelUrl);
      const blob = await response.blob();
      const buffer = await blob.arrayBuffer();

      const workbook = XLSX.read(buffer, { type: "array" });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      dadosExcel = XLSX.utils.sheet_to_json(worksheet, {
        range: 5,
        defval: ""
      });

      // 🆕 Salva os dados e marca o timestamp da atualização
      localStorage.setItem("dadosExcel", JSON.stringify(dadosExcel));
      localStorage.setItem("ultimaAtualizacaoExcel", new Date().toISOString());
      
      console.log("✅ Dados atualizados com sucesso!");
      
      return dadosExcel;
    }

    function gerarGraficoPorColuna(coluna, dados, canvasId, cor = CORES.roxo) {
    const contagem = {};
    dados.forEach(item => {
        const chave = item[coluna];
        if (chave) contagem[chave] = (contagem[chave] || 0) + 1;
    });

    const nomesCompletos = Object.keys(contagem);
    const valores = Object.values(contagem);

    const labels = nomesCompletos.map(nome => {
        return coluna.toLowerCase() === "responsável" ? nome.split(" ")[0] : nome;
    });

    const ctx = document.getElementById(canvasId).getContext("2d");

    const chart = new Chart(ctx, {
        type: "bar",
        data: {
        labels: labels,
        datasets: [{
            label: `Quantidade por ${coluna}`,
            data: valores,
            backgroundColor: cor.fundo,
            borderColor: cor.borda,
            borderWidth: 1
        }]
        },
        options: {
        responsive: true,
        onClick: (e, elements) => {
            if (coluna.toLowerCase() !== "responsável" || elements.length === 0) return;

            const index = elements[0].index;
            const primeiroNomeClicado = chart.data.labels[index];
            
            // Encontra nomes completos que batem com o primeiro nome
            const nomesCorrespondentes = nomesCompletos.filter(nome => nome.startsWith(primeiroNomeClicado));

            const resultados = dados.filter(item => 
            nomesCorrespondentes.includes(item["Responsável"])
            );

            exibirTabela(canvasId, resultados);
        },
        plugins: {
            datalabels: {
            anchor: 'end',
            align: 'top',
            color: '#000',
            font: { weight: 'bold', size: 12 },
            formatter: Math.round
            }
        },
        scales: {
            y: {
            beginAtZero: true,
            title: { display: true, text: 'Quantidade' }
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

    // 🆕 Função para formatar data e hora do Excel para formato brasileiro
    function formatarDataExcel(valorData) {
        if (!valorData || valorData === "-") return "-";
        
        // Se já estiver em formato de string de data, tenta converter
        if (typeof valorData === 'string' && valorData.includes('/')) {
            return valorData; // Já está formatado
        }
        
        // Se for número serial do Excel, converte
        if (typeof valorData === 'number') {
            // Excel conta dias desde 01/01/1900, mas com bug do ano 1900
            // JavaScript conta milissegundos desde 01/01/1970
            const diasDesde1900 = valorData - 25569; // Ajuste para JavaScript
            const data = new Date(diasDesde1900 * 86400 * 1000);
            
            // Formatar para DD/MM/AAAA HH:MM
            const dia = String(data.getDate()).padStart(2, '0');
            const mes = String(data.getMonth() + 1).padStart(2, '0');
            const ano = data.getFullYear();
            const horas = String(data.getHours()).padStart(2, '0');
            const minutos = String(data.getMinutes()).padStart(2, '0');
            
            return `${dia}/${mes}/${ano} ${horas}:${minutos}`;
        }
        
        // Tenta converter string diretamente
        try {
            const data = new Date(valorData);
            if (!isNaN(data.getTime())) {
                const dia = String(data.getDate()).padStart(2, '0');
                const mes = String(data.getMonth() + 1).padStart(2, '0');
                const ano = data.getFullYear();
                const horas = String(data.getHours()).padStart(2, '0');
                const minutos = String(data.getMinutes()).padStart(2, '0');
                
                return `${dia}/${mes}/${ano} ${horas}:${minutos}`;
            }
        } catch (e) {
            // Se não conseguir converter, retorna o valor original
        }
        
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


    async function adicionarGrafico(coluna) {
      const dados = await carregarExcel();

      const idUnico = `grafico_${Math.random().toString(36).substr(2, 9)}`;
      const container = document.createElement("div");
      container.className = "grafico-container";
      container.innerHTML = `
        <div class="grafico-header">
          <strong>Gráfico por ${coluna}</strong>
          <select onchange="trocarCor('${idUnico}', '${coluna}', this.value)">
            ${Object.keys(CORES).map(cor => `<option value="${cor}">${cor[0].toUpperCase() + cor.slice(1)}</option>`).join('')}
          </select>
        </div>
        <canvas id="${idUnico}"></canvas>
      `;

      document.getElementById("graficos").appendChild(container);
      const chart = gerarGraficoPorColuna(coluna, dados, idUnico, CORES.roxo);
      charts[idUnico] = { chart, coluna };
    }

    function trocarCor(canvasId, coluna, corSelecionada) {
      const { chart } = charts[canvasId];
      const novaCor = CORES[corSelecionada];

      chart.data.datasets[0].backgroundColor = novaCor.fundo;
      chart.data.datasets[0].borderColor = novaCor.borda;
      chart.update();
    }

    // pizzaaaaaaaaaaa, who doesn't love pizza? 🍕 
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
        '#FF6384', // Rosa
        '#36A2EB', // Azul
        '#FFCE56', // Amarelo
        '#4BC0C0', // Verde água
        '#9966FF', // Roxo
        '#FF9F40', // Laranja
        '#FF6384', // Rosa claro
        '#C9CBCF', // Cinza
        '#4BC0C0', // Verde
        '#FF6384'  // Rosa escuro
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
                        font: {
                            size: 12
                        }
                    }
                },
                datalabels: {
                    color: '#fff',
                    font: {
                        weight: 'bold',
                        size: 14
                    },
                    formatter: (value, context) => {
                        const total = context.dataset.data.reduce((a, b) => a + b, 0);
                        const porcentagem = ((value / total) * 100).toFixed(1);
                        return `${value}\n(${porcentagem}%)`;
                    },
                    textAlign: 'center'
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
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

                // Exibe a tabela modal (reutiliza a função existente)
                exibirTabela(canvasId, resultados);
            }
        },
        plugins: [ChartDataLabels]
    });

    return chart;
}

// 🚀 Função para adicionar o gráfico de pizza ao dashboard
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

// 📅 Função para gerar tabela de audiências
function gerarTabelaAudiencias(dados) {
    // Filtra registros do tipo "Audiência" ou similar
    const audiencias = dados.filter(item => {
        const tipo = item["Tipo"];
        return tipo && (
            tipo.toLowerCase().includes("audiência") ||
            tipo.toLowerCase().includes("audiencia") ||
            tipo.toLowerCase().includes("hearing")
        );
    });

    // Ordena por data de agendamento
    audiencias.sort((a, b) => {
        const dataA = new Date(a["Data do agendamento"]);
        const dataB = new Date(b["Data do agendamento"]);
        return dataA - dataB;
    });

    const container = document.createElement("div");
    container.className = "row";
    container.innerHTML = `
        <div class="col s12">
            <div class="card">
                <div class="card-content">
                    <span class="card-title">📅 Audiências Agendadas</span>
                    <div style="max-height: 400px; overflow-y: auto;">
                        <table class="striped highlight responsive-table">
                            <thead>
                                <tr>
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


function gerarTabelaAudiencias(dados) {
    // Filtra registros do tipo "Audiência" ou similar
    const audiencias = dados.filter(item => {
        const tipo = item["Tipo"];
        return tipo && (
            tipo.toLowerCase().includes("audiência") ||
            tipo.toLowerCase().includes("audiencia") ||
            tipo.toLowerCase().includes("hearing")
        );
    });

    // Ordena por data de agendamento
    audiencias.sort((a, b) => {
        const dataA = new Date(a["Data do agendamento"]);
        const dataB = new Date(b["Data do agendamento"]);
        return dataA - dataB;
    });

    const container = document.createElement("div");
    container.className = "row";
    container.id = "audiencias"; // ID para navegação
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

// 🚨 Função para gerar tabela de prazos fatais atrasados
function gerarTabelaPrazosFatais(dados) {
    // Filtra registros com prazo fatal "Atrasados"
    const atrasados = dados.filter(item => {
        const prazoFatal = item["Solicitação - Há Prazo Fatal"];
        return prazoFatal && prazoFatal.toLowerCase().includes("atrasados");
    });

    // Ordena por data de agendamento (mais antigos primeiro)
    atrasados.sort((a, b) => {
        const dataA = new Date(a["Data do agendamento"]);
        const dataB = new Date(b["Data do agendamento"]);
        return dataA - dataB;
    });

    const container = document.createElement("div");
    container.className = "row";
    container.id = "prazos-atrasados"; // <-- Necessário para rolarPara funcionar
    container.innerHTML = `
        <div class="col s12">
            <div class="card red lighten-5">
                <div class="card-content">
                    <span class="card-title red-text">🚨 Prazos Fatais Atrasados</span>
                    <div style="max-height: 400px; overflow-y: auto;">
                        <table class="striped highlight responsive-table">
                            <thead>
                                <tr>
                                    <th>Processo ID</th>
                                    <th>Data do Agendamento</th>
                                    <th>Responsável</th>
                                    <th>Empresa</th>
                                    <th>Tipo</th>
                                    <th>Status da Tarefa</th>
                                    <th>Observação</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${atrasados.map(reg => `
                                    <tr class="red lighten-4">
                                        <td><strong>${reg["Processo - ID"] || "-"}</strong></td>
                                        <td><strong class="red-text">${formatarDataExcel(reg["Data do agendamento"])}</strong></td>
                                        <td>${reg["Responsável"] || "-"}</td>
                                        <td>${reg["Empresa"] || "-"}</td>
                                        <td><span class="chip orange white-text">${reg["Tipo"] || "-"}</span></td>
                                        <td><span class="chip red white-text">${reg["Status da tarefa"] || "-"}</span></td>
                                        <td>${reg["Observação"] || "-"}</td>
                                    </tr>
                                `).join("")}
                            </tbody>
                        </table>
                    </div>
                    <div class="card-action red lighten-4">
                        <span><strong class="red-text">⚠️ Total de processos atrasados:</strong> ${atrasados.length}</span>
                    </div>
                </div>
            </div>
        </div>
    `;

    return container;
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

// 🚀 Função para adicionar ambas as tabelas ao dashboard
async function adicionarTabelasEspeciais() {
    const dados = await carregarExcel();
    
    // Cria container principal
    const containerPrincipal = document.createElement("div");
    containerPrincipal.className = "container";
    containerPrincipal.style.marginTop = "20px";
    
    // Adiciona título
    const titulo = document.createElement("h4");
    titulo.textContent = "📊 Relatórios Especiais";
    titulo.className = "center-align";
    containerPrincipal.appendChild(titulo);
    
    // Gera e adiciona tabela de audiências
    const tabelaAudiencias = gerarTabelaAudiencias(dados);
    containerPrincipal.appendChild(tabelaAudiencias);
    
    // Gera e adiciona tabela de prazos fatais
    const tabelaPrazos = gerarTabelaPrazosFatais(dados);
    containerPrincipal.appendChild(tabelaPrazos);
    
    // Adiciona ao final da página
    document.body.appendChild(containerPrincipal);
}
        async function atualizarEstatisticas() {
      const dados = await carregarExcel();
      
      // Total de processos
      document.getElementById('totalProcessos').textContent = dados.length;
      
      // Total de audiências
      const audiencias = dados.filter(item => {
        const tipo = item["Tipo"];
        return tipo && tipo.toLowerCase().includes("audiência");
      });
      document.getElementById('totalAudiencias').textContent = audiencias.length;
      
      // Tarefas pendentes
      const pendentes = dados.filter(item => {
        const status = item["Status da tarefa"];
        return status && (status.toLowerCase().includes("pendente") || status.toLowerCase().includes("ativo"));
      });
      document.getElementById('totalPendentes').textContent = pendentes.length;
      
      // Prazos atrasados
      const atrasados = dados.filter(item => {
        const prazoFatal = item["Solicitação - Há Prazo Fatal"];
        return prazoFatal && prazoFatal.toLowerCase().includes("atrasados");
      });
      document.getElementById('totalAtrasados').textContent = atrasados.length;
    }






  // 🚀 Inicialização melhorada
    window.onload = async () => {
      try {
        await carregarExcel();
        
        // Atualiza estatísticas
        await atualizarEstatisticas();
        
        // Gráficos existentes
        const colunas = ["Responsável", "Área do Direito"];
        for (const coluna of colunas) {
          await adicionarGrafico(coluna);
        }
        
        // Gráfico de pizza
        await adicionarGraficoPizza();
        
        // Tabelas especiais
        await adicionarTabelasEspeciais();
        
        // Adiciona animações aos elementos
        document.querySelectorAll('.grafico-container').forEach((el, index) => {
          el.classList.add('fade-in');
          el.style.animationDelay = `${index * 0.1}s`;
        });
        
        // Remove loading overlay
        document.getElementById('loadingOverlay').style.display = 'none';
        
        console.log('✅ Dashboard carregado com sucesso!');
        
      } catch (error) {
        console.error('❌ Erro ao carregar dashboard:', error);
        document.getElementById('loadingOverlay').innerHTML = `
          <div class="loading-content">
            <i class="material-icons" style="font-size: 4rem; color: #f44336;">error</i>
            <h5 style="color: #f44336;">Erro ao carregar dados</h5>
            <p>Verifique a conexão e tente novamente</p>
          </div>
        `;
      }
    };