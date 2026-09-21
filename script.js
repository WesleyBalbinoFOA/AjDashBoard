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
        <span title="${raw}">Atualizado em ${formatarDataCurta(dataB3Formatada)}</span>
        `;
    } catch (e) {
        console.warn("Erro ao formatar data:", e);
        el.innerHTML = `<span class="dot" style="background:var(--crit);"></span><em>Erro na data</em>`;
    }
}

// Encurta "DD/MM/AAAA HH:MM:SS" para "DD/MM/AA HH:MM" (indicador do topbar,
// onde espaço é curto); o valor completo continua no atributo title.
function formatarDataCurta(dataFormatada) {
    if (!dataFormatada) return dataFormatada;

    const match = dataFormatada.match(/^(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{2}):(\d{2}))?/);
    if (!match) return dataFormatada;

    const [, dia, mes, ano, hora, minuto] = match;
    const anoCurto = ano.slice(-2);

    return hora ? `${dia}/${mes}/${anoCurto} ${hora}:${minuto}` : `${dia}/${mes}/${anoCurto}`;
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

// 👥 Agrupa uma lista de itens por Responsável, ordenando cada grupo pela data
// informada por `extrairData` (por padrão "Data do agendamento") e os grupos
// por quantidade (usado pelos painéis e pelos modais dos tiles)
function agruparPorResponsavel(itens, extrairData = (item) => parseDataAgendamento(item["Data do agendamento"])) {
    const porResponsavel = {};
    itens.forEach(item => {
        const resp = (item["Responsável"] || "").trim() || "Sem responsável";
        if (!porResponsavel[resp]) porResponsavel[resp] = [];
        porResponsavel[resp].push(item);
    });

    const grupos = Object.entries(porResponsavel).map(([responsavel, registros]) => {
        const ordenados = registros.slice().sort((a, b) => {
            const dataA = extrairData(a);
            const dataB = extrairData(b);
            if (!dataA && !dataB) return 0;
            if (!dataA) return 1;
            if (!dataB) return -1;
            return dataA - dataB;
        });
        return { responsavel, registros: ordenados };
    });

    return grupos.sort((a, b) => b.registros.length - a.registros.length);
}

// 🚨 Extrai a data do prazo fatal de texto livre (ex.: "FATAL: 24/09 - ...",
// "Prazo Fatal: 24/09/2026"). O sistema de origem não tem uma coluna de data
// dedicada para o prazo fatal — só o marcador "Sim/Não" — então a data real
// vem embutida na descrição. Sem ano explícito no texto, assume o ano da
// "Data do agendamento" (ou o ano atual, se essa também faltar).
function extrairDataPrazoFatal(item) {
    const campos = [
        "Solicitação - Descrição da Solicitação",
        "Processo - Descrição",
        "Observação"
    ];

    const regex = /(?:prazo\s+)?fatal\s*:?\s*(\d{1,2})[\/\-](\d{1,2})(?:[\/\-](\d{2,4}))?/i;

    for (const campo of campos) {
        const texto = item[campo];
        if (!texto || typeof texto !== "string") continue;

        const match = texto.match(regex);
        if (!match) continue;

        const dia = parseInt(match[1], 10);
        const mes = parseInt(match[2], 10) - 1;
        let ano = match[3] ? parseInt(match[3], 10) : null;
        if (ano !== null && ano < 100) ano += 2000;

        if (ano === null) {
            const referencia = parseDataAgendamento(item["Data do agendamento"]);
            ano = referencia ? referencia.getFullYear() : new Date().getFullYear();
        }

        if (dia >= 1 && dia <= 31 && mes >= 0 && mes <= 11) {
            return new Date(ano, mes, dia);
        }
    }

    return null;
}

// 🎨 Monta o HTML de uma lista agrupada por responsável (seções com avatar + tabela),
// usado tanto no painel "Prazos e Riscos" quanto nos modais dos tiles do topo
function construirListaAgrupadaPorResponsavel(grupos, cabecalhos, renderizarLinha, rotuloItem = "item(ns)") {
    if (!grupos.length) {
        return `<p style="text-align:center; color:var(--ink-faint); padding:1rem 0;">Nenhum item encontrado</p>`;
    }

    return grupos.map(grupo => {
        const iniciais = grupo.responsavel.split(" ")
            .filter(Boolean)
            .slice(0, 2)
            .map(p => p[0].toUpperCase())
            .join("") || "?";

        const linhas = grupo.registros.map(renderizarLinha).join("");

        return `
            <div class="responsavel-grupo">
                <div class="responsavel-grupo-header">
                    <span class="avatar">${iniciais}</span>
                    <span class="responsavel-grupo-nome">${grupo.responsavel}</span>
                    <span class="responsavel-grupo-count">${grupo.registros.length} ${rotuloItem}</span>
                </div>
                <div style="overflow-x:auto;">
                    <table class="ranking-table">
                        <thead><tr>${cabecalhos.map(c => `<th>${c}</th>`).join("")}</tr></thead>
                        <tbody>${linhas}</tbody>
                    </table>
                </div>
            </div>
        `;
    }).join("");
}

// 🚨 Prazos fatais ainda pendentes, agrupados por responsável (com os itens individuais),
// ordenados pela data do prazo fatal extraída do texto (não pela Data do agendamento)
function calcularPrazosFatais(dados) {
    const pendentesComPrazo = dados.filter(item =>
        temPrazoFatalSim(item["Solicitação - Há Prazo Fatal"]) &&
        !statusIndicaConcluida(item["Status da tarefa"])
    );

    return agruparPorResponsavel(pendentesComPrazo, extrairDataPrazoFatal);
}

// 🎨 Linha padrão de uma tabela de prazos fatais (Processo ID | Área | Prazo Fatal | Situação)
function linhaPrazoFatal(item) {
    const dataFatal = extrairDataPrazoFatal(item);
    const situacao = situacaoPrazoItem(item);
    return `
        <tr>
            <td>${item["Processo - ID"] || "-"}</td>
            <td>${item["Área do Direito"] || "-"}</td>
            <td>${dataFatal ? formatarDataExcel(dataFatal) : "Não identificado"}</td>
            <td><span class="pill ${situacao.classe}">${situacao.texto}</span></td>
        </tr>
    `;
}

// 🚨 Situação de vencimento de um item individual de prazo fatal
// (com base na data do prazo fatal extraída do texto, não na Data do agendamento)
function situacaoPrazoItem(item) {
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    const data = extrairDataPrazoFatal(item);
    if (!data) return { texto: "Data não identificada", classe: "pill-neutral" };

    const diffDias = Math.round((data - hoje) / 86400000);

    if (diffDias < 0) return { texto: `${Math.abs(diffDias)} dia(s) atrasado`, classe: "pill-crit" };
    if (diffDias === 0) return { texto: "Vence hoje", classe: "pill-crit" };
    if (diffDias === 1) return { texto: "Vence amanhã", classe: "pill-warn" };
    if (diffDias <= 7) return { texto: `Vence em ${diffDias} dias`, classe: "pill-warn" };
    return { texto: `Vence em ${diffDias} dias`, classe: "pill-neutral" };
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

// 🎨 Renderiza os prazos com risco fatal em seções por responsável, listando cada item individualmente
function renderizarPainelPrazosFataisNovo(dados) {
    const grupos = calcularPrazosFatais(dados);
    const container = document.getElementById("prazosFataisPorResponsavel");
    if (!container) return;

    const meta = document.getElementById("prazosFataisMeta");
    if (meta) {
        const totalItens = grupos.reduce((acc, grupo) => acc + grupo.registros.length, 0);
        meta.textContent = `${totalItens} prazo(s) fatal(is) em aberto no momento do carregamento.`;
    }

    container.innerHTML = construirListaAgrupadaPorResponsavel(
        grupos,
        ["Processo ID", "Área", "Prazo Fatal", "Situação"],
        linhaPrazoFatal,
        "prazo(s) fatal(is)"
    );
}

// ============================================================
// 🆕 Modal genérico (usado pelos tiles clicáveis do topo)
// ============================================================

function abrirModal(titulo, corpoHtml) {
    const overlay = document.getElementById("modalOverlay");
    const tituloEl = document.getElementById("modalTitulo");
    const corpoEl = document.getElementById("modalCorpo");
    if (!overlay || !tituloEl || !corpoEl) return;

    tituloEl.textContent = titulo;
    corpoEl.innerHTML = corpoHtml;
    overlay.hidden = false;
}

function fecharModal() {
    const overlay = document.getElementById("modalOverlay");
    if (overlay) overlay.hidden = true;
}

document.addEventListener("click", (e) => {
    const overlay = document.getElementById("modalOverlay");
    if (overlay && !overlay.hidden && e.target === overlay) {
        fecharModal();
    }
});

document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") fecharModal();
});

// Tarefas ativas cuja "Data do agendamento" cai exatamente em `dia`
function filtrarAtivasPorDia(dados, dia) {
    return dados.filter(item => {
        if (statusIndicaConcluida(item["Status da tarefa"])) return false;
        const data = parseDataAgendamento(item["Data do agendamento"]);
        return data && mesmoDia(data, dia);
    });
}

// 🎨 Linha padrão de uma tarefa nos modais "Vencem hoje"/"Vencem amanhã"
function linhaModalTarefa(item) {
    return `
        <tr>
            <td>${item["Processo - ID"] || "-"}</td>
            <td>${item["Área do Direito"] || "-"}</td>
            <td>${item["Tipo"] || "-"}</td>
            <td><span class="pill ${pillClasseStatus(item["Status da tarefa"])}">${item["Status da tarefa"] || "-"}</span></td>
        </tr>
    `;
}

// Abre o modal do tile "Vencem hoje" (offsetDias=0) ou "Vencem amanhã" (offsetDias=1)
function abrirModalPorDia(offsetDias, titulo) {
    const dia = new Date();
    dia.setHours(0, 0, 0, 0);
    dia.setDate(dia.getDate() + offsetDias);

    const itens = filtrarAtivasPorDia(dadosExcel, dia);
    const grupos = agruparPorResponsavel(itens);
    const corpo = construirListaAgrupadaPorResponsavel(
        grupos,
        ["Processo ID", "Área", "Tipo", "Status"],
        linhaModalTarefa,
        "tarefa(s)"
    );

    abrirModal(`${titulo} (${itens.length})`, corpo);
}

// Abre o modal do tile "Prazo fatal em aberto"
function abrirModalPrazoFatal() {
    const grupos = calcularPrazosFatais(dadosExcel);
    const totalItens = grupos.reduce((acc, grupo) => acc + grupo.registros.length, 0);

    const corpo = construirListaAgrupadaPorResponsavel(
        grupos,
        ["Processo ID", "Área", "Prazo Fatal", "Situação"],
        linhaPrazoFatal,
        "prazo(s) fatal(is)"
    );

    abrirModal(`Prazo fatal em aberto (${totalItens})`, corpo);
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

// Abrevia o nome para caber na aba (primeiro + último nome); nomes com até
// duas palavras já são curtos o bastante e voltam sem alteração
function nomeResumido(nomeCompleto) {
    const partes = nomeCompleto.trim().split(/\s+/).filter(Boolean);
    if (partes.length <= 2) return nomeCompleto;
    return `${partes[0]} ${partes[partes.length - 1]}`;
}

// 🎨 Monta as abas de responsável da aba "Atividades por Responsável" e renderiza a primeira
function popularTabsResponsaveis(dados) {
    const container = document.getElementById("tabsResponsavelAtividades");
    if (!container) return;

    const nomes = obterListaResponsaveis(dados);
    container.innerHTML = "";

    nomes.forEach((nome, index) => {
        const aba = document.createElement("button");
        aba.type = "button";
        aba.className = "person-tab" + (index === 0 ? " active" : "");
        aba.textContent = nomeResumido(nome);
        aba.title = nome;
        aba.addEventListener("click", () => {
            container.querySelectorAll(".person-tab").forEach(el => el.classList.remove("active"));
            aba.classList.add("active");
            renderizarAtividadesPorResponsavel(dados, nome);
        });
        container.appendChild(aba);
    });

    if (nomes.length) {
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

// 🎨 Classe de pill correspondente ao status de uma tarefa
function pillClasseStatus(status) {
    if (!status) return "pill-neutral";
    const s = status.toLowerCase();

    if (s.includes("atras") || s.includes("venc")) return "pill-crit";
    if (s.includes("aguard")) return "pill-warn";
    if (s.includes("ativo") || s.includes("pendente")) return "pill-good";
    return "pill-neutral"; // concluído, em andamento, etc.
}

// 🎨 Renderiza a tabela de audiências agendadas (padrão visual dos novos painéis)
function renderizarPainelAudiencias(dados) {
    const corpo = document.getElementById("tabelaAudienciasBody");
    if (!corpo) return;

    const audiencias = dados
        .filter(item => {
            const tipo = item["Tipo"];
            return tipo && (
                tipo.toLowerCase().includes("audiência") ||
                tipo.toLowerCase().includes("audiencia") ||
                tipo.toLowerCase().includes("hearing")
            );
        })
        .slice()
        .sort((a, b) => {
            const dataA = parseDataAgendamento(a["Data do agendamento"]);
            const dataB = parseDataAgendamento(b["Data do agendamento"]);
            if (!dataA && !dataB) return 0;
            if (!dataA) return 1;
            if (!dataB) return -1;
            return dataA - dataB;
        });

    const meta = document.getElementById("audienciasMeta");
    if (meta) meta.textContent = `${audiencias.length} audiência(s) agendada(s)`;

    if (!audiencias.length) {
        corpo.innerHTML = `<tr><td colspan="5" style="text-align:center; color:var(--ink-faint);">Nenhuma audiência agendada</td></tr>`;
        return;
    }

    corpo.innerHTML = audiencias.map(item => {
        const nome = (item["Responsável"] || "").trim() || "-";
        const iniciais = nome !== "-"
            ? nome.split(" ").filter(Boolean).slice(0, 2).map(p => p[0].toUpperCase()).join("") || "?"
            : "?";

        return `
            <tr>
                <td>${formatarDataExcel(item["Data do agendamento"])}</td>
                <td>${item["Processo - ID"] || "-"}</td>
                <td>${item["Empresa"] || "-"}</td>
                <td>
                    <div class="ranking-row-name">
                        <span class="avatar">${iniciais}</span>
                        <span>${nome}</span>
                    </div>
                </td>
                <td><span class="pill ${pillClasseStatus(item["Status da tarefa"])}">${item["Status da tarefa"] || "-"}</span></td>
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

        exibirUltimaAtualizacao();

        // 🆕 Painel de Produtividade da Equipe
        renderizarTilesResumo(dados);
        renderizarRankingResponsaveis(dados);
        renderizarPainelPrazosFataisNovo(dados);
        gerarGraficoAreaDireito(dados, "graficoAreaDireito");
        gerarGraficoStatusTarefa(dados, "graficoStatusTarefa");

        // 🆕 Aba "Atividades por Responsável"
        popularTabsResponsaveis(dados);

        // 🆕 Aba "Audiências Agendadas"
        renderizarPainelAudiencias(dados);

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