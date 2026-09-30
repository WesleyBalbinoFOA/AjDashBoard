// ============================================================
// Model / excelService.js — carregamento da planilha e cache local
// ============================================================

import { excelUrl, dadosExcel, dataB3, dataB3Formatada, setDadosExcel, setDataB3, setDataB3Formatada } from "./state.js";
import { converterDataExcelParaPtBR, converterDataExcelParaPtBRComHora } from "./dateUtils.js";

export function limparLocalStorage() {
    localStorage.removeItem("dadosExcel");
    localStorage.removeItem("ultimaAtualizacaoExcel");
    localStorage.removeItem("dataB3");           // 🆕
    localStorage.removeItem("dataB3Formatada");  // 🆕
}

// 🆕 Função para capturar dados específicos de células
export function capturarDadosCelulas(worksheet) {
    try {
        // Captura a célula B3
        const celulaB3 = worksheet['B3'];

        if (celulaB3 && celulaB3.v !== undefined) {
            setDataB3(celulaB3.v); // Valor bruto da célula
            setDataB3Formatada(converterDataExcelParaPtBRComHora(celulaB3.v));

            console.log(`📅 Célula B3 capturada:`);
            console.log(`   Valor bruto: ${celulaB3.v}`);
            console.log(`   Formatado: ${converterDataExcelParaPtBRComHora(celulaB3.v)}`);
        } else {
            console.warn('⚠️ Célula B3 não encontrada ou vazia');
            setDataB3(null);
            setDataB3Formatada(null);
        }

        // 🆕 Você pode capturar outras células aqui se necessário
        // Exemplo: const celulaC3 = worksheet['C3'];

    } catch (error) {
        console.error('❌ Erro ao capturar dados das células:', error);
        setDataB3(null);
        setDataB3Formatada(null);
    }
}

// ⚠️ Erro com mensagem amigável para exibir na tela de carregamento
export class ErroPlanilha extends Error {
    constructor(titulo, orientacao, detalheTecnico) {
        super(`${titulo} ${orientacao}`);
        this.name = "ErroPlanilha";
        this.titulo = titulo;
        this.orientacao = orientacao;
        this.detalheTecnico = detalheTecnico;
    }
}

const ORIENTACAO_LINK =
    "No SharePoint, abra a planilha em Compartilhar → Configurações do link, " +
    "escolha \"Qualquer pessoa com o link\", copie o novo link e atualize " +
    "EXCEL_URL no arquivo env.js.";

// 📥 Baixa a planilha e garante que o conteúdo é realmente um .xlsx
async function baixarPlanilha() {
    if (!excelUrl) {
        throw new ErroPlanilha(
            "O link da planilha não está configurado.",
            "Crie o arquivo env.js a partir do env.example.js (na raiz do projeto) " +
            "e preencha EXCEL_URL com o link de download da planilha.",
            "window.ENV.EXCEL_URL vazio ou env.js não carregado"
        );
    }

    let response;
    try {
        response = await fetch(excelUrl);
    } catch (erroRede) {
        // fetch só lança exceção em falha de rede ou bloqueio de CORS —
        // no SharePoint isso normalmente significa que o link exige login.
        throw new ErroPlanilha(
            "Não foi possível acessar a planilha no SharePoint.",
            "Verifique sua conexão. Se a internet estiver ok, o link provavelmente " +
            "exige login ou não é público. " + ORIENTACAO_LINK,
            erroRede.message
        );
    }

    if (response.status === 401 || response.status === 403) {
        throw new ErroPlanilha(
            "Sem permissão para acessar a planilha.",
            "O link exige login ou o compartilhamento foi removido. " + ORIENTACAO_LINK,
            `HTTP ${response.status}`
        );
    }
    if (response.status === 404) {
        throw new ErroPlanilha(
            "A planilha não foi encontrada.",
            "O arquivo pode ter sido movido, renomeado ou excluído. " + ORIENTACAO_LINK,
            "HTTP 404"
        );
    }
    if (!response.ok) {
        throw new ErroPlanilha(
            "O SharePoint retornou um erro ao baixar a planilha.",
            "Tente novamente em alguns minutos. Se persistir, confira o link em env.js.",
            `HTTP ${response.status} ${response.statusText}`
        );
    }

    const buffer = await response.arrayBuffer();

    // Um .xlsx é um arquivo ZIP e sempre começa com os bytes "PK".
    // Se o SharePoint devolver uma página HTML (login, acesso negado,
    // link expirado), o XLSX.read tenta ler como tabela HTML e falha
    // com "Invalid HTML: could not find <table>".
    const bytes = new Uint8Array(buffer, 0, Math.min(2, buffer.byteLength));
    if (bytes[0] !== 0x50 || bytes[1] !== 0x4B) {
        throw new ErroPlanilha(
            "O link não retornou a planilha Excel — o compartilhamento provavelmente expirou.",
            "O SharePoint devolveu uma página (login ou acesso negado) em vez do arquivo. " +
            ORIENTACAO_LINK,
            `Content-Type recebido: ${response.headers.get("content-type") || "desconhecido"}`
        );
    }

    return buffer;
}

// 🔄 Função carregarExcel MODIFICADA para incluir captura da célula B3
export async function carregarExcel() {
    if (dadosExcel.length) {
        return dadosExcel;
    }

    const dadosSalvos = localStorage.getItem("dadosExcel");
    const dataB3Salva = localStorage.getItem("dataB3");
    const dataB3FormatadaSalva = localStorage.getItem("dataB3Formatada");

    if (dadosSalvos) {
        setDadosExcel(JSON.parse(dadosSalvos));
        setDataB3(dataB3Salva ? JSON.parse(dataB3Salva) : null);
        setDataB3Formatada(dataB3FormatadaSalva || null);

        console.log(`✅ Dados carregados do localStorage: ${dadosExcel.length} registros`);
        console.log(`📅 Data B3 recuperada: ${dataB3FormatadaSalva}`);
        return dadosExcel;
    }

    console.log("🔄 Carregando dados do Excel...");

    try {
        const buffer = await baixarPlanilha();

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
        const dadosConvertidos = dadosBrutos.map((registro, index) => {
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

        setDadosExcel(dadosConvertidos);

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
export function obterDataB3() {
    return {
        valorBruto: dataB3,
        formatada: dataB3Formatada
    };
}

// 🧹 Função para limpar o cache e forçar recarregamento dos dados
export function limparCacheERecarregar() {
    // console.log("🧹 Limpando cache e recarregando dados...");

    // Remove dados do localStorage
    localStorage.removeItem("dadosExcel");
    localStorage.removeItem("ultimaAtualizacaoExcel");

    // Limpa variável global
    setDadosExcel([]);

    // console.log("✅ Cache limpo! Recarregue a página para baixar dados atualizados.");

    // Opcionalmente, pode recarregar a página automaticamente:
    // window.location.reload();
}
