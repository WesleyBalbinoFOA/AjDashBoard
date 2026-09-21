// ============================================================
// Model / domain.js — regras de negócio sobre um registro da planilha
// ============================================================

import { parseDataAgendamento } from "./dateUtils.js";

// Considera "ativa" toda tarefa cujo status não indique conclusão/cancelamento
export function statusIndicaConcluida(status) {
    if (!status) return false;
    const s = status.toString().toLowerCase();
    return s.includes("concluíd") || s.includes("concluid") ||
        s.includes("finalizad") || s.includes("cancelad");
}

export function temPrazoFatalSim(valor) {
    if (!valor) return false;
    return valor.toString().trim().toLowerCase() === "sim";
}

// 🚨 Extrai a data do prazo fatal de texto livre (ex.: "FATAL: 24/09 - ...",
// "Prazo Fatal: 24/09/2026"). O sistema de origem não tem uma coluna de data
// dedicada para o prazo fatal — só o marcador "Sim/Não" — então a data real
// vem embutida na descrição. Sem ano explícito no texto, assume o ano da
// "Data do agendamento" (ou o ano atual, se essa também faltar).
export function extrairDataPrazoFatal(item) {
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

// 🚨 Situação de vencimento de um item individual de prazo fatal
// (com base na data do prazo fatal extraída do texto, não na Data do agendamento)
export function situacaoPrazoItem(item) {
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

// 📅 Situação de um item com base na "Data do agendamento" (data de cumprimento
// da atividade, diferente do prazo fatal extraído do texto). Usado na aba
// "Atividades por Responsável" para deixar claro o que já está vencido.
export function situacaoAgendamento(item) {
    if (statusIndicaConcluida(item["Status da tarefa"])) {
        return { texto: "Concluído", classe: "pill-neutral" };
    }

    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    const data = parseDataAgendamento(item["Data do agendamento"]);
    if (!data) return { texto: "Data não identificada", classe: "pill-neutral" };

    const diffDias = Math.round((data - hoje) / 86400000);

    if (diffDias < 0) return { texto: `${Math.abs(diffDias)} dia(s) em atraso`, classe: "pill-crit" };
    if (diffDias === 0) return { texto: "Vence hoje", classe: "pill-crit" };
    if (diffDias === 1) return { texto: "Vence amanhã", classe: "pill-warn" };
    return { texto: `Vence em ${diffDias} dias`, classe: "pill-neutral" };
}

// 🆕 Função para obter a descrição preenchida (ordem de prioridade)
export function obterDescricao(registro) {
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

// 🎨 Classe de pill correspondente ao status de uma tarefa
export function pillClasseStatus(status) {
    if (!status) return "pill-neutral";
    const s = status.toLowerCase();

    if (s.includes("atras") || s.includes("venc")) return "pill-crit";
    if (s.includes("aguard")) return "pill-warn";
    if (s.includes("ativo") || s.includes("pendente")) return "pill-good";
    return "pill-neutral"; // concluído, em andamento, etc.
}

// Abrevia o nome para caber na aba (primeiro + último nome); nomes com até
// duas palavras já são curtos o bastante e voltam sem alteração
export function nomeResumido(nomeCompleto) {
    const partes = nomeCompleto.trim().split(/\s+/).filter(Boolean);
    if (partes.length <= 2) return nomeCompleto;
    return `${partes[0]} ${partes[partes.length - 1]}`;
}
