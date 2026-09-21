// ============================================================
// Model / aggregations.js — agregações e agrupamentos usados pelos painéis
// ============================================================

import { parseDataAgendamento, mesmoDia } from "./dateUtils.js";
import { statusIndicaConcluida, temPrazoFatalSim, extrairDataPrazoFatal } from "./domain.js";

// 🧑‍💼 Ranking de tarefas ativas por responsável
export function calcularRankingResponsaveis(dados) {
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

// 📊 Resumo da fila para os tiles do topo
export function calcularResumoFila(dados) {
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

// 👥 Agrupa uma lista de itens por Responsável, ordenando cada grupo pela data
// informada por `extrairData` (por padrão "Data do agendamento") e os grupos
// por quantidade (usado pelos painéis e pelos modais dos tiles)
export function agruparPorResponsavel(itens, extrairData = (item) => parseDataAgendamento(item["Data do agendamento"])) {
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

// 🚨 Prazos fatais ainda pendentes, agrupados por responsável (com os itens individuais),
// ordenados pela data do prazo fatal extraída do texto (não pela Data do agendamento)
export function calcularPrazosFatais(dados) {
    const pendentesComPrazo = dados.filter(item =>
        temPrazoFatalSim(item["Solicitação - Há Prazo Fatal"]) &&
        !statusIndicaConcluida(item["Status da tarefa"])
    );

    return agruparPorResponsavel(pendentesComPrazo, extrairDataPrazoFatal);
}

// 🧑‍💼 Lista de responsáveis distintos presentes na planilha (ordenada)
export function obterListaResponsaveis(dados) {
    const nomes = new Set();
    dados.forEach(item => {
        const nome = (item["Responsável"] || "").trim();
        if (nome) nomes.add(nome);
    });
    return Array.from(nomes).sort((a, b) => a.localeCompare(b, 'pt-BR'));
}

// Tarefas ativas cuja "Data do agendamento" cai exatamente em `dia`
export function filtrarAtivasPorDia(dados, dia) {
    return dados.filter(item => {
        if (statusIndicaConcluida(item["Status da tarefa"])) return false;
        const data = parseDataAgendamento(item["Data do agendamento"]);
        return data && mesmoDia(data, dia);
    });
}
