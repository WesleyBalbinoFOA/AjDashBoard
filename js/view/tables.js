// ============================================================
// View / tables.js — tabelas e listas agrupadas por responsável
// ============================================================

import { parseDataAgendamento, formatarDataExcel } from "../model/dateUtils.js";
import { temPrazoFatalSim, extrairDataPrazoFatal, situacaoPrazoItem, obterDescricao, pillClasseStatus } from "../model/domain.js";
import { calcularRankingResponsaveis, calcularPrazosFatais } from "../model/aggregations.js";

// 🎨 Monta o HTML de uma lista agrupada por responsável (seções com avatar + tabela),
// usado tanto no painel "Prazos e Riscos" quanto nos modais dos tiles do topo
export function construirListaAgrupadaPorResponsavel(grupos, cabecalhos, renderizarLinha, rotuloItem = "item(ns)") {
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

// 🎨 Linha padrão de uma tabela de prazos fatais (Processo ID | Área | Prazo Fatal | Situação)
export function linhaPrazoFatal(item) {
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

// 🎨 Linha padrão de uma tarefa nos modais "Vencem hoje"/"Vencem amanhã"
export function linhaModalTarefa(item) {
    return `
        <tr>
            <td>${item["Processo - ID"] || "-"}</td>
            <td>${item["Área do Direito"] || "-"}</td>
            <td>${item["Tipo"] || "-"}</td>
            <td><span class="pill ${pillClasseStatus(item["Status da tarefa"])}">${item["Status da tarefa"] || "-"}</span></td>
        </tr>
    `;
}

// 🎨 Renderiza a tabela de ranking de tarefas por responsável
export function renderizarRankingResponsaveis(dados) {
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
export function renderizarPainelPrazosFataisNovo(dados) {
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

// 🎨 Renderiza a tabela de prazos/tarefas do responsável selecionado
export function renderizarAtividadesPorResponsavel(dados, nome) {
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

// 🎨 Renderiza a tabela de audiências agendadas (padrão visual dos novos painéis)
export function renderizarPainelAudiencias(dados) {
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
