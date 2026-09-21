// ============================================================
// View / tiles.js — tiles de resumo do topo (aba Visão Geral)
// ============================================================

import { calcularResumoFila } from "../model/aggregations.js";

// 🎨 Renderiza os tiles de resumo do topo
export function renderizarTilesResumo(dados) {
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
