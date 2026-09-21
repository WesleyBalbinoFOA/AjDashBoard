// ============================================================
// View / topbar.js — indicador "Atualizado em..." do topbar
// ============================================================

import { dataB3Formatada } from "../model/state.js";
import { formatarDataCurta } from "../model/dateUtils.js";

export function exibirUltimaAtualizacao() {
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
