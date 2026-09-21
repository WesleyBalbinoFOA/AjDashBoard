// ============================================================
// Controller / app.js — orquestração e inicialização da aplicação
// ============================================================

import { dadosExcel } from "../model/state.js";
import { limparLocalStorage, carregarExcel, limparCacheERecarregar } from "../model/excelService.js";
import { debugDatas, debugPrazosFatais } from "../model/debug.js";
import { agruparPorResponsavel, calcularPrazosFatais, filtrarAtivasPorDia } from "../model/aggregations.js";
import { exibirUltimaAtualizacao } from "../view/topbar.js";
import { renderizarTilesResumo } from "../view/tiles.js";
import {
    renderizarRankingResponsaveis,
    renderizarPainelPrazosFataisNovo,
    renderizarPainelAudiencias,
    construirListaAgrupadaPorResponsavel,
    linhaModalTarefa,
    linhaPrazoFatal
} from "../view/tables.js";
import { gerarGraficoAreaDireito, gerarGraficoStatusTarefa } from "../view/charts.js";
import { ativarAba, popularTabsResponsaveis, inicializarFiltroAtividades } from "../view/tabs.js";
import { abrirModal, fecharModal } from "../view/modal.js";
import { inicializarTema } from "../view/theme.js";

limparLocalStorage();
document.querySelector("#logo").addEventListener("click", () => location.reload());

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

// Funções referenciadas via atributos inline (onclick="...") no HTML precisam
// ser expostas explicitamente em `window`, já que módulos ES não criam globais.
window.ativarAba = ativarAba;
window.abrirModalPorDia = abrirModalPorDia;
window.abrirModalPrazoFatal = abrirModalPrazoFatal;
window.fecharModal = fecharModal;
window.debugDatas = debugDatas;
window.debugPrazosFatais = debugPrazosFatais;
window.limparCacheERecarregar = limparCacheERecarregar;

inicializarTema();

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
        inicializarFiltroAtividades(dados);
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
