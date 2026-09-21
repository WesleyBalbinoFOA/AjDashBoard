// ============================================================
// View / charts.js — gráficos do painel de produtividade (Chart.js)
// ============================================================

import { charts } from "../model/state.js";
import { statusIndicaConcluida } from "../model/domain.js";

// Lê um token de cor do :root (resolve o tema light/dark atual) para uso no Chart.js,
// que não entende `var(--x)` diretamente no canvas.
export function cssVar(nome) {
    return getComputedStyle(document.documentElement).getPropertyValue(nome).trim();
}

// 📊 Gráfico de barras horizontal: tarefas ativas por Área do Direito
export function gerarGraficoAreaDireito(dados, canvasId) {
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
export function gerarGraficoStatusTarefa(dados, canvasId) {
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
