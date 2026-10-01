// ============================================================
// Model / state.js — estado global dos dados carregados do Excel
// ============================================================

// 📦 URL da planilha "Pauta Diária"
// - Na Vercel: a função /api/planilha lê a variável de ambiente EXCEL_URL.
// - Localmente: se existir env.js (fora do Git) com window.ENV.EXCEL_URL, usa o link direto.
export const excelUrl = (window.ENV && window.ENV.EXCEL_URL) || "/api/planilha";
export const usandoApi = excelUrl === "/api/planilha";

// 🗂️ Dados carregados do Excel
export let dadosExcel = [];

// 📅 Data/hora da célula B3 (última atualização da pauta)
export let dataB3 = null;
export let dataB3Formatada = null;

// 📊 Registro das instâncias de gráficos Chart.js já criadas (por canvasId)
export const charts = {};

export function setDadosExcel(valor) {
    dadosExcel = valor;
}

export function setDataB3(valor) {
    dataB3 = valor;
}

export function setDataB3Formatada(valor) {
    dataB3Formatada = valor;
}
