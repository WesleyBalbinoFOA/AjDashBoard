// ============================================================
// Model / state.js — estado global dos dados carregados do Excel
// ============================================================

// 📦 URL da planilha "Pauta Diária" — definida em env.js (fora do Git)
export const excelUrl = (window.ENV && window.ENV.EXCEL_URL) || "";

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
