// ============================================================
// Model / state.js — estado global dos dados carregados do Excel
// ============================================================

// 📦 URL da planilha "Pauta Diária"
export const excelUrl = "https://fundacaooswaldoaranha-my.sharepoint.com/personal/wesley_balbino_foa_org_br/_layouts/15/download.aspx?share=EdsT2JkTPstFhYTAoyB0kWwB0T83o-R9AR4Wu2Yex8hxBw";

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
