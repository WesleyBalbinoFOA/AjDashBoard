// ============================================================
// View / theme.js — alternância manual de tema claro/escuro
// ============================================================

const CHAVE_TEMA = "temaPreferido";

function preferenciaSalva() {
    const valor = localStorage.getItem(CHAVE_TEMA);
    return valor === "light" || valor === "dark" ? valor : null;
}

function temaDoSistema() {
    return window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
}

function temaEfetivo() {
    return preferenciaSalva() || temaDoSistema();
}

function atualizarIcone(botao, tema) {
    const iconeSol = botao.querySelector(".icon-sun");
    const iconeLua = botao.querySelector(".icon-moon");
    if (!iconeSol || !iconeLua) return;
    iconeSol.hidden = tema === "dark";
    iconeLua.hidden = tema !== "dark";
}

// Aplica o tema salvo (se houver) e conecta o botão de alternância.
// Sem preferência salva, não força o atributo `data-theme`, preservando o
// comportamento atual de detecção automática via prefers-color-scheme.
export function inicializarTema() {
    const salvo = preferenciaSalva();
    if (salvo) {
        document.documentElement.setAttribute("data-theme", salvo);
    }

    const botao = document.getElementById("themeToggle");
    if (!botao) return;

    atualizarIcone(botao, temaEfetivo());

    botao.addEventListener("click", () => {
        const novoTema = temaEfetivo() === "dark" ? "light" : "dark";
        document.documentElement.setAttribute("data-theme", novoTema);
        localStorage.setItem(CHAVE_TEMA, novoTema);
        atualizarIcone(botao, novoTema);
    });

    window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change", () => {
        if (!preferenciaSalva()) {
            atualizarIcone(botao, temaEfetivo());
        }
    });
}
