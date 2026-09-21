// ============================================================
// View / tabs.js — abas do header e abas de responsável
// ============================================================

import { obterListaResponsaveis } from "../model/aggregations.js";
import { nomeResumido } from "../model/domain.js";
import { renderizarAtividadesPorResponsavel } from "./tables.js";

// Alterna as abas do header (Produtividade / Atividades / Prazos / Audiências)
export function ativarAba(nome) {
    document.querySelectorAll('.topbar-tab').forEach(botao => {
        botao.classList.toggle('active', botao.dataset.aba === nome);
    });

    document.querySelectorAll('.tab-panel').forEach(secao => {
        secao.hidden = secao.dataset.aba !== nome;
    });

    window.scrollTo({ top: 0, behavior: 'smooth' });
}

// Estado da aba "Atividades por Responsável": pessoa e filtro de situação
// selecionados atualmente (compartilhado entre as abas de pessoa e os botões
// de filtro pendentes/futuras/todos).
let nomeAtividadesAtual = null;
let filtroAtividadesAtual = "pendentes";

// 🎨 Monta as abas de responsável da aba "Atividades por Responsável" e renderiza a primeira
export function popularTabsResponsaveis(dados) {
    const container = document.getElementById("tabsResponsavelAtividades");
    if (!container) return;

    const nomes = obterListaResponsaveis(dados);
    container.innerHTML = "";

    nomes.forEach((nome, index) => {
        const aba = document.createElement("button");
        aba.type = "button";
        aba.className = "person-tab" + (index === 0 ? " active" : "");
        aba.textContent = nomeResumido(nome);
        aba.title = nome;
        aba.addEventListener("click", () => {
            container.querySelectorAll(".person-tab").forEach(el => el.classList.remove("active"));
            aba.classList.add("active");
            nomeAtividadesAtual = nome;
            renderizarAtividadesPorResponsavel(dados, nome, filtroAtividadesAtual);
        });
        container.appendChild(aba);
    });

    if (nomes.length) {
        nomeAtividadesAtual = nomes[0];
        renderizarAtividadesPorResponsavel(dados, nomes[0], filtroAtividadesAtual);
    }
}

// 🎨 Liga os botões de filtro (Pendentes/Futuros/Todos) da aba "Atividades por
// Responsável": clicar em um desmarca os demais e refiltra a tabela da pessoa
// atualmente selecionada.
export function inicializarFiltroAtividades(dados) {
    const botoes = document.querySelectorAll(".filtro-atividades-btn");
    if (!botoes.length) return;

    botoes.forEach(botao => {
        botao.classList.toggle("active", botao.dataset.filtro === filtroAtividadesAtual);
        botao.addEventListener("click", () => {
            filtroAtividadesAtual = botao.dataset.filtro;
            botoes.forEach(b => b.classList.toggle("active", b === botao));
            if (nomeAtividadesAtual) {
                renderizarAtividadesPorResponsavel(dados, nomeAtividadesAtual, filtroAtividadesAtual);
            }
        });
    });
}
