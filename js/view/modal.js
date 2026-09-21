// ============================================================
// View / modal.js — modal genérico (listagens dos tiles clicáveis)
// ============================================================

export function abrirModal(titulo, corpoHtml) {
    const overlay = document.getElementById("modalOverlay");
    const tituloEl = document.getElementById("modalTitulo");
    const corpoEl = document.getElementById("modalCorpo");
    if (!overlay || !tituloEl || !corpoEl) return;

    tituloEl.textContent = titulo;
    corpoEl.innerHTML = corpoHtml;
    overlay.hidden = false;
}

export function fecharModal() {
    const overlay = document.getElementById("modalOverlay");
    if (overlay) overlay.hidden = true;
}

document.addEventListener("click", (e) => {
    const overlay = document.getElementById("modalOverlay");
    if (overlay && !overlay.hidden && e.target === overlay) {
        fecharModal();
    }
});

document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") fecharModal();
});
