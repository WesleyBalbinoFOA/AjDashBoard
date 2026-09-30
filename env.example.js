// ============================================================
// env.example.js — modelo de configuração do ambiente
// ============================================================
// Copie este arquivo para "env.js" (na mesma pasta) e preencha os valores.
// O env.js está no .gitignore e NÃO deve ser commitado.
//
// Atenção: por ser um site estático, o navegador precisa ler esse valor,
// então quem abrir o dashboard ainda consegue ver o link no DevTools.
// O env.js apenas tira o link do repositório no GitHub.

window.ENV = {
    // Link de DOWNLOAD da planilha "Pauta Diária".
    // Formato recomendado (link "Qualquer pessoa com o link"):
    // https://<tenant>-my.sharepoint.com/personal/<usuario>/_layouts/15/download.aspx?share=<codigo>
    EXCEL_URL: ""
};
