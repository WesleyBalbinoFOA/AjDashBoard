// ============================================================
// api/planilha.js — função serverless (Vercel) que entrega a planilha
// ============================================================
// Lê o link da variável de ambiente EXCEL_URL (Vercel → Settings →
// Environment Variables), baixa o .xlsx no servidor e repassa ao dashboard.
// Assim o link não aparece no navegador e não há bloqueio de CORS.
// Em caso de erro, responde JSON { codigo, detalhe } para o front-end
// exibir a mensagem adequada (ver js/model/excelService.js).

module.exports = async (req, res) => {
    const excelUrl = process.env.EXCEL_URL;

    const responderErro = (status, codigo, detalhe) => {
        res.setHeader("Cache-Control", "no-store");
        res.status(status).json({ codigo, detalhe });
    };

    if (!excelUrl) {
        return responderErro(500, "SEM_CONFIG", "Variável de ambiente EXCEL_URL não definida na Vercel");
    }

    let resposta;
    try {
        resposta = await fetch(excelUrl, { redirect: "follow" });
    } catch (erro) {
        return responderErro(502, "FALHA_REDE", erro.message);
    }

    if (!resposta.ok) {
        return responderErro(502, "HTTP", `SharePoint respondeu HTTP ${resposta.status} ${resposta.statusText}`);
    }

    const buffer = Buffer.from(await resposta.arrayBuffer());

    // Um .xlsx é um ZIP e começa com "PK". Se vier HTML, o link exige
    // login (ex.: link direto .../Documents/...) ou o compartilhamento expirou.
    if (buffer.length < 2 || buffer[0] !== 0x50 || buffer[1] !== 0x4B) {
        return responderErro(502, "NAO_E_XLSX", `Content-Type recebido: ${resposta.headers.get("content-type") || "desconhecido"}`);
    }

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    // Cache curto na CDN da Vercel para não baixar do SharePoint a cada acesso
    res.setHeader("Cache-Control", "s-maxage=60, stale-while-revalidate=300");
    res.status(200).send(buffer);
};
