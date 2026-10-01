// ============================================================
// api/planilha.js — função serverless (Vercel) que entrega a planilha
// ============================================================
// Lê o link da variável de ambiente EXCEL_URL (Vercel → Settings →
// Environment Variables), baixa o .xlsx no servidor e repassa ao dashboard.
// Assim o link não aparece no navegador e não há bloqueio de CORS.
// Em caso de erro, responde JSON { codigo, detalhe } para o front-end
// exibir a mensagem adequada (ver js/model/excelService.js).

const MAX_REDIRECIONAMENTOS = 10;

// Aceita tanto o link de download (…/_layouts/15/download.aspx?share=…)
// quanto o link copiado do botão Compartilhar (…/:x:/g/…?e=…).
function montarUrlDownload(url) {
    const u = new URL(url);
    if (/\/:[a-z]:\//i.test(u.pathname)) {
        u.searchParams.set("download", "1");
    }
    return u.toString();
}

// Segue os redirecionamentos manualmente, guardando os cookies.
// Links "Qualquer pessoa com o link" do SharePoint definem um cookie de
// acesso (FedAuth) no meio dos redirecionamentos; sem reenviá-lo, o
// SharePoint responde 401.
async function baixarComCookies(url) {
    const cookies = new Map();
    let atual = url;

    for (let i = 0; i <= MAX_REDIRECIONAMENTOS; i++) {
        const resposta = await fetch(atual, {
            redirect: "manual",
            headers: {
                "User-Agent": "Mozilla/5.0 (compatible; AjDashBoard/1.0)",
                ...(cookies.size && {
                    Cookie: [...cookies].map(([nome, valor]) => `${nome}=${valor}`).join("; ")
                })
            }
        });

        for (const linha of resposta.headers.getSetCookie?.() || []) {
            const [par] = linha.split(";");
            const separador = par.indexOf("=");
            if (separador > 0) {
                cookies.set(par.slice(0, separador).trim(), par.slice(separador + 1).trim());
            }
        }

        const destino = resposta.headers.get("location");
        if (resposta.status >= 300 && resposta.status < 400 && destino) {
            atual = new URL(destino, atual).toString();
            continue;
        }
        return resposta;
    }
    throw new Error("Redirecionamentos demais ao acessar o SharePoint");
}

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
        resposta = await baixarComCookies(montarUrlDownload(excelUrl.trim()));
    } catch (erro) {
        return responderErro(502, "FALHA_REDE", erro.message);
    }

    if (resposta.status === 401 || resposta.status === 403) {
        return responderErro(502, "SEM_PERMISSAO", `SharePoint respondeu HTTP ${resposta.status} ${resposta.statusText}`);
    }
    if (resposta.status === 404) {
        return responderErro(502, "NAO_ENCONTRADA", "SharePoint respondeu HTTP 404");
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
