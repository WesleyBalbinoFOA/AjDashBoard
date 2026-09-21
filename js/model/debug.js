// ============================================================
// Model / debug.js — utilitários de depuração via console do navegador
// ============================================================

// 🔍 Função de debug para verificar como as datas estão sendo processadas
// Para usar: debugDatas(dadosExcel) no console do navegador
export function debugDatas(dados) {
    // console.log("🔍 === DEBUG DE DATAS ===");

    const hoje = new Date();
    hoje.setHours(23, 59, 59, 999);
    // console.log(`📅 Data de referência (hoje): ${hoje.toLocaleDateString('pt-BR')}`);

    // Analisa os primeiros 20 registros
    const amostra = dados.slice(0, 20);

    // console.log("\n📊 Análise das primeiras 20 datas:");
    amostra.forEach((item, i) => {
        const dataStr = item["Data do agendamento"];
        let status = "❌ Inválida";
        let dataParsed = null;
        let incluiNoFiltro = false;

        if (dataStr && typeof dataStr === 'string' && dataStr.includes('/')) {
            const parteData = dataStr.split(' ')[0];
            const partes = parteData.split('/');

            if (partes.length === 3) {
                const dia = parseInt(partes[0]);
                const mes = parseInt(partes[1]) - 1;
                const ano = parseInt(partes[2]);

                if (dia >= 1 && dia <= 31 && mes >= 0 && mes <= 11 && ano >= 2000) {
                    dataParsed = new Date(ano, mes, dia);
                    incluiNoFiltro = dataParsed <= hoje;

                    if (incluiNoFiltro) {
                        status = "✅ Incluída";
                    } else {
                        status = "🔮 Futura";
                    }
                }
            }
        }

        // console.log(`${i + 1}. "${dataStr}" -> ${status} ${dataParsed ? `(${dataParsed.toLocaleDateString('pt-BR')})` : ''}`);
    });

    // Conta quantas são futuras vs passadas
    const futuras = dados.filter(item => {
        const dataStr = item["Data do agendamento"];
        if (!dataStr || typeof dataStr !== 'string' || !dataStr.includes('/')) return false;

        const parteData = dataStr.split(' ')[0];
        const partes = parteData.split('/');

        if (partes.length === 3) {
            const dia = parseInt(partes[0]);
            const mes = parseInt(partes[1]) - 1;
            const ano = parseInt(partes[2]);

            if (dia >= 1 && dia <= 31 && mes >= 0 && mes <= 11 && ano >= 2000) {
                const data = new Date(ano, mes, dia);
                return data > hoje;
            }
        }
        return false;
    });

    // console.log(`\n📈 Resumo:`);
    // console.log(`   Total de registros: ${dados.length}`);
    // console.log(`   Registros com datas futuras: ${futuras.length}`);
    // console.log(`   Registros que devem passar no filtro: ${dados.length - futuras.length}`);

    // console.log("🔍 === FIM DEBUG ===\n");
}

// 🔍 Função de debug para analisar prazos fatais
// Para usar: debugPrazosFatais(dadosExcel) no console do navegador
export function debugPrazosFatais(dados) {
    // console.log("🚨 === DEBUG PRAZOS FATAIS ===");

    const hoje = new Date();
    hoje.setHours(23, 59, 59, 999);
    const seteDiasDepois = new Date(hoje);
    seteDiasDepois.setDate(hoje.getDate() + 7);

    // console.log(`📅 Período de análise: ${hoje.toLocaleDateString('pt-BR')} até ${seteDiasDepois.toLocaleDateString('pt-BR')}`);

    // Campos que podem indicar prazo fatal
    const camposPrazoFatal = [
        "Solicitação - Há Prazo Fatal",
        "Há Prazo Fatal",
        "Prazo Fatal",
        "Prazo Crítico",
        "Urgente"
    ];

    // Campos que podem conter datas de vencimento
    const camposData = [
        "Data do agendamento",
        "Data de Vencimento",
        "Data Limite",
        "Prazo"
    ];

    // console.log("\n🔍 Analisando campos de prazo fatal disponíveis:");
    camposPrazoFatal.forEach(campo => {
        const valores = dados
            .map(item => item[campo])
            .filter(valor => valor && valor !== "" && valor !== "-")
            .slice(0, 10); // Primeiros 10 valores únicos

        if (valores.length > 0) {
            // console.log(`   ${campo}: ${[...new Set(valores)].join(', ')}`);
        }
    });

    // console.log("\n🔍 Analisando campos de data disponíveis:");
    camposData.forEach(campo => {
        const count = dados.filter(item => item[campo] && item[campo] !== "" && item[campo] !== "-").length;
        if (count > 0) {
            // console.log(`   ${campo}: ${count} registros com data`);
        }
    });

    // Analisa registros com "Sim" nos campos de prazo fatal
    const comPrazoFatal = dados.filter(item => {
        return camposPrazoFatal.some(campo => {
            const valor = item[campo];
            if (!valor) return false;
            const valorLower = valor.toString().toLowerCase().trim();
            return valorLower === 'sim' ||
                valorLower === 's' ||
                valorLower === 'yes' ||
                valorLower === 'y' ||
                valorLower === 'true' ||
                valorLower === '1';
        });
    });

    // console.log(`\n📊 Registros com prazo fatal = "Sim": ${comPrazoFatal.length}`);

    // Mostra alguns exemplos
    // console.log("\n🔍 Primeiros 5 exemplos de prazos fatais:");
    comPrazoFatal.slice(0, 5).forEach((item, i) => {
        const processoId = item["Processo - ID"] || "N/A";

        // Encontra qual campo tem o prazo fatal
        const campoComPrazo = camposPrazoFatal.find(campo => {
            const valor = item[campo];
            if (!valor) return false;
            const valorLower = valor.toString().toLowerCase().trim();
            return valorLower === 'sim' || valorLower === 's' || valorLower === 'yes' || valorLower === 'y' || valorLower === 'true' || valorLower === '1';
        });

        // Encontra a data
        const campoComData = camposData.find(campo => item[campo] && item[campo] !== "" && item[campo] !== "-");
        const data = campoComData ? item[campoComData] : "Sem data";

        // console.log(`   ${i + 1}. Processo ${processoId}: ${campoComPrazo} = "${item[campoComPrazo]}" | Data: ${data}`);
    });

    // console.log("🚨 === FIM DEBUG ===\n");

    return {
        totalComPrazoFatal: comPrazoFatal.length,
        camposEncontrados: {
            prazoFatal: camposPrazoFatal.filter(campo =>
                dados.some(item => item[campo] && item[campo] !== "" && item[campo] !== "-")
            ),
            datas: camposData.filter(campo =>
                dados.some(item => item[campo] && item[campo] !== "" && item[campo] !== "-")
            )
        }
    };
}
