// ============================================================
// Model / dateUtils.js — conversão e formatação de datas do Excel
// ============================================================

export function converterDataExcelParaPtBR(valorData) {
    if (!valorData || valorData === "-" || valorData === "") return valorData;

    // Se já estiver em formato de string pt-BR, retorna como está
    if (typeof valorData === 'string' && valorData.includes('/')) {
        return valorData;
    }

    // Se for número serial do Excel, converte
    if (typeof valorData === 'number' && valorData > 0) {
        try {
            // Excel conta dias desde 01/01/1900, mas com bug do ano 1900
            // JavaScript conta milissegundos desde 01/01/1970
            const diasDesde1900 = valorData - 25569; // Ajuste para JavaScript
            const data = new Date(diasDesde1900 * 86400 * 1000);

            // Verifica se a data é válida
            if (isNaN(data.getTime())) {
                console.warn(`Data inválida do Excel: ${valorData}`);
                return valorData;
            }

            // Formatar para DD/MM/AAAA HH:MM
            const dia = String(data.getDate()).padStart(2, '0');
            const mes = String(data.getMonth() + 1).padStart(2, '0');
            const ano = data.getFullYear();
            const horas = String(data.getHours()).padStart(2, '0');
            const minutos = String(data.getMinutes()).padStart(2, '0');

            // Se não tem horário específico (00:00), retorna só a data
            if (horas === '00' && minutos === '00') {
                return `${dia}/${mes}/${ano}`;
            } else {
                return `${dia}/${mes}/${ano} ${horas}:${minutos}`;
            }
        } catch (e) {
            console.warn(`Erro ao converter data do Excel: ${valorData}`, e);
            return valorData;
        }
    }

    // Tenta converter outros formatos de data
    if (valorData instanceof Date) {
        const dia = String(valorData.getDate()).padStart(2, '0');
        const mes = String(valorData.getMonth() + 1).padStart(2, '0');
        const ano = valorData.getFullYear();
        const horas = String(valorData.getHours()).padStart(2, '0');
        const minutos = String(valorData.getMinutes()).padStart(2, '0');

        if (horas === '00' && minutos === '00') {
            return `${dia}/${mes}/${ano}`;
        } else {
            return `${dia}/${mes}/${ano} ${horas}:${minutos}`;
        }
    }

    // Se for string, tenta converter
    if (typeof valorData === 'string') {
        try {
            const data = new Date(valorData);
            if (!isNaN(data.getTime())) {
                const dia = String(data.getDate()).padStart(2, '0');
                const mes = String(data.getMonth() + 1).padStart(2, '0');
                const ano = data.getFullYear();
                const horas = String(data.getHours()).padStart(2, '0');
                const minutos = String(data.getMinutes()).padStart(2, '0');

                if (horas === '00' && minutos === '00') {
                    return `${dia}/${mes}/${ano}`;
                } else {
                    return `${dia}/${mes}/${ano} ${horas}:${minutos}`;
                }
            }
        } catch (e) {
            // Se não conseguir converter, retorna o valor original
        }
    }

    return valorData;
}

// 🆕 Função para converter data do Excel para pt-BR com hora
export function converterDataExcelParaPtBRComHora(valorData) {
    if (!valorData || valorData === "-" || valorData === "") return valorData;

    // Se já estiver em formato de string pt-BR, retorna como está
    if (typeof valorData === 'string' && valorData.includes('/')) {
        return valorData;
    }

    // Se for número serial do Excel, converte
    if (typeof valorData === 'number' && valorData > 0) {
        try {
            // Excel conta dias desde 01/01/1900, mas com bug do ano 1900
            // Ajuste para JavaScript (que conta desde 01/01/1970)
            const diasDesde1900 = valorData - 25569;

            // 🔧 CORREÇÃO: Cria a data em UTC primeiro para evitar problemas de fuso horário
            const dataUTC = new Date(diasDesde1900 * 86400 * 1000);

            // 🔧 Ajusta o fuso horário para o Brasil (UTC-3)
            // Como o Excel não considera fuso horário, precisamos ajustar manualmente
            const offsetBrasil = +6 * 60; // -3 horas em minutos
            const offsetLocal = dataUTC.getTimezoneOffset(); // Offset local em minutos
            const diferencaOffset = offsetBrasil - offsetLocal;

            // Aplica o ajuste
            const dataAjustada = new Date(dataUTC.getTime() + (diferencaOffset * 60 * 1000));

            // Verifica se a data é válida
            if (isNaN(dataAjustada.getTime())) {
                console.warn(`Data inválida do Excel: ${valorData}`);
                return valorData;
            }

            // Formatar para DD/MM/AAAA HH:MM:SS
            const dia = String(dataAjustada.getDate()).padStart(2, '0');
            const mes = String(dataAjustada.getMonth() + 1).padStart(2, '0');
            const ano = dataAjustada.getFullYear();
            const horas = String(dataAjustada.getHours()).padStart(2, '0');
            const minutos = String(dataAjustada.getMinutes()).padStart(2, '0');
            const segundos = String(dataAjustada.getSeconds()).padStart(2, '0');

            // Sempre inclui hora, minuto e segundo para B3
            return `${dia}/${mes}/${ano} ${horas}:${minutos}:${segundos}`;

        } catch (e) {
            console.warn(`Erro ao converter data do Excel: ${valorData}`, e);
            return valorData;
        }
    }

    // Tenta converter outros formatos de data
    if (valorData instanceof Date) {
        const dia = String(valorData.getDate()).padStart(2, '0');
        const mes = String(valorData.getMonth() + 1).padStart(2, '0');
        const ano = valorData.getFullYear();
        const horas = String(valorData.getHours()).padStart(2, '0');
        const minutos = String(valorData.getMinutes()).padStart(2, '0');
        const segundos = String(valorData.getSeconds()).padStart(2, '0');

        return `${dia}/${mes}/${ano} ${horas}:${minutos}:${segundos}`;
    }

    // Se for string, tenta converter
    if (typeof valorData === 'string') {
        try {
            const data = new Date(valorData);
            if (!isNaN(data.getTime())) {
                const dia = String(data.getDate()).padStart(2, '0');
                const mes = String(data.getMonth() + 1).padStart(2, '0');
                const ano = data.getFullYear();
                const horas = String(data.getHours()).padStart(2, '0');
                const minutos = String(data.getMinutes()).padStart(2, '0');
                const segundos = String(data.getSeconds()).padStart(2, '0');

                return `${dia}/${mes}/${ano} ${horas}:${minutos}:${segundos}`;
            }
        } catch (e) {
            // Se não conseguir converter, retorna o valor original
        }
    }

    return valorData;
}

// Converte "Data do agendamento" (pt-BR, DD/MM/YYYY) em Date, sem hora
export function parseDataAgendamento(dataStr) {
    if (!dataStr || dataStr === "-" || dataStr === "") return null;

    if (typeof dataStr === 'string' && dataStr.includes('/')) {
        const parteData = dataStr.split(' ')[0];
        const partes = parteData.split('/');

        if (partes.length === 3) {
            const dia = parseInt(partes[0]);
            const mes = parseInt(partes[1]) - 1;
            const ano = parseInt(partes[2]);

            if (dia >= 1 && dia <= 31 && mes >= 0 && mes <= 11 && ano >= 2000) {
                return new Date(ano, mes, dia);
            }
        }
    }

    return null;
}

export function mesmoDia(a, b) {
    return a.getFullYear() === b.getFullYear() &&
        a.getMonth() === b.getMonth() &&
        a.getDate() === b.getDate();
}

// 🆕 Função simplificada para formatar data (dados já estão em pt-BR)
export function formatarDataExcel(valorData) {
    if (!valorData || valorData === "-" || valorData === "") return "-";

    // Se já estiver em formato pt-BR, retorna como está
    if (typeof valorData === 'string' && valorData.includes('/')) {
        return valorData;
    }

    // Fallback: se por algum motivo ainda vier como número do Excel, converte
    if (typeof valorData === 'number' && valorData > 0) {
        return converterDataExcelParaPtBR(valorData);
    }

    // Fallback: se vier como Date object, converte
    if (valorData instanceof Date && !isNaN(valorData.getTime())) {
        const dia = String(valorData.getDate()).padStart(2, '0');
        const mes = String(valorData.getMonth() + 1).padStart(2, '0');
        const ano = valorData.getFullYear();
        const horas = String(valorData.getHours()).padStart(2, '0');
        const minutos = String(valorData.getMinutes()).padStart(2, '0');

        if (horas === '00' && minutos === '00') {
            return `${dia}/${mes}/${ano}`;
        } else {
            return `${dia}/${mes}/${ano} ${horas}:${minutos}`;
        }
    }

    // Se não conseguir converter, retorna o valor original
    return valorData;
}

// Encurta "DD/MM/AAAA HH:MM:SS" para "DD/MM/AA HH:MM" (indicador do topbar,
// onde espaço é curto); o valor completo continua no atributo title.
export function formatarDataCurta(dataFormatada) {
    if (!dataFormatada) return dataFormatada;

    const match = dataFormatada.match(/^(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{2}):(\d{2}))?/);
    if (!match) return dataFormatada;

    const [, dia, mes, ano, hora, minuto] = match;
    const anoCurto = ano.slice(-2);

    return hora ? `${dia}/${mes}/${anoCurto} ${hora}:${minuto}` : `${dia}/${mes}/${anoCurto}`;
}
