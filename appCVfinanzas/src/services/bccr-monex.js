const BCCR_API_URL =
    'https://apim.bccr.fi.cr/SDDE/api/' +
    'Bccr.GE.SDDE.Publico.Indicadores.API/cuadro/219/series';

const BCCR_INDICATORS = {
    3436: 'minimo',
    3437: 'maximo',
    3439: 'promedio_ponderado',
    3446: 'monto_total'
};

function parseBccrMonexPayload(payload, { fecha, sesion, capturadoEn }) {
    if (!payload || payload.estado !== true) {
        throw new Error(`La API del BCCR rechazo la consulta: ${payload?.mensaje || 'respuesta invalida'}`);
    }

    const values = {};

    for (const block of payload.datos || []) {
        for (const indicator of block.indicadores || []) {
            const field = BCCR_INDICATORS[String(indicator.codigoIndicador)];
            if (!field) continue;

            const point = (indicator.series || []).find(
                (item) => String(item.fecha || '').slice(0, 10) === fecha
            );

            if (point) {
                const value = point.valorDatoPorPeriodo;
                if (typeof value !== 'number' || !Number.isFinite(value)) {
                    throw new Error(`El indicador ${field} no contiene un valor numerico.`);
                }
                values[field] = value;
            }
        }
    }

    const missing = Object.values(BCCR_INDICATORS).filter((field) => values[field] === undefined);
    if (missing.length) {
        return null;
    }

    if (values.promedio_ponderado <= 0 || values.monto_total <= 0) {
        return null;
    }

    if (!(values.minimo > 0 &&
        values.minimo <= values.promedio_ponderado &&
        values.promedio_ponderado <= values.maximo)) {
        throw new Error('El minimo, promedio ponderado y maximo del BCCR son inconsistentes.');
    }

    return {
        fecha,
        sesion,
        promedio_ponderado: values.promedio_ponderado,
        monto_total: values.monto_total,
        minimo: values.minimo,
        maximo: values.maximo,
        capturado_en: capturadoEn
    };
}

async function fetchBccrMonex({ fecha, sesion, capturadoEn, fetchImpl = global.fetch }) {
    const token = process.env.BCCR_API_BEARER_TOKEN;
    if (!token) {
        throw new Error('Falta configurar BCCR_API_BEARER_TOKEN.');
    }
    if (typeof fetchImpl !== 'function') {
        throw new Error('El runtime de Node.js no incluye fetch.');
    }

    const url = new URL(BCCR_API_URL);
    const apiDate = fecha.replaceAll('-', '/');
    url.searchParams.set('fechaInicio', apiDate);
    url.searchParams.set('fechaFin', apiDate);
    url.searchParams.set('idioma', 'ES');

    const response = await fetchImpl(url, {
        headers: {
            Accept: 'application/json',
            Authorization: `Bearer ${token}`,
            'User-Agent': 'CVFinanzas/1.0'
        },
        signal: AbortSignal.timeout(60_000)
    });

    if (!response.ok) {
        const detail = (await response.text()).slice(0, 500);
        throw new Error(`BCCR respondio ${response.status}${detail ? `: ${detail}` : ''}`);
    }

    return parseBccrMonexPayload(await response.json(), { fecha, sesion, capturadoEn });
}

module.exports = {
    BCCR_API_URL,
    fetchBccrMonex,
    parseBccrMonexPayload
};
