const { queryPostgres, quoteTableName } = require('../lib/postgres');

const monexTableName = process.env.AZURE_PG_MONEX_TABLE ||
    process.env.PG_MONEX_TABLE ||
    'monex_tipo_cambio';

async function listMonexExchangeRates(query = queryPostgres) {
    const sql = `
        SELECT
            to_char(fecha, 'YYYY-MM-DD') AS fecha,
            promedio_ponderado::double precision AS promedio_ponderado,
            monto_total::double precision AS monto_total,
            minimo::double precision AS minimo,
            maximo::double precision AS maximo,
            to_char(sesion, 'HH24:MI') AS sesion,
            to_char(
                capturado_en AT TIME ZONE 'America/Costa_Rica',
                'YYYY-MM-DD HH24:MI:SS'
            ) AS timestamp,
            to_char(
                ultima_actualizacion AT TIME ZONE 'America/Costa_Rica',
                'YYYY-MM-DD HH24:MI:SS'
            ) AS ultima_actualizacion
        FROM ${quoteTableName(monexTableName)}
        ORDER BY fecha ASC, sesion ASC
    `;

    const rows = await query(sql);
    const actualizado = rows.reduce((latest, row) => {
        const value = row.ultima_actualizacion || '';
        return value > latest ? value : latest;
    }, '');

    return {
        actualizado: actualizado || null,
        datos: rows.map(({ ultima_actualizacion, ...row }) => row)
    };
}

module.exports = { listMonexExchangeRates };
