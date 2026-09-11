const { queryPostgres, quoteTableName } = require('../lib/postgres');

const monexTableName = quoteTableName(
    process.env.AZURE_PG_MONEX_TABLE || process.env.PG_MONEX_TABLE || 'monex_tipo_cambio'
);

async function monexSessionExists({ fecha, sesion }, query = queryPostgres) {
    const rows = await query(
        `SELECT EXISTS (
            SELECT 1
            FROM ${monexTableName}
            WHERE fecha = $1::date AND sesion = $2::time
        ) AS existe`,
        [fecha, sesion]
    );

    return rows[0]?.existe === true;
}

async function upsertMonexExchangeRate(data, query = queryPostgres) {
    const rows = await query(
        `INSERT INTO ${monexTableName} AS destino (
            fecha, sesion, promedio_ponderado, monto_total,
            minimo, maximo, capturado_en
        ) VALUES ($1::date, $2::time, $3, $4, $5, $6, $7::timestamptz)
        ON CONFLICT (fecha, sesion) DO UPDATE SET
            promedio_ponderado = EXCLUDED.promedio_ponderado,
            monto_total = EXCLUDED.monto_total,
            minimo = EXCLUDED.minimo,
            maximo = EXCLUDED.maximo,
            capturado_en = EXCLUDED.capturado_en
        RETURNING
            to_char(fecha, 'YYYY-MM-DD') AS fecha,
            to_char(sesion, 'HH24:MI') AS sesion`,
        [
            data.fecha,
            data.sesion,
            data.promedio_ponderado,
            data.monto_total,
            data.minimo,
            data.maximo,
            data.capturado_en
        ]
    );

    return rows[0];
}

module.exports = {
    monexSessionExists,
    upsertMonexExchangeRate
};
