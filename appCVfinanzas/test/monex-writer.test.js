const test = require('node:test');
const assert = require('node:assert/strict');

const { monexSessionExists, upsertMonexExchangeRate } = require('../src/services/monex-writer');

test('consulta existencia por fecha y sesión', async () => {
    let params;
    const exists = await monexSessionExists(
        { fecha: '2026-09-10', sesion: '13:05' },
        async (sql, values) => {
            assert.match(sql, /FROM "monex_tipo_cambio"/);
            params = values;
            return [{ existe: true }];
        }
    );

    assert.equal(exists, true);
    assert.deepEqual(params, ['2026-09-10', '13:05']);
});

test('hace UPSERT usando la llave fecha y sesión', async () => {
    let executedSql;
    let params;
    const data = {
        fecha: '2026-09-10',
        sesion: '17:00',
        promedio_ponderado: 501,
        monto_total: 2000000,
        minimo: 500,
        maximo: 502,
        capturado_en: '2026-09-10T23:20:00.000Z'
    };

    await upsertMonexExchangeRate(data, async (sql, values) => {
        executedSql = sql;
        params = values;
        return [{ fecha: data.fecha, sesion: data.sesion }];
    });

    assert.match(executedSql, /ON CONFLICT \(fecha, sesion\) DO UPDATE/);
    assert.deepEqual(params, [
        data.fecha,
        data.sesion,
        data.promedio_ponderado,
        data.monto_total,
        data.minimo,
        data.maximo,
        data.capturado_en
    ]);
});
