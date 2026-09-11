const test = require('node:test');
const assert = require('node:assert/strict');

const { parseBccrMonexPayload } = require('../src/services/bccr-monex');

function payloadFor(date = '2026-09-10') {
    return {
        estado: true,
        datos: [{
            indicadores: [
                { codigoIndicador: '3436', series: [{ fecha: date, valorDatoPorPeriodo: 499.5 }] },
                { codigoIndicador: '3437', series: [{ fecha: date, valorDatoPorPeriodo: 502 }] },
                { codigoIndicador: '3439', series: [{ fecha: date, valorDatoPorPeriodo: 500.25 }] },
                { codigoIndicador: '3446', series: [{ fecha: date, valorDatoPorPeriodo: 1250000 }] }
            ]
        }]
    };
}

test('convierte la respuesta BCCR al registro PostgreSQL', () => {
    const result = parseBccrMonexPayload(payloadFor(), {
        fecha: '2026-09-10',
        sesion: '13:05',
        capturadoEn: '2026-09-10T19:20:00.000Z'
    });

    assert.deepEqual(result, {
        fecha: '2026-09-10',
        sesion: '13:05',
        promedio_ponderado: 500.25,
        monto_total: 1250000,
        minimo: 499.5,
        maximo: 502,
        capturado_en: '2026-09-10T19:20:00.000Z'
    });
});

test('devuelve null mientras falten indicadores del día', () => {
    const payload = payloadFor();
    payload.datos[0].indicadores.pop();

    assert.equal(parseBccrMonexPayload(payload, {
        fecha: '2026-09-10',
        sesion: '13:05',
        capturadoEn: '2026-09-10T19:20:00.000Z'
    }), null);
});
