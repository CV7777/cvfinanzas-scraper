const test = require('node:test');
const assert = require('node:assert/strict');

const {
    createMonexBackgroundService,
    getDueSchedule
} = require('../src/background/monex-background-service');

test('selecciona las sesiones usando la hora de Costa Rica', () => {
    assert.equal(getDueSchedule(new Date('2026-09-10T19:19:00Z')), null);
    assert.deepEqual(getDueSchedule(new Date('2026-09-10T19:20:00Z')), {
        fecha: '2026-09-10',
        sesion: '13:05'
    });
    assert.deepEqual(getDueSchedule(new Date('2026-09-10T23:20:00Z')), {
        fecha: '2026-09-10',
        sesion: '17:00'
    });
});

test('omite fines de semana y feriados', () => {
    assert.equal(getDueSchedule(new Date('2026-09-12T19:20:00Z')), null);
    assert.equal(getDueSchedule(new Date('2026-09-15T19:20:00Z')), null);
});

test('guarda una sesión pendiente una sola vez', async () => {
    let fetchCalls = 0;
    let saveCalls = 0;
    const fixedDate = new Date('2026-09-10T19:20:00Z');
    const logger = { log() {}, error() {} };
    const service = createMonexBackgroundService({
        now: () => fixedDate,
        logger,
        executionLogger: async () => {},
        sessionExists: async () => false,
        fetchRate: async (request) => {
            fetchCalls += 1;
            return { ...request, promedio_ponderado: 500, monto_total: 1000000 };
        },
        saveRate: async () => {
            saveCalls += 1;
        }
    });

    await service.check();
    await service.check();

    assert.equal(fetchCalls, 1);
    assert.equal(saveCalls, 1);
    assert.equal(service.getStatus().state, 'success');
    assert.equal(service.getStatus().history[0].message, 'Datos guardados correctamente.');
});

test('expone el fallo y la fecha del siguiente reintento', async () => {
    const fixedDate = new Date('2026-09-10T19:20:00Z');
    const service = createMonexBackgroundService({
        now: () => fixedDate,
        logger: { log() {}, error() {} },
        executionLogger: async () => {},
        sessionExists: async () => false,
        fetchRate: async () => {
            throw new Error('BCCR no disponible');
        }
    });

    await service.check();

    const status = service.getStatus();
    assert.equal(status.state, 'error');
    assert.equal(status.lastRunAt, '2026-09-10T19:20:00.000Z');
    assert.equal(status.lastError, 'BCCR no disponible');
    assert.equal(status.nextRetryAt, '2026-09-10T19:25:00.000Z');
    assert.deepEqual(status.history[0], {
        state: 'error',
        fecha: '2026-09-10',
        sesion: '13:05',
        timestamp: '2026-09-10T19:20:00.000Z',
        message: 'BCCR no disponible'
    });
});
