const test = require('node:test');
const assert = require('node:assert/strict');

const { buildDailyActivity } = require('../src/services/profiles');

test('agrupa la actividad diaria y completa dias sin resultados', () => {
    const activity = buildDailyActivity([
        { createdAt: '2026-09-09T08:00:00.000Z' },
        { createdAt: '2026-09-10T08:00:00.000Z' },
        { createdAt: '2026-09-10T09:00:00.000Z' }
    ], 3, new Date('2026-09-11T12:00:00.000Z'));

    assert.deepEqual(activity, [
        { date: '2026-09-09', total: 1 },
        { date: '2026-09-10', total: 2 },
        { date: '2026-09-11', total: 0 }
    ]);
});
