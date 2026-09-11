const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');

const {
    getLogDirectory,
    readBackgroundLogs,
    writeApiLog,
    writeBackgroundLog
} = require('../src/lib/app-logger');
const { getMonexBackgroundStatus } = require('../src/background/monex-background-service');

test('escribe logs JSON diarios y recupera el historial del background', async (t) => {
    const previousLogDirectory = process.env.APP_LOG_DIR;
    const temporaryDirectory = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'cvfinanzas-logs-'));
    process.env.APP_LOG_DIR = temporaryDirectory;

    t.after(async () => {
        if (previousLogDirectory === undefined) {
            delete process.env.APP_LOG_DIR;
        } else {
            process.env.APP_LOG_DIR = previousLogDirectory;
        }
        await fs.promises.rm(temporaryDirectory, { recursive: true, force: true });
    });

    await writeApiLog({
        timestamp: '2026-09-11T10:00:00.000Z',
        method: 'GET',
        path: '/api/example',
        statusCode: 200
    });
    await writeBackgroundLog({
        timestamp: '2026-09-11T10:01:00.000Z',
        event: 'execution',
        state: 'success',
        fecha: '2026-09-11',
        sesion: '13:05'
    });

    const files = await fs.promises.readdir(getLogDirectory());
    assert.deepEqual(files.sort(), [
        'api-2026-09-11.log',
        'background-service-2026-09-11.log'
    ]);

    const entries = await readBackgroundLogs({ limit: 10 });
    assert.equal(entries.length, 1);
    assert.equal(entries[0].state, 'success');
    assert.equal(entries[0].type, 'background_service');

    const dashboardStatus = await getMonexBackgroundStatus();
    assert.equal(dashboardStatus.lastRunAt, '2026-09-11T10:01:00.000Z');
    assert.equal(dashboardStatus.history[0].state, 'success');
});
