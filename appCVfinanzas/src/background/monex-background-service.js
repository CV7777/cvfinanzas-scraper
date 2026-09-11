const { fetchBccrMonex } = require('../services/bccr-monex');
const { monexSessionExists, upsertMonexExchangeRate } = require('../services/monex-writer');
const { readBackgroundLogs, writeBackgroundLog } = require('../lib/app-logger');

const COSTA_RICA_TIME_ZONE = 'America/Costa_Rica';
const CHECK_INTERVAL_MS = 60_000;
const RETRY_INTERVAL_MS = 5 * 60_000;
const HOLIDAYS = new Set([
    '2026-01-01',
    '2026-04-02',
    '2026-04-03',
    '2026-04-11',
    '2026-05-01',
    '2026-07-25',
    '2026-08-02',
    '2026-08-15',
    '2026-08-31',
    '2026-09-15',
    '2026-12-01',
    '2026-12-25'
]);
const SCHEDULES = [
    { hour: 13, minute: 20, session: '13:05' },
    { hour: 17, minute: 20, session: '17:00' }
];
const MAX_HISTORY_ITEMS = 12;

function createInitialStatus(overrides = {}) {
    return {
        enabled: true,
        state: 'idle',
        startedAt: null,
        lastCheckAt: null,
        lastRunAt: null,
        lastSuccessAt: null,
        lastErrorAt: null,
        lastError: null,
        nextRetryAt: null,
        currentSession: null,
        history: [],
        ...overrides
    };
}

let inactiveStatus = createInitialStatus({
    enabled: false,
    state: 'not_started'
});
let activeService = null;

function getCostaRicaParts(date = new Date()) {
    const formatter = new Intl.DateTimeFormat('en-CA', {
        timeZone: COSTA_RICA_TIME_ZONE,
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        hourCycle: 'h23',
        weekday: 'short'
    });
    const parts = Object.fromEntries(
        formatter.formatToParts(date)
            .filter((part) => part.type !== 'literal')
            .map((part) => [part.type, part.value])
    );

    return {
        fecha: `${parts.year}-${parts.month}-${parts.day}`,
        hour: Number(parts.hour),
        minute: Number(parts.minute),
        weekday: parts.weekday
    };
}

function getDueSchedule(date = new Date()) {
    const now = getCostaRicaParts(date);
    if (now.weekday === 'Sat' || now.weekday === 'Sun' || HOLIDAYS.has(now.fecha)) {
        return null;
    }

    const currentMinute = now.hour * 60 + now.minute;
    const schedule = [...SCHEDULES].reverse().find(
        (item) => currentMinute >= item.hour * 60 + item.minute
    );

    return schedule ? { fecha: now.fecha, sesion: schedule.session } : null;
}

function createMonexBackgroundService({
    fetchRate = fetchBccrMonex,
    sessionExists = monexSessionExists,
    saveRate = upsertMonexExchangeRate,
    now = () => new Date(),
    logger = console,
    executionLogger = writeBackgroundLog
} = {}) {
    const completed = new Set();
    const nextRetryAt = new Map();
    const status = createInitialStatus();
    let running = false;
    let timer = null;

    async function persistEvent(entry) {
        try {
            await executionLogger(entry);
        } catch (error) {
            logger.error('[MONEX] No se pudo escribir el log:', error.message);
        }
    }

    async function addHistory({ state, due, message, at }) {
        const historyItem = {
            state,
            fecha: due.fecha,
            sesion: due.sesion,
            timestamp: at.toISOString(),
            message
        };

        status.history.unshift(historyItem);
        status.history = status.history.slice(0, MAX_HISTORY_ITEMS);
        await persistEvent({
            event: 'execution',
            ...historyItem,
            nextRetryAt: status.nextRetryAt
        });
    }

    function getStatus() {
        return {
            ...status,
            currentSession: status.currentSession ? { ...status.currentSession } : null,
            history: status.history.map((item) => ({ ...item }))
        };
    }

    async function check() {
        if (running) return;

        const currentDate = now();
        status.lastCheckAt = currentDate.toISOString();
        const due = getDueSchedule(currentDate);
        if (!due) return;

        const key = `${due.fecha}|${due.sesion}`;
        if (completed.has(key) || (nextRetryAt.get(key) || 0) > currentDate.getTime()) return;

        running = true;
        status.state = 'running';
        status.lastRunAt = currentDate.toISOString();
        status.currentSession = { ...due };
        try {
            if (await sessionExists(due)) {
                completed.add(key);
                nextRetryAt.delete(key);
                status.state = 'success';
                status.lastSuccessAt = currentDate.toISOString();
                status.lastError = null;
                status.nextRetryAt = null;
                await addHistory({
                    state: 'success',
                    due,
                    message: 'La sesion ya estaba guardada.',
                    at: currentDate
                });
                return;
            }

            const data = await fetchRate({
                ...due,
                capturadoEn: currentDate.toISOString()
            });

            if (!data) {
                const retryAt = currentDate.getTime() + RETRY_INTERVAL_MS;
                nextRetryAt.set(key, retryAt);
                status.state = 'waiting';
                status.lastError = null;
                status.nextRetryAt = new Date(retryAt).toISOString();
                await addHistory({
                    state: 'waiting',
                    due,
                    message: 'El BCCR aun no publica datos completos.',
                    at: currentDate
                });
                logger.log(`[MONEX] El BCCR aun no publica datos completos para ${key}.`);
                return;
            }

            await saveRate(data);
            completed.add(key);
            nextRetryAt.delete(key);
            status.state = 'success';
            status.lastSuccessAt = currentDate.toISOString();
            status.lastError = null;
            status.nextRetryAt = null;
            await addHistory({
                state: 'success',
                due,
                message: 'Datos guardados correctamente.',
                at: currentDate
            });
            logger.log(`[MONEX] Sesion ${key} guardada en PostgreSQL.`);
        } catch (error) {
            const retryAt = currentDate.getTime() + RETRY_INTERVAL_MS;
            nextRetryAt.set(key, retryAt);
            status.state = 'error';
            status.lastErrorAt = currentDate.toISOString();
            status.lastError = error.message;
            status.nextRetryAt = new Date(retryAt).toISOString();
            await addHistory({
                state: 'error',
                due,
                message: error.message,
                at: currentDate
            });
            logger.error(`[MONEX] Error procesando ${key}:`, error.message);
        } finally {
            running = false;
            status.currentSession = null;
        }
    }

    function start() {
        if (timer) return timer;
        status.enabled = true;
        status.state = 'idle';
        status.startedAt = now().toISOString();
        void persistEvent({
            event: 'lifecycle',
            state: 'started',
            timestamp: status.startedAt,
            message: 'Background service iniciado.'
        });
        logger.log('[MONEX] Background service activo (13:20 y 17:20 Costa Rica).');
        void check();
        timer = setInterval(() => void check(), CHECK_INTERVAL_MS);
        return timer;
    }

    function stop() {
        if (timer) clearInterval(timer);
        timer = null;
        status.state = 'stopped';
        void persistEvent({
            event: 'lifecycle',
            state: 'stopped',
            timestamp: now().toISOString(),
            message: 'Background service detenido.'
        });
    }

    return { check, getStatus, start, stop };
}

function startMonexBackgroundService() {
    if (String(process.env.MONEX_BACKGROUND_ENABLED).toLowerCase() !== 'true') {
        activeService = null;
        inactiveStatus = createInitialStatus({
            enabled: false,
            state: 'disabled'
        });
        void writeBackgroundLog({
            event: 'lifecycle',
            state: 'disabled',
            message: 'Background service desactivado.'
        }).catch((error) => console.error('[LOGS] No se pudo escribir el log:', error.message));
        console.log('[MONEX] Background service desactivado.');
        return null;
    }

    if (!process.env.BCCR_API_BEARER_TOKEN) {
        activeService = null;
        inactiveStatus = createInitialStatus({
            enabled: true,
            state: 'misconfigured',
            lastError: 'Falta configurar BCCR_API_BEARER_TOKEN.'
        });
        void writeBackgroundLog({
            event: 'lifecycle',
            state: 'misconfigured',
            level: 'error',
            message: inactiveStatus.lastError
        }).catch((error) => console.error('[LOGS] No se pudo escribir el log:', error.message));
        console.error('[MONEX] No se inicio: falta BCCR_API_BEARER_TOKEN.');
        return null;
    }

    const service = createMonexBackgroundService();
    activeService = service;
    service.start();
    return service;
}

async function getMonexBackgroundStatus() {
    const currentStatus = activeService ? activeService.getStatus() : {
        ...inactiveStatus,
        history: inactiveStatus.history.map((item) => ({ ...item }))
    };

    try {
        const entries = await readBackgroundLogs({ limit: 100 });
        const persistedHistory = entries
            .filter((entry) => entry.event === 'execution')
            .slice(0, MAX_HISTORY_ITEMS)
            .map((entry) => ({
                state: entry.state,
                fecha: entry.fecha,
                sesion: entry.sesion,
                timestamp: entry.timestamp,
                message: entry.message,
                nextRetryAt: entry.nextRetryAt || null
            }));

        if (persistedHistory.length === 0) return currentStatus;

        const latest = persistedHistory[0];
        const lastSuccess = persistedHistory.find((entry) => entry.state === 'success');
        const lastError = persistedHistory.find((entry) => entry.state === 'error');
        const canUsePersistedState = activeService && !currentStatus.lastRunAt;
        const retryIsPending = latest.nextRetryAt && new Date(latest.nextRetryAt).getTime() > Date.now();

        return {
            ...currentStatus,
            state: canUsePersistedState ? latest.state : currentStatus.state,
            lastRunAt: currentStatus.lastRunAt || latest.timestamp,
            lastSuccessAt: currentStatus.lastSuccessAt || lastSuccess?.timestamp || null,
            lastErrorAt: currentStatus.lastErrorAt || lastError?.timestamp || null,
            lastError: currentStatus.lastError || (latest.state === 'error' ? latest.message : null),
            nextRetryAt: currentStatus.nextRetryAt || (retryIsPending ? latest.nextRetryAt : null),
            history: persistedHistory
        };
    } catch (error) {
        console.error('[LOGS] No se pudo leer el historial del background service:', error.message);
        return currentStatus;
    }
}

module.exports = {
    COSTA_RICA_TIME_ZONE,
    SCHEDULES,
    createMonexBackgroundService,
    getCostaRicaParts,
    getDueSchedule,
    getMonexBackgroundStatus,
    startMonexBackgroundService
};
