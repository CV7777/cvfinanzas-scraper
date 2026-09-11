const fs = require('fs');
const path = require('path');

const DEFAULT_LOG_DIRECTORY = path.join(__dirname, '../../logs');
const MAX_LOG_FILES_TO_READ = 14;
const writeQueues = new Map();

function getLogDirectory() {
    return path.resolve(process.env.APP_LOG_DIR || DEFAULT_LOG_DIRECTORY);
}

function getLogFile(channel, date = new Date()) {
    const day = date.toISOString().slice(0, 10);
    return path.join(getLogDirectory(), `${channel}-${day}.log`);
}

function serializeValue(value) {
    if (value instanceof Error) {
        return {
            name: value.name,
            message: value.message,
            stack: value.stack
        };
    }

    return value;
}

async function appendLog(channel, entry) {
    const timestamp = entry.timestamp || new Date().toISOString();
    const logFile = getLogFile(channel, new Date(timestamp));
    const line = `${JSON.stringify({ ...entry, timestamp }, (_key, value) => serializeValue(value))}\n`;
    const previousWrite = writeQueues.get(logFile) || Promise.resolve();
    const currentWrite = previousWrite
        .catch(() => {})
        .then(async () => {
            await fs.promises.mkdir(getLogDirectory(), { recursive: true });
            await fs.promises.appendFile(logFile, line, 'utf8');
        });

    writeQueues.set(logFile, currentWrite);
    await currentWrite;
}

async function readRecentLogs(channel, { limit = 50 } = {}) {
    const directory = getLogDirectory();

    await Promise.all(
        [...writeQueues.entries()]
            .filter(([file]) => path.dirname(file) === directory)
            .map(([, pendingWrite]) => pendingWrite.catch(() => {}))
    );

    let files;
    try {
        files = await fs.promises.readdir(directory);
    } catch (error) {
        if (error.code === 'ENOENT') return [];
        throw error;
    }

    const matchingFiles = files
        .filter((file) => file.startsWith(`${channel}-`) && file.endsWith('.log'))
        .sort()
        .reverse()
        .slice(0, MAX_LOG_FILES_TO_READ);
    const entries = [];

    for (const file of matchingFiles) {
        const content = await fs.promises.readFile(path.join(directory, file), 'utf8');
        const lines = content.trim().split('\n').filter(Boolean);

        for (const line of lines) {
            try {
                entries.push(JSON.parse(line));
            } catch (error) {
                console.error(`[LOGS] Se omitio una linea invalida en ${file}.`);
            }
        }
    }

    return entries
        .sort((a, b) => String(b.timestamp).localeCompare(String(a.timestamp)))
        .slice(0, limit);
}

function writeApiLog(entry) {
    return appendLog('api', { type: 'api_request', ...entry });
}

function writeBackgroundLog(entry) {
    return appendLog('background-service', { type: 'background_service', ...entry });
}

function writeApplicationLog(entry) {
    return appendLog('application', { type: 'application', ...entry });
}

function readBackgroundLogs(options) {
    return readRecentLogs('background-service', options);
}

module.exports = {
    appendLog,
    getLogDirectory,
    readBackgroundLogs,
    readRecentLogs,
    writeApiLog,
    writeApplicationLog,
    writeBackgroundLog
};
