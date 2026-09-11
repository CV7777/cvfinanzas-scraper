const { writeApiLog } = require('../lib/app-logger');

const LEGACY_API_PATHS = new Set([
    '/auth/login',
    '/auth/logout',
    '/auth/me',
    '/dashboard-stats',
    '/search-results',
    '/test-postgres',
    '/test-supabase'
]);

function isApiRequest(req) {
    return req.path.startsWith('/api/') || LEGACY_API_PATHS.has(req.path);
}

function requestLogger(req, res, next) {
    if (!isApiRequest(req)) return next();

    const startedAt = Date.now();
    let logged = false;

    function logRequest(outcome) {
        if (logged) return;
        logged = true;

        void writeApiLog({
            level: res.statusCode >= 500 ? 'error' : res.statusCode >= 400 ? 'warn' : 'info',
            method: req.method,
            path: req.path,
            queryKeys: Object.keys(req.query || {}),
            statusCode: res.statusCode,
            durationMs: Date.now() - startedAt,
            outcome,
            user: req.user?.usuario || null,
            ip: req.ip || null,
            userAgent: req.get('user-agent') || null
        }).catch((error) => {
            console.error('[LOGS] No se pudo registrar la llamada API:', error.message);
        });
    }

    res.once('finish', () => logRequest('completed'));
    res.once('close', () => logRequest(res.writableEnded ? 'completed' : 'aborted'));
    return next();
}

module.exports = { isApiRequest, requestLogger };
