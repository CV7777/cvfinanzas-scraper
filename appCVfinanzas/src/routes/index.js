const express = require('express');
const router = express.Router();
const { getDashboardStats, getFinancialProfile, getFinancialProfileByEmail, searchResults } = require('../services/profiles');
const { hasSupabaseConfig, selectFromSupabase } = require('../lib/supabase');
const { hasPostgresConfig, postgresConfig, queryPostgres } = require('../lib/postgres');
const { clearSessionCookies, getAccessToken, getUserFromToken, setSessionCookies, signInWithPassword } = require('../services/auth');
const { requireAuth } = require('../middleware/auth');
const {
    createExchangeRateComment,
    listExchangeRateComments
} = require('../services/tipo-cambio-comentarios');

const showResult = async (req, res, next) => {
    try {
        const email = req.query.email || req.query.user_email;
        const slug = req.params.slug || req.query.profile || 'el-ambicioso';
        const profile = email
            ? await getFinancialProfileByEmail(email)
            : getFinancialProfile(slug);

        res.render('result', { profile });
    } catch (error) {
        next(error);
    }
};

// Ruta de prueba de conexión a Supabase
router.get('/test-supabase', async (req, res) => {
    try {
        const status = {
            hasConfig: hasSupabaseConfig,
            timestamp: new Date().toISOString()
        };

        if (!hasSupabaseConfig) {
            status.message = 'Supabase no está configurado (faltan variables de entorno)';
            return res.json(status);
        }

        // Intentar consultar la tabla de resultados con un email de prueba
        const tableName = process.env.SUPABASE_RESULTS_TABLE || 'quiz_honda_results';
        const params = new URLSearchParams({
            select: 'count'
        });

        try {
            const data = await selectFromSupabase(tableName, params);
            status.message = 'Conexión exitosa a Supabase';
            status.table = tableName;
            status.connected = true;
        } catch (supabaseErr) {
            status.message = 'Tabla no existe o no accesible, pero Supabase está configurado';
            status.table = tableName;
            status.connected = false;
            status.supabaseError = supabaseErr.message;
        }

        return res.json(status);
    } catch (error) {
        console.error('test-supabase error:', error);
        return res.status(500).json({
            message: 'Error en la prueba de Supabase',
            error: error.message,
            timestamp: new Date().toISOString()
        });
    }
});

router.post('/auth/login', async (req, res) => {
    const usuario = (req.body.email || req.body.usuario || '').trim();
    const password = req.body.password || '';

    if (!usuario || !password) {
        return res.status(400).json({ error: 'usuario and password are required' });
    }

    try {
        const session = await signInWithPassword(usuario, password);
        setSessionCookies(res, session);
        return res.json({
            user: session.user,
            tokenType: session.token_type,
            expiresIn: session.expires_in,
            redirectTo: '/search'
        });
    } catch (error) {
        console.error('auth/login error:', error.message);
        return res.status(error.status || 401).json({
            error: 'invalid_credentials',
            message: 'Correo o contrasena incorrectos'
        });
    }
});

router.post('/auth/logout', (req, res) => {
    clearSessionCookies(res);
    return res.json({ ok: true, redirectTo: '/login' });
});

router.get('/auth/me', async (req, res, next) => {
    try {
        const user = await getUserFromToken(getAccessToken(req));

        if (!user) {
            return res.status(401).json({ error: 'unauthorized' });
        }

        return res.json({ user });
    } catch (error) {
        return next(error);
    }
});

router.get('/test-postgres', async (req, res) => {
    const status = {
        hasConfig: hasPostgresConfig,
        host: postgresConfig.host,
        database: postgresConfig.database,
        timestamp: new Date().toISOString()
    };

    if (!hasPostgresConfig) {
        return res.json({
            ...status,
            connected: false,
            message: 'Azure PostgreSQL no esta configurado'
        });
    }

    try {
        const rows = await queryPostgres('select now() as now');
        return res.json({
            ...status,
            connected: true,
            message: 'Conexion exitosa a Azure PostgreSQL',
            now: rows[0]?.now
        });
    } catch (error) {
        return res.status(500).json({
            ...status,
            connected: false,
            message: 'Error conectando a Azure PostgreSQL',
            error: error.message
        });
    }
});

// Ruta API para buscar todos los resultados o filtrar por email/tipo de perfil
router.get('/search-results', requireAuth, async (req, res) => {
    const email = (req.query.email || '').trim();
    const profileType = (
        req.query.profile_type ||
        req.query.profileType ||
        req.query.tipo_perfil ||
        req.query.tipo ||
        ''
    ).trim();

    try {
        const results = await searchResults({ email, profileType });
        return res.json(results);
    } catch (err) {
        console.error('search-results error', err);
        return res.status(500).json({ error: 'internal_error' });
    }
});

router.get('/dashboard-stats', requireAuth, async (req, res) => {
    try {
        return res.json(await getDashboardStats());
    } catch (err) {
        console.error('dashboard-stats error', err);
        return res.status(500).json({ error: 'internal_error' });
    }
});

router.get('/api/tipo-cambio/comentarios', requireAuth, async (req, res) => {
    try {
        const comentarios = await listExchangeRateComments();
        return res.json({ comentarios });
    } catch (error) {
        console.error('tipo-cambio comentarios GET error:', error.message);
        return res.status(500).json({
            error: 'internal_error',
            message: 'No se pudieron cargar los comentarios.'
        });
    }
});

router.post('/api/tipo-cambio/comentarios', requireAuth, async (req, res) => {
    const comentario = String(req.body.comentario || '').trim();

    if (!comentario) {
        return res.status(400).json({
            error: 'validation_error',
            message: 'El comentario es requerido.'
        });
    }

    if (comentario.length > 2000) {
        return res.status(400).json({
            error: 'validation_error',
            message: 'El comentario no puede superar los 2000 caracteres.'
        });
    }

    try {
        const nuevoComentario = await createExchangeRateComment({
            comentario,
            usuario: req.user.usuario
        });
        return res.status(201).json({ comentario: nuevoComentario });
    } catch (error) {
        console.error('tipo-cambio comentarios POST error:', error.message);
        return res.status(500).json({
            error: 'internal_error',
            message: 'No se pudo guardar el comentario.'
        });
    }
});

// Ruta principal y ruta explicita para mostrar los resultados
router.get('/', showResult);
router.get('/result', showResult);
router.get('/result/:slug', showResult);

module.exports = router;
