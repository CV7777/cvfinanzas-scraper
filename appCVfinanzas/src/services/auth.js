const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const { hasPostgresConfig, queryPostgres, quoteTableName } = require('../lib/postgres');

const authCookieName = 'cvf_access_token';
const usersTableName = process.env.AZURE_PG_USERS_TABLE || process.env.PG_USERS_TABLE || 'usuarios';
const tokenIssuer = 'cvfinanzas-api';
const tokenAudience = 'cvfinanzas-admin';
const defaultTokenTtl = process.env.AUTH_TOKEN_TTL || '8h';
const developmentJwtSecret = require('crypto').randomBytes(32).toString('hex');
const dummyBcryptHash = bcrypt.hashSync('cvfinanzas-invalid-password', 10);

function parseCookies(req) {
    return (req.headers.cookie || '')
        .split(';')
        .map((cookie) => cookie.trim())
        .filter(Boolean)
        .reduce((cookies, cookie) => {
            const separatorIndex = cookie.indexOf('=');
            if (separatorIndex === -1) {
                return cookies;
            }

            const key = cookie.slice(0, separatorIndex);
            const value = cookie.slice(separatorIndex + 1);
            cookies[key] = decodeURIComponent(value);
            return cookies;
        }, {});
}

function getJwtSecret() {
    const secret = process.env.JWT_SECRET || process.env.AUTH_TOKEN_SECRET;

    if (secret) {
        return secret;
    }

    if (process.env.NODE_ENV === 'production') {
        throw new Error('JWT_SECRET debe estar configurado en produccion.');
    }

    return developmentJwtSecret;
}

function parseDurationSeconds(value) {
    if (!value) {
        return 8 * 60 * 60;
    }

    if (/^\d+$/.test(value)) {
        return Number(value);
    }

    const match = String(value).trim().match(/^(\d+)\s*([smhd])$/i);

    if (!match) {
        return 8 * 60 * 60;
    }

    const amount = Number(match[1]);
    const unit = match[2].toLowerCase();
    const multipliers = {
        s: 1,
        m: 60,
        h: 60 * 60,
        d: 24 * 60 * 60
    };

    return amount * multipliers[unit];
}

function getJwtExpiresIn() {
    if (/^\d+$/.test(defaultTokenTtl)) {
        return Number(defaultTokenTtl);
    }

    return defaultTokenTtl;
}

function publicUser(row) {
    if (!row) {
        return null;
    }

    return {
        id: row.id,
        usuario: row.usuario,
        createdAt: row.fecha_creacion
    };
}

function setSessionCookies(res, session) {
    const isProduction = process.env.NODE_ENV === 'production';
    const maxAgeSeconds = session.expires_in || parseDurationSeconds(defaultTokenTtl);
    const cookieOptions = [
        'HttpOnly',
        'Path=/',
        'SameSite=Lax',
        `Max-Age=${maxAgeSeconds}`
    ];

    if (isProduction) {
        cookieOptions.push('Secure');
    }

    res.append('Set-Cookie', `${authCookieName}=${encodeURIComponent(session.access_token)}; ${cookieOptions.join('; ')}`);
}

function clearSessionCookies(res) {
    const expired = 'HttpOnly; Path=/; SameSite=Lax; Max-Age=0';
    res.append('Set-Cookie', `${authCookieName}=; ${expired}`);
}

function getAccessToken(req) {
    const authorization = req.headers.authorization || '';
    const bearerMatch = authorization.match(/^Bearer\s+(.+)$/i);

    if (bearerMatch) {
        return bearerMatch[1].trim();
    }

    return parseCookies(req)[authCookieName];
}

async function findUserByUsuario(usuario) {
    const sql = `
        select id, usuario, password_hash, fecha_creacion
        from ${quoteTableName(usersTableName)}
        where lower(usuario) = lower($1)
        limit 1
    `;
    const rows = await queryPostgres(sql, [usuario]);

    return rows[0] || null;
}

async function findUserById(id) {
    const sql = `
        select id, usuario, fecha_creacion
        from ${quoteTableName(usersTableName)}
        where id = $1
        limit 1
    `;
    const rows = await queryPostgres(sql, [id]);

    return rows[0] || null;
}

async function verifyPassword(password, passwordHash) {
    const hashToCompare = passwordHash || dummyBcryptHash;

    if (!hashToCompare.startsWith('$2')) {
        const error = new Error('Formato de password_hash no soportado. Usa bcrypt.');
        error.status = 500;
        throw error;
    }

    return bcrypt.compare(password, hashToCompare);
}

async function signInWithPassword(usuario, password) {
    if (!hasPostgresConfig) {
        const error = new Error('Azure PostgreSQL no esta configurado.');
        error.status = 500;
        throw error;
    }

    const user = await findUserByUsuario(usuario);
    const passwordMatches = await verifyPassword(password, user?.password_hash);

    if (!user || !passwordMatches) {
        const error = new Error('Credenciales invalidas');
        error.status = 401;
        throw error;
    }

    const expiresIn = getJwtExpiresIn();
    const accessToken = jwt.sign(
        {
            sub: user.id,
            usuario: user.usuario
        },
        getJwtSecret(),
        {
            expiresIn,
            issuer: tokenIssuer,
            audience: tokenAudience
        }
    );

    return {
        access_token: accessToken,
        expires_in: parseDurationSeconds(defaultTokenTtl),
        token_type: 'Bearer',
        user: publicUser(user)
    };
}

async function getUserFromToken(accessToken) {
    if (!accessToken) {
        return null;
    }

    let payload;

    try {
        payload = jwt.verify(accessToken, getJwtSecret(), {
            issuer: tokenIssuer,
            audience: tokenAudience
        });
    } catch (error) {
        return null;
    }

    if (!payload?.sub) {
        return null;
    }

    return publicUser(await findUserById(payload.sub));
}

module.exports = {
    authCookieName,
    clearSessionCookies,
    getAccessToken,
    getUserFromToken,
    setSessionCookies,
    signInWithPassword
};
