const fs = require('fs');
const path = require('path');

let Pool;

try {
    ({ Pool } = require('pg'));
} catch (error) {
    Pool = null;
}

const postgresConfig = {
    host: process.env.AZURE_PG_HOST || process.env.PGHOST,
    user: process.env.AZURE_PG_USER || process.env.PGUSER,
    password: process.env.AZURE_PG_PASSWORD || process.env.PGPASSWORD,
    database: process.env.AZURE_PG_DATABASE || process.env.PGDATABASE || 'postgres',
    port: Number(process.env.AZURE_PG_PORT || process.env.PGPORT || 5432),
    sslCaPath: process.env.AZURE_PG_SSL_CA_PATH || process.env.PGSSLROOTCERT,
    sslCa: process.env.AZURE_PG_SSL_CA
};

const hasPostgresConfig = Boolean(
    postgresConfig.host &&
    postgresConfig.user &&
    postgresConfig.password &&
    postgresConfig.database
);

function getSslConfig() {
    if (postgresConfig.sslCa) {
        return { ca: postgresConfig.sslCa };
    }

    if (postgresConfig.sslCaPath) {
        const caPath = path.isAbsolute(postgresConfig.sslCaPath)
            ? postgresConfig.sslCaPath
            : path.join(process.cwd(), postgresConfig.sslCaPath);

        return { ca: fs.readFileSync(caPath, 'utf8') };
    }

    return { rejectUnauthorized: true };
}

let pool;

function getPool() {
    if (!Pool) {
        throw new Error('El paquete pg no esta instalado. Ejecuta npm install.');
    }

    if (!hasPostgresConfig) {
        throw new Error('Azure PostgreSQL no esta configurado.');
    }

    if (!pool) {
        pool = new Pool({
            host: postgresConfig.host,
            user: postgresConfig.user,
            password: postgresConfig.password,
            database: postgresConfig.database,
            port: postgresConfig.port,
            ssl: getSslConfig()
        });
    }

    return pool;
}

async function queryPostgres(text, values = []) {
    const result = await getPool().query(text, values);
    return result.rows;
}

function quoteIdentifier(identifier) {
    if (!/^[a-zA-Z_][a-zA-Z0-9_]*$/.test(identifier)) {
        throw new Error(`Identificador SQL invalido: ${identifier}`);
    }

    return `"${identifier}"`;
}

function quoteTableName(tableName) {
    return tableName
        .split('.')
        .map(quoteIdentifier)
        .join('.');
}

module.exports = {
    hasPostgresConfig,
    postgresConfig,
    queryPostgres,
    quoteTableName
};
