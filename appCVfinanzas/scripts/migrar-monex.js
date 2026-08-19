require('dotenv').config();

const fs = require('fs');
const path = require('path');
const { queryPostgres } = require('../src/lib/postgres');

async function main() {
    const migrationPath = path.join(
        __dirname,
        '..',
        'db',
        'migrations',
        '001_create_monex_tipo_cambio.sql'
    );
    const sql = fs.readFileSync(migrationPath, 'utf8');

    await queryPostgres(sql);
    console.log('Migracion MONEX aplicada correctamente.');
}

main().catch((error) => {
    console.error(error.message);
    process.exitCode = 1;
});
