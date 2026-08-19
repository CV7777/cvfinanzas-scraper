require('dotenv').config();

const bcrypt = require('bcryptjs');
const { randomUUID } = require('crypto');
const { queryPostgres, quoteTableName } = require('../src/lib/postgres');

const usersTableName = process.env.AZURE_PG_USERS_TABLE || process.env.PG_USERS_TABLE || 'usuarios';
const [usuario, password] = process.argv.slice(2);

async function main() {
    if (!usuario || !password) {
        console.error('Uso: npm run upsert-user -- usuario contrasena');
        process.exit(1);
    }

    const passwordHash = await bcrypt.hash(password, 12);
    const tableName = quoteTableName(usersTableName);
    const existing = await queryPostgres(
        `select id from ${tableName} where lower(usuario) = lower($1) limit 1`,
        [usuario]
    );

    if (existing[0]) {
        await queryPostgres(
            `update ${tableName} set password_hash = $1 where id = $2`,
            [passwordHash, existing[0].id]
        );
        console.log(`Usuario actualizado: ${usuario}`);
        return;
    }

    await queryPostgres(
        `insert into ${tableName} (id, usuario, password_hash, fecha_creacion) values ($1, $2, $3, now())`,
        [randomUUID(), usuario, passwordHash]
    );
    console.log(`Usuario creado: ${usuario}`);
}

main()
    .then(() => process.exit(0))
    .catch((error) => {
        console.error(error.message);
        process.exit(1);
    });
