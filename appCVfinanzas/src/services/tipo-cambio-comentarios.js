const { queryPostgres, quoteTableName } = require('../lib/postgres');

const commentsTableName = process.env.AZURE_PG_EXCHANGE_COMMENTS_TABLE ||
    process.env.PG_EXCHANGE_COMMENTS_TABLE ||
    'comentarios_tipo_cambio';

async function listExchangeRateComments() {
    const sql = `
        select id, comentario, usuario, fecha
        from ${quoteTableName(commentsTableName)}
        order by fecha desc, id desc
    `;

    return queryPostgres(sql);
}

async function createExchangeRateComment({ comentario, usuario }) {
    const sql = `
        insert into ${quoteTableName(commentsTableName)} (comentario, usuario)
        values ($1, $2)
        returning id, comentario, usuario, fecha
    `;
    const rows = await queryPostgres(sql, [comentario, usuario]);

    return rows[0];
}

module.exports = {
    createExchangeRateComment,
    listExchangeRateComments
};
