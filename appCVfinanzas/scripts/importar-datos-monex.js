require('dotenv').config();

const fs = require('fs');
const path = require('path');
const { queryPostgres, quoteTableName } = require('../src/lib/postgres');

const FECHA_MINIMA = '2024-01-01';
const COSTA_RICA_TZ = 'America/Costa_Rica';
const tableName = quoteTableName(
    process.env.AZURE_PG_MONEX_TABLE || process.env.PG_MONEX_TABLE || 'monex_tipo_cambio'
);

function esNumeroFinito(value) {
    return typeof value === 'number' && Number.isFinite(value);
}

function esFechaValida(value) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(value || '')) {
        return false;
    }
    const [year, month, day] = value.split('-').map(Number);
    const date = new Date(Date.UTC(year, month - 1, day));
    return date.getUTCFullYear() === year &&
        date.getUTCMonth() === month - 1 &&
        date.getUTCDate() === day;
}

function esHoraValida(value) {
    const match = /^(\d{2}):(\d{2})(?::(\d{2}))?$/.exec(value || '');
    return Boolean(match && Number(match[1]) < 24 && Number(match[2]) < 60 &&
        Number(match[3] || 0) < 60);
}

function leerYValidar(ruta) {
    const contenido = JSON.parse(fs.readFileSync(ruta, 'utf8'));

    if (!contenido || !Array.isArray(contenido.datos)) {
        throw new Error('El JSON debe tener la estructura { "actualizado": ..., "datos": [...] }.');
    }

    const claves = new Set();
    const errores = [];
    let valoresFaltantesConvertidos = 0;

    const datos = contenido.datos.map((fila, indice) => {
        const numeroFila = indice + 1;
        const clave = `${fila.fecha}|${fila.sesion}`;

        if (!esFechaValida(fila.fecha) || fila.fecha < FECHA_MINIMA) {
            errores.push(`fila ${numeroFila}: fecha invalida o anterior a ${FECHA_MINIMA}`);
        }
        if (!esHoraValida(fila.sesion)) {
            errores.push(`fila ${numeroFila}: sesion invalida`);
        }
        const partesTimestamp = /^(\d{4}-\d{2}-\d{2}) (\d{2}:\d{2}(?::\d{2})?)$/.exec(
            fila.timestamp || ''
        );
        if (!partesTimestamp || !esFechaValida(partesTimestamp[1]) || !esHoraValida(partesTimestamp[2])) {
            errores.push(`fila ${numeroFila}: timestamp invalido`);
        }
        if (claves.has(clave)) {
            errores.push(`fila ${numeroFila}: fecha y sesion duplicadas (${clave})`);
        }
        claves.add(clave);

        for (const campo of ['promedio_ponderado', 'monto_total']) {
            if (!esNumeroFinito(fila[campo])) {
                errores.push(`fila ${numeroFila}: ${campo} no es numerico`);
            }
        }
        for (const campo of ['minimo', 'maximo']) {
            if (fila[campo] !== null && !esNumeroFinito(fila[campo])) {
                errores.push(`fila ${numeroFila}: ${campo} debe ser numerico o null`);
            }
        }
        if (fila.promedio_ponderado <= 0 || fila.monto_total <= 0) {
            errores.push(`fila ${numeroFila}: promedio_ponderado y monto_total deben ser positivos`);
        }

        const minimo = fila.minimo === 0 ? null : fila.minimo;
        const maximo = fila.maximo === 0 ? null : fila.maximo;
        valoresFaltantesConvertidos += Number(fila.minimo === 0) + Number(fila.maximo === 0);

        if (minimo !== null && minimo > fila.promedio_ponderado) {
            errores.push(`fila ${numeroFila}: minimo es mayor que promedio_ponderado`);
        }
        if (maximo !== null && maximo < fila.promedio_ponderado) {
            errores.push(`fila ${numeroFila}: maximo es menor que promedio_ponderado`);
        }

        return { ...fila, minimo, maximo };
    });

    if (errores.length) {
        throw new Error(`No se puede importar:\n- ${errores.slice(0, 20).join('\n- ')}`);
    }

    return { datos, actualizado: contenido.actualizado, valoresFaltantesConvertidos };
}

async function main() {
    const argumentos = process.argv.slice(2);
    const soloValidar = argumentos.includes('--validar');
    const rutaIndicada = argumentos.find((argumento) => argumento !== '--validar');
    const ruta = path.resolve(
        rutaIndicada || path.join(__dirname, '..', '..', 'datos-json', 'datos.json')
    );
    const { datos, actualizado, valoresFaltantesConvertidos } = leerYValidar(ruta);

    console.log(`${datos.length} registros validos desde ${datos[0]?.fecha || 'sin datos'}.`);
    console.log(`${valoresFaltantesConvertidos} valores cero de minimo/maximo se importaran como NULL.`);

    if (soloValidar) {
        console.log('Validacion finalizada; no se modifico PostgreSQL.');
        return;
    }

    const rows = await queryPostgres(
        `
        WITH fuente AS (
            SELECT *
            FROM jsonb_to_recordset($1::jsonb) AS x(
                fecha date,
                promedio_ponderado numeric,
                monto_total numeric,
                minimo numeric,
                maximo numeric,
                sesion time,
                "timestamp" text
            )
        ), importados AS (
            INSERT INTO ${tableName} AS destino (
                fecha, sesion, promedio_ponderado, monto_total,
                minimo, maximo, capturado_en
            )
            SELECT
                fecha,
                sesion,
                promedio_ponderado,
                monto_total,
                minimo,
                maximo,
                ("timestamp" || ' ${COSTA_RICA_TZ}')::timestamptz
            FROM fuente
            ON CONFLICT (fecha, sesion) DO UPDATE SET
                promedio_ponderado = EXCLUDED.promedio_ponderado,
                monto_total = EXCLUDED.monto_total,
                minimo = EXCLUDED.minimo,
                maximo = EXCLUDED.maximo,
                capturado_en = EXCLUDED.capturado_en
            WHERE (destino.promedio_ponderado,
                   destino.monto_total,
                   destino.minimo,
                   destino.maximo,
                   destino.capturado_en)
                  IS DISTINCT FROM
                  (EXCLUDED.promedio_ponderado,
                   EXCLUDED.monto_total,
                   EXCLUDED.minimo,
                   EXCLUDED.maximo,
                   EXCLUDED.capturado_en)
            RETURNING (xmax = 0) AS insertado
        )
        SELECT
            count(*) FILTER (WHERE insertado)::int AS insertados,
            count(*) FILTER (WHERE NOT insertado)::int AS actualizados
        FROM importados
        `,
        [JSON.stringify(datos)]
    );

    const resultado = rows[0] || { insertados: 0, actualizados: 0 };
    const sinCambios = datos.length - resultado.insertados - resultado.actualizados;
    console.log(
        `Importacion terminada: ${resultado.insertados} insertados, ` +
        `${resultado.actualizados} actualizados y ${sinCambios} sin cambios.`
    );
    if (actualizado) {
        console.log(`El archivo indica que fue actualizado: ${actualizado}.`);
    }
}

main().catch((error) => {
    console.error(error.message);
    process.exitCode = 1;
});
