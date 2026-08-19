BEGIN;

CREATE TABLE IF NOT EXISTS monex_tipo_cambio (
    fecha DATE NOT NULL,
    sesion TIME WITHOUT TIME ZONE NOT NULL,
    promedio_ponderado NUMERIC(10, 4) NOT NULL,
    monto_total NUMERIC(18, 2) NOT NULL,
    minimo NUMERIC(10, 4),
    maximo NUMERIC(10, 4),
    capturado_en TIMESTAMPTZ NOT NULL,
    ultima_actualizacion TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT monex_tipo_cambio_pk PRIMARY KEY (fecha, sesion),
    CONSTRAINT monex_fecha_desde_2024_chk CHECK (fecha >= DATE '2024-01-01'),
    CONSTRAINT monex_promedio_positivo_chk CHECK (promedio_ponderado > 0),
    CONSTRAINT monex_monto_positivo_chk CHECK (monto_total > 0),
    CONSTRAINT monex_minimo_positivo_chk CHECK (minimo IS NULL OR minimo > 0),
    CONSTRAINT monex_maximo_positivo_chk CHECK (maximo IS NULL OR maximo > 0),
    CONSTRAINT monex_rango_chk CHECK (
        (minimo IS NULL OR minimo <= promedio_ponderado)
        AND (maximo IS NULL OR promedio_ponderado <= maximo)
        AND (minimo IS NULL OR maximo IS NULL OR minimo <= maximo)
    )
);

COMMENT ON TABLE monex_tipo_cambio IS
    'Resultados por sesion del mercado MONEX publicados por el BCCR.';
COMMENT ON COLUMN monex_tipo_cambio.capturado_en IS
    'Fecha y hora de Costa Rica en que la fuente fue consultada.';
COMMENT ON COLUMN monex_tipo_cambio.ultima_actualizacion IS
    'Fecha y hora en que este registro fue insertado o modificado en PostgreSQL.';

-- La llave primaria ya sirve para consultas por fecha. Este indice
-- mejora el caso comun de una API que solicita primero los datos mas recientes.
CREATE INDEX IF NOT EXISTS monex_tipo_cambio_recientes_idx
    ON monex_tipo_cambio (capturado_en DESC);

CREATE OR REPLACE FUNCTION monex_tipo_cambio_set_ultima_actualizacion()
RETURNS TRIGGER
LANGUAGE plpgsql
AS $$
BEGIN
    NEW.ultima_actualizacion = CURRENT_TIMESTAMP;
    RETURN NEW;
END;
$$;

DROP TRIGGER IF EXISTS monex_tipo_cambio_actualizado_trg ON monex_tipo_cambio;
CREATE TRIGGER monex_tipo_cambio_actualizado_trg
BEFORE UPDATE ON monex_tipo_cambio
FOR EACH ROW
EXECUTE FUNCTION monex_tipo_cambio_set_ultima_actualizacion();

COMMIT;
