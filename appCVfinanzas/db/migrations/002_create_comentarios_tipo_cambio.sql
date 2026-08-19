BEGIN;

CREATE TABLE IF NOT EXISTS comentarios_tipo_cambio (
    id BIGSERIAL PRIMARY KEY,
    comentario TEXT NOT NULL,
    usuario VARCHAR(255) NOT NULL,
    fecha TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT comentarios_tipo_cambio_comentario_chk
        CHECK (char_length(btrim(comentario)) BETWEEN 1 AND 2000),
    CONSTRAINT comentarios_tipo_cambio_usuario_chk
        CHECK (char_length(btrim(usuario)) BETWEEN 1 AND 255)
);

COMMENT ON TABLE comentarios_tipo_cambio IS
    'Bitacora de comentarios ingresados desde el mantenimiento de tipo de cambio.';
COMMENT ON COLUMN comentarios_tipo_cambio.usuario IS
    'Usuario autenticado que registro el comentario.';
COMMENT ON COLUMN comentarios_tipo_cambio.fecha IS
    'Fecha y hora en que se registro el comentario.';

CREATE INDEX IF NOT EXISTS comentarios_tipo_cambio_fecha_idx
    ON comentarios_tipo_cambio (fecha DESC);

COMMIT;
