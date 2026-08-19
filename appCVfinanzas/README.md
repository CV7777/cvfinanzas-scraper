# CVFinanzas Results App

Este proyecto es una aplicación web que muestra los resultados de un cuestionario sobre la relación con el dinero. Utiliza Node.js y Express para el backend y Tailwind CSS para el diseño frontend.

## Estructura del Proyecto

```
cvfinanzas-results-app
├── src
│   ├── server.js          # Punto de entrada de la aplicación
│   ├── routes
│   │   └── index.js       # Definición de rutas
│   ├── views
│   │   └── result.ejs     # Plantilla EJS para mostrar resultados
│   └── styles
│       └── tailwind.css    # Estilos de Tailwind CSS
├── public
│   └── js
│       └── app.js         # Código JavaScript del lado del cliente
├── package.json           # Configuración de npm
├── tailwind.config.js     # Configuración de Tailwind CSS
├── postcss.config.js      # Configuración para PostCSS
├── .gitignore             # Archivos y carpetas a ignorar por Git
└── README.md              # Documentación del proyecto
```

## Instalación

1. Clona el repositorio:

   ```
   git clone <URL_DEL_REPOSITORIO>
   ```

2. Navega al directorio del proyecto:

   ```
   cd cvfinanzas-results-app
   ```

3. Instala las dependencias:
   ```
   npm install
   ```

## Uso

1. Inicia el servidor:

   ```
   npm start
   ```

2. Abre tu navegador y visita `http://localhost:3000` para ver la aplicación en acción.

## Tailwind CSS

El proyecto compila Tailwind desde `src/styles/tailwind.css` hacia `public/styles/tailwind.css`.

```
npm run build:css
```

`npm start` y `npm run dev` ejecutan ese build antes de levantar Express.

## Panel administrativo

El panel privado usa una estructura adaptada del Flowbite Admin Dashboard: sidebar, navbar superior, contenido principal y footer. La plantilla original es open-source MIT y esta basada en Tailwind CSS + Flowbite.

Pantallas protegidas:

```
http://localhost:3000/dashboard
http://localhost:3000/search
http://localhost:3000/gastos
http://localhost:3000/tipo-cambio
```

La pantalla `/gastos` permite probar una interfaz temporal para registrar gastos, subir/tomar foto de una factura y simular extraccion de datos. Por ahora no guarda en base de datos.

La pantalla `/tipo-cambio` permite registrar comentarios de mantenimiento. El usuario se toma de la sesion activa y la fecha se genera en PostgreSQL. Para crear la tabla requerida:

```sh
npm run migrar-comentarios-tipo-cambio
```

## Conexion con Azure PostgreSQL

1. Instala las dependencias si todavia no estan disponibles:

   ```
   npm install
   ```

2. Crea tu archivo `.env` a partir de `.env.example`:

   ```
   cp .env.example .env
   ```

3. Configura Azure PostgreSQL para resultados y usuarios:

   ```
   AZURE_PG_HOST=cvfinanzas.postgres.database.azure.com
   AZURE_PG_USER=cvfinanzas
   AZURE_PG_PASSWORD=your_password
   AZURE_PG_DATABASE=postgres
   AZURE_PG_PORT=5432
   AZURE_PG_RESULTS_TABLE=quiz_personalidad_results
   AZURE_PG_USERS_TABLE=usuarios
   JWT_SECRET=change_this_to_a_long_random_secret
   AUTH_TOKEN_TTL=8h
   # AZURE_PG_SSL_CA_PATH=certs/DigiCertGlobalRootCA.crt.pem
   ```

   `JWT_SECRET` debe ser un valor largo y privado en produccion. En desarrollo, si no lo configuras, Express crea uno temporal y las sesiones se cierran al reiniciar el servidor.

   Para usar Azure PostgreSQL y login con hash de contrasena, instala las dependencias:

   ```
   npm install
   ```

   Si `npm install` falla por permisos en `node_modules`, corrige el dueño de la carpeta o reinstala dependencias antes de correrlo.

4. Reinicia Express y prueba:

   ```
   npm start
   ```

   Rutas disponibles:

   ```
   http://localhost:3000
   http://localhost:3000/result
   http://localhost:3000/result/el-ambicioso
   http://localhost:3000/result?email=prueba@gmail.com
   http://localhost:3000/test-postgres
   ```

## API de busqueda

La ruta `/search-results` permite traer todos los resultados o filtrar por email, por tipo de perfil, o por ambos filtros:

```
/search-results
/search-results?email=usuario@ejemplo.com
/search-results?profile_type=G
/search-results?email=usuario@ejemplo.com&profile_type=G
```

Tambien acepta `profileType`, `tipo_perfil` o `tipo` como nombre del parametro para el tipo de perfil.

## Datos MONEX del BCCR

La migracion crea una tabla normalizada usando la configuracion PostgreSQL del
archivo `.env`:

```sh
npm run migrar-monex
```

Antes de importar se puede validar el JSON sin conectarse a PostgreSQL:

```sh
npm run importar-monex -- --validar
```

Para importar `../datos-json/datos.json`, usando la configuracion PostgreSQL del
archivo `.env`:

```sh
npm run importar-monex
```

Tambien se puede indicar otra ruta. La importacion es idempotente: inserta filas
nuevas, actualiza solo las que cambiaron y usa `(fecha, sesion)` como llave unica.

```sh
npm run importar-monex -- /ruta/al/datos.json
```

Ejemplo de consulta para una API con paginacion por fecha:

```sql
SELECT fecha, sesion, promedio_ponderado, monto_total, minimo, maximo,
       capturado_en, ultima_actualizacion
FROM monex_tipo_cambio
WHERE fecha BETWEEN $1::date AND $2::date
ORDER BY fecha DESC, sesion DESC
LIMIT $3;
```

## Login

El buscador esta protegido con usuarios guardados en Azure PostgreSQL.

```
http://localhost:3000/login
```

La tabla de login debe llamarse `usuarios` o configurarse con `AZURE_PG_USERS_TABLE`. El backend espera estas columnas:

```
id uuid
usuario varchar
password_hash text
fecha_creacion timestamptz
```

`password_hash` debe ser un hash bcrypt. Para generar uno:

```
node -e "const bcrypt = require('bcryptjs'); bcrypt.hash('tu_contrasena', 12).then(console.log)"
```

Tambien puedes crear o resetear un usuario directamente desde el proyecto:

```
npm run upsert-user -- prueba@gmail.com tu_contrasena
```

Al iniciar sesion correctamente, Express crea un JWT en una cookie `HttpOnly` y redirige a:

```
http://localhost:3000/search
```
