#!/bin/bash
set -euo pipefail

# ==============================
# Deploy automático Node.js + Express en Azure App Service
# ==============================

RESOURCE_GROUP="${AZURE_RESOURCE_GROUP:-AppCv77}"
APP_NAME="${AZURE_WEBAPP_NAME:-cvfinanzas}"
SUBSCRIPTION_ID="${AZURE_SUBSCRIPTION_ID:-5c361cc3-7542-4d13-8267-1d18527a425c}"
APP_SERVICE_PLAN="${AZURE_APP_SERVICE_PLAN:-ASP-AppCv77-b392}"
APP_SERVICE_SKU="${AZURE_APP_SERVICE_SKU:-F1}"
LOCATION="${AZURE_LOCATION:-centralus}"
NODE_RUNTIME="${AZURE_NODE_RUNTIME:-NODE|20-lts}"
WEBAPP_CREATE_RUNTIME="${AZURE_WEBAPP_CREATE_RUNTIME:-NODE:20-lts}"
PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
ZIP_FILE="${PROJECT_DIR}/azure-node-express.zip"
ENV_FILE="${PROJECT_DIR}/.env"

required_settings=(
  AZURE_PG_HOST
  AZURE_PG_USER
  AZURE_PG_PASSWORD
  AZURE_PG_DATABASE
  AZURE_PG_PORT
  AZURE_PG_RESULTS_TABLE
  AZURE_PG_USERS_TABLE
  JWT_SECRET
  AUTH_TOKEN_TTL
)

cd "$PROJECT_DIR"

echo "🧹 Limpiando ZIP previo..."
rm -f "$ZIP_FILE"

echo "🔎 Verificando Azure CLI..."
if ! command -v az >/dev/null 2>&1; then
    echo "❌ Error: Azure CLI no esta instalado."
    exit 1
fi

echo "✅ Verificando package.json..."
if [ ! -f "package.json" ]; then
    echo "❌ Error: no existe package.json"
    exit 1
fi

echo "✅ Verificando script start..."
if ! grep -q '"start"' package.json; then
    echo "❌ Error: package.json no tiene script start"
    exit 1
fi

echo "📥 Instalando dependencias..."
npm install

echo "🎨 Compilando Tailwind..."
npm run build:css

if [ -f "$ENV_FILE" ]; then
    echo "🔐 Cargando variables desde .env para App Settings..."
    set -a
    # shellcheck disable=SC1090
    source "$ENV_FILE"
    set +a
else
    echo "⚠️ No existe .env local. Usare variables exportadas en la terminal."
fi

for setting in "${required_settings[@]}"; do
    if [ -z "${!setting:-}" ]; then
        echo "❌ Error: falta $setting en .env o en variables de entorno."
        exit 1
    fi
done

echo "📦 Empaquetando proyecto..."
zip -rq "$ZIP_FILE" . \
  -x "node_modules/*" \
  -x ".env" \
  -x ".git/*" \
  -x ".DS_Store" \
  -x "npm-debug.log*" \
  -x "azure-node-express.zip"

echo "🔐 Iniciando sesión en Azure..."
echo "   Usa la cuenta donde existe el App Service $APP_NAME."
az login 

echo "🔁 Usando suscripción $SUBSCRIPTION_ID..."
if ! az account set --subscription "$SUBSCRIPTION_ID"; then
    echo "❌ Azure CLI no tiene acceso a la suscripción $SUBSCRIPTION_ID."
    echo "   Inicia sesión con la misma cuenta que usas en el portal de Azure:"
    echo "   az login --use-device-code"
    echo "   Luego vuelve a correr: ./deployAzure.sh"
    exit 1
fi

echo "🏗️ Verificando resource group..."
if ! az group show --name "$RESOURCE_GROUP" --output none >/dev/null 2>&1; then
    az group create \
      --name "$RESOURCE_GROUP" \
      --location "$LOCATION" \
      --output none
fi

echo "🏗️ Verificando App Service Plan..."
if ! az appservice plan show --resource-group "$RESOURCE_GROUP" --name "$APP_SERVICE_PLAN" --output none >/dev/null 2>&1; then
    az appservice plan create \
      --resource-group "$RESOURCE_GROUP" \
      --name "$APP_SERVICE_PLAN" \
      --location "$LOCATION" \
      --is-linux \
      --sku "$APP_SERVICE_SKU" \
      --output none
fi

echo "🏗️ Verificando Web App..."
if ! az webapp show --resource-group "$RESOURCE_GROUP" --name "$APP_NAME" --output none >/dev/null 2>&1; then
    az webapp create \
      --resource-group "$RESOURCE_GROUP" \
      --plan "$APP_SERVICE_PLAN" \
      --name "$APP_NAME" \
      --runtime "$WEBAPP_CREATE_RUNTIME" \
      --output none
fi

echo "⚙️ Configurando runtime Node.js..."
az webapp config set \
  --resource-group "$RESOURCE_GROUP" \
  --name "$APP_NAME" \
  --linux-fx-version "$NODE_RUNTIME" \
  --output none

echo "⚙️ Configurando startup command..."
az webapp config set \
  --resource-group "$RESOURCE_GROUP" \
  --name "$APP_NAME" \
  --startup-file "npm start" \
  --output none

echo "🔐 Configurando App Settings..."
az webapp config appsettings set \
  --resource-group "$RESOURCE_GROUP" \
  --name "$APP_NAME" \
  --settings \
    NODE_ENV="production" \
    WEBSITE_NODE_DEFAULT_VERSION="~20" \
    SCM_DO_BUILD_DURING_DEPLOYMENT="true" \
    ENABLE_ORYX_BUILD="true" \
    AZURE_PG_HOST="$AZURE_PG_HOST" \
    AZURE_PG_USER="$AZURE_PG_USER" \
    AZURE_PG_PASSWORD="$AZURE_PG_PASSWORD" \
    AZURE_PG_DATABASE="$AZURE_PG_DATABASE" \
    AZURE_PG_PORT="$AZURE_PG_PORT" \
    AZURE_PG_RESULTS_TABLE="$AZURE_PG_RESULTS_TABLE" \
    AZURE_PG_USERS_TABLE="$AZURE_PG_USERS_TABLE" \
    JWT_SECRET="$JWT_SECRET" \
    AUTH_TOKEN_TTL="$AUTH_TOKEN_TTL" \
  --output none

echo "🚀 Subiendo a Azure App Service..."
az webapp deployment source config-zip \
  --resource-group "$RESOURCE_GROUP" \
  --name "$APP_NAME" \
  --src "$ZIP_FILE" \
  --output none

echo "📝 Esperando 5s para validación..."
sleep 5

echo "🩺 Revisando disponibilidad..."
curl -I "https://${APP_NAME}.azurewebsites.net/login" || echo "⚠️ No se pudo contactar el sitio."

echo "✅ Despliegue completado correctamente."
echo "🌐 URL: https://${APP_NAME}.azurewebsites.net/login"
