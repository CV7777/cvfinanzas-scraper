/**
 * PROXY para GoHighLevel CRM — Google Apps Script
 * Soporta acciones: "register" (upsert) y "search" (buscar por email)
 */

var GHL_API_KEY = getRequiredScriptProperty_("GHL_API_KEY");
var GHL_LOCATION_ID = "ortc5ChhiliYRLpw9ktA";
var TOKENGIT = getRequiredScriptProperty_("GITHUB_TOKEN");
var REPO = "CV7777/cvfinanzas-scraper";
var FILECREDEN = "datos-json/creden.json";
var GHL_UPSERT_URL = "https://services.leadconnectorhq.com/contacts/upsert";
var GHL_SEARCH_URL = "https://services.leadconnectorhq.com/contacts/search";

/**
 * Lee secretos desde Configuración del proyecto > Propiedades de la
 * secuencia de comandos. Los tokens nunca deben guardarse en este archivo.
 */
function getRequiredScriptProperty_(name) {
  var value = PropertiesService.getScriptProperties().getProperty(name);
  if (!value) {
    throw new Error("Falta configurar la propiedad de Apps Script: " + name);
  }
  return value;
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action || "register";

    // -- ACCION: buscar contacto por email (login / verificacion de duplicado) --
    if (action === "search") {
      return searchContact(data.email);
    }

    // -- ACCION: registrar nuevo contacto --
    if (action === "register") {
      return registerContact(data);
    }

    // -- ACCION: validar credenciales y actualizar --
    if (action === "validateCredentialsAndUpdate") {
      var username = data.username;
      var password = data.password;
      var result = validateCredentialsAndUpdate(username, password);
      return jsonResponse(result);
    }

    return jsonResponse({ error: "Accion no reconocida: " + action });
  } catch (err) {
    Logger.log("Error: " + err.message);
    return jsonResponse({ error: err.message });
  }
}

/**
 * Busca un contacto en GHL por email.
 * Retorna { found: true, contact: {...} } o { found: false }
 */
function searchContact(email) {
  if (!email) {
    return jsonResponse({ error: "Email requerido para buscar" });
  }

  try {
    var searchBody = {
      locationId: GHL_LOCATION_ID,
      page: 1,
      pageLimit: 1,
      filters: [
        {
          field: "email",
          operator: "eq",
          value: email,
        },
      ],
    };

    var options = {
      method: "post",
      contentType: "application/json",
      headers: {
        Authorization: "Bearer " + GHL_API_KEY,
        Version: "2021-07-28",
      },
      payload: JSON.stringify(searchBody),
      muteHttpExceptions: true,
    };

    var response = UrlFetchApp.fetch(GHL_SEARCH_URL, options);
    var code = response.getResponseCode();
    var body = response.getContentText();

    Logger.log("GHL Search [" + code + "]: " + body);

    if (code >= 200 && code < 300) {
      var result = JSON.parse(body);
      var contacts = result.contacts || [];

      if (contacts.length > 0) {
        var c = contacts[0];
        return jsonResponse({
          found: true,
          contact: {
            id: c.id,
            firstName: c.firstName || "",
            lastName: c.lastName || "",
            email: c.email || "",
          },
        });
      } else {
        return jsonResponse({ found: false });
      }
    } else {
      return jsonResponse({
        error: "GHL Search codigo " + code,
        details: body,
      });
    }
  } catch (err) {
    Logger.log("Search Error: " + err.message);
    return jsonResponse({ error: err.message });
  }
}

/**
 * Registra (upsert) un contacto en GHL.
 */
function registerContact(data) {
  if (!data.email || !data.firstName) {
    return jsonResponse({ error: "Faltan campos requeridos" });
  }

  var customFields =
    data.customFields && data.customFields.length > 0
      ? data.customFields.slice()
      : [];

  // Si viene edad, calcular fecha de nacimiento y asignar al campo estándar
  var dateOfBirth = data.dateOfBirth || "";
  if (data.edad && !data.dateOfBirth) {
    var edad = parseInt(data.edad, 10);
    if (!isNaN(edad) && edad > 0 && edad < 120) {
      var hoy = new Date();
      var fechaNacimiento = new Date(
        hoy.getFullYear() - edad,
        hoy.getMonth(),
        hoy.getDate(),
      );
      dateOfBirth = fechaNacimiento.toISOString().split("T")[0];
    }
  }

  var ghlBody = {
    firstName: data.firstName,
    lastName: data.lastName || "",
    email: data.email,
    locationId: GHL_LOCATION_ID,
    source: data.source || "CV Finanzas - Tipo de Cambio",
    tags: data.tags || ["cvfinanzas", "tipo-cambio", "lead-web"],
  };

  if (dateOfBirth) {
    ghlBody.dateOfBirth = dateOfBirth;
  }
  if (customFields.length > 0) {
    ghlBody.customFields = customFields;
  }

  var options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + GHL_API_KEY,
      Version: "2021-07-28",
    },
    payload: JSON.stringify(ghlBody),
    muteHttpExceptions: true,
  };

  var response = UrlFetchApp.fetch(GHL_UPSERT_URL, options);
  var code = response.getResponseCode();
  var body = response.getContentText();

  Logger.log("GHL Register [" + code + "]: " + body);

  if (code >= 200 && code < 300) {
    var result = JSON.parse(body);
    return jsonResponse({
      success: true,
      contactId: result.contact ? result.contact.id : null,
      message: "Contacto sincronizado",
    });
  } else {
    return jsonResponse({
      success: false,
      error: "GHL codigo " + code,
      details: body,
    });
  }
}

function doGet(e) {
  return jsonResponse({
    status: "ok",
    message: "Proxy GHL activo. Acciones: register, search",
  });
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(
    ContentService.MimeType.JSON,
  );
}

function validateCredentials(username, password) {
  try {
    // URL del archivo creden.json en el repositorio
    var repoUrl =
      "https://api.github.com/repos/" + REPO + "/contents/" + FILECREDEN;

    // Configurar la solicitud con el token de GitHub
    var options = {
      method: "get",
      headers: {
        Authorization: "Bearer " + TOKENGIT,
        Accept: "application/vnd.github+json",
      },
      muteHttpExceptions: true,
    };

    // Realizar la solicitud para obtener el archivo creden.json
    var response = UrlFetchApp.fetch(repoUrl, options);
    var code = response.getResponseCode();
    var body = response.getContentText();

    if (code >= 200 && code < 300) {
      // Decodificar el contenido del archivo
      var fileData = JSON.parse(body);
      var decodedContent = JSON.parse(
        Utilities.newBlob(
          Utilities.base64Decode(fileData.content),
        ).getDataAsString(),
      );
      // Buscar el usuario en el archivo creden.json
      var user = decodedContent.usuarios.find(function (u) {
        return u.user === username && u.pass === password;
      });

      if (user) {
        Logger.log("Credenciales válidas");
        return { success: true, message: "Credenciales válidas" };
      } else {
        Logger.log("Credenciales inválidas");
        return { success: false, message: "Credenciales inválidas" };
      }
    } else {
      Logger.log("Error al acceder al repositorio: " + code);
      return {
        success: false,
        message: "Error al acceder al repositorio: " + code,
      };
    }
  } catch (err) {
    Logger.log("Error en validateCredentials: " + err.message);
    return { success: false, message: "Error interno: " + err.message };
  }
}

function validateCredentialsAndUpdate(username, password) {
  try {
    // Validar credenciales usando validateCredentials
    var validationResult = validateCredentials(username, password);

    if (!validationResult.success) {
      return { success: false, message: "Credenciales inválidas" };
    }

    // Si las credenciales son válidas, actualizar el campo ultacc
    var repoUrl =
      "https://api.github.com/repos/" + REPO + "/contents/" + FILECREDEN;

    var options = {
      method: "get",
      headers: {
        Authorization: "Bearer " + TOKENGIT,
        Accept: "application/vnd.github+json",
      },
      muteHttpExceptions: true,
    };

    var response = UrlFetchApp.fetch(repoUrl, options);
    var code = response.getResponseCode();
    var body = response.getContentText();

    if (code >= 200 && code < 300) {
      var fileData = JSON.parse(body);
      var decodedContent = JSON.parse(
        Utilities.newBlob(
          Utilities.base64Decode(fileData.content),
        ).getDataAsString(),
      );

      var user = decodedContent.usuarios.find(function (u) {
        return u.user === username;
      });

      if (user) {
        user.ultacc = new Date().toISOString();

        var updatedContent = Utilities.base64Encode(
          Utilities.newBlob(JSON.stringify(decodedContent, null, 2)).getBytes(),
        );

        var updateOptions = {
          method: "put",
          contentType: "application/json",
          headers: {
            Authorization: "Bearer " + TOKENGIT,
            Accept: "application/vnd.github+json",
          },
          payload: JSON.stringify({
            message: "metodo validateCredentialsAndUpdate: " + username,
            content: updatedContent,
            sha: fileData.sha,
          }),
          muteHttpExceptions: true,
        };

        var updateResponse = UrlFetchApp.fetch(repoUrl, updateOptions);
        var updateCode = updateResponse.getResponseCode();

        if (updateCode >= 200 && updateCode < 300) {
          Logger.log("Campo ultacc actualizado correctamente para " + username);
        } else {
          Logger.log("Error al actualizar creden.json: " + updateCode);
        }
      }
    }

    return { success: true, message: "Credenciales válidas" };
  } catch (err) {
    Logger.log("Error en validateCredentialsAndUpdate: " + err.message);
    return { success: false, message: "Error interno: " + err.message };
  }
}

function updateAnalisis(fecha, texto, hora) {
  try {
    // URL del archivo analisis.json en el repositorio
    var repoUrl =
      "https://api.github.com/repos/" +
      REPO +
      "/contents/datos-json/analisis.json";

    // Configurar la solicitud con el token de GitHub
    var options = {
      method: "get",
      headers: {
        Authorization: "Bearer " + TOKENGIT,
        Accept: "application/vnd.github+json",
      },
      muteHttpExceptions: true,
    };

    // Realizar la solicitud para obtener el archivo analisis.json
    var response = UrlFetchApp.fetch(repoUrl, options);
    var code = response.getResponseCode();
    var body = response.getContentText();

    if (code >= 200 && code < 300) {
      // Decodificar el contenido del archivo
      var fileData = JSON.parse(body);
      var decodedContent = JSON.parse(
        Utilities.newBlob(
          Utilities.base64Decode(fileData.content),
        ).getDataAsString(),
      );

      // Agregar el nuevo análisis al inicio del array
      decodedContent.analisis.unshift({
        fecha: fecha,
        texto: texto,
        hora: hora,
      });

      // Limitar el historial a 30 entradas
      if (decodedContent.analisis.length > 30) {
        decodedContent.analisis = decodedContent.analisis.slice(0, 30);
      }

      // Codificar el contenido actualizado
      var updatedContent = Utilities.base64Encode(
        Utilities.newBlob(JSON.stringify(decodedContent, null, 2)).getBytes(),
      );

      // Enviar los cambios al repositorio
      var updateOptions = {
        method: "put",
        contentType: "application/json",
        headers: {
          Authorization: "Bearer " + TOKENGIT,
          Accept: "application/vnd.github+json",
        },
        payload: JSON.stringify({
          message: "Actualizar analisis.json con nuevo análisis",
          content: updatedContent,
          sha: fileData.sha,
        }),
        muteHttpExceptions: true,
      };

      var updateResponse = UrlFetchApp.fetch(repoUrl, updateOptions);
      var updateCode = updateResponse.getResponseCode();

      if (updateCode >= 200 && updateCode < 300) {
        Logger.log("Archivo analisis.json actualizado correctamente");
      } else {
        Logger.log("Error al actualizar analisis.json: " + updateCode);
      }
    } else {
      Logger.log("Error al acceder al repositorio: " + code);
    }
  } catch (err) {
    Logger.log("Error en updateAnalisis: " + err.message);
  }
}

// -- FUNCION DE PRUEBA (ejecutar desde el editor) --
function testDoPost() {
  var fakeEvent = {
    postData: {
      contents: JSON.stringify({
        action: "register",
        firstName: "Kevin",
        lastName: "Test",
        email: "kevin.test@ejemplo.com",
      }),
    },
  };
  var result = doPost(fakeEvent);
  Logger.log(result.getContent());
}

function testSearch() {
  var fakeEvent = {
    postData: {
      contents: JSON.stringify({
        action: "search",
        email: "prueba@gmail.com",
      }),
    },
  };
  var result = doPost(fakeEvent);
  Logger.log(result.getContent());
}

function testValidateCredentials() {
  var username = getRequiredScriptProperty_("TEST_USERNAME"); // Reemplaza con un usuario de prueba
  var password = getRequiredScriptProperty_("TEST_PASSWORD"); // Reemplaza con la contraseña correspondiente

  var result = validateCredentials(username, password);
  Logger.log(result); // Esto imprimirá el resultado en los logs de Apps Script
}

function testValidateCredentialsandUpdate() {
  var username = getRequiredScriptProperty_("TEST_USERNAME"); // Reemplaza con un usuario de prueba
  var password = getRequiredScriptProperty_("TEST_PASSWORD"); // Reemplaza con la contraseña correspondiente

  var result = validateCredentialsAndUpdate(username, password);
  Logger.log(result); // Esto imprimirá el resultado en los logs de Apps Script
}

function testUpdateAnalisis() {
  var fecha = "2023-10-10";
  var texto = "Este es un análisis de prueba";
  var hora = "10:00";

  updateAnalisis(fecha, texto, hora);
}
