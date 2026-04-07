/**
 * PROXY para GoHighLevel CRM — Google Apps Script
 * Soporta acciones: "register" (upsert) y "search" (buscar por email)
 */

var GHL_API_KEY = "pit-b9b523ad-4d10-4e90-9877-d3e290fc1ab7";
var GHL_LOCATION_ID = "ortc5ChhiliYRLpw9ktA";
var GHL_UPSERT_URL = "https://services.leadconnectorhq.com/contacts/upsert";
var GHL_SEARCH_URL = "https://services.leadconnectorhq.com/contacts/search";

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

  var ghlBody = {
    firstName: data.firstName,
    lastName: data.lastName || "",
    email: data.email,
    locationId: GHL_LOCATION_ID,
    source: data.source || "CV Finanzas - Tipo de Cambio",
    tags: data.tags || ["cvfinanzas", "tipo-cambio", "lead-web"],
  };

  if (data.customFields && data.customFields.length > 0) {
    ghlBody.customFields = data.customFields;
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
        email: "kevin.test@ejemplo.com",
      }),
    },
  };
  var result = doPost(fakeEvent);
  Logger.log(result.getContent());
}
