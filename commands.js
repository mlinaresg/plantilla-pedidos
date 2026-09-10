/* ============================================================
   Plantilla reenvío pedidos - lógica del botón
   Inserta los campos de pedido en el cuerpo SOLO si
   pedidosclientes@plexus.es está entre los destinatarios (Para/CC).

   Robustez: cada callback anidado tiene su propio try/catch, y hay
   un temporizador de seguridad que garantiza que el botón nunca se
   quede "trabajando" indefinidamente aunque algo falle de forma
   inesperada.
   ============================================================ */

var CORREO_OBJETIVO = "pedidosclientes@plexus.es";
var TIMEOUT_SEGURIDAD_MS = 8000;

function insertarPlantilla(event) {
  var finalizado = false;

  function finalizar() {
    if (finalizado) return;
    finalizado = true;
    try {
      event.completed();
    } catch (e) {
      // Si ni siquiera esto funciona, no hay nada más que hacer.
    }
  }

  // Red de seguridad: si nada más llama a esto en unos segundos,
  // lo forzamos igualmente en vez de dejar el botón colgado.
  var temporizador = setTimeout(function () {
    if (!finalizado) {
      try {
        mostrarAviso(
          Office.context.mailbox.item,
          "errorMessage",
          "error-timeout",
          "La operación ha tardado demasiado y se ha cancelado. Vuelve a intentarlo."
        );
      } catch (e) {
        // no-op
      }
      finalizar();
    }
  }, TIMEOUT_SEGURIDAD_MS);

  function finalizarYLimpiar() {
    clearTimeout(temporizador);
    finalizar();
  }

  try {
    var item = Office.context.mailbox.item;

    Promise.all([obtenerDestinatarios(item.to), obtenerDestinatarios(item.cc)])
      .then(function (listas) {
        try {
          var para = listas[0];
          var cc = listas[1];
          var todos = para.concat(cc).map(function (r) {
            return (r.emailAddress || "").toLowerCase();
          });
          var yaEstaba = todos.indexOf(CORREO_OBJETIVO.toLowerCase()) !== -1;

          if (yaEstaba) {
            insertarEnCuerpo(item, finalizarYLimpiar);
            return;
          }

          // No estaba entre los destinatarios: lo añadimos nosotros a "Para"
          // en vez de exigir que el comercial lo escriba a mano.
          item.to.addAsync([CORREO_OBJETIVO], function (addResult) {
            try {
              if (addResult.status === Office.AsyncResultStatus.Failed) {
                mostrarAviso(
                  item,
                  "errorMessage",
                  "error-add-destinatario",
                  "No se pudo añadir " + CORREO_OBJETIVO + " como destinatario: " + addResult.error.message
                );
                finalizarYLimpiar();
                return;
              }
              insertarEnCuerpo(item, finalizarYLimpiar);
            } catch (errorInterno) {
              mostrarAviso(item, "errorMessage", "error-interno-add", "Error al añadir el destinatario: " + errorInterno.message);
              finalizarYLimpiar();
            }
          });
        } catch (errorInterno) {
          mostrarAviso(item, "errorMessage", "error-interno", "Error al procesar destinatarios: " + errorInterno.message);
          finalizarYLimpiar();
        }
      })
      .catch(function (err) {
        try {
          mostrarAviso(
            item,
            "errorMessage",
            "error-destinatarios",
            "No se pudieron leer los destinatarios: " + (err && err.message ? err.message : err)
          );
        } catch (e) {
          // no-op
        }
        finalizarYLimpiar();
      });
  } catch (errorInesperado) {
    try {
      Office.context.mailbox.item.notificationMessages.replaceAsync("error-inesperado", {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: "Error inesperado al insertar la plantilla: " + errorInesperado.message
      });
    } catch (e) {
      // no-op
    }
    finalizarYLimpiar();
  }
}

function obtenerDestinatarios(campoDestinatarios) {
  return new Promise(function (resolve, reject) {
    try {
      campoDestinatarios.getAsync(function (result) {
        try {
          if (result.status === Office.AsyncResultStatus.Failed) {
            reject(result.error);
          } else {
            resolve(result.value || []);
          }
        } catch (e) {
          reject(e);
        }
      });
    } catch (e) {
      reject(e);
    }
  });
}

function insertarEnCuerpo(item, callbackFinal) {
  try {
    item.body.getTypeAsync(function (tipoResult) {
      try {
        if (tipoResult.status === Office.AsyncResultStatus.Failed) {
          mostrarAviso(item, "errorMessage", "error-tipo-cuerpo", "No se pudo determinar el formato del correo: " + tipoResult.error.message);
          callbackFinal();
          return;
        }

        var esHtml = tipoResult.value === Office.CoercionType.Html;
        var contenido = esHtml ? plantillaHtml() : plantillaTexto();
        var coercion = esHtml ? Office.CoercionType.Html : Office.CoercionType.Text;

        item.body.prependAsync(contenido, { coercionType: coercion }, function (insertResult) {
          try {
            if (insertResult.status === Office.AsyncResultStatus.Failed) {
              mostrarAviso(item, "errorMessage", "error-insercion", "No se pudo insertar la plantilla en el cuerpo: " + insertResult.error.message);
            } else {
              mostrarAviso(item, "informationalMessage", "plantilla-insertada", "Destinatario y plantilla de pedido listos.");
            }
          } catch (e) {
            // no-op, se finaliza igualmente abajo
          }
          callbackFinal();
        });
      } catch (errorInterno) {
        mostrarAviso(item, "errorMessage", "error-interno-cuerpo", "Error al insertar la plantilla: " + errorInterno.message);
        callbackFinal();
      }
    });
  } catch (errorSincrono) {
    mostrarAviso(item, "errorMessage", "error-sincrono", "Error al preparar la inserción: " + errorSincrono.message);
    callbackFinal();
  }
}

function mostrarAviso(item, tipo, id, mensaje) {
  try {
    var esError = tipo === "errorMessage";
    var tipoEnum = esError
      ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
      : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage;

    var opciones = { type: tipoEnum, message: mensaje };
    if (!esError) {
      opciones.icon = "icon-16";
      opciones.persistent = false;
    }

    item.notificationMessages.replaceAsync(id, opciones, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error("No se pudo mostrar la notificación:", result.error);
      }
    });
  } catch (e) {
    console.error("Error mostrando aviso:", e);
  }
}

function plantillaHtml() {
  return (
    '<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;margin-bottom:12px;">' +
    '<p style="margin:0;"><strong>NOMBRE CLIENTE:</strong>&nbsp;</p>' +
    '<p style="margin:12px 0 0 0;"><strong>PEDIDO DIRECTO S/N:</strong>&nbsp;</p>' +
    '<p style="margin:0;"><span style="font-weight:normal;font-style:italic;">(Si es S: se lanza sin ok de JP ; Si es N o no se pone nada sigue ciclo redmine)</span>&nbsp;</p>' +
    '<p style="margin:12px 0 0 0;"><strong>FECHA DE FACTURACIÓN APROXIMADA:</strong>&nbsp;</p>' +
    '<p style="margin:12px 0 0 0;"><strong>DIRECCIÓN DE ENTREGA:</strong>&nbsp;</p>' +
    '<p style="margin:0;"><span style="font-weight:normal;font-style:italic;">(si no se indica nada, la dirección de entrega será Santiago)</span>&nbsp;</p>' +
    '<p style="margin:12px 0 0 0;"><strong> </strong>&nbsp;</p>' +
    "</div>"
  );
}

function plantillaTexto() {
  return (
    "NOMBRE CLIENTE: \n" +
    "PEDIDO DIRECTO S/N: (Si es S: se lanza sin ok de JP ; Si es N o no se pone nada sigue ciclo redmine) \n" +
    "FECHA DE FACTURACIÓN APROXIMADA: \n" +
    "DIRECCIÓN DE ENTREGA: (si no se indica nada, la dirección de entrega será Santiago) \n" +
    "------------------------------------\n\n"
  );
}

/* Le confirmamos a Office.js que la página ha terminado de cargar.
   Sin esto, Office.js se queda esperando indefinidamente y el comando
   nunca llega a reconocerse como listo (causa del "colgado"). */
Office.onReady();

Office.actions.associate("insertarPlantilla", insertarPlantilla);
