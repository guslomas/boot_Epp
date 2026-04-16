/**
 * CONFIGURACIÓN GLOBAL
 */
const scriptProperties = PropertiesService.getScriptProperties();
const CONFIG = {
  TOKEN: scriptProperties.getProperty('apiKey'),
  URL_BASE: scriptProperties.getProperty('URLAPI_TELEGRAM'),
  WEBHOOK: scriptProperties.getProperty('URL_APPSCRIPT1'),
  get API_URL() { return `${this.URL_BASE}${this.TOKEN}`; }
};

const SS = SpreadsheetApp.getActiveSpreadsheet();
const sheet = SS.getActiveSheet();
const CONFIG_HOJAS = {
  LOG: SS.getActiveSheet(),
  PEDIDOS: SS.getSheetByName("pedidos") || SS.insertSheet("pedidos")
};

/**
 * MAPEADO DE ACCIONES (Global)
 * Centraliza los nombres de los productos para que sean fáciles de mantener.
 */
const ACCIONES_BOTONES = {
  "k_Subproductos": "K Subproductos",
  "K_Coque": "K Coque",
  "ver_datos": "VER_DATOS",
  "c_Candados": "C Candados",
  "epp": "EPP",
  "lentes": "Lentes",
  "guantes": "Guantes",
  "barbijos": "Barbijos",
  "p_audit": "P Audit",
  "D_Gases": "D Gases",
};

function doPost(e) {
  try {
    const contents = JSON.parse(e.postData.contents);
    const msg = contents.message;

    if (msg && msg.text) {
      const chatId = msg.chat.id;
      const text = msg.text;
      const name = msg.from.first_name;

      // --- NUEVO: DETECTOR DE COMANDOS ---
      if (text.startsWith("/")) {
        return handleCommand(chatId, text);
      }

      // --- DETECTAR RESPUESTA AL REGISTRO DE KITS ---
      if (msg.reply_to_message) {
        const textoBot = msg.reply_to_message.text;

        // Verificamos si es nuestro mensaje de registro
        if (textoBot.includes("Modo Registro de Kits")) {
          // EXTRAEMOS EL PRODUCTO: Buscamos lo que está entre corchetes [ ]
          const match = textoBot.match(/\[(.*?)\]/);
          const producto = match ? match[1] : "Desconocido";

          return gestionarRegistroKit(chatId, producto, name, text);
        }
      }

      // --- FLUJO NORMAL ---
      CONFIG_HOJAS.LOG.appendRow([new Date(), name, text, chatId]);
      sendMessage(chatId, `¡Bienvenido ${name}!`);
      enviarMenuPrincipal(chatId);
    }

    // --- MANEJO DE BOTONES ---
    if (contents.callback_query) {
      const cb = contents.callback_query;
      const data = cb.data;
      const chatId = cb.message.chat.id;
      const messageId = cb.message.message_id;
      const name = cb.from.first_name;

      // --- LÓGICA DE ELIMINACIÓN ---
      if (data.startsWith("borrar_")) {
        const filaABorrar = data.split("_")[1]; // Extraemos el número de fila
        eliminarPedido(chatId, filaABorrar, cb.message.message_id);
      }
      else if (data.startsWith("prepS_")) {
        const partes = data.split("_");
        const fila = partes[1];
        const nombreProd = partes[2];
        prepararCambioStatus(chatId, fila, nombreProd, messageId);

      }
      else if (data.startsWith("updDs_")) {
        const partes = data.split("_"); // updDs_FILA_STATUS
        const product=partes[3];
        ejecutarCambioStatus(chatId, partes[1], partes[2], product, messageId, name);
      }
      // Verificamos si la key existe en nuestro mapeo global
      if (data in ACCIONES_BOTONES) {
        if (data === "ver_datos") {
          sendMessage(chatId, "📊 Buscando tus últimos pedidos...");
          ejecutarDatos(chatId);
          // Aquí puedes agregar: ejecutarConsultaPedidos(chatId);
        } else if (data === "epp") {
          MenuEpp(chatId);
        } else {
          const nombreProducto = ACCIONES_BOTONES[data];
          solicitarNumeroKit(chatId, nombreProducto);
        }
      } /* else {
        sendMessage(chatId, "⚠️ Acción no reconocida.");
      } */
      answerCallback(cb.id, "Processing...");
    }
  } catch (err) {
    console.error("Error en doPost: " + err.toString());
  }

}
/**
 * Mapeo de acciones: Cada callback_data apunta a una función real.
 *
 * PASO 1: Solicitar cantidad.
 * Metemos el nombre del producto entre corchetes para recuperarlo luego.
 */
function solicitarNumeroKit(chatId, tipoProducto) {
  const payload = {
    chat_id: chatId,
    text: `💰 <b>Modo Registro de Kits</b>\nProducto: [${tipoProducto}]\n\nEscribe la cant que deseas agregar:`,
    parse_mode: "HTML",
    reply_markup: JSON.stringify({
      force_reply: true,
      selective: true,
      input_field_placeholder: "Ejemplo: 5"
    })
  };
  fetchTelegram("sendMessage", payload);
}

/**
 * PASO 2: Guardar en la hoja de Pedidos
 */
function gestionarRegistroKit(chatId, producto, name, cantidadTexto) {
  const cantidad = parseFloat(cantidadTexto.replace(',', '.'));
  const ped = "Pedido";
  if (isNaN(cantidad)) {
    return sendMessage(chatId, "❌ <b>Error:</b> Por favor envía un número válido.");
  }

  // Guardar: Fecha | Nombre | Cantidad | Producto | ID Chat
  CONFIG_HOJAS.PEDIDOS.appendRow([new Date(), name, cantidad, producto, chatId, "Pendiente"]);

  sendMessage(chatId, `✅ <b>Pedido Registrado</b>\nProducto: ${producto}\nCantidad: ${cantidad}`);
}
/**
 * INTERFAZ Y COMUNICACIONES
 */
function enviarMenuPrincipal(chatId) {
  const keyboard = {
    inline_keyboard: [
      [{ text: "💱 K Subproductos", callback_data: "k_Subproductos" }, { text: "⚫ K Coque", callback_data: "K_Coque" }, { text: "🎆 K Sold", callback_data: "K_Sold" }],
      [{ text: "🔐🧰 C Candados", callback_data: "c_Candados" }, { text: "♨ D Gases", callback_data: "D_Gases" }, { text: "⛑🚧 Epp", callback_data: "epp" }],
      [{ text: "📊 Mis Pedidos", callback_data: "ver_datos" }]
    ]
  };
  sendMessage(chatId, "<b>Menú Principal</b>\nSelecciona una opción:", keyboard);
}
// menu Epp
function MenuEpp(chatId) {
  const keyboard = {
    inline_keyboard: [
      [{ text: "🥽 Lentes", callback_data: "lentes" }, { text: "🥊 Guantes", callback_data: "guantes" }],
      [{ text: " 😷 Barbijos", callback_data: "barbijos" }, { text: "🎧 P Audit", callback_data: "p_audit" }],
      [{ text: "📊 Mis Pedidos", callback_data: "ver_datos" }]
    ]
  };
  sendMessage(chatId, "<b>Menú EPP</b>\nSelecciona una opción:", keyboard);
}

function sendMessage(chatId, text, keyboard = null) {
  const payload = { chat_id: chatId, text: text, parse_mode: "HTML" };
  if (keyboard) payload.reply_markup = JSON.stringify(keyboard);
  return fetchTelegram("sendMessage", payload);
}

function fetchTelegram(metodo, payload) {
  return UrlFetchApp.fetch(`${CONFIG.API_URL}/${metodo}`, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

function answerCallback(id) {
  return fetchTelegram("answerCallbackQuery", { callback_query_id: id });
}
/**
 * Busca datos específicos del usuario en la hoja activa.
 * @param {number|string} chatId - El ID del chat para filtrar.
 * @return {string} Un mensaje formateado con los datos encontrados.
 */
function ejecutarDatos(chatId) {

  try {
    const nombreHoja = CONFIG_HOJAS.LOG.getName();
    const data_Sheet = CONFIG_HOJAS.PEDIDOS.getDataRange().getValues(); // Leemos todo una sola vez (Eficiencia)
    // Ejemplo de lectura de la hoja activa

    sendMessage(chatId, `📊 Estás registrado en la hoja de sheet: <b>${nombreHoja}</b>`);
    // 2. OMITIR ENCABEZADOS: Filtramos desde la fila 2 en adelante
    // .slice(1) elimina el primer elemento (índice 0) del array
    const dataSinTitulos = data_Sheet.slice(1);

    // Buscamos la fila donde la columna D (índice 3) coincida con el chatId
    // Usamos map para conservar el número de fila (index + 1)
    const registrosConID = dataSinTitulos.map((fila, index) => ({ datos: fila, fila: index + 2 }));
    // Buscamos la fila donde la columna D (índice 3) coincida con el chatId
    const filaUsuario = registrosConID.filter(item => item.datos[4] == chatId);
    if (filaUsuario.length > 0) {
      let mensaje = `📊 <b>Tus Pedidos Registrados (${filaUsuario.length}):</b>\n\n`;
      sendMessage(chatId, `📂 <b>Tus últimos pedidos:</b>\n<i>Pulsa eliminar si hubo un error.</i>`);
      // 3. Recorremos los registros encontrados (mostramos los últimos 10 para no saturar Telegram)
      // Si quieres ver todos, quita el .slice()
      const ultimosRegistros = filaUsuario.reverse().slice(0, 10);
      ultimosRegistros.forEach((item, index) => {
        const fechaFormateada = Utilities.formatDate(new Date(item.datos[0]), "GMT-3", "dd/MM/yyyy HH:mm");
        const cantidad = item.datos;
        const producto = item.fila; // Este es nuestro ID de eliminación 
        // Construimos el contenido de la burbuja actual
        const burbuja = `📦 <b>Pedido #${producto}</b>\n` +
          `━━━━━━━━━━━━━━━━━━\n` +
          `🔹 <b>Product:</b> ${cantidad[3]} <b>Cant<b>🔢 ${cantidad[2]}</b>\n` +
          ` <b>Estado:</b>🚦${cantidad[5]}</b>\n` +
          `📅 <b>Fecha:</b> ${fechaFormateada}`;

        // Agregamos el botón de eliminar pasando el número de fila
        const tecladoEliminar = {
          inline_keyboard: [[
            { text: "❌ Eliminar este pedido", callback_data: `borrar_${producto}` }
          ]]
        };

        // ENVIAMOS UNA BURBUJA INDEPENDIENTE
        sendMessage(chatId, burbuja, tecladoEliminar);
      });
    } else {
      sendMessage(chatId, "❌ No encontré registros vinculados a tu ID.");
    }
  } catch (err) {
    sendMessage(chatId, "⚠️ Error al acceder a la base de datos.");
  }
}
function eliminarPedido(chatId, numFila, messageId) {
  try {
    const fila = parseInt(numFila);

    // 1. Borrar la fila en la hoja de Google Sheets
    CONFIG_HOJAS.PEDIDOS.deleteRow(fila);

    // 2. Editar el mensaje en Telegram para que el usuario vea que ya no existe
    const payload = {
      chat_id: chatId,
      message_id: messageId,
      text: "🗑️ <b>Este pedido ha sido eliminado de la base de datos.</b>",
      parse_mode: "HTML"
    };

    fetchTelegram("editMessageText", payload);
  } catch (e) {
    sendMessage(chatId, "❌ No se pudo eliminar. Es posible que la lista se haya actualizado.");
  }
}
/**
 * Gestiona los comandos que empiezan con /
 */
function handleCommand(chatId, text) {
  const parts = text.split(" ");
  const command = parts[0].toLowerCase();
  const args = parts.slice(1);

  switch (command) {
    case "/start":
      enviarMenuPrincipal(chatId);
      break;
    case "/find":
      if (args.length === 0) {
        sendMessage(chatId, "⚠️ Uso: <code>/find [nombre]</code>\nEjemplo: <code>/find Juan</code>");
      } else {
        handleFindCommand(chatId, args.join(" "));
      }
      break;
    case "/help":
      sendMessage(chatId, "📖 <b>Comandos disponibles:</b>\n/start - Menú principal\n/find [nombre] - Busca pedidos por nombre.");
      break;
    default:
      sendMessage(chatId, "❌ Comando no reconocido.");
  }
}

/**
 * Busca pedidos por nombre en la hoja PEDIDOS
 */
function handleFindCommand(chatId, nombreBusqueda) {
  try {
    const dataSheet = CONFIG_HOJAS.PEDIDOS.getDataRange().getValues();
    const dataSinTitulos = dataSheet.slice(1);

    // Filtramos: Columna B (índice 1) es el Nombre. 
    // Usamos toLowerCase() para que no importe si escriben en mayúsculas o minúsculas.
    const resultados = dataSinTitulos
      .map((fila, index) => ({ datos: fila, filaOriginal: index + 2 }))
      .filter(item => item.datos[1].toString().toLowerCase().includes(nombreBusqueda.toLowerCase()));

    if (resultados.length > 0) {
      sendMessage(chatId, `🔍 <b>Resultados para "${nombreBusqueda}" (${resultados.length}):</b>`);

      // Limitamos a los últimos 5 para no saturar
      resultados.reverse().slice(0, 5).forEach(item => {
        const productoActual = item.datos[3]; // Columna D
        const filaActual = item.filaOriginal;
        const fechaForm = Utilities.formatDate(new Date(item.datos[0]), "GMT-3", "dd/MM/yyyy HH:mm");
        const burbuja = `👤 <b>Usuario:</b> ${item.datos[1]}\n` +
          `📦 <b>Pedido:</b> ${item.datos[3]}\n` +
          `🔢 <b>Cant:</b> ${item.datos[2]}\n` +
          `🚦 <b>Estado actual:</b> ${item.datos[5]}\n` +
          `📅 <b>Fecha:</b> ${fechaForm}`;

        // Botón para cambiar status pasando la fila
        const teclado = {
          inline_keyboard: [[
            { text: "🔄 Cambiar Status", callback_data: `prepS_${filaActual}_${productoActual.substring(0, 15)}` }
          ]]
        };

        sendMessage(chatId, burbuja, teclado);
      });
    } else {
      sendMessage(chatId, `❌ No se encontraron pedidos para: <b>${nombreBusqueda}</b>`);
    }
  } catch (e) {
    sendMessage(chatId, "⚠️ Error al realizar la búsqueda.");
  }
}
/**
 * Paso 1: Muestra las opciones de status al presionar el botón
 */
function prepararCambioStatus(chatId, numFila, productoNombre, messageId) {
  const teclado = {
    inline_keyboard: [
      [
        { text: "🟢 Entregado", callback_data: `updDs_${numFila}_Entregado_${productoNombre}` },
        { text: "🟡 Pendiente", callback_data: `updDs_${numFila}_Pendiente_${productoNombre}` }
      ],
      [{ text: "🔴 Cancelado", callback_data: `updDs_${numFila}_Cancelado_${productoNombre}` }]
    ]
  };

  const payload = {
    chat_id: chatId,
    message_id: messageId,
    text: `⚠️ <b>Selecciona el nuevo estado para:</b>\n📦 <code>${productoNombre}</code> (Fila ${numFila})`,
    parse_mode: "HTML",
    reply_markup: JSON.stringify(teclado)
  };
  fetchTelegram("editMessageText", payload);
}
/**
 * Paso 2: Realiza el cambio físico en la celda de la hoja
 */
function ejecutarCambioStatus(chatId, numFila, nuevoStatus, productoNombre, messageId, name) {
  try {
    const fila = parseInt(numFila);
    // Columna F es índice 5 (Estado)
    CONFIG_HOJAS.PEDIDOS.getRange(fila, 6).setValue(nuevoStatus);
    CONFIG_HOJAS.PEDIDOS.getRange(fila, 7).setValue(name);

    const payload = {
      chat_id: chatId,
      message_id: messageId,
      text: `✅ <b>Status Actualizado</b>\nProducto: <b>${productoNombre}</b>\nFila: ${fila}\nEstado: ${nuevoStatus}`,
      parse_mode: "HTML"
    };
    fetchTelegram("editMessageText", payload);
  } catch (e) {
    sendMessage(chatId, "❌ Error al actualizar el status.");
  }
}