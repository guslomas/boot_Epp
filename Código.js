/**
 * CONFIGURACIÓN GLOBAL
 */
const scriptProperties = PropertiesService.getScriptProperties();
const CONFIG = {
  /* TOKEN: scriptProperties.getProperty('apiKey'),
  URL_BASE: scriptProperties.getProperty('URLAPI_TELEGRAM'),
  WEBHOOK: scriptProperties.getProperty('URL_APPSCRIPT1'),
  get API_URL() { return `${this.URL_BASE}${this.TOKEN}`; } */


  TELEGRAM: {
    TOKEN: scriptProperties.getProperty('apiKey'),
    URL_BASE: scriptProperties.getProperty('URLAPI_TELEGRAM'),
    get API_URL() { return `${this.URL_BASE}${this.TOKEN}`; }
  },
  COLUMNAS: {
    FECHA: 1, NOMBRE: 2, CANTIDAD: 3, PRODUCTO: 4, CHAT_ID: 5, ESTADO: 6, RESPONSABLE: 7
  }
};
/** * Acceso a Hojas con inicialización perezosa (Lazy Loading)
 */
const getSheets = () => {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  /* const sheet = SS.getActiveSheet();
  const CONFIG_HOJAS = {*/
  return {
    LOG: SS.getActiveSheet(),
    PEDIDOS: SS.getSheetByName("pedidos") || SS.insertSheet("pedidos")
  };
};
/**
 * MAPEADO DE ACCIONES (Global)
 * Centraliza los nombres de los productos para que sean fáciles de mantener.
 */
const ACCIONES_BOTONES = {
  "k_Subproductos": "K Subproductos",
  "K_Sold": "K Sold",
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


    // 1. MANEJO DE MENSAJES DE TEXTO
    if (contents.message) {
      const msg = contents.message;
      const chatId = msg.chat.id;
      const text = msg.text || "";
      const name = msg.from.first_name || "Usuario";


      if (text.startsWith("/")) return handleCommand(chatId, text);


      // Respuesta al ForceReply
      if (msg.reply_to_message && msg.reply_to_message.text.includes("Modo Registro de Kits")) {
        const match = msg.reply_to_message.text.match(/\[(.*?)\]/);
        const producto = match ? match[1] : "Desconocido";
        return gestionarRegistroKit(chatId, producto, name, text);
      }


      // Registro en Log y Menú Principal
      getSheets().LOG.appendRow([new Date(), name, text, chatId]);
      TelegramService.sendMessage(chatId, `¡Bienvenido <b>${name}</b>!`);
      enviarMenuPrincipal(chatId);
    }


    // 2. MANEJO DE CALLBACK QUERY (BOTONES)
    if (contents.callback_query) {
      const cb = contents.callback_query;
      const data = cb.data;
      const chatId = cb.message.chat.id;
      const messageId = cb.message.message_id;
      const name = cb.from.first_name;


      const userId = cb.from.id; // ID numérico de Telegram de quien pulsa el botón


      // Lógica por prefijos
      if (data.startsWith("borrar_")) {
        eliminarPedido(chatId, data.split("_")[1], messageId);
      }
      else if (data.startsWith("prepS_")) {
        const [, fila, prod, user, originalOwnerId] = data.split("_");


        // REGLA DE NEGOCIO: No puede cambiar el estado el mismo usuario que pidió
        if (userId.toString() === originalOwnerId) {
          return TelegramService.answerCallback(cb.id, "⚠️ No puedes procesar tu propio pedido.", true);
        } else {
          prepararCambioStatus(chatId, fila, prod, messageId, user);
        }

      }
      else if (data.startsWith("updDs_")) {
        const [, fila, status, prod] = data.split("_");
        ejecutarCambioStatus(chatId, fila, status, prod, messageId, name);
      }


      // Lógica de mapeo ACCIONES_BOTONES
      else if (data in ACCIONES_BOTONES) {
        if (data === "ver_datos") {
          TelegramService.sendMessage(chatId, "📊 Buscando tus últimos pedidos...");
          ejecutarDatos(chatId);
        } else if (data === "epp") {
          MenuEpp(chatId);
        } else {
          // Esta es la parte que solicita la cantidad
          solicitarNumeroKit(chatId, ACCIONES_BOTONES[data]);
        }
      }


      TelegramService.answerCallback(cb.id, "Procesando...");
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
    /* reply_markup: JSON.stringify({ */
    reply_markup: ({
      force_reply: true,
      selective: true,
      input_field_placeholder: "Ejemplo: 5"
    })
  };
  TelegramService.fetch("sendMessage", payload);
}


/**
 * PASO 2: Guardar en la hoja de Pedidos
 */
function gestionarRegistroKit(chatId, producto, name, cantidadTexto) {
  const cantidad = parseFloat(cantidadTexto.replace(',', '.'));
  const ped = "Pedido";
  if (isNaN(cantidad)) {
    return TelegramService.sendMessage(chatId, "❌ <b>Error:</b> Por favor envía un número válido.");
  }


  // Guardar: Fecha | Nombre | Cantidad | Producto | ID Chat
  getSheets().PEDIDOS.appendRow([new Date(), name, cantidad, producto, chatId, "Pendiente"]);


  TelegramService.sendMessage(chatId, `✅ <b>Pedido Registrado</b>\nProducto: ${producto}\nCantidad: ${cantidad}`);
}
/**
 * INTERFAZ Y COMUNICACIONES
 */
function enviarMenuPrincipal(chatId) {
  const keyboard = {
    inline_keyboard: [
      [{ text: "🛠HERRAMIENTAS", callback_data: "herramientas" }],
      [{ text: "💱 K Subproductos", callback_data: "k_Subproductos" }, { text: "⚫ K Coque", callback_data: "K_Coque" }, { text: "🎆 K Sold", callback_data: "K_Sold" }],
      [{ text: "🔐🧰 C Candados", callback_data: "c_Candados" }, { text: "♨ D Gases", callback_data: "D_Gases" }, { text: "⛑🚧 Epp", callback_data: "epp" }],
      [{ text: "📊 Mis Pedidos", callback_data: "ver_datos" }]
    ]
  };
  TelegramService.sendMessage(chatId, "<b>Menú Principal</b>\nSelecciona una opción:", keyboard);
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
  TelegramService.sendMessage(chatId, "<b>Menú EPP</b>\nSelecciona una opción:", keyboard);
}


/* function sendMessage(chatId, text, keyboard = null) {
  const payload = { chat_id: chatId, text: text, parse_mode: "HTML", reply_markup: keyboard ? JSON.stringify(keyboard) : undefined };
  if (keyboard) payload.reply_markup = JSON.stringify(keyboard);
  return fetchTelegram("sendMessage", payload);
} */


/* function fetchTelegram(metodo, payload) {
  return UrlFetchApp.fetch(`${CONFIG.API_URL}/${metodo}`, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
} */


/* function answerCallback(id) {
  return fetchTelegram("answerCallbackQuery", { callback_query_id: id });
} */
/**
 * Busca datos específicos del usuario en la hoja activa.
 * @param {number|string} chatId - El ID del chat para filtrar.
 * @return {string} Un mensaje formateado con los datos encontrados.
 */
function ejecutarDatos(chatId) {

  try {
    const nombreHoja = getSheets().LOG.getName();
    const data_Sheet = getSheets().PEDIDOS.getDataRange().getValues(); // Leemos todo una sola vez (Eficiencia)
    // Ejemplo de lectura de la hoja activa
    TelegramService.sendMessage(chatId, `📊 Estás registrado en la hoja de sheet: <b>${nombreHoja}</b>`);
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
      TelegramService.sendMessage(chatId, `📂 <b>Tus últimos pedidos:</b>\n<i>Pulsa eliminar si hubo un error.</i>`);
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
        let tecladoEliminar = null;
        if (cantidad[5].toString().trim().toLowerCase() === "pendiente") {
          tecladoEliminar = {
            inline_keyboard: [[
              { text: "❌ Eliminar Pedido", callback_data: `borrar_${producto}` }
            ]]
          };
        };
        /* const tecladoEliminar = {
          inline_keyboard: [[
            { text: "❌ Eliminar este pedido", callback_data: `borrar_${producto}` }
          ]]
        }; */


        // ENVIAMOS UNA BURBUJA INDEPENDIENTE
        TelegramService.sendMessage(chatId, burbuja, tecladoEliminar);
      });
    } else {
      TelegramService.sendMessage(chatId, "❌ No encontré registros vinculados a tu ID.");
    }
  } catch (err) {
    TelegramService.sendMessage(chatId, "⚠️ Error al acceder a la base de datos.");
  }
}
function eliminarPedido(chatId, numFila, messageId) {
  try {
    const fila = parseInt(numFila);


    // 1. Borrar la fila en la hoja de Google Sheets
    getSheets().PEDIDOS.deleteRow(fila);


    // 2. Editar el mensaje en Telegram para que el usuario vea que ya no existe
    const payload = {
      chat_id: chatId,
      message_id: messageId,
      text: "🗑️ <b>Este pedido ha sido eliminado de la base de datos.</b>",
      parse_mode: "HTML"
    };


    TelegramService.fetch("editMessageText", payload);
  } catch (e) {
    TelegramService.sendMessage(chatId, "❌ No se pudo eliminar. Es posible que la lista se haya actualizado.");
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
        TelegramService.sendMessage(chatId, "⚠️ Uso: <code>/find [nombre]</code>\nEjemplo: <code>/find Juan</code>");
      } else {
        handleFindCommand(chatId, args.join(" "));
      }
      break;
    case "/help":
      TelegramService.sendMessage(chatId, "📖 <b>Comandos disponibles:</b>\n/start - Menú principal\n/find [nombre] - Busca pedidos por nombre.");
      break;
    default:
      TelegramService.sendMessage(chatId, "❌ Comando no reconocido.");
  }
}


/**
 * Busca pedidos por nombre en la hoja PEDIDOS
 */
function handleFindCommand(chatId, nombreBusqueda) {
  try {
    const dataSheet = getSheets().PEDIDOS.getDataRange().getValues();
    const dataSinTitulos = dataSheet.slice(1);


    // Filtramos: Columna B (índice 1) es el Nombre.
    // Usamos toLowerCase() para que no importe si escriben en mayúsculas o minúsculas.
    const resultados = dataSinTitulos
      .map((fila, index) => ({ datos: fila, filaOriginal: index + 2 }))
      .filter(item => item.datos[1].toString().toLowerCase().includes(nombreBusqueda.toLowerCase()));


    if (resultados.length > 0) {
      TelegramService.sendMessage(chatId, `🔍 <b>Resultados para "${nombreBusqueda}" (${resultados.length}):</b>`);


      // Limitamos a los últimos 5 para no saturar
      resultados.reverse().slice(0, 5).forEach(item => {
        const productoActual = item.datos[3]; // Columna D
        const usuarioActual = item.datos[1];
        const ownerChatId = item.datos[4];
        const filaActual = item.filaOriginal;
        const fechaForm = Utilities.formatDate(new Date(item.datos[0]), "GMT-3", "dd/MM/yyyy HH:mm");
        const burbuja = `👤 <b>Usuario:</b> ${usuarioActual}\n` +
          `📦 <b>Pedido:</b> ${item.datos[3]}\n` +
          `🔢 <b>Cant:</b> ${item.datos[2]}\n` +
          `🚦 <b>Estado actual:</b> ${item.datos[5]}\n` +
          `📅 <b>Fecha:</b> ${fechaForm}`;


        // Botón para cambiar status pasando la fila
        const teclado = {
          inline_keyboard: [[
            {
              text: "🔄 Cambiar Status",
              callback_data: `prepS_${filaActual}_${productoActual.substring(0, 15)}_${usuarioActual}_${ownerChatId}`
            }
          ]]
        };


        TelegramService.sendMessage(chatId, burbuja, teclado);
      });
    } else {
      TelegramService.sendMessage(chatId, `❌ No se encontraron pedidos para: <b>${nombreBusqueda}</b>`);
    }
  } catch (e) {
    TelegramService.sendMessage(chatId, "⚠️ Error al realizar la búsqueda.");
  }
}
/**
 * Paso 1: Muestra las opciones de status al presionar el botón
 */
function prepararCambioStatus(chatId, numFila, productoNombre, messageId, name) {
  const teclado = {
    inline_keyboard: [
      [
        { text: "🟢 Entregado", callback_data: `updDs_${numFila}_Entregado_${productoNombre}` },
        { text: "🟡 Proceso", callback_data: `updDs_${numFila}_Proceso_${productoNombre}` }
      ],
      [{ text: "🔴 Cancelado", callback_data: `updDs_${numFila}_Cancelado_${productoNombre}` }]
    ]
  };


  const payload = {
    chat_id: chatId,
    message_id: messageId,
    text: `⚠️ <b>Cambio de Estado:</b>${name ? `Usuario: ${name}` : ""}\n📦 <code>${productoNombre}</code> (Fila ${numFila})`,
    parse_mode: "HTML",
    reply_markup: JSON.stringify(teclado)
  };
  TelegramService.fetch("editMessageText", payload);
}
/**
 * Paso 2: Realiza el cambio físico en la celda de la hoja
 */
function ejecutarCambioStatus(chatId, numFila, nuevoStatus, productoNombre, messageId, name) {
  try {
    const fila = parseInt(numFila);
    // Columna F es índice 5 (Estado)
    getSheets().PEDIDOS.getRange(fila, 6).setValue(nuevoStatus);
    getSheets().PEDIDOS.getRange(fila, 7).setValue(name);


    const payload = {
      chat_id: chatId,
      message_id: messageId,
      text: `✅ <b>Status Actualizado</b>\nProducto: <b>${productoNombre}</b>\nFila: ${fila}\nEstado: ${nuevoStatus}`,
      parse_mode: "HTML"
    };
    TelegramService.fetch("editMessageText", payload);
  } catch (e) {
    TelegramService.sendMessage(chatId, "❌ Error al actualizar el status.");
  }
}




/**
 * SERVICIO DE COMUNICACIÓN CON TELEGRAM (Optimizado)
 */
const TelegramService = {
  fetch: function (method, payload) {
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    return UrlFetchApp.fetch(`${CONFIG.TELEGRAM.API_URL}/${method}`, options);
  },


  sendMessage: function (chatId, text, keyboard = null) {
    const payload = {
      chat_id: chatId,
      text: text,
      parse_mode: "HTML"
    };
    if (keyboard) payload.reply_markup = keyboard; // Se pasa el objeto, fetch lo stringifica
    return this.fetch("sendMessage", payload);
  },


  answerCallback: function (callbackQueryId, text = "") {
    return this.fetch("answerCallbackQuery", {
      callback_query_id: callbackQueryId,
      text: text
    });
  },


  editMessageText: function (payload) {
    return this.fetch("editMessageText", payload);
  }
};

