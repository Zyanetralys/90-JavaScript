/**
 * ============================================================
 * BUSCADOR Y GENERADOR DE BORRADORES DE ACCESO
 * ============================================================
 *
 * Busca solicitudes de acceso relacionadas con un usuario
 * espejo y genera borradores de respuesta.
 *
 * Criterios para considerar una solicitud válida:
 *
 * 1. El mensaje debe proceder del remitente configurado.
 * 2. El mismo mensaje debe contener:
 *      - Nombre y apellidos del usuario espejo.
 *      - Identificador del usuario.
 *      - Correo electrónico del usuario espejo.
 * 3. Los tres datos deben aparecer dentro del mismo mensaje.
 * 4. Los mensajes reenviados por terceros no se consideran
 *    solicitudes independientes.
 *
 * Por cada aparición completa de los tres identificadores
 * se crea un borrador Reply All.
 *
 * IMPORTANTE:
 * - El script NO envía correos.
 * - Solo crea borradores.
 * - Toda la configuración editable está al principio.
 * - Sustituye los valores de ejemplo por los correspondientes
 *   antes de utilizar el script.
 * ============================================================
 */


/* ============================================================
 * CONFIGURACIÓN EDITABLE
 * ============================================================
 *
 * Esta es la única sección que debería ser necesario modificar
 * para adaptar el script a otro entorno.
 */


/* ------------------------------------------------------------
 * REMITENTE AUTORIZADO
 * ------------------------------------------------------------
 *
 * Solo los mensajes cuyo remitente real coincida exactamente
 * con esta dirección serán considerados solicitudes válidas.
 */

const REMITENTE_SOLICITUDES = "REMITENTE_AUTORIZADO@DOMINIO.COM";


/* ------------------------------------------------------------
 * USUARIO ESPEJO
 * ------------------------------------------------------------ */

const USUARIO_ESPEJO = {

  // Identificador interno del usuario.
  usuario: "USUARIO_ESPEJO_ID",

  // Nombre y apellidos completos.
  nombre: "NOMBRE APELLIDOS",

  // Dirección de correo del usuario.
  email: "usuario.espejo@dominio.com"
};


/* ------------------------------------------------------------
 * USUARIO NUEVO
 * ------------------------------------------------------------ */

const USUARIO_NUEVO = {

  // Identificador del nuevo usuario.
  usuario: "USUARIO_NUEVO_ID",

  // Nombre y apellidos del nuevo usuario.
  nombre: "NUEVO USUARIO",

  // Dirección de correo del nuevo usuario.
  email: "nuevo.usuario@dominio.com"
};


/* ------------------------------------------------------------
 * USUARIO QUE EJECUTA EL SCRIPT
 * ------------------------------------------------------------
 *
 * Se utiliza para excluir mensajes enviados por la propia
 * cuenta cuando sea necesario.
 */

const USUARIO_PROPIO = {

  email: "USUARIO_PROPIO@DOMINIO.COM"
};


/* ------------------------------------------------------------
 * PRODUCTOS
 * ------------------------------------------------------------
 *
 * Lista completamente editable.
 *
 * Añade, elimina o modifica elementos según las aplicaciones,
 * servicios o productos que deban identificarse.
 */

const PRODUCTOS = [

  "PRODUCTO_01",
  "PRODUCTO_02",
  "PRODUCTO_03",
  "PRODUCTO_04",
  "PRODUCTO_05",
  "PRODUCTO_06",
  "PRODUCTO_07",
  "PRODUCTO_08",
  "PRODUCTO_09",
  "PRODUCTO_10",
  "PRODUCTO_11",
  "PRODUCTO_12",
  "PRODUCTO_13",
  "PRODUCTO_14",
  "PRODUCTO_15",
  "PRODUCTO_16",
  "PRODUCTO_17",
  "PRODUCTO_18",
  "PRODUCTO_19",
  "PRODUCTO_20"
];


/* ------------------------------------------------------------
 * MENSAJE ESTÁNDAR
 * ------------------------------------------------------------
 *
 * Esta plantilla puede modificarse completamente.
 *
 * Las variables se sustituyen automáticamente utilizando
 * los valores definidos en la configuración superior.
 */

const MENSAJE_ESTANDAR = `Buenos días,

Se solicita el acceso para un nuevo usuario.

El usuario espejo sería:

Identificador: ${USUARIO_ESPEJO.usuario}
Nombre: ${USUARIO_ESPEJO.nombre}
Correo: ${USUARIO_ESPEJO.email}

El usuario a dar de alta sería:

Identificador: ${USUARIO_NUEVO.usuario}
Nombre: ${USUARIO_NUEVO.nombre}
Correo: ${USUARIO_NUEVO.email}

Muchas gracias,
Saludos`;


/* ------------------------------------------------------------
 * MENSAJE PERSONALIZADO
 * ------------------------------------------------------------
 */

const MENSAJE_PERSONALIZADO = `Buenos días,

Solicitamos acceso para el nuevo usuario.

Usuario espejo: ${USUARIO_ESPEJO.nombre} (${USUARIO_ESPEJO.email})
Usuario nuevo: ${USUARIO_NUEVO.nombre} (${USUARIO_NUEVO.email})

Gracias.`;


/* ------------------------------------------------------------
 * CONFIGURACIÓN GENERAL
 * ------------------------------------------------------------ */

const CONFIG = {

  // Número máximo de hilos que se recuperarán en la búsqueda.
  maxResultados: 100,

  // Número máximo de días hacia atrás que se revisarán.
  diasAntiguedad: 730,

  // Categoría utilizada cuando no se detecta ningún producto.
  productoPorDefecto: "General",

  // Pausa entre operaciones de creación de borradores.
  pausaEntreBorradores: 300,

  /*
   * Tamaño de la ventana utilizada para determinar si las tres
   * identificaciones pertenecen a la misma solicitud.
   *
   * Se puede aumentar si los correos utilizan bloques más
   * extensos.
   */
  ventanaAntes: 500,
  ventanaDespues: 1000
};


/* ============================================================
 * FUNCIÓN PRINCIPAL
 * ============================================================
 */

function crearBorradoresRespuesta() {

  const ui = SpreadsheetApp.getUi();


  const respuesta = ui.prompt(

    "SELECCIÓN DE MENSAJE",

    "¿Qué mensaje quieres utilizar?\n\n" +
    "1 = Mensaje estándar\n" +
    "2 = Mensaje personalizado\n\n" +
    "Escribe 1 o 2:",

    ui.ButtonSet.OK_CANCEL
  );


  if (
    respuesta.getSelectedButton() ===
    ui.Button.CANCEL
  ) {

    Logger.log("Proceso cancelado.");

    return;
  }


  const opcion =
    respuesta.getResponseText().trim();


  if (opcion === "1") {

    procesarSolicitudes(
      MENSAJE_ESTANDAR,
      "ESTÁNDAR"
    );

    return;
  }


  if (opcion === "2") {

    procesarSolicitudes(
      MENSAJE_PERSONALIZADO,
      "PERSONALIZADO"
    );

    return;
  }


  ui.alert(
    "Opción no válida. Debes introducir 1 o 2."
  );
}


/* ============================================================
 * EJECUCIONES DIRECTAS
 * ============================================================
 */

function crearBorradores_MensajeEstandar() {

  procesarSolicitudes(
    MENSAJE_ESTANDAR,
    "ESTÁNDAR"
  );
}


function crearBorradores_MensajePersonalizado() {

  procesarSolicitudes(
    MENSAJE_PERSONALIZADO,
    "PERSONALIZADO"
  );
}


/* ============================================================
 * PROCESAMIENTO PRINCIPAL
 * ============================================================
 */

function procesarSolicitudes(
  mensaje,
  tipoMensaje
) {

  const spreadsheet =
    SpreadsheetApp.create(
      "Borradores (" +
      tipoMensaje +
      ") - " +
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        "yyyy-MM-dd HH-mm"
      )
    );


  const sheet =
    spreadsheet.getSheets()[0];


  /* ----------------------------------------------------------
   * CABECERA
   * ---------------------------------------------------------- */

  sheet
    .getRange(1, 1, 1, 8)
    .setValues([[
      "Producto(s)",
      "Asunto",
      "Remitente",
      "Fecha",
      "Aparición",
      "Hilo",
      "Borrador",
      "Estado"
    ]]);


  sheet
    .getRange(1, 1, 1, 8)
    .setFontWeight("bold")
    .setBackground("#1a1a2e")
    .setFontColor("#ffffff");


  let fila = 2;

  let totalBorradores = 0;

  let totalSolicitudes = 0;

  let totalMensajesValidos = 0;


  /* ----------------------------------------------------------
   * FECHA MÍNIMA
   * ---------------------------------------------------------- */

  const fechaLimite =
    new Date();


  fechaLimite.setDate(
    fechaLimite.getDate() -
    CONFIG.diasAntiguedad
  );


  const fechaStr =
    Utilities.formatDate(
      fechaLimite,
      Session.getScriptTimeZone(),
      "yyyy/MM/dd"
    );


  /* ----------------------------------------------------------
   * BÚSQUEDA INICIAL
   * ----------------------------------------------------------
   *
   * La búsqueda recupera posibles candidatos.
   *
   * La validación definitiva se realiza posteriormente sobre
   * cada mensaje individual.
   */

  const query =
    `"${USUARIO_ESPEJO.usuario}" ` +
    `OR "${USUARIO_ESPEJO.nombre}" ` +
    `OR "${USUARIO_ESPEJO.email}" ` +
    `after:${fechaStr}`;


  Logger.log(
    "Búsqueda inicial: " +
    query
  );


  let threads;


  try {

    threads =
      GmailApp.search(
        query,
        0,
        CONFIG.maxResultados
      );

  } catch (error) {

    Logger.log(
      "Error realizando la búsqueda: " +
      error.message
    );


    SpreadsheetApp
      .getUi()
      .alert(
        "Error realizando la búsqueda:\n\n" +
        error.message
      );


    return;
  }


  Logger.log(
    "Hilos candidatos: " +
    threads.length
  );


  /* ----------------------------------------------------------
   * RECORRER HILOS
   * ---------------------------------------------------------- */

  threads.forEach(thread => {

    const mensajes =
      thread.getMessages();


    if (
      !mensajes ||
      mensajes.length === 0
    ) {

      return;
    }


    /*
     * Se revisan individualmente todos los mensajes del hilo.
     *
     * No se utiliza el contenido completo del hilo para decidir
     * si existe una solicitud.
     */

    mensajes.forEach(message => {

      const remitente =
        obtenerDireccionRemitente(
          message.getFrom()
        );


      /* ------------------------------------------------------
       * VALIDAR REMITENTE
       * ------------------------------------------------------
       *
       * Solo se aceptan mensajes enviados directamente por
       * el remitente configurado.
       *
       * Un mensaje reenviado por otra persona no cumple esta
       * condición.
       */

      if (
        remitente.toLowerCase() !==
        REMITENTE_SOLICITUDES.toLowerCase()
      ) {

        return;
      }


      totalMensajesValidos++;


      const asunto =
        thread.getFirstMessageSubject();


      const fecha =
        message.getDate();


      const cuerpo =
        message.getPlainBody() || "";


      /* ------------------------------------------------------
       * DETECTAR SOLICITUDES
       * ------------------------------------------------------
       */

      const solicitudes =
        contarSolicitudes(
          cuerpo
        );


      if (
        solicitudes.length === 0
      ) {

        return;
      }


      /* ------------------------------------------------------
       * DETECTAR PRODUCTOS
       * ------------------------------------------------------ */

      const textoBusqueda =
        (
          asunto +
          "\n" +
          cuerpo
        ).toLowerCase();


      const productosDetectados =
        PRODUCTOS.filter(
          producto =>
            textoBusqueda.includes(
              producto.toLowerCase()
            )
        );


      if (
        productosDetectados.length === 0
      ) {

        productosDetectados.push(
          CONFIG.productoPorDefecto
        );
      }


      /* ------------------------------------------------------
       * CREAR UN BORRADOR POR APARICIÓN
       * ------------------------------------------------------
       *
       * Una aparición válida significa:
       *
       *     Nombre + Identificador + Correo
       *
       * encontrados dentro de la misma zona del mensaje.
       *
       * Si hay dos solicitudes independientes en un mismo
       * mensaje, se crean dos borradores sobre el mismo hilo.
       */

      solicitudes.forEach(
        (solicitud, posicion) => {

          let borrador;


          try {

            borrador =
              GmailApp.createDraftReplyAll(
                thread.getId(),
                {
                  htmlBody:
                    mensaje.replace(
                      /\n/g,
                      "<br>"
                    )
                }
              );


          } catch (error) {

            Logger.log(
              "Error creando borrador: " +
              error.message
            );


            sheet
              .getRange(
                fila,
                1,
                1,
                8
              )
              .setValues([[
                productosDetectados.join(", "),
                asunto,
                remitente,
                fecha,
                posicion + 1,
                thread.getId(),
                "",
                "ERROR: " +
                  error.message
              ]]);


            fila++;

            return;
          }


          const urlBorrador =
            "https://mail.google.com/mail/u/0/#drafts/" +
            borrador.getId();


          sheet
            .getRange(
              fila,
              1,
              1,
              8
            )
            .setValues([[
              productosDetectados.join(", "),
              asunto,
              remitente,
              fecha,
              posicion + 1,
              thread.getId(),
              urlBorrador,
              "CREADO"
            ]]);


          fila++;

          totalBorradores++;

          totalSolicitudes++;


          if (
            CONFIG.pausaEntreBorradores > 0
          ) {

            Utilities.sleep(
              CONFIG.pausaEntreBorradores
            );
          }
        }
      );
    });
  });


  /* ----------------------------------------------------------
   * FORMATO DE LA HOJA
   * ---------------------------------------------------------- */

  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(2, 350);
  sheet.setColumnWidth(3, 280);
  sheet.setColumnWidth(4, 160);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 250);
  sheet.setColumnWidth(7, 450);
  sheet.setColumnWidth(8, 250);


  /* ----------------------------------------------------------
   * RESUMEN
   * ---------------------------------------------------------- */

  const filaResumen =
    fila + 2;


  sheet
    .getRange(
      filaResumen,
      1,
      3,
      2
    )
    .setValues([
      [
        "MENSAJES DEL REMITENTE VÁLIDO",
        totalMensajesValidos
      ],
      [
        "SOLICITUDES DETECTADAS",
        totalSolicitudes
      ],
      [
        "BORRADORES CREADOS",
        totalBorradores
      ]
    ])
    .setFontWeight("bold");


  /* ----------------------------------------------------------
   * RESULTADO
   * ---------------------------------------------------------- */

  SpreadsheetApp
    .getUi()
    .alert(
      "PROCESO COMPLETADO\n\n" +

      "Mensajes del remitente válido: " +
      totalMensajesValidos +

      "\n\nSolicitudes detectadas: " +
      totalSolicitudes +

      "\n\nBorradores creados: " +
      totalBorradores +

      "\n\nNo se ha enviado ningún correo."
    );


  Logger.log(
    "Proceso finalizado."
  );


  Logger.log(
    "Borradores creados: " +
    totalBorradores
  );


  Logger.log(
    "Hoja de registro: " +
    spreadsheet.getUrl()
  );
}


/* ============================================================
 * DETECTAR SOLICITUDES
 * ============================================================
 *
 * Busca cada aparición del nombre del usuario y comprueba
 * si el identificador y el correo aparecen dentro de la misma
 * ventana de texto.
 */

function contarSolicitudes(texto) {

  const nombre =
    escaparRegex(
      USUARIO_ESPEJO.nombre
    );


  const usuario =
    escaparRegex(
      USUARIO_ESPEJO.usuario
    );


  const email =
    escaparRegex(
      USUARIO_ESPEJO.email
    );


  const regexNombre =
    new RegExp(
      nombre,
      "gi"
    );


  const resultados = [];

  let coincidencia;


  while (
    (
      coincidencia =
        regexNombre.exec(texto)
    ) !== null
  ) {

    const inicio =
      coincidencia.index;


    const inicioVentana =
      Math.max(
        0,
        inicio -
          CONFIG.ventanaAntes
      );


    const finVentana =
      Math.min(
        texto.length,
        inicio +
          CONFIG.ventanaDespues
      );


    const ventana =
      texto.substring(
        inicioVentana,
        finVentana
      );


    const tieneUsuario =
      new RegExp(
        usuario,
        "i"
      ).test(ventana);


    const tieneEmail =
      new RegExp(
        email,
        "i"
      ).test(ventana);


    if (
      tieneUsuario &&
      tieneEmail
    ) {

      resultados.push({

        posicion: inicio

      });
    }
  }


  return resultados;
}


/* ============================================================
 * OBTENER DIRECCIÓN REAL DEL REMITENTE
 * ============================================================
 *
 * Gmail puede devolver:
 *
 * Nombre del remitente <correo@dominio.com>
 *
 * Por eso se extrae únicamente la dirección.
 */

function obtenerDireccionRemitente(
  remitente
) {

  const match =
    remitente.match(
      /<([^>]+)>/
    );


  if (match) {

    return match[1].trim();
  }


  return remitente.trim();
}


/* ============================================================
 * ESCAPAR TEXTO PARA REGEX
 * ============================================================ */

function escaparRegex(texto) {

  return texto.replace(
    /[.*+?^${}()|[\]\\]/g,
    "\\$&"
  );
}


/* ============================================================
 * COMPROBACIÓN SIN CREAR BORRADORES
 * ============================================================
 *
 * Esta función permite revisar qué solicitudes detectaría
 * el algoritmo antes de crear ningún borrador.
 */

function comprobarSolicitudes() {

  const fechaLimite =
    new Date();


  fechaLimite.setDate(
    fechaLimite.getDate() -
    CONFIG.diasAntiguedad
  );


  const fechaStr =
    Utilities.formatDate(
      fechaLimite,
      Session.getScriptTimeZone(),
      "yyyy/MM/dd"
    );


  const query =
    `"${USUARIO_ESPEJO.usuario}" ` +
    `OR "${USUARIO_ESPEJO.nombre}" ` +
    `OR "${USUARIO_ESPEJO.email}" ` +
    `after:${fechaStr}`;


  const threads =
    GmailApp.search(
      query,
      0,
      CONFIG.maxResultados
    );


  let total = 0;


  threads.forEach(
    thread => {

      thread
        .getMessages()
        .forEach(
          message => {

            const remitente =
              obtenerDireccionRemitente(
                message.getFrom()
              );


            if (
              remitente.toLowerCase() !==
              REMITENTE_SOLICITUDES.toLowerCase()
            ) {

              return;
            }


            const cuerpo =
              message.getPlainBody() || "";


            const solicitudes =
              contarSolicitudes(
                cuerpo
              );


            if (
              solicitudes.length === 0
            ) {

              return;
            }


            total +=
              solicitudes.length;


            Logger.log(
              "----------------------------------------"
            );


            Logger.log(
              "Hilo: " +
              thread.getId()
            );


            Logger.log(
              "Asunto: " +
              thread.getFirstMessageSubject()
            );


            Logger.log(
              "Remitente: " +
              remitente
            );


            Logger.log(
              "Solicitudes detectadas: " +
              solicitudes.length
            );
          }
        );
    }
  );


  SpreadsheetApp
    .getUi()
    .alert(
      "COMPROBACIÓN FINALIZADA\n\n" +

      "Solicitudes detectadas: " +
      total +

      "\n\nNo se ha creado ni enviado ningún correo."
    );
}
