// REEMPLAZO Y COMPLETADO
function reemplazarDatosContrato() {
  const ui = DocumentApp.getUi();
  const htmlInput = HtmlService.createHtmlOutput(
    '<p>Pega el bloque completo con los datos del contrato:</p>' +
    '<textarea id="datos" style="width:100%; height:250px;"></textarea><br><br>' +
    '<button onclick="google.script.run.withSuccessHandler(google.script.host.close).procesarTexto(document.getElementById(\'datos\').value)">Aceptar</button>'
  ).setWidth(600).setHeight(400);
  ui.showModalDialog(htmlInput, 'Pegar datos del contrato');
}

function procesarTexto(texto) {
  // 1. Extraer cliente y teléfono desde el mismo input
  const clienteMatch = texto.match(/cliente[:\s]+([^,]+),\s*(\+56\d{8,9})/i);
  let clienteNombre = '', clienteTelefono = '';
  if (clienteMatch) {
    clienteNombre = clienteMatch[1].trim();
    clienteTelefono = clienteMatch[2].trim();
  }

  // 2. Eliminar “teléfono vendedor” y “teléfono comprador” para que no aparezcan en el doc
  texto = texto
    .replace(/teléfono\s+vendedor:\s*\+56\d{8,9}/gi, '')
    .replace(/teléfono\s+comprador:\s*\+56\d{8,9}/gi, '');

  // 3. Crear copia del documento y obtener cuerpo
  const ui = DocumentApp.getUi();
  const plantillaId = DocumentApp.getActiveDocument().getId();
  const copia = DriveApp.getFileById(plantillaId)
    .makeCopy("Contrato generado – " + generarFechaHoy());
  const docCopia = DocumentApp.openById(copia.getId());
  const body = docCopia.getBody();
  const props = PropertiesService.getDocumentProperties();
  const envioAuto = props.getProperty("envioWhatsAppAuto") === "true";

  // 4. Generar textoPlano para extraer datos
  const textoPlano = texto
    .replace(/\s*\n\s*/g, ', ')
    .replace(/\t/g, ' ')
    .replace(/,{2,}/g, ',');

  // 5. Extraer datos de vendedor/comprador y vehicular
let datos = extraerDatosDesdeTextoPlano(texto);  // <-- TEXTO ORIGINAL, no textoPlano
  datos['CLIENTE_NOMBRE'] = clienteNombre;
  datos['CLIENTE_TELEFONO'] = clienteTelefono;
  datos = completarDesdeBloqueVehicular(texto, datos);
  datos = validarYCompletarDatos(datos);

  // 6. Reemplazar marcadores en el cuerpo del contrato
  for (const clave in datos) {
    const marcador = `{${clave}}`;
    const valor = datos[clave];
    const parrafos = body.getParagraphs();
    parrafos.forEach(p => {
      if (p.getText().includes(marcador)) {
        p.replaceText(`\\{${clave}\\}`, valor);
      }
    });
    const tablas = body.getTables();
    tablas.forEach(tabla => {
      for (let i = 0; i < tabla.getNumRows(); i++) {
        const fila = tabla.getRow(i);
        for (let j = 0; j < fila.getNumCells(); j++) {
          const celda = fila.getCell(j);
          if (celda.getText().includes(marcador)) {
            celda.replaceText(`\\{${clave}\\}`, valor);
          }
        }
      }
    });
  }

  // 7. Verificar marcadores faltantes antes de cerrar
  const textoActualizado = body.getText();
  const marcadoresNoReemplazados = [...textoActualizado.matchAll(/\{([A-Z0-9_]+)\}/g)].map(m => m[1]);
  if (marcadoresNoReemplazados.length > 0) {
    let mensaje = "⚠️ Antes de aplicar los datos, faltan campos obligatorios:\n\n";
    marcadoresNoReemplazados.forEach(m => {
      const sugerencia = datos[m] || "[SIN VALOR]";
      mensaje += `• ${m} → sugerido: ${sugerencia}\n`;
    });
    ui.alert("Faltan datos para completar el contrato", mensaje, ui.ButtonSet.OK);
    throw new Error("Detenido por campos vacíos antes de aplicar texto.");
  }

  docCopia.saveAndClose();
  Logger.log(datos);
  validarMarcadoresPendientes(body, datos);

  // 8. Mostrar validación visual
  mostrarValidacionVisual(docCopia, datos, envioAuto);
}


function extraerDatosDesdeTextoPlano(textoOriginal) {
  const datos = {};

  // Extraer fechas de firma (si existen) antes de limpiar el texto
  const fechaVMatch = textoOriginal.match(/fecha\s+firma\s+vendedor\s+(\d{1,2}\s+\w+\s+\d{4})/i);
  if (fechaVMatch) {
    datos['FECHA_FIRMA_VENDEDOR'] = fechaVMatch[1].trim();
  }
  const fechaCMatch = textoOriginal.match(/fecha\s+firma\s+comprador\s+(\d{1,2}\s+\w+\s+\d{4})/i);
  if (fechaCMatch) {
    datos['FECHA_FIRMA_COMPRADOR'] = fechaCMatch[1].trim();
  }

  // Eliminar frases "FECHA FIRMA VENDEDOR ..." y "FECHA FIRMA COMPRADOR ..." para
  // que no interfieran con el bloque vendedor/comprador
  const textoSinFechas = textoOriginal
    .replace(/fecha\s+firma\s+vendedor\s+\d{1,2}\s+\w+\s+\d{4}/i, '')
    .replace(/fecha\s+firma\s+comprador\s+\d{1,2}\s+\w+\s+\d{4}/i, '');

  // Extraer vendedor y comprador desde el texto limpio
  const bloques = textoSinFechas.match(
    /vendedor[:\s]+(.+?)\s+comprador[:\s]+(.+?)(?=información vehicular|patente|tipo|marca|modelo|año|color|tasacion|$)/is
  );
  if (!bloques || bloques.length < 3) {
    DocumentApp.getUi().alert("❌ No se pudo identificar bloque vendedor/comprador.");
    throw new Error("Formato incorrecto en bloque vendedor/comprador.");
  }

  const vendedorTxt = bloques[1].trim();
  const compradorTxt = bloques[2].trim();
  if (!compradorTxt || compradorTxt.length < 10) {
    DocumentApp.getUi().alert("⚠️ No se logró extraer correctamente los datos del comprador.");
  }

  const partesV = vendedorTxt.split(',');
  datos['NOMBRE_VENDEDOR'] = partesV[0]?.trim();
  datos['ESTADO_CIVIL_VENDEDOR'] = partesV[1]?.trim();
  datos['PROFESION_U_OFICIO_VENDEDOR'] = partesV[2]?.trim();
  datos['RUT_VENDEDOR'] =
    partesV[3]?.match(/\d{1,2}\.\d{3}\.\d{3}-[\dkK]|\d{7,8}-[\dkK]/)?.[0] || partesV[3]?.trim();
  datos['DIRECCION_VENDEDOR'] = partesV
    .slice(4)
    .join(',')
    .replace(/.*?domicilio en /i, '')
    .replace(/quien actúa.*/i, '')
    .trim()
    .toUpperCase();


  const partesC = compradorTxt.split(',');
  datos['NOMBRE_COMPRADOR'] = partesC[0]?.trim();
  datos['ESTADO_CIVIL_COMPRADOR'] = partesC[1]?.trim();
  datos['PROFESION_U_OFICIO_COMPRADOR'] = partesC[2]?.trim();
  datos['RUT_COMPRADOR'] =
    partesC[3]?.match(/\d{1,2}\.\d{3}\.\d{3}-[\dkK]|\d{7,8}-[\dkK]/)?.[0] || partesC[3]?.trim();
  datos['DIRECCION_COMPRADOR'] = partesC
    .slice(4)
    .join(',')
    .replace(/.*?domicilio en /i, '')
    .replace(/quien actúa.*/i, '')
    .trim()
    .toUpperCase();


  return datos;
}

function completarDesdeBloqueVehicular(textoOriginal, datos) {
  const lineas = textoOriginal.split(/\n+/).map(l => l.trim());
  const mapeo = {
    Patente: 'VEHICULO_PATENTE',
    Tipo: 'VEHICULO_TIPO',
    Marca: 'VEHICULO_MARCA',
    Modelo: 'VEHICULO_MODELO',
    Año: 'VEHICULO_ANO',
    Color: 'VEHICULO_COLOR',
    'N° Motor': 'VEHICULO_NRO_MOTOR',
    'N° Chasis': 'VEHICULO_NRO_CHASIS',
    TASACION: 'VALOR_TASACION'
  };

  for (const linea of lineas) {
    for (const etiqueta in mapeo) {
      const regex = new RegExp(`^${etiqueta}\\s*[:\t ]*(.*)$`, 'i');
      const match = linea.match(regex);
      if (match && match[1]) {
        datos[mapeo[etiqueta]] = match[1].trim();
      }
    }
  }

  // Conversión a letras
  if (datos['VALOR_TASACION'] && !datos['VALOR_TASACION_LETRAS']) {
    datos['VALOR_TASACION_LETRAS'] = convertirNumeroALetras(datos['VALOR_TASACION']);
  }

  if (!datos['VEHICULO_COMBUSTIBLE']) {
    const matchCombustible = textoOriginal.match(/Combustible\s*[:\s]*([^\n\r]+)/i);
    if (matchCombustible) {
      datos['VEHICULO_COMBUSTIBLE'] = matchCombustible[1].trim().toUpperCase();
    }
  }

  return datos;
}


function validarYCompletarDatos(datos) {
  const ui = DocumentApp.getUi();
  const camposObligatorios = [
    'NOMBRE_VENDEDOR',
    'RUT_VENDEDOR',
    'DIRECCION_VENDEDOR',
    'NOMBRE_COMPRADOR',
    'RUT_COMPRADOR',
    'DIRECCION_COMPRADOR',
    'FECHA_FIRMA_VENDEDOR',
    'FECHA_FIRMA_COMPRADOR',
    'VEHICULO_MARCA',
    'VEHICULO_MODELO',
    'VEHICULO_TIPO',
    'VEHICULO_ANO',
    'VEHICULO_COLOR',
    'VEHICULO_COMBUSTIBLE',
    'VEHICULO_PATENTE',
    'VEHICULO_NRO_MOTOR',
    'VEHICULO_NRO_CHASIS',
    'VALOR_TASACION',
    'VALOR_TASACION_LETRAS'
  ];
  for (const campo of camposObligatorios) {
    if (campo === 'FECHA_CONTRATO') {
      datos[campo] = generarFechaHoy();
      continue;
    }
    if (!datos[campo] || datos[campo] === '') {
      let mensaje = `Falta el dato: ${campo}.\n\nPor favor ingrésalo`;
      const ejemplos = {
        RUT_VENDEDOR: 'Ej: 12.345.678-9',
        RUT_COMPRADOR: 'Ej: 21.987.654-3',
        FECHA_FIRMA_VENDEDOR: 'Ej: 1 junio 2025',
        FECHA_FIRMA_COMPRADOR: 'Ej: 5 junio 2025',
        VEHICULO_PATENTE: 'Ej: STTP66-4',
        VEHICULO_ANO: 'Ej: 2023',
        VALOR_TASACION: 'Ej: 4.870.000',
        VALOR_TASACION_LETRAS: 'Ej: CUATRO MILLONES OCHOCIENTOS SETENTA MIL PESOS'
      };
      if (ejemplos[campo]) mensaje += `\nFormato sugerido: ${ejemplos[campo]}`;
      const input = ui.prompt(mensaje);
      if (input.getSelectedButton() === ui.Button.OK) {
        datos[campo] = input.getResponseText().trim();
      }
    }
  }
  datos['FECHA_CONTRATO'] = generarFechaHoy();
  if (datos['VALOR_TASACION']) {
    datos['VALOR_VENTA'] = datos['VALOR_TASACION'];
    datos['VALOR_VENTA_LETRAS'] = convertirNumeroALetras(datos['VALOR_TASACION']);
    datos['VALOR_TASACION_LETRAS'] = convertirNumeroALetras(datos['VALOR_TASACION']);
  }
  return datos;
}

function generarFechaHoy() {
  const hoy = new Date();
  const meses = [
    'enero',
    'febrero',
    'marzo',
    'abril',
    'mayo',
    'junio',
    'julio',
    'agosto',
    'septiembre',
    'octubre',
    'noviembre',
    'diciembre'
  ];
  return hoy.getDate() + ' de ' + meses[hoy.getMonth()] + ' de ' + hoy.getFullYear();
}


// --------------------------------------------
// guardarYEnviarContrato sin pedir cliente en loop
// --------------------------------------------
function guardarYEnviarContrato(patente, envioAuto, datos, doc) {
  try {
    const ui = DocumentApp.getUi();

    // 1. Determinar patente limpia
    const patenteLimpia = (patente || '')
      .replace(/[\.\-\t]/g, '')
      .trim()
      .substring(0, 6)
      .toUpperCase();
    if (!patenteLimpia || patenteLimpia.trim() === '') {
      throw new Error("No se pudo determinar una patente válida para guardar el archivo.");
    }

    // 2. Extraer nombreCliente y numeroCliente del objeto datos
    const nombreCliente = datos['CLIENTE_NOMBRE'] || '';
    const numeroCliente = datos['CLIENTE_TELEFONO'] || '';

    // 3. Carpeta y creación de PDF
    const carpetaRaiz = DriveApp.getFolderById("10_YmNNv8X31dP1lo1KJ9mvlLR8emN8gh");
    const letra = patenteLimpia[0];
    const carpetaLetra = obtenerOCrearSubcarpeta(carpetaRaiz, letra);
    const carpetaPatente = obtenerOCrearSubcarpeta(carpetaLetra, patenteLimpia);

    const nombreArchivo = "Contrato Compraventa - " + patenteLimpia + ".pdf";
    const archivosExistentes = carpetaPatente.getFilesByName(nombreArchivo);
    if (archivosExistentes.hasNext()) {
      ui.alert("⚠️ Ya existe un archivo PDF con ese nombre en la carpeta de la patente.");
      return;
    }
    const pdf = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF);
    const pdfFile = carpetaPatente.createFile(pdf).setName(nombreArchivo);

    // 4. Abrir hoja y preparar datos para la fila
    const spreadsheetId = "1Xl6rDN9dVXuTTQCUblKka_coJxFtsHLCqVblz7q0c7k";
    const hoja = SpreadsheetApp.openById(spreadsheetId).getSheetByName("2024");

    // 5. Calcular valores
    const marca = datos['VEHICULO_MARCA'];
    const vendedor =
      datos['NOMBRE_VENDEDOR']?.split(" ").slice(0, 2).join(" ") || '';
    const comprador =
      datos['NOMBRE_COMPRADOR']?.split(" ").slice(0, 2).join(" ") || "";
    const telefonoVendedor = datos['TELEFONO_VENDEDOR'] || ""; // si extrajiste ese campo o lo pasaste en datos
    const telefonoComprador = datos['TELEFONO_COMPRADOR'] || ""; // idem

    const tasacionTexto = datos['VALOR_TASACION'] || "0";
    const valorTasacion =
      parseInt(tasacionTexto.replace(/\D/g, ""), 10) || 0;
    let valorTramite = Math.round(valorTasacion * 0.015) + 37000 + 36000 + 5000;

    // 6. Descuentos especiales (opcional)
    const clienteEspecial = [
      { nombre: "matias neumann", telefono: "+56926199778", descuento: 10000 },
      { nombre: "yeison milano", telefono: "942589489", descuento: 10000 },
      { nombre: "ariel israel", telefono: "+56977162699", descuento: 10000 },
      { nombre: "danilo", telefono: "+56985044922", descuento: 136000 }
    ];
    const clienteMatch = clienteEspecial.find(
      c =>
        nombreCliente.toLowerCase().includes(c.nombre) ||
        numeroCliente.replace(/\D/g, "").endsWith(
          c.telefono.replace(/\D/g, "")
        )
    );
    if (clienteMatch) {
      const confirm = ui.alert(
        `Se detectó que el cliente es "${clienteMatch.nombre}". ¿Aplicar descuento de $${clienteMatch.descuento}?`,
        ui.ButtonSet.YES_NO
      );
      if (confirm === ui.Button.YES) valorTramite -= clienteMatch.descuento;
    }

    const linkContrato = doc.getUrl();

    // 7. Agregar fila a la hoja sin volver a preguntar nada
    hoja.appendRow([
      patenteLimpia, // Columna B
      marca, // Columna C
      nombreCliente, // Columna D
      "cerrada",
      "",
      "",
      "",
      "",
      `${vendedor} - ${telefonoVendedor}\n${comprador} - ${telefonoComprador}`, // columna I
      "",
      "",
      "",
      "",
      valorTasacion, // columna N
      valorTramite // columna O
    ]);
    const ultimaFila = hoja.getLastRow();
    hoja
      .getRange(ultimaFila, 2)
      .setFormula(`=HYPERLINK("${linkContrato}", "${patenteLimpia}")`);

    // 8. Envío de WhatsApp (solo si envioAuto = true)
    if (envioAuto) {
      enviarMensajesWhatsapp(datos, patenteLimpia, pdfFile.getUrl());
    } else {
      ui.alert("✅ Documento guardado. Recuerda enviar mensajes si hace falta.");
    }
  } catch (error) {
    DocumentApp.getUi().alert("❌ Error: " + error.message);
    throw error;
  }
}


function obtenerOCrearSubcarpeta(padre, nombre) {
  const carpetas = padre.getFoldersByName(nombre);
  return carpetas.hasNext() ? carpetas.next() : padre.createFolder(nombre);
}


// --------------------------------------------
// enviarMensajesWhatsapp sin loops adicionales
// --------------------------------------------
function enviarMensajesWhatsapp(datos, patente, link) {
  const primerNombreVendedor =
    datos["NOMBRE_VENDEDOR"]?.split(" ")[0] || "";
  const primerNombreComprador =
    datos["NOMBRE_COMPRADOR"]?.split(" ")[0] || "";
  const nombreCliente = datos["CLIENTE_NOMBRE"]?.split(" ")[0] || "";

  // Ya no volvemos a preguntar teléfono: 
  // se asume que CLIENTE_TELEFONO, TELEFONO_VENDEDOR y TELEFONO_COMPRADOR 
  // fueron extraídos antes y están en datos.
  const numeroCliente = datos["CLIENTE_TELEFONO"] || "";
  const numeroVendedor = datos["TELEFONO_VENDEDOR"] || "";
  const numeroComprador = datos["TELEFONO_COMPRADOR"] || "";

  const mensajeVendedor = `Hola ${primerNombreVendedor}, contrato vehículo ${patente}: ${link}`;
  const mensajeComprador = `Hola ${primerNombreComprador}, contrato vehículo ${patente}: ${link}`;
  const mensajeCliente = `Hola ${nombreCliente}, enviamos contrato vehículo ${patente}. ${link}`;

  if (numeroCliente) enviarWhatsAppUltraMsg(numeroCliente, mensajeCliente);
  if (numeroVendedor) enviarWhatsAppUltraMsg(numeroVendedor, mensajeVendedor);
  if (numeroComprador)
    enviarWhatsAppUltraMsg(numeroComprador, mensajeComprador);
}


// --------------------------------------------
// solicitarTelefono ya no se invoca en loop
// --------------------------------------------
function solicitarTelefono(nombre) {
  const ui = DocumentApp.getUi();
  if (!nombre || nombre === "undefined") {
    nombre = "Cliente";
  }
  const respuesta = ui.prompt(
    `¿Cuál es el número de WhatsApp de ${nombre}? (con +56)`
  );
  return respuesta.getResponseText();
}


function convertirNumeroALetras(numero) {
  numero = parseInt(numero.toString().replace(/\D/g, ""), 10);
  if (isNaN(numero)) return "";
  const enLetras = convertirEnPalabras(numero).toUpperCase();
  return `${enLetras}`;
}


function validarMarcadoresPendientes(body, datos) {
  const ui = DocumentApp.getUi();
  const texto = body.getText();
  const marcadoresPendientes = [...texto.matchAll(/\{([A-Z0-9_]+)\}/g)].map(
    (m) => m[1]
  );
  const ejemplos = {
    RUT_VENDEDOR: "Ej: 12.345.678-9",
    RUT_COMPRADOR: "Ej: 21.987.654-3",
    FECHA_FIRMA_VENDEDOR: "Ej: 1 junio 2025",
    FECHA_FIRMA_COMPRADOR: "Ej: 5 junio 2025",
    VEHICULO_PATENTE: "Ej: STTP66-4",
    VEHICULO_ANO: "Ej: 2023",
    VALOR_TASACION: "Ej: 4.870.000",
    VALOR_TASACION_LETRAS: "Ej: CUATRO MILLONES OCHOCIENTOS SETENTA MIL PESOS",
  };

  let algunoRellenado = false;
  for (const marcador of marcadoresPendientes) {
    let mensaje = `Falta el campo: ${marcador}.\n\nPor favor ingrésalo.`;
    if (ejemplos[marcador]) mensaje += `\nFormato sugerido: ${ejemplos[marcador]}`;
    const input = ui.prompt(mensaje);
    if (input.getSelectedButton() === ui.Button.OK) {
      datos[marcador] = input.getResponseText().trim();
      algunoRellenado = true;
    }
  }
  if (algunoRellenado) {
    validarMarcadoresPendientes(body, datos);
    return;
  }
  const textoFinal = body.getText();
  const sinReemplazar = [...textoFinal.matchAll(/\{([A-Z0-9_]+)\}/g)].map((m) => m[1]);
  if (sinReemplazar.length > 0) {
    let mensaje = "⚠️ Los siguientes campos no fueron reemplazados:\n\n";
    sinReemplazar.forEach((m) => {
      const sugerencia = datos[m] || "[PENDIENTE]";
      mensaje += `• ${m} → sugerido: ${sugerencia}\n`;
    });
    mensaje += "\nEdita manualmente el documento o completa los datos para continuar.";
    ui.alert("Faltan datos", mensaje, ui.ButtonSet.OK);
    throw new Error("Contrato incompleto – contiene marcadores sin reemplazo.");
  }
}


function convertirEnPalabras(numero) {
  const unidades = [
    "cero",
    "uno",
    "dos",
    "tres",
    "cuatro",
    "cinco",
    "seis",
    "siete",
    "ocho",
    "nueve",
  ];
  const decenas = [
    "diez",
    "once",
    "doce",
    "trece",
    "catorce",
    "quince",
    "dieciséis",
    "diecisiete",
    "dieciocho",
    "diecinueve",
  ];
  const decenas2 = [
    "",
    "",
    "veinte",
    "treinta",
    "cuarenta",
    "cincuenta",
    "sesenta",
    "setenta",
    "ochenta",
    "noventa",
  ];
  const centenas = [
    "",
    "cien",
    "doscientos",
    "trescientos",
    "cuatrocientos",
    "quinientos",
    "seiscientos",
    "setecientos",
    "ochocientos",
    "novecientos",
  ];

  function seccion(num) {
    if (num < 10) return unidades[num];
    if (num < 20) return decenas[num - 10];
    if (num < 100)
      return (
        decenas2[Math.floor(num / 10)] +
        (num % 10 !== 0 ? " y " + unidades[num % 10] : "")
      );
    if (num < 1000)
      return (
        centenas[Math.floor(num / 100)] +
        (num % 100 !== 0 ? " " + seccion(num % 100) : "")
      );
    if (num < 1000000) {
      const miles = Math.floor(num / 1000);
      const resto = num % 1000;
      const milesTxt = miles === 1 ? "mil" : seccion(miles) + " mil";
      return milesTxt + (resto !== 0 ? " " + seccion(resto) : "");
    }
    const millones = Math.floor(num / 1000000);
    const resto = num % 1000000;
    const millonesTxt = millones === 1 ? "un millón" : seccion(millones) + " millones";
    return millonesTxt + (resto !== 0 ? " " + seccion(resto) : "");
  }

  return seccion(numero) + " pesos";
}


function enviarWhatsAppUltraMsg(numero, mensaje) {
  const instanceId = "111615";
  let token = PropertiesService.getDocumentProperties().getProperty("ultramsg_token");
  if (!token) {
    token = "1y361vgl7me45pcl";
    PropertiesService.getDocumentProperties().setProperty("ultramsg_token", token);
  }
  const url = `https://api.ultramsg.com/instance${instanceId}/messages/chat`;
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      token: token,
      to: numero,
      body: mensaje,
    }),
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    Logger.log(`✅ WhatsApp enviado a ${numero}: ${mensaje}`);
    Logger.log(`🔄 Respuesta API: ${response.getContentText()}`);
  } catch (e) {
    Logger.log(`❌ Error al enviar WhatsApp a ${numero}: ${e.message}`);
    DocumentApp.getUi().alert(`Error al enviar WhatsApp a ${numero}: ${e.message}`);
  }
}

function mostrarValidacionVisual(docCopia, datos, envioAuto) {
  const htmlContrato = construirHtmlDesdeBody(docCopia.getBody());
  const html = HtmlService.createTemplateFromFile("ValidadorVisual");
  html.texto = htmlContrato;
  html.datos = datos;
  html.envioAuto = envioAuto;
  html.docId = docCopia.getId();
  DocumentApp.getUi().showModalDialog(html.evaluate().setWidth(900).setHeight(600), "Validación visual del contrato");
}

function construirHtmlDesdeBody(body) {
  let html = "";
  const elements = body.getNumChildren();
  for (let i = 0; i < elements; i++) {
    const element = body.getChild(i);
    const type = element.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      const parrafo = element.asParagraph();
      let alineacion = parrafo.getAlignment();
      let clase = "";
      if (alineacion === DocumentApp.HorizontalAlignment.CENTER)
        clase = ' style="text-align:center;"';
      else if (alineacion === DocumentApp.HorizontalAlignment.RIGHT)
        clase = ' style="text-align:right;"';
      let texto = "";
      const numRuns = parrafo.getNumChildren();
      for (let r = 0; r < numRuns; r++) {
        const run = parrafo.getChild(r);
        if (run.getType() === DocumentApp.ElementType.TEXT) {
          const text = run.asText();
          let parte = text.getText();
          const attrs = text.getAttributes();
          if (attrs.BOLD) parte = `<strong>${parte}</strong>`;
          if (attrs.ITALIC) parte = `<em>${parte}</em>`;
          if (attrs.UNDERLINE) parte = `<u>${parte}</u>`;
          texto += parte;
        }
      }
      html += `<p${clase}>${texto}</p>`;
    } else if (type === DocumentApp.ElementType.TABLE) {
      const table = element.asTable();
      html += '<table border="0" cellspacing="0" cellpadding="6" style="width:100%; border-collapse:collapse; border:0;">';
      for (let r = 0; r < table.getNumRows(); r++) {
        html += "<tr>";
        const row = table.getRow(r);
        for (let c = 0; c < row.getNumCells(); c++) {
          const cell = row.getCell(c);
          let cellText = "";
          const numChildren = cell.getNumChildren();
          for (let k = 0; k < numChildren; k++) {
            const cellChild = cell.getChild(k);
            if (cellChild.getType() === DocumentApp.ElementType.PARAGRAPH) {
              cellText += cellChild.asParagraph().getText() + "<br>";
            }
          }
          html += `<td>${cellText}</td>`;
        }
        html += "</tr>";
      }
      html += "</table><br>";
    }
  }
  return html;
}

function continuarDespuesValidacion(docId, datosJSON, envioAuto) {
  const datosBrutos = JSON.parse(datosJSON);
  const datos = {};
  Object.keys(datosBrutos).forEach((key) => {
    const claveLimpia = key.replace(/\s|\u200B|\uFEFF/g, "").toUpperCase();
    datos[claveLimpia] = datosBrutos[key];
  });
  Logger.log("Objeto normalizado:", datos);
  let patente = datos["VEHICULO_PATENTE"] || "";
  if (!patente) {
    const doc = DocumentApp.openById(docId);
    const textoCompleto = doc.getBody().getText();
    const match = textoCompleto.match(/patente\s+([^\s-]+)-/i);
    if (match) {
      patente = match[1].toUpperCase();
      datos["VEHICULO_PATENTE"] = patente;
      Logger.log("Patente extraída automáticamente:", patente);
    }
  }
  const doc = DocumentApp.openById(docId);
  guardarYEnviarContrato(datos["VEHICULO_PATENTE"], envioAuto, datos, doc);
}

function reportarErrorSegmento(textoError, marcadorDetectado) {
  const admin = "+56984497385";
  const msg = `⚠️ Error en contrato detectado:\n\nSegmento:\n"${textoError}"\n\nCampo posiblemente relacionado: ${marcadorDetectado || "No detectado"}`;
  enviarWhatsAppUltraMsg(admin, msg);
}
