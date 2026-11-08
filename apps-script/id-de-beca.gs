/**
 * IDs de beca y correcciones desde el formulario (reemplaza a ids-automaticos.gs).
 *
 * Qué hace:
 *   - Cada beca nueva que llega por el formulario recibe un ID (BPH-0463, BPH-0464, …) en la columna "id_beca".
 *   - A quien la envió le llega un correo con ese ID.
 *   - Si alguien se equivocó, vuelve a llenar el formulario y pone el ID en la pregunta "ID de beca a corregir":
 *     la fila de ese ID se SOBREESCRIBE con los datos nuevos y la respuesta repetida se borra.
 *   - Lo que envía un correo de CORREOS_QUE_PUBLICAN (Jenny) se marca "aprobada" y se publica solo.
 *     Lo que envía cualquier otra persona queda "Pendiente de revisión", y solo esos correos pueden corregir por ID.
 *   - El formulario pide iniciar sesión con Google, así nadie puede escribir el correo de Jenny haciéndose pasar por ella.
 *
 * Campos nuevos (15 de octubre): pega este archivo completo y ejecuta "actualizarCampos" una vez (ver abajo).
 *
 * Cómo usarlo (una sola vez):
 *   script.google.com (cuenta raudalesg99@gmail.com) > Nuevo proyecto > pega este archivo > Guardar >
 *   elige la función "configurarIds" > Ejecutar > acepta los permisos (incluye enviar correos).
 */
const IDS_FORMULARIO = "1Eh02U-z-BddnM5JcXh4vhYgvagMu1H1c1AS2bWNwfm4";
const IDS_HOJA = "1rzRIiZHC14KqA4zYR-TMpgCVQ5k6dEJOI2XWYQsE8Qw";
const IDS_PREFIJO = "BPH-";
const TITULO_CORREGIR = "ID de beca a corregir";
// Correos que publican solos y pueden corregir por ID (en minúsculas).
// raudalesg99@gmail.com está solo para las pruebas: quitarlo cuando se le entregue a Jenny.
const CORREOS_QUE_PUBLICAN = ["becasparahn@gmail.com", "raudalesg99@gmail.com"];
// Quién recibe copia de cada corrección. Durante las pruebas solo Mariaj; al entregar, cambiar a ["becasparahn@gmail.com"].
const CORREOS_COPIA_CORRECCIONES = ["raudalesg99@gmail.com"];
const COLUMNAS_PROPIAS = ["aprobada", "estado", "id_beca", "plazos", "archivada"]; // columnas de la hoja que no vienen del formulario

// Archivar o publicar desde el formulario (8 de noviembre): ver "activarArchivar" más abajo.
const TITULO_ACCION = "¿Qué quieres hacer?";
const ACCION_REGISTRAR = "Registrar una beca nueva o corregir una";
const ACCION_ARCHIVAR = "Archivar una beca (quitarla del buscador)";
const ACCION_PUBLICAR = "Publicar o volver a publicar una beca";
const TITULO_CODIGO = "Código de la beca";
const SECCION_DATOS = "Datos de la beca";
const SECCION_ARCHIVAR = "Archivar o publicar una beca";
const COL_ARCHIVADA = "archivada"; // fecha en que se archivó (vacía = no archivada)
const SOLO_FORMULARIO = [TITULO_ACCION, TITULO_CODIGO]; // preguntas que no son datos de la beca

function configurarIds() {
  const form = FormApp.openById(IDS_FORMULARIO);

  // Pregunta opcional al inicio para corregir una beca ya enviada
  let item = form.getItems().find((i) => i.getTitle() === TITULO_CORREGIR);
  if (!item) {
    item = form.addTextItem().setTitle(TITULO_CORREGIR)
      .setHelpText("Déjalo VACÍO si es una beca nueva. Si te equivocaste en una beca que ya enviaste, escribe aquí su ID " +
        "(por ejemplo BPH-0463) y llena todo de nuevo: la beca anterior se reemplaza con estos datos.")
      .setValidation(FormApp.createTextValidation()
        .requireTextMatchesPattern("^\\s*([Bb][Pp][Hh]-\\d{4})?\\s*$")
        .setHelpText("Usa el formato BPH-0000, o déjalo vacío.").build())
      .setRequired(false);
  }
  form.moveItem(item.getIndex(), 0);

  // Pedir el correo verificado (iniciar sesión con Google) para mandar el ID y saber quién publica
  try { form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED); }
  catch (e) { form.setCollectEmail(true); }
  form.setConfirmationMessage("¡Gracias! Te enviamos un correo con el ID de esta beca (BPH-…). Guárdalo: si te equivocaste, " +
    "vuelve a llenar el formulario y escribe ese ID en la primera pregunta para corregirla.");

  // Disparador: un solo "alEnviar" (quita los de versiones anteriores)
  ScriptApp.getProjectTriggers()
    .filter((t) => ["alEnviar", "asignarIds"].includes(t.getHandlerFunction()))
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("alEnviar").forSpreadsheet(IDS_HOJA).onFormSubmit().create();

  SpreadsheetApp.flush();
  Utilities.sleep(4000); // que aparezcan las columnas nuevas
  completarIdsFaltantes();
  Logger.log("Listo: IDs automáticos, correo con el ID y correcciones por ID activados.");
}

// Se ejecuta sola con cada respuesta del formulario.
function alEnviar(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const hoja = e.range.getSheet();
    const fila = e.range.getRow();
    const cab = encabezados(hoja);
    const ancho = cab.length;
    const nueva = hoja.getRange(fila, 1, 1, ancho).getValues()[0];
    const c = (titulo) => cab.indexOf(titulo); // índice 0-based o -1
    const iId = c("id_beca"), iAprob = c("aprobada"), iCorr = c(TITULO_CORREGIR);
    const iNombre = c("Nombre de la beca");
    const iCorreo = cab.findIndex((h) => /correo|email/i.test(h));
    const nombre = String(nueva[iNombre] || "");
    const correo = iCorreo >= 0 ? String(nueva[iCorreo] || "").trim() : "";
    const autorizada = CORREOS_QUE_PUBLICAN.includes(correo.toLowerCase());
    const idPedido = iCorr >= 0 ? String(nueva[iCorr] || "").trim().toUpperCase() : "";
    const datos = () => resumen(cab, nueva, iCorr, iCorreo);
    const comoQueda = autorizada ? "Ya está publicada en el buscador." : "Quedó pendiente de revisión antes de publicarse.";

    // Archivar o publicar: la respuesta no es una beca, solo cambia "aprobada" de la beca con ese código
    const iAccion = c(TITULO_ACCION);
    const accion = iAccion >= 0 ? String(nueva[iAccion] || "").trim() : "";
    if (accion === ACCION_ARCHIVAR || accion === ACCION_PUBLICAR) {
      archivarOPublicar(hoja, fila, cab, nueva, accion === ACCION_ARCHIVAR, correo, autorizada);
      return;
    }

    if (idPedido && autorizada) {
      const ids = hoja.getRange(2, iId + 1, hoja.getLastRow() - 1, 1).getValues().map((r) => String(r[0]).trim().toUpperCase());
      const destino = ids.findIndex((v, i) => v === idPedido && i + 2 !== fila) + 2; // 1 si no se encontró
      if (destino >= 2) {
        // Sobrescribir la fila de ese ID con lo que vino del formulario (sin tocar aprobada/estado/id)
        cab.forEach((h, i) => {
          if (COLUMNAS_PROPIAS.includes(h) || SOLO_FORMULARIO.includes(h) || i === iCorr) return;
          hoja.getRange(destino, i + 1).setValue(nueva[i]);
        });
        if (iCorr >= 0) hoja.getRange(destino, iCorr + 1).clearContent();
        if (iAprob >= 0) hoja.getRange(destino, iAprob + 1).setValue(true); // se publica sola
        if (c(COL_ARCHIVADA) >= 0) hoja.getRange(destino, c(COL_ARCHIVADA) + 1).clearContent(); // corregida = ya no archivada
        escribirPlazos(hoja, destino);
        hoja.deleteRow(fila);
        asegurarEstado(hoja);
        const copias = [...new Set([correo.toLowerCase(), ...CORREOS_COPIA_CORRECCIONES])].filter(Boolean).join(",");
        avisar(copias, "Beca corregida: " + idPedido,
          "Se reemplazaron los datos de la beca " + idPedido + " (" + nombre + ").\nCorrección enviada por: " + correo + "\n" +
          "Ya está publicada con los datos nuevos.\n\nAsí quedó registrada:\n" + datos() +
          textoEnlaceCorregir(hoja, destino > fila ? destino - 1 : destino));
        return;
      }
    }

    // Beca nueva (o corrección que no se pudo aplicar)
    let aviso = "";
    if (idPedido) {
      aviso = autorizada
        ? "No encontramos el ID " + idPedido + ", así que esta beca se guardó como nueva.\n\n"
        : "Solo Becas para Hondureños puede corregir becas por ID, así que esta respuesta se guardó como beca nueva.\n\n";
    }
    const nuevoId = siguienteId(hoja, iId);
    hoja.getRange(fila, iId + 1).setValue(nuevoId);
    if (iAprob >= 0) hoja.getRange(fila, iAprob + 1).setValue(autorizada); // solo se publica sola si la envía BpH
    escribirPlazos(hoja, fila);
    asegurarEstado(hoja);
    avisar(correo, "Beca registrada: " + nuevoId,
      aviso + "ID: " + nuevoId + "\n" + comoQueda + "\n\nLo que registraste:\n" + datos() + "\n\n" +
      (autorizada
        ? "Revisa los datos y guarda este ID." + textoEnlaceCorregir(hoja, fila)
        : "Guarda este ID por si necesitas consultarlo."));
  } finally {
    lock.releaseLock();
  }
}

// Pone ID a cualquier fila con nombre que todavía no tenga.
function completarIdsFaltantes() {
  const hoja = SpreadsheetApp.openById(IDS_HOJA).getSheetByName("Respuestas");
  const cab = encabezados(hoja);
  const iId = cab.indexOf("id_beca"), iNombre = cab.indexOf("Nombre de la beca");
  const filas = hoja.getLastRow() - 1;
  if (iId < 0 || iNombre < 0 || filas < 1) return;
  const nombres = hoja.getRange(2, iNombre + 1, filas, 1).getValues();
  const rango = hoja.getRange(2, iId + 1, filas, 1);
  const ids = rango.getValues();
  let mayor = mayorId(ids), cambios = 0;
  ids.forEach((r, i) => { if (!r[0] && String(nombres[i][0]).trim()) { r[0] = formatoId(++mayor); cambios++; } });
  if (cambios) rango.setValues(ids);
  asegurarEstado(hoja);
}

// La fórmula de "estado" vive en la fila 2; si se borró esa fila, se vuelve a poner. Con forzar=true se reescribe siempre.
function asegurarEstado(hoja, forzar) {
  const cab = encabezados(hoja);
  const iEst = cab.indexOf("estado"), iAprob = cab.indexOf("aprobada"), iNombre = cab.indexOf("Nombre de la beca");
  const iArch = cab.indexOf(COL_ARCHIVADA);
  if (iEst < 0) return;
  const celda = hoja.getRange(2, iEst + 1);
  if (celda.getFormula() && !forzar) return;
  const L = (i) => letra(i + 1);
  const rango = (i) => L(i) + '2:' + L(i);
  const segunAprobada = 'IF(' + rango(iAprob) + '=TRUE,"Publicada","Pendiente de revisión")';
  const estado = iArch >= 0 ? 'IF(' + rango(iArch) + '<>"","Archivada",' + segunAprobada + ')' : segunAprobada;
  hoja.getRange(2, iEst + 1, hoja.getMaxRows() - 1, 1).clearContent();
  celda.setFormula('=ARRAYFORMULA(IF(' + rango(iNombre) + '="","",' + estado + '))');
}

// Archiva (aprobada = FALSE, "archivada" = fecha) o publica (aprobada = TRUE) la beca con el código pedido.
// La respuesta del formulario se borra porque no es una beca. Solo CORREOS_QUE_PUBLICAN pueden hacerlo.
function archivarOPublicar(hoja, fila, cab, nueva, archivar, correo, autorizada) {
  const iId = cab.indexOf("id_beca"), iAprob = cab.indexOf("aprobada"), iArch = cab.indexOf(COL_ARCHIVADA);
  const iNombre = cab.indexOf(N.nombre), iCodigo = cab.indexOf(TITULO_CODIGO);
  const codigo = iCodigo >= 0 ? String(nueva[iCodigo] || "").trim().toUpperCase() : "";
  const verbo = archivar ? "archivar" : "publicar";
  hoja.deleteRow(fila);
  asegurarEstado(hoja);
  if (!autorizada) {
    avisar(correo, "No se pudo " + verbo + " la beca " + codigo,
      "Solo Becas para Hondureños puede archivar o publicar becas. No se cambió nada.");
    return;
  }
  const filas = Math.max(hoja.getLastRow() - 1, 1);
  const ids = hoja.getRange(2, iId + 1, filas, 1).getValues().map((r) => String(r[0]).trim().toUpperCase());
  const i = codigo ? ids.indexOf(codigo) : -1;
  if (i < 0) {
    avisar(correo, "No se pudo " + verbo + " la beca " + codigo,
      "No encontramos el código " + codigo + ". Revisa que esté escrito igual que al pie de la ficha en el buscador. No se cambió nada.");
    return;
  }
  const destino = i + 2;
  const nombre = String(hoja.getRange(destino, iNombre + 1).getValue());
  hoja.getRange(destino, iAprob + 1).setValue(!archivar);
  if (iArch >= 0) {
    if (archivar) hoja.getRange(destino, iArch + 1).setValue(new Date());
    else hoja.getRange(destino, iArch + 1).clearContent();
  }
  const copias = [...new Set([correo.toLowerCase(), ...CORREOS_COPIA_CORRECCIONES])].filter(Boolean).join(",");
  avisar(copias, (archivar ? "Beca archivada: " : "Beca publicada: ") + codigo,
    codigo + " (" + nombre + ")\n" +
    (archivar
      ? "Ya no aparece en el buscador. Sus datos se guardan: para mostrarla de nuevo, elige \"" + ACCION_PUBLICAR + "\" en el formulario."
      : "Ya aparece en el buscador.") +
    "\nCambio enviado por: " + correo);
}

// Las fechas de la hoja son medianoche en la zona horaria de la HOJA. Si se leen con la zona del script
// (que puede ser otra), salen un día antes. Por eso todo se formatea con la zona de la hoja.
let zonaHojaGuardada = "";
function zonaHoja() {
  if (!zonaHojaGuardada) zonaHojaGuardada = SpreadsheetApp.openById(IDS_HOJA).getSpreadsheetTimeZone();
  return zonaHojaGuardada;
}
// Fecha de la hoja -> mismo día calendario en la zona del script (lo que esperan las respuestas del formulario).
function mismoDia(fecha, zona) {
  const [a, m, d] = Utilities.formatDate(fecha, zona, "yyyy-MM-dd").split("-").map(Number);
  return new Date(a, m - 1, d, 12); // mediodía: lejos de cualquier cambio de hora
}
function encabezados(hoja) { return hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(String); }
function mayorId(valores) {
  let m = 0;
  valores.forEach(([v]) => { const x = String(v).match(/(\d+)$/); if (x) m = Math.max(m, Number(x[1])); });
  return m;
}
function siguienteId(hoja, iId) { return formatoId(mayorId(hoja.getRange(2, iId + 1, hoja.getLastRow() - 1, 1).getValues()) + 1); }
function formatoId(n) { return IDS_PREFIJO + String(n).padStart(4, "0"); }
function avisar(correo, asunto, texto) {
  if (!correo || !/@/.test(correo)) return;
  MailApp.sendEmail(correo, asunto + " · Buscador de becas BpH", texto + "\n\nBecas para Hondureños");
}
// Lista "Pregunta: respuesta" de todo lo que se envió (sin marca temporal, correo ni columnas propias de la hoja).
function resumen(cab, fila, iCorr, iCorreo) {
  const zona = zonaHoja();
  return cab.map((h, i) => {
    if (i === 0 || i === iCorr || i === iCorreo || COLUMNAS_PROPIAS.includes(h) || SOLO_FORMULARIO.includes(h)) return null;
    let v = fila[i];
    if (v instanceof Date) v = Utilities.formatDate(v, zona, v.getFullYear() < 1901 ? "HH:mm" : "dd/MM/yyyy"); // 1899 = solo hora
    v = String(v === null || v === undefined ? "" : v).trim();
    if (!v && (/^Plazo \d|^¿Agregar/.test(h) || h === N.fechaVieja || h === N.variasVieja || Object.values(U).includes(h))) return null; // secciones que no se usaron
    return "- " + h + ": " + (v || "(vacío)");
  }).filter(Boolean).join("\n");
}
function letra(n) {
  let s = "";
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/* =====================================================================================================
 * CAMPOS NUEVOS (15 de octubre): ejecutar UNA vez la función "actualizarCampos".
 *
 * Fechas límite sin duplicar becas: cada beca sigue siendo UNA fila. El formulario tiene hasta
 * MAX_PLAZOS secciones "Fecha límite N" (ronda, tipo de plazo, fecha, hora, zona horaria, periodo de
 * ingreso, confirmada/estimada). Al final de cada una pregunta si hay otra fecha. El script junta esas
 * columnas en la columna "plazos" (texto JSON) que lee el buscador.
 * ===================================================================================================== */
const MAX_PLAZOS = 4;
const TIPOS_CIERRE = ["Fecha única", "Varias rondas o fechas", "Solo mes", "Abierta todo el año", "Sin fecha publicada"];
const CIERRE_VIEJO_A_NUEVO = { "fecha": "Fecha única", "varias fechas": "Varias rondas o fechas", "solo mes": "Solo mes", "abierta todo el año": "Abierta todo el año" };
const TIPOS_PLAZO = ["Solicitud de beca", "Admisión universitaria", "Nominación", "Entrega de documentación"];
const PERIODOS_INGRESO = ["Primavera", "Otoño", "Año académico", "No especificado"];
const TIPOS_POSGRADO = ["Máster", "Doctorado", "Otros estudios de posgrado"];
const DURACIONES = ["1 año", "2 años", "3 años", "4 años", "Duración del programa", "Variable", "No especificada"];
const IMPORTE_SEGUN = ["Necesidad económica", "Mérito académico", "Evaluación individual", "Monto fijo", "No especificado"];
const MESES_CIERRE = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];

// Títulos (= encabezados en la hoja)
const N = {
  nombre: "Nombre de la beca", paisViejo: "País", pais: "País de destino", nacionalidades: "Nacionalidades elegibles",
  nivel: "Nivel", posgrado: "Tipo de posgrado", area: "Área de estudio", subarea: "Subárea o tema específico",
  edadMin: "Edad mínima", edadMax: "Edad máxima", cobertura: "Cobertura", duracion: "Duración de la financiación",
  importe: "El importe depende de", plazas: "Número de plazas", idioma: "Idioma", modalidad: "Modalidad",
  enlace: "Enlace oficial", verificado: "Fecha de última verificación", tipo: "Tipo de cierre", mes: "Mes de cierre",
  fechaVieja: "Fecha de cierre", variasVieja: "Fechas de cierre",
};
const P = (n) => ({
  seccion: "Fecha límite " + n,
  ronda: "Plazo " + n + ": nombre de la ronda",
  tipo: "Plazo " + n + ": tipo de plazo",
  fecha: "Plazo " + n + ": fecha",
  hora: "Plazo " + n + ": hora límite",
  zona: "Plazo " + n + ": zona horaria",
  ingreso: "Plazo " + n + ": periodo de ingreso",
  estado: "Plazo " + n + ": fecha confirmada o estimada",
  otra: "¿Agregar la fecha límite " + (n + 1) + "?",
});

// Sección para "Fecha única": una sola fecha, sin nombre de ronda ni "¿agregar otra?".
const U = {
  seccion: "Fecha de cierre", fecha: "Fecha de cierre", hora: "Hora límite", zona: "Zona horaria",
  tipo: "Tipo de plazo", ingreso: "Periodo de ingreso", estado: "Fecha confirmada o estimada",
};

// Ejecutar UNA vez (16 de octubre): "Fecha única" deja de llevar a "Fecha límite 1" y va a "Fecha de cierre".
// Las rondas solo aparecen con "Varias rondas o fechas". Se puede volver a ejecutar sin duplicar nada.
function agregarFechaUnica(desdeActualizar) {
  const form = FormApp.openById(IDS_FORMULARIO);
  const T_ = FormApp.ItemType;
  const comoPB = (x) => x.asPageBreakItem ? x.asPageBreakItem() : x;
  const comoMC = (x) => x.asMultipleChoiceItem ? x.asMultipleChoiceItem() : x;
  const reintentar = (f) => { for (let k = 1; ; k++) { try { return f(); } catch (e) { if (k >= 5) throw e; Utilities.sleep(3000 * k); } } };
  const cache = reintentar(() => form.getItems().map((i) => ({ i, titulo: i.getTitle(), tipo: i.getType() })));
  const buscar = (titulo, tipo) => { const x = cache.find((c) => c.titulo === titulo && (!tipo || c.tipo === tipo)); return x ? x.i : null; };

  const mes = buscar(N.mes, T_.LIST), pbMes = buscar(N.mes, T_.PAGE_BREAK), pb1 = buscar(P(1).seccion, T_.PAGE_BREAK);
  const tipoItem = buscar(N.tipo, T_.MULTIPLE_CHOICE);
  if (!mes || !pbMes || !pb1 || !tipoItem) throw new Error("Primero hay que haber ejecutado actualizarCampos.");

  let pb = buscar(U.seccion, T_.PAGE_BREAK);
  let items;
  if (!pb) {
    pb = form.addPageBreakItem().setTitle(U.seccion).setHelpText("La beca tiene una sola fecha límite.");
    items = [
      form.addDateItem().setTitle(U.fecha).setRequired(true),
      form.addTimeItem().setTitle(U.hora).setHelpText("Solo si se publica.").setRequired(false),
      form.addTextItem().setTitle(U.zona).setHelpText("Solo si se publica la hora. Ej.: hora de Honduras, CET, ET (EE. UU.).").setRequired(false),
      form.addListItem().setTitle(U.tipo).setHelpText("Opcional.").setChoiceValues(TIPOS_PLAZO).setRequired(false),
      form.addListItem().setTitle(U.ingreso).setHelpText("Opcional.").setChoiceValues(PERIODOS_INGRESO).setRequired(false),
      form.addMultipleChoiceItem().setTitle(U.estado).setChoiceValues(["Confirmada", "Estimada"]).setRequired(true),
    ];
  } else {
    items = [U.fecha, U.hora, U.zona, U.tipo, U.ingreso, U.estado].map((t) => buscar(t)).filter(Boolean);
  }

  // Orden: … Mes de cierre | Fecha de cierre | Fecha límite 1 …
  let pos = reintentar(() => mes.getIndex()) + 1;
  [pb, ...items].forEach((it) => {
    reintentar(() => { if (it.getIndex() !== pos) form.moveItem(it.getIndex(), pos); });
    pos++;
  });

  // Navegación: después de "Mes de cierre" y de "Fecha de cierre" se envía; nunca se cae en la ronda 1.
  comoPB(pb).setGoToPage(FormApp.PageNavigationType.SUBMIT);
  comoPB(pb1).setGoToPage(FormApp.PageNavigationType.SUBMIT);
  const tipo = comoMC(tipoItem);
  tipo.setHelpText("Fecha única: una sola fecha de cierre. Varias rondas o fechas: registras cada ronda o plazo.").setChoices([
    tipo.createChoice(TIPOS_CIERRE[0], comoPB(pb)),
    tipo.createChoice(TIPOS_CIERRE[1], comoPB(pb1)),
    tipo.createChoice(TIPOS_CIERRE[2], comoPB(pbMes)),
    tipo.createChoice(TIPOS_CIERRE[3], FormApp.PageNavigationType.SUBMIT),
    tipo.createChoice(TIPOS_CIERRE[4], FormApp.PageNavigationType.SUBMIT),
  ]);
  comoPB(pb1).setHelpText("Registra una ronda o plazo por sección (nominación, admisión, documentos…).");

  if (desdeActualizar === true) return;
  SpreadsheetApp.flush();
  Utilities.sleep(8000); // que aparezcan las columnas nuevas en la hoja
  if (encabezados(SpreadsheetApp.openById(IDS_HOJA).getSheetByName("Respuestas")).indexOf(U.fecha) < 0) {
    Logger.log("Formulario listo. Las columnas nuevas todavía no aparecen en la hoja: espera un minuto y ejecuta prepararHoja.");
    return;
  }
  prepararHoja();
  Logger.log("Listo: \"Fecha única\" ahora lleva a \"Fecha de cierre\".");
}

// Ejecutar UNA vez (16 de octubre): deja una sola columna "Fecha de cierre".
// Guarda la columna vieja "Fecha de cierre" (K) y "Fechas de cierre" (M) en la pestaña "Respaldo fechas viejas",
// borra esas dos columnas y renombra "Fecha límite" (la que llena "Fecha única") a "Fecha de cierre".
function unificarFechaCierre() {
  const libro = SpreadsheetApp.openById(IDS_HOJA);
  const hoja = libro.getSheetByName("Respuestas");
  const NUEVA = "Fecha límite";
  let cab = encabezados(hoja);
  const iNueva = cab.indexOf(NUEVA);
  const viejas = cab.map((h, i) => (h === N.fechaVieja || h === N.variasVieja) && i !== iNueva ? i : -1).filter((i) => i >= 0);
  if (iNueva < 0) throw new Error("No encuentro la columna \"" + NUEVA + "\". ¿Ya se ejecutó unificarFechaCierre?");
  const filas = hoja.getLastRow() - 1;

  if (viejas.length && filas > 0) {
    const datos = hoja.getRange(2, 1, filas, cab.length).getValues();
    // Respaldo antes de borrar
    let resp = libro.getSheetByName("Respaldo fechas viejas");
    if (!resp) resp = libro.insertSheet("Respaldo fechas viejas");
    resp.clear();
    const iId = cab.indexOf("id_beca"), iNom = cab.indexOf(N.nombre);
    const filasResp = [["id_beca", "Nombre de la beca"].concat(viejas.map((i) => cab[i]))];
    datos.forEach((r) => { if (viejas.some((i) => String(r[i]).trim())) filasResp.push([r[iId], r[iNom]].concat(viejas.map((i) => r[i]))); });
    resp.getRange(1, 1, filasResp.length, filasResp[0].length).setValues(filasResp);

    // Borrar las columnas viejas (de derecha a izquierda para no correr índices)
    viejas.sort((a, b) => b - a).forEach((i) => hoja.deleteColumn(i + 1));
    Logger.log("Columnas borradas: " + viejas.map((i) => cab[i]).join(", ") + ". Respaldo en la pestaña \"Respaldo fechas viejas\".");
  }

  // Renombrar la pregunta y su columna
  const form = FormApp.openById(IDS_FORMULARIO);
  const item = form.getItems(FormApp.ItemType.DATE).find((i) => i.getTitle() === NUEVA);
  if (item) item.setTitle(U.fecha);
  cab = encabezados(hoja);
  if (cab.indexOf(NUEVA) >= 0) hoja.getRange(1, cab.indexOf(NUEVA) + 1).setValue(U.fecha);
  SpreadsheetApp.flush();
  prepararHoja(); // listas, plazos y Publicadas
  Logger.log("Listo: una sola columna \"" + U.fecha + "\".");
}

function actualizarCampos() {
  const form = FormApp.openById(IDS_FORMULARIO);
  const T_ = FormApp.ItemType;
  const comoPB = (x) => x.asPageBreakItem ? x.asPageBreakItem() : x;   // Item genérico o PageBreakItem
  const comoMC = (x) => x.asMultipleChoiceItem ? x.asMultipleChoiceItem() : x;
  // Google Forms a veces falla si se le piden las preguntas muchas veces seguidas: se leen una vez y se reintenta.
  const leerItems = () => {
    for (let intento = 1; ; intento++) {
      try { return form.getItems().map((i) => ({ i, titulo: i.getTitle(), tipo: i.getType() })); }
      catch (e) { if (intento >= 5) throw e; Utilities.sleep(3000 * intento); }
    }
  };
  let cache = leerItems();
  const buscar = (titulo, tipo) => { const x = cache.find((c) => c.titulo === titulo && (!tipo || c.tipo === tipo)); return x ? x.i : null; };
  const texto = (titulo, ayuda, validacion) => {
    let it = buscar(titulo, T_.TEXT);
    if (!it) it = form.addTextItem().setTitle(titulo).setRequired(false);
    if (validacion) (it.asTextItem ? it.asTextItem() : it).setValidation(validacion); // Item genérico o TextItem
    if (ayuda) it.setHelpText(ayuda);
    return it;
  };
  const casillas = (titulo, ayuda, opciones) => buscar(titulo, T_.CHECKBOX) ||
    form.addCheckboxItem().setTitle(titulo).setHelpText(ayuda).setChoiceValues(opciones).setRequired(false);
  const lista = (titulo, ayuda, opciones, obligatoria) => buscar(titulo, T_.LIST) ||
    form.addListItem().setTitle(titulo).setHelpText(ayuda).setChoiceValues(opciones).setRequired(!!obligatoria);

  form.setDescription("Formulario para cargar y corregir becas del buscador de Becas para Hondureños.");

  // País -> País de destino (mismo item, mismas respuestas)
  const pais = buscar(N.paisViejo) || buscar(N.pais);
  pais.setTitle(N.pais).setHelpText("País donde se estudia. Si son varios, elige una región o \"Varios países\" (al inicio de la lista).");

  const nacionalidades = buscar(N.nacionalidades); // ya no se usa (8 de noviembre): si se borró del formulario, no se vuelve a crear
  const posgrado = casillas(N.posgrado, "Solo si el nivel incluye posgrado.", TIPOS_POSGRADO);
  const edad = FormApp.createTextValidation().requireNumberBetween(10, 100).setHelpText("Escribe solo el número de años.").build();
  const edadMin = texto(N.edadMin, "Si aplica. Solo el número.", edad);
  const edadMax = texto(N.edadMax, "Si aplica. Solo el número.", edad);
  const duracion = buscar(N.duracion, T_.MULTIPLE_CHOICE) || form.addMultipleChoiceItem().setTitle(N.duracion)
    .setHelpText("Cuánto tiempo cubre la financiación.").setChoiceValues(DURACIONES).showOtherOption(true).setRequired(false);
  const importe = casillas(N.importe, "Si el monto cambia según la persona, marca de qué depende.", IMPORTE_SEGUN);
  const plazas = texto(N.plazas, "Solo si la convocatoria lo publica.",
    FormApp.createTextValidation().requireWholeNumber().setHelpText("Escribe solo el número.").build());
  const verificado = buscar(N.verificado, T_.DATE) || form.addDateItem().setTitle(N.verificado)
    .setHelpText("Fecha en que revisaste la convocatoria oficial.").setRequired(true);
  buscar(N.idioma).setHelpText("Solo idiomas de enseñanza confirmados para los programas elegibles.");

  // Secciones de fechas límite
  const pbMes = buscar(N.mes, T_.PAGE_BREAK);
  const mes = buscar(N.mes, T_.LIST);
  const secciones = [];
  for (let n = 1; n <= MAX_PLAZOS; n++) {
    const p = P(n);
    let pb = buscar(p.seccion, T_.PAGE_BREAK);
    const items = [];
    if (!pb) {
      pb = form.addPageBreakItem().setTitle(p.seccion)
        .setHelpText(n === 1 ? "Si hay varias fechas (rondas, nominación, admisión, documentos…), registra una por sección." : "");
      items.push(form.addTextItem().setTitle(p.ronda).setHelpText("Ej.: Early Decision, Regular Decision, Segunda convocatoria. Opcional.").setRequired(false));
      items.push(form.addListItem().setTitle(p.tipo).setChoiceValues(TIPOS_PLAZO).setRequired(true));
      items.push(form.addDateItem().setTitle(p.fecha).setRequired(true));
      items.push(form.addTimeItem().setTitle(p.hora).setHelpText("Solo si se publica.").setRequired(false));
      items.push(form.addTextItem().setTitle(p.zona).setHelpText("Solo si se publica la hora. Ej.: hora de Honduras, CET, ET (EE. UU.).").setRequired(false));
      items.push(form.addListItem().setTitle(p.ingreso).setChoiceValues(PERIODOS_INGRESO).setRequired(false));
      items.push(form.addMultipleChoiceItem().setTitle(p.estado).setChoiceValues(["Confirmada", "Estimada"]).setRequired(true));
      if (n < MAX_PLAZOS) items.push(form.addMultipleChoiceItem().setTitle(p.otra).setRequired(true));
    } else {
      [p.ronda, p.tipo, p.fecha, p.hora, p.zona, p.ingreso, p.estado, p.otra].forEach((t) => { const it = buscar(t); if (it) items.push(it); });
    }
    secciones.push({ pb, items });
  }
  // Navegación: "¿Agregar otra?" Sí -> siguiente sección, No -> enviar. Después de "Mes de cierre" -> enviar.
  secciones.forEach((s, i) => {
    const otra = s.items.find((it) => it.getTitle() === P(i + 1).otra);
    if (otra) {
      const mc = comoMC(otra);
      mc.setChoices([mc.createChoice("Sí", comoPB(secciones[i + 1].pb)), mc.createChoice("No", FormApp.PageNavigationType.SUBMIT)]);
    }
  });
  comoPB(secciones[0].pb).setGoToPage(FormApp.PageNavigationType.SUBMIT);

  // Tipo de cierre con las opciones nuevas
  if (!pbMes || !mes) throw new Error("No encontré la sección \"" + N.mes + "\" en el formulario.");
  const tipo = comoMC(buscar(N.tipo, T_.MULTIPLE_CHOICE));
  tipo.setHelpText("Fecha única y Varias rondas te llevan a registrar cada fecha límite.").setChoices([
    tipo.createChoice(TIPOS_CIERRE[0], comoPB(secciones[0].pb)),
    tipo.createChoice(TIPOS_CIERRE[1], comoPB(secciones[0].pb)),
    tipo.createChoice(TIPOS_CIERRE[2], comoPB(pbMes)),
    tipo.createChoice(TIPOS_CIERRE[3], FormApp.PageNavigationType.SUBMIT),
    tipo.createChoice(TIPOS_CIERRE[4], FormApp.PageNavigationType.SUBMIT),
  ]);

  // Quitar las preguntas de fecha viejas (sus columnas y datos se quedan en la hoja).
  // Va después de cambiar "Tipo de cierre": Forms no deja borrar una sección a la que todavía apunta una opción.
  [buscar(N.variasVieja, T_.TEXT), buscar("Varias fechas de cierre", T_.PAGE_BREAK)].forEach((it) => { if (it) form.deleteItem(it); });

  // Orden final de la primera sección y luego las secciones
  const primero = [buscar(TITULO_CORREGIR), buscar(N.nombre), pais, nacionalidades, buscar(N.nivel), posgrado, buscar(N.area),
    buscar(N.subarea), edadMin, edadMax, buscar(N.cobertura), duracion, importe, plazas, buscar(N.idioma), buscar(N.modalidad),
    buscar(N.enlace), verificado].filter(Boolean);
  const ids = new Set();
  const orden = [...primero];
  const resto = [tipo, pbMes, mes].concat(...secciones.map((s) => [s.pb, ...s.items]));
  [...orden, ...resto].forEach((it) => ids.add(it.getId()));
  leerItems().forEach(({ i: it }) => { if (!ids.has(it.getId())) orden.push(it); }); // preguntas que no conozco: quedan en la 1ª sección
  [...orden, ...resto].forEach((it, i) => {
    for (let intento = 1; ; intento++) {
      try { if (it.getIndex() !== i) form.moveItem(it.getIndex(), i); break; }
      catch (e) { if (intento >= 5) throw e; Utilities.sleep(3000 * intento); }
    }
  });

  agregarFechaUnica(true); // "Fecha única" va a su propia sección (sin rondas)
  if (buscar(TITULO_ACCION)) activarArchivar(true); // el orden de arriba movió "¿Qué quieres hacer?": se vuelve a poner primero
  SpreadsheetApp.flush();
  Utilities.sleep(6000); // que aparezcan las columnas nuevas en la hoja
  prepararHoja();
  Logger.log("Listo: formulario y hoja actualizados.");
}

// Se puede volver a ejecutar sola: encabezados, listas, valores viejos de "Tipo de cierre", columna plazos y Publicadas.
function prepararHoja() {
  const libro = SpreadsheetApp.openById(IDS_HOJA);
  const hoja = libro.getSheetByName("Respuestas");
  let cab = encabezados(hoja);
  if (cab.indexOf(N.paisViejo) >= 0 && cab.indexOf(N.pais) < 0) hoja.getRange(1, cab.indexOf(N.paisViejo) + 1).setValue(N.pais);
  cab = encabezados(hoja);
  const faltan = [N.verificado, P(1).fecha].filter((t) => cab.indexOf(t) < 0);
  if (faltan.length) throw new Error("Todavía no aparecen en la hoja: " + faltan.join(", ") + ". Espera un minuto y ejecuta prepararHoja.");

  // Columna "plazos" (la llena el script) al final
  if (cab.indexOf("plazos") < 0) {
    const n = hoja.getLastColumn() + 1;
    if (hoja.getMaxColumns() < n) hoja.insertColumnsAfter(hoja.getMaxColumns(), 1);
    hoja.getRange(1, n).setValue("plazos").setFontWeight("bold").setBackground("#e8f0fe");
    cab = encabezados(hoja);
  }
  const ultima = hoja.getMaxRows();
  const col = (t) => cab.indexOf(t) + 1;
  const regla = (t, opciones, libre) => { if (col(t)) hoja.getRange(2, col(t), ultima - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(opciones, true).setAllowInvalid(!!libre).build()); };
  const fecha = (t) => { if (col(t)) hoja.getRange(2, col(t), ultima - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build()); };

  // Tipo de cierre: primero la lista nueva (la vieja rechazaría los valores nuevos), luego valores viejos -> nuevos
  regla(N.tipo, TIPOS_CIERRE.concat(Object.keys(CIERRE_VIEJO_A_NUEVO)), true);
  const iTipo = col(N.tipo);
  const tipos = hoja.getRange(2, iTipo, ultima - 1, 1).getValues();
  let cambio = false;
  tipos.forEach((r) => { const nuevo = CIERRE_VIEJO_A_NUEVO[String(r[0]).trim().toLowerCase()]; if (nuevo) { r[0] = nuevo; cambio = true; } });
  if (cambio) hoja.getRange(2, iTipo, ultima - 1, 1).setValues(tipos);

  regla(N.tipo, TIPOS_CIERRE);
  regla(N.posgrado, TIPOS_POSGRADO, true);
  regla(N.duracion, DURACIONES, true);
  regla(N.importe, IMPORTE_SEGUN, true);
  regla(N.mes, MESES_CIERRE);
  fecha(N.verificado);
  regla(U.tipo, TIPOS_PLAZO); regla(U.ingreso, PERIODOS_INGRESO); regla(U.estado, ["Confirmada", "Estimada"]); fecha(U.fecha);
  for (let n = 1; n <= MAX_PLAZOS; n++) {
    regla(P(n).tipo, TIPOS_PLAZO); regla(P(n).ingreso, PERIODOS_INGRESO); regla(P(n).estado, ["Confirmada", "Estimada"]); fecha(P(n).fecha);
  }

  recalcularPlazos();
  rehacerPublicadas();
}

// Junta las columnas de fechas límite de una fila en un texto JSON para el buscador.
function plazosDeFila(cab, fila) {
  const zona = zonaHoja();
  const v = (t) => { const i = cab.indexOf(t); return i < 0 ? "" : fila[i]; };
  const dia = (x) => x instanceof Date ? Utilities.formatDate(x, zona, "yyyy-MM-dd") : String(x || "").trim();
  const hora = (x) => x instanceof Date ? Utilities.formatDate(x, zona, "HH:mm") : String(x || "").trim();
  const lista = [];
  const fu = dia(v(U.fecha));
  if (fu) lista.push({ ronda: "", tipo: String(v(U.tipo)).trim(), fecha: fu, hora: hora(v(U.hora)),
    zona: String(v(U.zona)).trim(), ingreso: String(v(U.ingreso)).trim(), estado: String(v(U.estado)).trim() });
  for (let n = 1; n <= MAX_PLAZOS; n++) {
    const p = P(n);
    const f = dia(v(p.fecha));
    if (!f) continue;
    lista.push({ ronda: String(v(p.ronda)).trim(), tipo: String(v(p.tipo)).trim(), fecha: f, hora: hora(v(p.hora)),
      zona: String(v(p.zona)).trim(), ingreso: String(v(p.ingreso)).trim(), estado: String(v(p.estado)).trim() });
  }
  return lista.length ? JSON.stringify(lista) : "";
}
function escribirPlazos(hoja, fila) {
  const cab = encabezados(hoja);
  const iPl = cab.indexOf("plazos");
  if (iPl < 0) return;
  const valores = hoja.getRange(fila, 1, 1, cab.length).getValues()[0];
  hoja.getRange(fila, iPl + 1).setValue(plazosDeFila(cab, valores));
}
function recalcularPlazos() {
  const hoja = SpreadsheetApp.openById(IDS_HOJA).getSheetByName("Respuestas");
  const cab = encabezados(hoja);
  const iPl = cab.indexOf("plazos");
  const filas = hoja.getLastRow() - 1;
  if (iPl < 0 || filas < 1) return;
  const datos = hoja.getRange(2, 1, filas, cab.length).getValues();
  hoja.getRange(2, iPl + 1, filas, 1).setValues(datos.map((r) => [plazosDeFila(cab, r)]));
}

// Pestaña Publicadas: solo aprobadas, con los encabezados que lee config.js. Se arma por nombre de columna.
function rehacerPublicadas() {
  const libro = SpreadsheetApp.openById(IDS_HOJA);
  const cab = encabezados(libro.getSheetByName("Respuestas"));
  const r = (t) => { const i = cab.indexOf(t); if (i < 0) throw new Error("Falta la columna " + t); const L = letra(i + 1); return "Respuestas!" + L + ":" + L; };
  const comoTexto = (t) => "IF(ISNUMBER(" + r(t) + "),TEXT(" + r(t) + ',"yyyy-mm-dd"),' + r(t) + ")";
  const columnas = [
    ["nombre", r(N.nombre)], ["pais", r(N.pais)], ["nivel", r(N.nivel)],
    ["tipo_posgrado", r(N.posgrado)], ["area", r(N.area)], ["subarea", r(N.subarea)], ["edad_min", r(N.edadMin)],
    ["edad_max", r(N.edadMax)], ["cobertura", r(N.cobertura)], ["duracion", r(N.duracion)], ["importe_segun", r(N.importe)],
    ["plazas", r(N.plazas)], ["idioma", r(N.idioma)], ["modalidad", r(N.modalidad)], ["tipo_cierre", r(N.tipo)],
    ["cierre", r(N.mes)], // las fechas van en "plazos"; aquí solo queda el mes de "Solo mes"
    ["plazos", r("plazos")], ["enlace", r(N.enlace)], ["verificado", comoTexto(N.verificado)],
    ["aprobada", "IF(" + r("aprobada") + '=TRUE,"sí","")'], ["id", r("id_beca")],
  ];
  const nombres = "{" + columnas.map((c) => '"' + c[0] + '"').join(",") + "}";
  const vacios = "{" + columnas.map(() => '""').join(",") + "}";
  const formula = "=" + "{" + nombres.slice(1, -1) + ";IFERROR(ARRAYFORMULA(FILTER({" + columnas.map((c) => c[1]).join(",") + "}," +
    r("aprobada") + "=TRUE," + r(N.nombre) + '<>"")),' + vacios + ")}";
  const pub = libro.getSheetByName("Publicadas");
  pub.getRange("A1").setFormula(formula);
}

/* =====================================================================================================
 * CORREGIR SIN EMPEZAR DE CERO (15 de octubre)
 *
 * Google Forms no puede llenarse solo cuando alguien escribe un ID dentro del formulario. Lo que sí
 * permite es abrirlo con un ENLACE PRELLENADO: todas las respuestas de la beca ya puestas, incluido su ID.
 * Jenny cambia solo lo que haga falta y envía; el resto sigue igual que antes (la fila se reemplaza).
 *
 *   1. Cada correo de "Beca registrada" o "Beca corregida" trae el enlace "Corregir esta beca".
 *   2. Para cualquier beca (también las 461 de las guías) hay una página "Corregir beca": se escribe el ID
 *      y aparece el botón que abre el formulario prellenado.
 *
 * Activar la página (una sola vez): Implementar > Nueva implementación > tipo "App web" >
 *   Ejecutar como: Yo · Quién tiene acceso: Cualquier usuario con una cuenta de Google > Implementar.
 *   Copia la URL que termina en /exec y pégala en URL_PAGINA_CORREGIR (abajo). Guardar.
 * ===================================================================================================== */
const URL_PAGINA_CORREGIR = ""; // p. ej. "https://script.google.com/macros/s/AKfycbyCg4jQCpvE5na4slb-oAwOLy7PCYGOqJgb9BS2S1P05OlZCYU_DCOHJR8eNVixiyZi/exec"

// Arma el enlace prellenado para una fila de la hoja. Devuelve { url, faltan: [títulos que no se pudieron poner] }.
function enlaceParaCorregir(cab, fila, form) {
  form = form || FormApp.openById(IDS_FORMULARIO);
  const T_ = FormApp.ItemType;
  const zona = zonaHoja();
  const respuesta = form.createResponse();
  const faltan = [];
  const id = String(fila[cab.indexOf("id_beca")] || "").trim();

  form.getItems().forEach((item) => {
    const titulo = item.getTitle();
    const tipo = item.getType();
    let v;
    const deUnica = Object.keys(U).find((k) => U[k] === titulo && k !== "seccion");
    if (titulo === TITULO_ACCION) v = ACCION_REGISTRAR;
    else if (titulo === TITULO_CODIGO) return;
    else if (titulo === TITULO_CORREGIR) v = id;
    else if (deUnica && !valorDe(cab, fila, U.fecha) && valorDe(cab, fila, N.tipo) === TIPOS_CIERRE[0])
      v = valorDe(cab, fila, P(1)[deUnica]) || (deUnica === "fecha" ? valorDe(cab, fila, N.fechaVieja) : ""); // datos de antes del cambio
    else if (titulo === N.verificado) v = new Date(); // al corregir se está verificando de nuevo
    else {
      const i = cab.indexOf(titulo);
      if (i < 0) return;
      v = fila[i];
    }
    if (v === "" || v === null || v === undefined) return;
    try {
      let r;
      if (tipo === T_.TEXT) r = item.asTextItem().createResponse(String(v instanceof Date ? Utilities.formatDate(v, zona, "dd/MM/yyyy") : v));
      else if (tipo === T_.PARAGRAPH_TEXT) r = item.asParagraphTextItem().createResponse(String(v));
      else if (tipo === T_.MULTIPLE_CHOICE) r = item.asMultipleChoiceItem().createResponse(String(v));
      else if (tipo === T_.LIST) r = item.asListItem().createResponse(String(v));
      else if (tipo === T_.CHECKBOX) {
        const it = item.asCheckboxItem();
        r = it.createResponse(separarOpciones(String(v), it.getChoices().map((c) => c.getValue())));
      }
      else if (tipo === T_.DATE) r = item.asDateItem().createResponse(v instanceof Date ? mismoDia(v, zona) : new Date(v));
      else if (tipo === T_.TIME) {
        const d = v instanceof Date ? v : null;
        const m = d ? null : String(v).match(/(\d{1,2}):(\d{2})/);
        r = item.asTimeItem().createResponse(d ? Number(Utilities.formatDate(d, zona, "H")) : Number(m[1]),
                                             d ? Number(Utilities.formatDate(d, zona, "m")) : Number(m[2]));
      }
      else if (tipo === T_.SCALE) r = item.asScaleItem().createResponse(Number(v));
      if (r) respuesta.withItemResponse(r);
    } catch (e) {
      faltan.push(titulo); // valor que ya no existe en las opciones, etc.: hay que llenarlo a mano
    }
  });
  return { url: respuesta.toPrefilledUrl(), faltan };
}

function valorDe(cab, fila, titulo) { const i = cab.indexOf(titulo); return i < 0 ? "" : fila[i]; }

// "Alemania, Corea del Sur, Otro valor" -> ["Alemania", "Corea del Sur", "Otro valor"], respetando opciones que llevan comas.
function separarOpciones(texto, opciones) {
  const partes = texto.split(", ").map((s) => s.trim()).filter(Boolean);
  const elegidas = [], otras = [];
  for (let i = 0; i < partes.length; i++) {
    let hasta = -1;
    for (let j = partes.length - 1; j >= i; j--) if (opciones.includes(partes.slice(i, j + 1).join(", "))) { hasta = j; break; }
    if (hasta >= 0) { elegidas.push(partes.slice(i, hasta + 1).join(", ")); i = hasta; }
    else otras.push(partes[i]);
  }
  if (otras.length) elegidas.push(otras.join(", ")); // va a "Otro" si la pregunta lo permite
  return elegidas;
}

// Texto para los correos: enlace de la beca que se acaba de guardar.
function textoEnlaceCorregir(hoja, filaNum) {
  try {
    const cab = encabezados(hoja);
    const fila = hoja.getRange(filaNum, 1, 1, cab.length).getValues()[0];
    const { url, faltan } = enlaceParaCorregir(cab, fila);
    return "\n\nPara corregirla sin llenar todo de nuevo, abre este enlace (ya trae todos los datos), cambia lo que haga falta y envía:\n" +
      url + (faltan.length ? "\nOjo: estos campos no se pudieron prellenar y hay que llenarlos: " + faltan.join(", ") : "") +
      (URL_PAGINA_CORREGIR ? "\n\nPara corregir cualquier otra beca, busca su ID aquí: " + URL_PAGINA_CORREGIR : "");
  } catch (e) {
    console.error("No se pudo armar el enlace para corregir: " + e);
    return "";
  }
}

// Página "Corregir beca" (app web).
function doGet(e) {
  const id = String((e && e.parameter && e.parameter.id) || "").trim().toUpperCase();
  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  let cuerpo = "";
  if (id) {
    const hoja = SpreadsheetApp.openById(IDS_HOJA).getSheetByName("Respuestas");
    const cab = encabezados(hoja);
    const iId = cab.indexOf("id_beca");
    const filas = hoja.getRange(2, 1, Math.max(hoja.getLastRow() - 1, 1), cab.length).getValues();
    const fila = filas.find((f) => String(f[iId]).trim().toUpperCase() === id);
    if (!fila) cuerpo = '<p class="error">No encontramos el ID ' + esc(id) + '. Revisa que esté bien escrito (por ejemplo BPH-0463).</p>';
    else {
      const { url, faltan } = enlaceParaCorregir(cab, fila);
      cuerpo = '<p><b>' + esc(id) + '</b> · ' + esc(fila[cab.indexOf(N.nombre)]) + '</p>' +
        '<a class="boton" href="' + esc(url) + '" target="_blank" rel="noopener">Abrir el formulario con sus datos</a>' +
        '<p class="nota">Cambia solo lo que haga falta y envía. La beca se reemplaza con lo que envíes.</p>' +
        (faltan.length ? '<p class="error">Estos campos no se pudieron prellenar; llénalos en el formulario: ' + esc(faltan.join(", ")) + '</p>' : '');
    }
  }
  const html = '<!doctype html><html><head><meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<style>body{font-family:Arial,sans-serif;max-width:560px;margin:40px auto;padding:0 16px;color:#16305c}' +
    'input{font-size:18px;padding:10px;width:200px;text-transform:uppercase}button,.boton{font-size:18px;padding:10px 18px;' +
    'background:#003cd2;color:#fff;border:0;border-radius:8px;text-decoration:none;display:inline-block;cursor:pointer}' +
    '.nota{color:#55627a}.error{color:#a12a2a}</style></head><body>' +
    '<h1>Corregir una beca</h1><form method="get" action="' + esc(ScriptApp.getService().getUrl()) + '" target="_top">' +
    '<input name="id" placeholder="BPH-0000" value="' + esc(id) + '" required> <button>Buscar</button></form>' + cuerpo +
    '</body></html>';
  return HtmlService.createHtmlOutput(html).setTitle("Corregir beca · BpH")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Prueba rápida desde el editor: muestra en el registro el enlace de una beca.
function probarEnlace() {
  const hoja = SpreadsheetApp.openById(IDS_HOJA).getSheetByName("Respuestas");
  const r = enlaceParaCorregir(encabezados(hoja), hoja.getRange(2, 1, 1, hoja.getLastColumn()).getValues()[0]);
  Logger.log(r.url);
  Logger.log("Sin prellenar: " + (r.faltan.join(", ") || "ninguno"));
}

/* =====================================================================================================
 * ARCHIVAR O PUBLICAR DESDE EL FORMULARIO (8 de noviembre): ejecutar UNA vez la función "activarArchivar".
 *
 * El formulario empieza con "¿Qué quieres hacer?":
 *   - Registrar una beca nueva o corregir una  -> sigue igual que antes (sección "Datos de la beca").
 *   - Archivar una beca (quitarla del buscador) -> solo pide el código y envía. La beca deja de verse en el
 *     buscador, pero sus datos se quedan en la hoja (columna "archivada" con la fecha, estado "Archivada").
 *   - Publicar o volver a publicar una beca     -> solo pide el código. Sirve para una beca archivada y
 *     también para publicar una "Pendiente de revisión" sin abrir la hoja.
 * Solo los correos de CORREOS_QUE_PUBLICAN pueden archivar o publicar; a cualquier otra persona le llega
 * un correo diciendo que no se cambió nada.
 *
 * También quita la regla que obligaba a escribir "BPH-0000" en "ID de beca a corregir", porque las becas
 * de las guías tienen códigos como B1AMCA001 o BMEUE+001 y no se podían corregir.
 *
 * Se puede volver a ejecutar sin duplicar nada. Si alguna vez se ejecuta "actualizarCampos", esta función
 * se vuelve a aplicar sola al final.
 * ===================================================================================================== */
function activarArchivar(desdeActualizar) {
  const form = FormApp.openById(IDS_FORMULARIO);
  const T_ = FormApp.ItemType;
  const reintentar = (f) => { for (let k = 1; ; k++) { try { return f(); } catch (e) { if (k >= 5) throw e; Utilities.sleep(3000 * k); } } };
  const items = reintentar(() => form.getItems());
  const buscar = (titulo, tipo) => items.find((i) => i.getTitle() === titulo && (!tipo || i.getType() === tipo)) || null;
  const mover = (it, a) => reintentar(() => { if (it.getIndex() !== a) form.moveItem(it.getIndex(), a); });
  // Lo que se acaba de crear ya es PageBreakItem/MultipleChoiceItem/TextItem; lo que se encontró es un Item genérico
  const comoPB = (x) => x.asPageBreakItem ? x.asPageBreakItem() : x;
  const comoMC = (x) => x.asMultipleChoiceItem ? x.asMultipleChoiceItem() : x;
  const comoTX = (x) => x.asTextItem ? x.asTextItem() : x;
  const ultimo = () => reintentar(() => form.getItems().length) - 1;

  // 1. Sección final: solo el código de la beca
  let pbArch = buscar(SECCION_ARCHIVAR, T_.PAGE_BREAK);
  if (!pbArch) pbArch = form.addPageBreakItem().setTitle(SECCION_ARCHIVAR)
    .setHelpText("Escribe el código que aparece al pie de la ficha en el buscador (por ejemplo BPH-0007 o B1AMCA001) y envía.");
  let codigo = buscar(TITULO_CODIGO, T_.TEXT);
  if (!codigo) codigo = form.addTextItem().setTitle(TITULO_CODIGO).setRequired(true);
  mover(pbArch, ultimo());
  mover(codigo, ultimo());
  // Quien llega al final de las fechas límite envía: nunca cae en esta sección por accidente
  comoPB(pbArch).setGoToPage(FormApp.PageNavigationType.SUBMIT);

  // 2. Primera sección: "¿Qué quieres hacer?" y después el salto a "Datos de la beca"
  let pbDatos = buscar(SECCION_DATOS, T_.PAGE_BREAK);
  if (!pbDatos) pbDatos = form.addPageBreakItem().setTitle(SECCION_DATOS);
  let accion = buscar(TITULO_ACCION, T_.MULTIPLE_CHOICE);
  if (!accion) accion = form.addMultipleChoiceItem().setTitle(TITULO_ACCION).setRequired(true);
  mover(accion, 0);
  mover(pbDatos, 1);
  const mc = comoMC(accion);
  mc.setChoices([
    mc.createChoice(ACCION_REGISTRAR, comoPB(pbDatos)),
    mc.createChoice(ACCION_ARCHIVAR, comoPB(pbArch)),
    mc.createChoice(ACCION_PUBLICAR, comoPB(pbArch)),
  ]);

  // 3. "ID de beca a corregir" acepta cualquier código (no solo BPH-0000)
  const corr = buscar(TITULO_CORREGIR, T_.TEXT);
  if (corr) comoTX(corr).clearValidation().setHelpText("Déjalo VACÍO si es una beca nueva. Para corregir una beca, escribe su código " +
    "(aparece al pie de la ficha en el buscador, por ejemplo BPH-0007 o B1AMCA001) y llena todo de nuevo: la beca se reemplaza con estos datos.");

  if (desdeActualizar === true) return;

  // 4. Hoja: columna "archivada", estado con "Archivada" y Publicadas armada por nombre de columna
  SpreadsheetApp.flush();
  Utilities.sleep(6000); // que aparezcan las columnas nuevas del formulario
  const hoja = SpreadsheetApp.openById(IDS_HOJA).getSheetByName("Respuestas");
  if (encabezados(hoja).indexOf(COL_ARCHIVADA) < 0) {
    const n = hoja.getLastColumn() + 1;
    if (hoja.getMaxColumns() < n) hoja.insertColumnsAfter(hoja.getMaxColumns(), 1);
    hoja.getRange(1, n).setValue(COL_ARCHIVADA).setFontWeight("bold").setBackground("#e8f0fe");
    hoja.getRange(2, n, hoja.getMaxRows() - 1, 1).setNumberFormat("dd/mm/yyyy");
  }
  asegurarEstado(hoja, true);
  rehacerPublicadas();
  Logger.log("Listo: el formulario ya permite archivar y publicar becas por código.");
}
