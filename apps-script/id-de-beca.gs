/**
 * IDs de beca y correcciones desde el formulario (reemplaza a ids-automaticos.gs).
 *
 * Qué hace:
 *   - Cada beca nueva que llega por el formulario recibe un ID (BPH-0463, BPH-0464, …) en la columna "id_beca".
 *   - Cada envío queda anotado en la "Bitácora del buscador de becas" (un Google Docs de solo lectura para Jenny),
 *     con los datos, el código y el enlace para corregir. Ya no se mandan correos de confirmación.
 *   - El ÚNICO correo que se envía es cuando otra persona (no BpH) registra una beca: les llega a CORREOS_AVISO_PENDIENTES.
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
 *   elige la función "configurarIds" > Ejecutar > acepta los permisos.
 *
 * Bitácora (9 de noviembre): pega este archivo completo, elige la función "crearBitacora" > Ejecutar >
 *   acepta los permisos nuevos (Google Docs y Drive). En "Registro de ejecución" aparece el enlace del documento.
 */
// Versión de este archivo: si la que corre en Apps Script no coincide con GitHub, ejecuta "verVersion".
const VERSION_SCRIPT = "2026-11-23 (género y seguridad)";
function verVersion() { Logger.log("Versión del script: " + VERSION_SCRIPT); }
const IDS_FORMULARIO = "1Eh02U-z-BddnM5JcXh4vhYgvagMu1H1c1AS2bWNwfm4";
const IDS_HOJA = "1rzRIiZHC14KqA4zYR-TMpgCVQ5k6dEJOI2XWYQsE8Qw";
const IDS_PREFIJO = "BPH-";
const TITULO_CORREGIR = "ID de beca a corregir";
// Correos que publican solos y pueden corregir por ID (en minúsculas).
// raudalesg99@gmail.com está solo para las pruebas: quitarlo cuando se le entregue a Jenny.
const CORREOS_QUE_PUBLICAN = ["becasparahn@gmail.com", "raudalesg99@gmail.com"];
// Quién recibe el aviso cuando otra persona (no BpH) registra una beca que queda pendiente de revisión.
// Es el único correo que envía el script. raudalesg99@gmail.com está para las pruebas: quitarlo al entregar.
const CORREOS_AVISO_PENDIENTES = ["becasparahn@gmail.com", "raudalesg99@gmail.com"];
// Bitácora: Google Docs donde se anota cada envío del formulario (el más reciente arriba).
const TITULO_BITACORA = "Bitácora del buscador de becas";
const CARPETA_BITACORA = "1JvVsCO7F_mM18OsN3Gqceed4ysDh0-Yp"; // carpeta "Buscador de becas BpH" en Drive
const LECTORES_BITACORA = ["becasparahn@gmail.com"]; // pueden ver la bitácora pero no editarla
const LIMITE_BITACORA = 800000; // caracteres; al pasarlo, lo viejo se guarda en una copia y la bitácora sigue con el mismo enlace
const COLUMNAS_PROPIAS = ["aprobada", "estado", "id_beca", "plazos", "archivada"]; // columnas de la hoja que no vienen del formulario

// Archivar o publicar desde el formulario (8 de noviembre): ver "activarArchivar" más abajo.
const TITULO_ACCION = "¿Qué quieres hacer?";
const ACCION_REGISTRAR = "Registrar una beca nueva o corregir una"; // opción vieja (respuestas anteriores al 8 de noviembre)
const ACCION_NUEVA = "Añadir una beca nueva";
const ACCION_CORREGIR = "Corregir una beca";
const SECCION_CORREGIR = "Corregir una beca";
const AYUDA_CORREGIR = "Para corregir una beca, escribe su código (aparece al pie de la ficha en el buscador, por ejemplo BPH-0007 o B1AMCA001) " +
  "y llena todo de nuevo: la beca se reemplaza con estos datos.";
const ACCION_ARCHIVAR = "Archivar una beca (quitarla del buscador)";
const ACCION_PUBLICAR = "Publicar o volver a publicar una beca";
const TITULO_CODIGO = "Código de la beca";
const SECCION_DATOS = "Datos de la beca";
const SECCION_ARCHIVAR = "Archivar o publicar una beca";
const COL_ARCHIVADA = "archivada"; // fecha en que se archivó (vacía = no archivada)
const SOLO_FORMULARIO = [TITULO_ACCION, TITULO_CODIGO]; // preguntas que no son datos de la beca
// Corrección enviada por alguien que no es de BpH: queda en su propia fila, sin publicar, con este código
// (por ejemplo "BPASAE001_1 (corrección pendiente)"). Al aprobarla reemplaza la beca original y conserva su código.
const SUFIJO_PENDIENTE = " (corrección pendiente)";
function codigoPendiente(id) { return String(id).trim().toUpperCase() + SUFIJO_PENDIENTE.toUpperCase(); }
function codigoOriginal(id) {
  const t = String(id).trim();
  return t.toUpperCase().endsWith(SUFIJO_PENDIENTE.toUpperCase()) ? t.slice(0, t.length - SUFIJO_PENDIENTE.length).trim() : t;
}
function idsDe(hoja, iId) {
  return hoja.getRange(2, iId + 1, Math.max(hoja.getLastRow() - 1, 1), 1).getValues().map((r) => String(r[0]).trim().toUpperCase());
}
// Seguridad: un texto que empieza con = + - @ se guarda como texto (con apóstrofo delante), nunca como fórmula.
// Así nadie puede colar una fórmula en una corrección (p. ej. para juntar los correos de la hoja y publicarlos).
function seguro(v) { return typeof v === "string" && /^\s*[=+\-@]/.test(v) ? "'" + v : v; }
// Copia los datos de la beca (no las columnas propias ni las del formulario) a la fila destino.
function copiarDatos(hoja, cab, valores, destino, iCorr) {
  cab.forEach((h, i) => {
    if (COLUMNAS_PROPIAS.includes(h) || SOLO_FORMULARIO.includes(h) || i === iCorr) return;
    hoja.getRange(destino, i + 1).setValue(seguro(valores[i]));
  });
}
// Si en la fila recién enviada quedó alguna fórmula en las columnas del formulario, se convierte en texto.
function neutralizarFormulas(hoja, fila, cab) {
  const formulas = hoja.getRange(fila, 1, 1, cab.length).getFormulas()[0];
  formulas.forEach((f, i) => {
    if (f && !COLUMNAS_PROPIAS.includes(cab[i])) hoja.getRange(fila, i + 1).setValue("'" + f);
  });
}

// Seguridad: busca la fila de ESTA respuesta por su marca temporal. El número de fila del disparador puede quedar
// viejo si otro envío borró filas antes (por ejemplo, tres becas archivadas seguidas). Devuelve -1 si no la encuentra.
function filaDeRespuesta(hoja, e) {
  const esperada = e.range.getRow();
  const marca = e.values && e.values[0] ? String(e.values[0]).trim() : "";
  if (!marca) return esperada;
  const ultima = hoja.getLastRow();
  if (esperada <= ultima && hoja.getRange(esperada, 1).getDisplayValue().trim() === marca) return esperada;
  if (ultima < 2) return -1;
  const marcas = hoja.getRange(2, 1, ultima - 1, 1).getDisplayValues();
  for (let i = marcas.length - 1; i >= 0; i--) if (marcas[i][0].trim() === marca) return i + 2;
  const t = new Date(marca).getTime(); // por si la hoja muestra la fecha con otro formato
  if (!isNaN(t)) {
    const fechas = hoja.getRange(2, 1, ultima - 1, 1).getValues();
    for (let i = fechas.length - 1; i >= 0; i--) if (fechas[i][0] instanceof Date && Math.abs(fechas[i][0].getTime() - t) < 1500) return i + 2;
  }
  return -1;
}

// Seguridad: solo se confía en el correo si el formulario pide iniciar sesión con Google (correo verificado).
// Se revisa como mucho cada 10 minutos para no gastar cuota.
function correoVerificado() {
  const cache = CacheService.getScriptCache();
  const guardado = cache.get("correoVerificado");
  if (guardado) return guardado === "si";
  let ok = false;
  try { ok = FormApp.openById(IDS_FORMULARIO).getEmailCollectionType() === FormApp.EmailCollectionType.VERIFIED; }
  catch (err) { console.error("No se pudo revisar cómo recoge correos el formulario: " + err); }
  cache.put("correoVerificado", ok ? "si" : "no", 600);
  return ok;
}

// Fecha de verificación en el futuro = casi seguro se escribió ahí la fecha de cierre. Se cambia a hoy y se avisa.
function revisarVerificacion(hoja, fila, cab) {
  const i = cab.indexOf(N.verificado);
  if (i < 0) return "";
  const v = hoja.getRange(fila, i + 1).getValue();
  if (!(v instanceof Date)) return "";
  const manana = new Date(); manana.setDate(manana.getDate() + 1);
  if (v.getTime() <= manana.getTime()) return "";
  hoja.getRange(fila, i + 1).setValue(new Date());
  return "⚠️ La «" + N.verificado + "» venía en el futuro (" + textoValor(v, zonaHoja()) + "), así que se cambió a hoy. " +
    "Si esa era la fecha de cierre, corrige la beca y ponla en «" + U.fecha + "» o en el plazo que corresponda.";
}
// Borra las correcciones pendientes de un código (de abajo hacia arriba). Devuelve cuántas borró.
function borrarPendientes(hoja, iId, id) {
  const pend = codigoPendiente(id);
  const ids = idsDe(hoja, iId);
  let n = 0;
  for (let i = ids.length - 1; i >= 0; i--) if (ids[i] === pend) { hoja.deleteRow(i + 2); n++; }
  return n;
}

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
  // Si ya existe "¿Qué quieres hacer?", esa pregunta va primero y "ID de beca a corregir" en "Datos de la beca"
  if (form.getItems().some((i) => i.getTitle() === TITULO_ACCION)) activarArchivar(true);

  // Pedir el correo verificado (iniciar sesión con Google) para saber quién publica
  form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED); // si falla, que se vea el error: nunca correos sin verificar
  form.setConfirmationMessage("¡Gracias! Tu respuesta quedó registrada.");

  // Disparador: un solo "alEnviar" (quita los de versiones anteriores)
  ScriptApp.getProjectTriggers()
    .filter((t) => ["alEnviar", "asignarIds"].includes(t.getHandlerFunction()))
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("alEnviar").forSpreadsheet(IDS_HOJA).onFormSubmit().create();

  SpreadsheetApp.flush();
  Utilities.sleep(4000); // que aparezcan las columnas nuevas
  completarIdsFaltantes();
  Logger.log("Listo: IDs automáticos y correcciones por ID activados.");
}

// Se ejecuta sola con cada respuesta del formulario.
function alEnviar(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(120000); // espera hasta 2 minutos a que termine el envío anterior
  try {
    procesarEnvio(e);
  } catch (err) {
    // Que el error no pase en silencio: queda en la bitácora con ⚠️ y Google avisa del fallo del disparador.
    anotar("Error al procesar un envío", "", "Algo falló y este envío no se procesó completo.\nDetalle técnico: " + err +
      "\nRevisa la hoja (la respuesta puede haber quedado sin código) o vuelve a enviarlo.");
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function procesarEnvio(e) {
  {
    const hoja = e.range.getSheet();
    const fila = filaDeRespuesta(hoja, e);
    if (fila < 2) {
      anotar("No se pudo procesar un envío", "", "No encontramos en la hoja la respuesta enviada el " + (e.values ? e.values[0] : "") +
        ". No se cambió nada. Si era tuya, vuelve a enviarla.");
      return;
    }
    const cab = encabezados(hoja);
    neutralizarFormulas(hoja, fila, cab);
    const avisoVerif = revisarVerificacion(hoja, fila, cab);
    const ancho = cab.length;
    const nueva = hoja.getRange(fila, 1, 1, ancho).getValues()[0];
    const c = (titulo) => cab.indexOf(titulo); // índice 0-based o -1
    const iId = c("id_beca"), iAprob = c("aprobada"), iCorr = c(TITULO_CORREGIR);
    const iCorreo = cab.findIndex((h) => /correo|email/i.test(h));
    const correo = iCorreo >= 0 ? String(nueva[iCorreo] || "").trim() : "";
    // Solo publica sola si el correo es de BpH Y el formulario exige correo verificado (si no, cualquiera podría escribirlo).
    const enLista = CORREOS_QUE_PUBLICAN.includes(correo.toLowerCase());
    const autorizada = enLista && correoVerificado();
    if (enLista && !autorizada) {
      anotar("Formulario sin correo verificado", "", "⚠️ El formulario no está pidiendo iniciar sesión con Google, así que no se puede " +
        "confirmar quién envía. Este envío quedó pendiente de revisión. Arréglalo en el formulario: Configuración > Respuestas > " +
        "Recopilar direcciones de correo electrónico > Verificadas.");
    }
    const iAccion = c(TITULO_ACCION);
    const accion = iAccion >= 0 ? String(nueva[iAccion] || "").trim() : "";
    const idPedido = iCorr >= 0 && accion !== ACCION_NUEVA ? String(nueva[iCorr] || "").trim().toUpperCase() : "";

    // Archivar o publicar: la respuesta no es una beca, solo cambia "aprobada" de la beca con ese código
    if (accion === ACCION_ARCHIVAR || accion === ACCION_PUBLICAR) {
      archivarOPublicar(hoja, fila, cab, nueva, accion === ACCION_ARCHIVAR, correo, autorizada);
      return;
    }

    if (idPedido && autorizada) {
      const destino = idsDe(hoja, iId).findIndex((v, i) => v === idPedido && i + 2 !== fila) + 2; // 1 si no se encontró
      if (destino >= 2) {
        const antes = hoja.getRange(destino, 1, 1, ancho).getValues()[0];
        // Sobrescribir la fila de ese ID con lo que vino del formulario (sin tocar aprobada/estado/id)
        copiarDatos(hoja, cab, nueva, destino, iCorr);
        if (iCorr >= 0) hoja.getRange(destino, iCorr + 1).clearContent();
        if (iAprob >= 0) hoja.getRange(destino, iAprob + 1).setValue(true); // se publica sola
        if (c(COL_ARCHIVADA) >= 0) hoja.getRange(destino, c(COL_ARCHIVADA) + 1).clearContent(); // corregida = ya no archivada
        escribirPlazos(hoja, destino);
        aprenderOpciones(cab, nueva);
        hoja.deleteRow(fila);
        const pendientes = borrarPendientes(hoja, iId, idPedido); // la corrección de otra persona ya quedó resuelta
        asegurarEstado(hoja);
        const filaFinal = idsDe(hoja, iId).indexOf(idPedido) + 2;
        anotar("Beca corregida", idPedido,
          ficha(cab, nueva, idPedido) + "\nCorregida por: " + correo + "\n" +
          textoCambios(cab, antes, nueva, "Cambios:") +
          (avisoVerif ? "\n" + avisoVerif : "") +
          (pendientes ? "\nSe quitó la corrección pendiente que había enviado otra persona para esta beca." : "") +
          lineaEnlaces(hoja, filaFinal), "corregida");
        return;
      }
    }

    // Corrección de alguien que no es de BpH: no se crea una beca nueva; queda pendiente ligada al código original
    const iOriginal = idPedido && !autorizada ? idsDe(hoja, iId).findIndex((v, i) => v === idPedido && i + 2 !== fila) : -1;
    if (iOriginal >= 0) {
      const original = hoja.getRange(iOriginal + 2, 1, 1, ancho).getValues()[0];
      hoja.getRange(fila, iId + 1).setValue(codigoPendiente(idPedido));
      if (iAprob >= 0) hoja.getRange(fila, iAprob + 1).setValue(false);
      escribirPlazos(hoja, fila);
      asegurarEstado(hoja);
      avisarPendiente(hoja, fila, idPedido, correo, "", original);
      return;
    }

    // Beca nueva (o corrección que no se pudo aplicar)
    let aviso = "";
    if (idPedido) aviso = "No encontramos el código " + idPedido + ", así que esta respuesta se guardó como beca nueva.\n";
    const nuevoId = siguienteId(hoja, iId);
    hoja.getRange(fila, iId + 1).setValue(nuevoId);
    if (iAprob >= 0) hoja.getRange(fila, iAprob + 1).setValue(autorizada); // solo se publica sola si la envía BpH
    escribirPlazos(hoja, fila);
    aprenderOpciones(cab, nueva);
    asegurarEstado(hoja);
    if (autorizada) {
      anotar("Beca registrada", nuevoId,
        aviso + ficha(cab, nueva, nuevoId) + "\nRegistrada por: " + correo + "\n" + detalleCorto(cab, nueva) +
        (avisoVerif ? "\n" + avisoVerif : "") +
        "\nYa está publicada en el buscador." + lineaEnlaces(hoja, fila));
    } else {
      avisarPendiente(hoja, fila, nuevoId, correo, aviso); // correo a BpH + anotación en la bitácora
    }
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
  const codigo = iCodigo >= 0 ? codigoOriginal(String(nueva[iCodigo] || "")).toUpperCase() : "";
  const verbo = archivar ? "archivar" : "publicar";
  hoja.deleteRow(fila);
  asegurarEstado(hoja);
  if (!autorizada) {
    anotar("No se pudo " + verbo, limpio(codigo),
      "Pedido por: " + (correo || "(sin correo)") + "\nSolo Becas para Hondureños puede archivar o publicar becas. No se cambió nada.");
    return;
  }
  let ids = idsDe(hoja, iId);
  // Publicar una beca con una corrección pendiente = aprobar esa corrección tal cual (la más reciente)
  let corregidaPor = "", cambiosAprobados = "";
  const iPend = !archivar && codigo ? ids.lastIndexOf(codigoPendiente(codigo)) : -1;
  if (iPend >= 0) {
    const ancho = cab.length, iCorr = cab.indexOf(TITULO_CORREGIR), iCorreo = cab.findIndex((h) => /correo|email/i.test(h));
    const pend = hoja.getRange(iPend + 2, 1, 1, ancho).getValues()[0];
    corregidaPor = iCorreo >= 0 ? String(pend[iCorreo] || "") : "";
    const iOrig = ids.indexOf(codigo);
    if (iOrig >= 0) {
      cambiosAprobados = textoCambios(cab, hoja.getRange(iOrig + 2, 1, 1, ancho).getValues()[0], pend, "Cambios aprobados:");
      copiarDatos(hoja, cab, pend, iOrig + 2, iCorr);
      if (iCorr >= 0) hoja.getRange(iOrig + 2, iCorr + 1).clearContent();
      escribirPlazos(hoja, iOrig + 2);
      borrarPendientes(hoja, iId, codigo);
    } else {
      hoja.getRange(iPend + 2, iId + 1).setValue(codigo); // la original ya no existe: la corrección pasa a ser la beca
    }
    ids = idsDe(hoja, iId);
  }
  const i = codigo ? ids.indexOf(codigo) : -1;
  if (i < 0) {
    anotar("No se pudo " + verbo, codigo,
      "Pedido por: " + correo + "\nNo encontramos el código " + codigo + ". Revisa que esté escrito igual que al pie de la ficha en el buscador. No se cambió nada.");
    return;
  }
  const destino = i + 2;
  const nombre = String(hoja.getRange(destino, iNombre + 1).getValue());
  hoja.getRange(destino, iAprob + 1).setValue(!archivar);
  if (iArch >= 0) {
    if (archivar) hoja.getRange(destino, iArch + 1).setValue(new Date());
    else hoja.getRange(destino, iArch + 1).clearContent();
  }
  anotar(archivar ? "Beca archivada" : "Beca publicada", codigo,
    ficha(cab, hoja.getRange(destino, 1, 1, cab.length).getValues()[0], codigo) + "\n" +
    (archivar
      ? "Ya no aparece en el buscador. Sus datos se guardan: para mostrarla de nuevo, elige \"" + ACCION_PUBLICAR + "\" en el formulario."
      : "Ya aparece en el buscador.") +
    (iPend >= 0 ? "\nSe aprobó tal cual la corrección propuesta por " + (corregidaPor || "otra persona") + "; la beca conserva su código." +
      (cambiosAprobados ? "\n" + cambiosAprobados : "") : "") +
    "\nCambio enviado por: " + correo, archivar ? "" : "publicada");
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
  texto += "\n\nBecas para Hondureños";
  const plano = texto.replace(RE_ENLACE, "$1 ($2)");
  const esc = (t) => t.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const html = esc(texto).replace(RE_ENLACE, (_, t, u) => '<a href="' + u.replace(/"/g, "&quot;") + '">' + t + "</a>")
    .replace(/\n/g, "<br>");
  MailApp.sendEmail(correo, asunto + " · Buscador de becas BpH", plano, { htmlBody: html });
}
// Lista "Pregunta: respuesta" de todo lo que se envió (sin marca temporal, correo ni columnas propias de la hoja).
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
const TIPOS_POSGRADO = ["Maestría", "Doctorado", "Postdoctorado", "Otros estudios de posgrado"];
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
  fechaVieja: "Fecha de cierre", variasVieja: "Fechas de cierre", genero: "Género",
};
// Para quién es la beca (23 de noviembre). Vacío o "Todos" = cualquier persona. "Otro" permite escribir p. ej. "Mujeres y personas no binarias".
const GENEROS = ["Todos", "Mujeres", "Hombres"];
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
    buscar(N.genero), buscar(N.enlace), verificado].filter(Boolean);
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
  // Con "Varias rondas o fechas" y plazos 1–4 llenos, la fecha única se ignora para no repetir fechas.
  let hayPlazos = false;
  for (let n = 1; n <= MAX_PLAZOS; n++) if (dia(v(P(n).fecha))) hayPlazos = true;
  const varias = String(v(N.tipo)).trim() === TIPOS_CIERRE[1] && hayPlazos;
  const fu = varias ? "" : dia(v(U.fecha));
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
  limpiarFechaUnica(hoja, cab, fila, valores);
  hoja.getRange(fila, iPl + 1).setValue(plazosDeFila(cab, valores));
}
// Si la beca es de "Varias rondas o fechas" y tiene plazos 1–4, vacía las columnas de fecha única
// (Fecha de cierre, Hora límite, Zona horaria, Tipo de plazo, Periodo de ingreso, Fecha confirmada o estimada).
function limpiarFechaUnica(hoja, cab, fila, valores) {
  const v = (t) => { const i = cab.indexOf(t); return i < 0 ? "" : valores[i]; };
  if (String(v(N.tipo)).trim() !== TIPOS_CIERRE[1]) return;
  let hayPlazos = false;
  for (let n = 1; n <= MAX_PLAZOS; n++) if (String(v(P(n).fecha) || "").trim()) hayPlazos = true;
  if (!hayPlazos) return; // si no hay plazos, la fecha única es la única que hay: no se borra
  [U.fecha, U.hora, U.zona, U.tipo, U.ingreso, U.estado].forEach((t) => {
    const i = cab.indexOf(t);
    if (i >= 0 && String(valores[i] || "").trim()) { hoja.getRange(fila, i + 1).clearContent(); valores[i] = ""; }
  });
}

// Columnas de la hoja que forman la columna "plazos" (la que lee el buscador).
function columnasDeFechas() {
  const t = [U.fecha, U.hora, U.zona, U.tipo, U.ingreso, U.estado];
  for (let n = 1; n <= MAX_PLAZOS; n++) { const p = P(n); t.push(p.ronda, p.tipo, p.fecha, p.hora, p.zona, p.ingreso, p.estado); }
  return t;
}

// Se ejecuta sola cuando alguien edita la hoja a mano. Si el cambio toca fechas límite en "Respuestas",
// rehace "plazos" de esas filas para que el buscador muestre la fecha nueva (sin pasar por el formulario).
function alEditarHoja(e) {
  try {
    const rango = e && e.range;
    if (!rango || rango.getSheet().getName() !== "Respuestas") return;
    const hoja = rango.getSheet();
    const cab = encabezados(hoja);
    const fechas = columnasDeFechas().map((t) => cab.indexOf(t) + 1).filter((c) => c > 0);
    const c1 = rango.getColumn(), c2 = rango.getLastColumn();
    if (!fechas.some((c) => c >= c1 && c <= c2)) return;
    const iPl = cab.indexOf("plazos");
    if (iPl < 0) return;
    const f1 = Math.max(2, rango.getRow()), f2 = Math.min(rango.getLastRow(), hoja.getLastRow());
    if (f2 < f1) return;
    const datos = hoja.getRange(f1, 1, f2 - f1 + 1, cab.length).getValues();
    hoja.getRange(f1, iPl + 1, datos.length, 1).setValues(datos.map((r) => [plazosDeFila(cab, r)]));
  } catch (err) {
    console.error("No se pudo actualizar plazos tras editar la hoja: " + err);
  }
}

// Ejecutar UNA vez: activa la actualización automática de fechas al editar la hoja a mano.
// Se puede volver a ejecutar: deja un solo disparador "alEditarHoja".
function activarEdicionDeFechas() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === "alEditarHoja")
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("alEditarHoja").forSpreadsheet(IDS_HOJA).onEdit().create();
  recalcularPlazos(); // por si ya se cambiaron fechas a mano antes de activarlo
  Logger.log("Listo: las fechas que se cambien en la hoja se verán en el buscador.");
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
  if (cab.indexOf(N.genero) >= 0) columnas.push(["genero", r(N.genero)]); // existe después de ejecutar "agregarGenero"
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
const URL_PAGINA_CORREGIR = ""; // déjalo vacío: se usa sola la URL de la app web de ESTE proyecto
function urlPaginaCorregir() {
  try { return ScriptApp.getService().getUrl() || URL_PAGINA_CORREGIR; } catch (e) { return URL_PAGINA_CORREGIR; }
} // p. ej. "https://script.google.com/macros/s/AKfycbyCg4jQCpvE5na4slb-oAwOLy7PCYGOqJgb9BS2S1P05OlZCYU_DCOHJR8eNVixiyZi/exec"

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
    if (titulo === TITULO_ACCION) v = ACCION_CORREGIR;
    else if (titulo === TITULO_CODIGO) return;
    else if (titulo === TITULO_CORREGIR) v = codigoOriginal(id); // una corrección pendiente se aprueba sobre la beca original
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
    if (!fila) cuerpo = '<p class="error">No encontramos el ID ' + esc(id) + '. Revisa que esté escrito igual que al pie de la ficha en el buscador (por ejemplo BPH-0007 o B1AMCA001).</p>';
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
    '<input name="id" placeholder="Código" value="' + esc(id) + '" required> <button>Buscar</button></form>' + cuerpo +
    '</body></html>';
  return HtmlService.createHtmlOutput(html).setTitle("Corregir beca · BpH"); // sin ALLOWALL: otras páginas no la pueden incrustar
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
 * Solo los correos de CORREOS_QUE_PUBLICAN pueden archivar o publicar; si lo intenta otra persona no se cambia
 * nada y queda anotado en la bitácora.
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

  // 2. Orden: ¿Qué quieres hacer? | Corregir una beca (enlace + código) | Datos de la beca | … | Archivar o publicar
  //    "Añadir una beca nueva" salta directo a Datos (en blanco, sin la pregunta del código).
  //    "Corregir una beca" muestra el enlace a la página de corrección y el código; después sigue a Datos.
  let pbDatos = buscar(SECCION_DATOS, T_.PAGE_BREAK);
  if (!pbDatos) pbDatos = form.addPageBreakItem().setTitle(SECCION_DATOS);
  let pbCorr = buscar(SECCION_CORREGIR, T_.PAGE_BREAK);
  if (!pbCorr) pbCorr = form.addPageBreakItem().setTitle(SECCION_CORREGIR);
  let corr = buscar(TITULO_CORREGIR, T_.TEXT);
  if (!corr) corr = form.addTextItem().setTitle(TITULO_CORREGIR);
  let accion = buscar(TITULO_ACCION, T_.MULTIPLE_CHOICE);
  if (!accion) accion = form.addMultipleChoiceItem().setTitle(TITULO_ACCION).setRequired(true);
  mover(accion, 0);
  mover(pbCorr, 1);
  mover(corr, 2);
  mover(pbDatos, 3);
  const url = urlPaginaCorregir();
  comoPB(pbCorr).setHelpText(url
    ? "Para que el formulario se abra con todos los datos de la beca, búscala por su código en esta página:\n" + url +
      "\n\nSi llegaste desde esa página o desde el correo, el código ya está puesto: toca Siguiente."
    : "Si llegaste desde el enlace del correo, el código ya está puesto: toca Siguiente.");
  comoPB(pbDatos).setGoToPage(FormApp.PageNavigationType.CONTINUE);
  const mc = comoMC(accion);
  mc.setChoices([
    mc.createChoice(ACCION_NUEVA, comoPB(pbDatos)),
    mc.createChoice(ACCION_CORREGIR, comoPB(pbCorr)),
    mc.createChoice(ACCION_ARCHIVAR, comoPB(pbArch)),
    mc.createChoice(ACCION_PUBLICAR, comoPB(pbArch)),
  ]);

  // 3. El código: obligatorio al corregir, acepta cualquier formato (BPH-0007, B1AMCA001…)
  comoTX(corr).clearValidation().setRequired(true).setHelpText(AYUDA_CORREGIR);

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


/* =====================================================================================================
 * OPCIONES DEL FORMULARIO = VALORES DE LA HOJA (8 de noviembre): ejecutar UNA vez "sincronizarOpciones".
 *
 * - Agrega a cada pregunta de selección las opciones que ya existen en la hoja (por ejemplo "Otros niveles",
 *   "Postdoctorado" o idiomas como Ruso y Catalán). Así, al corregir una beca, el formulario se abre lleno.
 * - Activa "Otro" (con espacio para escribir) en PREGUNTAS_CON_OTRO. Lo que Jenny escriba ahí se agrega
 *   solo como opción nueva después de enviar (función aprenderOpciones), para elegirlo la próxima vez.
 * - Las listas desplegables no permiten "Otro" en Google Forms: solo se completan sus opciones.
 * Se puede volver a ejecutar cuando se quiera; no duplica opciones (no distingue mayúsculas).
 * ===================================================================================================== */
const PREGUNTAS_CON_OTRO = [N.nivel, N.area, N.idioma, N.posgrado, N.importe, N.duracion, N.genero];
const PREGUNTAS_SINCRONIZADAS = PREGUNTAS_CON_OTRO.concat([N.pais, N.cobertura, N.modalidad, N.mes, U.tipo, U.ingreso])
  .concat([1, 2, 3, 4].map((n) => [P(n).tipo, P(n).ingreso]).flat());
const RENOMBRAR_OPCIONES = { [N.posgrado]: { "Máster": "Maestría" } }; // nombre viejo en el formulario -> nombre de la hoja
const OPCIONES_NUEVAS = { [N.nivel]: ["Otros niveles"], [N.posgrado]: ["Postdoctorado"] };

function itemDeSeleccion(it) {
  const T_ = FormApp.ItemType, t = it.getType();
  if (t === T_.CHECKBOX) return { tipo: "casillas", x: it.asCheckboxItem ? it.asCheckboxItem() : it };
  if (t === T_.MULTIPLE_CHOICE) return { tipo: "opcion", x: it.asMultipleChoiceItem ? it.asMultipleChoiceItem() : it };
  if (t === T_.LIST) return { tipo: "lista", x: it.asListItem ? it.asListItem() : it };
  return null;
}
// Valores de una celda: las casillas vienen separadas por coma. Las fechas sueltas no son opciones.
function partesDe(valor, tipo) {
  if (valor === "" || valor === null || valor === undefined || valor instanceof Date) return [];
  const s = String(valor).trim();
  if (!s || /^\d{4}-\d{2}-\d{2}$/.test(s)) return [];
  return tipo === "casillas" ? s.split(/,\s*/).map((x) => x.trim()).filter(Boolean) : [s];
}
// Agrega a la pregunta los valores que todavía no son opción. Devuelve la lista de opciones agregadas.
function agregarOpciones(sel, titulo, valores, conOtro) {
  const ren = RENOMBRAR_OPCIONES[titulo] || {};
  const originales = sel.x.getChoices().map((c) => c.getValue());
  const actuales = originales.map((v) => ren[v] || v);
  const ya = new Set(actuales.map((v) => v.toLowerCase()));
  const nuevas = [];
  (OPCIONES_NUEVAS[titulo] || []).concat(valores).forEach((v) => {
    v = ren[v] || v; // un valor viejo de la hoja (p. ej. "Máster") no vuelve a entrar como opción
    const k = String(v).toLowerCase();
    if (k && !ya.has(k)) { ya.add(k); nuevas.push(String(v)); }
  });
  nuevas.sort((a, b) => a.localeCompare(b, "es"));
  if (nuevas.length || originales.some((v) => ren[v])) sel.x.setChoiceValues(actuales.concat(nuevas));
  if (conOtro && sel.tipo !== "lista") sel.x.showOtherOption(true);
  return nuevas;
}

function sincronizarOpciones() {
  const form = FormApp.openById(IDS_FORMULARIO);
  const hoja = SpreadsheetApp.openById(IDS_HOJA).getSheetByName("Respuestas");
  const cab = encabezados(hoja);
  const datos = hoja.getRange(2, 1, Math.max(hoja.getLastRow() - 1, 1), cab.length).getValues();
  const items = form.getItems();
  const informe = [];
  PREGUNTAS_SINCRONIZADAS.forEach((titulo) => {
    const it = items.find((i) => i.getTitle() === titulo && itemDeSeleccion(i));
    if (!it) { informe.push("- " + titulo + ": no está en el formulario"); return; }
    const sel = itemDeSeleccion(it);
    const col = cab.indexOf(titulo);
    const valores = col < 0 ? [] : datos.map((r) => partesDe(r[col], sel.tipo)).flat();
    const conOtro = PREGUNTAS_CON_OTRO.includes(titulo);
    const nuevas = agregarOpciones(sel, titulo, valores, conOtro);
    informe.push("- " + titulo + ": " + (nuevas.length ? "se agregó " + nuevas.join(", ") : "ya estaba completa") +
      (conOtro ? (sel.tipo === "lista" ? " (es lista desplegable: Google no permite \"Otro\")" : " · \"Otro\" activado") : ""));
  });
  // Lo que se renombró en el formulario también se renombra en la hoja (por ejemplo Máster -> Maestría)
  Object.keys(RENOMBRAR_OPCIONES).forEach((titulo) => {
    const col = cab.indexOf(titulo);
    if (col < 0) return;
    Object.entries(RENOMBRAR_OPCIONES[titulo]).forEach(([viejo, nuevo]) =>
      hoja.getRange(2, col + 1, Math.max(hoja.getLastRow() - 1, 1), 1).createTextFinder(viejo).matchCase(true).replaceAllWith(nuevo));
  });
  Logger.log("Opciones del formulario actualizadas:\n" + informe.join("\n"));
}

/* =====================================================================================================
 * GÉNERO (23 de noviembre) — ejecutar UNA vez "agregarGenero".
 *  Agrega al formulario la pregunta "Género" (Todos / Mujeres / Hombres / Otro) después de "Modalidad",
 *  crea su columna en la hoja y la suma a Publicadas como "genero". Las becas que ya existen quedan vacías = para todos.
 *  Se puede volver a ejecutar sin duplicar nada.
 * ===================================================================================================== */
function agregarGenero() {
  const form = FormApp.openById(IDS_FORMULARIO);
  let it = form.getItems().find((i) => i.getTitle() === N.genero);
  if (!it) {
    it = form.addMultipleChoiceItem().setTitle(N.genero)
      .setHelpText("¿Para quién es la beca? Si es para cualquier persona, elige «Todos». Usa «Otro» si es, por ejemplo, para mujeres y personas no binarias.")
      .setChoiceValues(GENEROS).showOtherOption(true).setRequired(false);
  }
  const mod = form.getItems().find((i) => i.getTitle() === N.modalidad);
  if (mod && it.getIndex() !== mod.getIndex() + 1) form.moveItem(it.getIndex(), mod.getIndex() + (it.getIndex() > mod.getIndex() ? 1 : 0));
  // La columna aparece en la hoja cuando Google la crea; se espera un poco y se arma Publicadas.
  const hoja = SpreadsheetApp.openById(IDS_HOJA).getSheetByName("Respuestas");
  for (let k = 0; k < 10 && encabezados(hoja).indexOf(N.genero) < 0; k++) { SpreadsheetApp.flush(); Utilities.sleep(3000); }
  if (encabezados(hoja).indexOf(N.genero) < 0) {
    Logger.log("La pregunta quedó en el formulario, pero la columna \"" + N.genero + "\" todavía no aparece en la hoja. Espera un minuto y vuelve a ejecutar agregarGenero.");
    return;
  }
  rehacerPublicadas();
  Logger.log("Listo: \"" + N.genero + "\" está en el formulario, en la hoja y en Publicadas (columna genero).");
}

/* =====================================================================================================
 * ALINEAR FORMULARIO Y HOJA — ejecutar UNA vez, DESPUÉS de cambiar "Nivel" a "Casillas" en el editor del formulario.
 *  1. En la hoja: "Pregrado y posgrado" pasa a "Pregrado, Posgrado"; en las columnas de varias opciones
 *     se quitan repetidos y textos sueltos que no son opción (p. ej. "Arabe e ingles").
 *  2. En el formulario: se quita la opción "Pregrado y posgrado" y se agregan todas las opciones que usa la hoja.
 * ===================================================================================================== */
const REEMPLAZOS_HOJA = { [N.nivel]: { "Pregrado y posgrado": "Pregrado, Posgrado" } };
const QUITAR_OPCIONES = { [N.nivel]: ["Pregrado y posgrado"] };
const BASURA_EN_CASILLAS = { [N.idioma]: ["Arabe e ingles"] };

function alinearFormularioYHoja() {
  const form = FormApp.openById(IDS_FORMULARIO);
  const nivel = form.getItems().find((i) => i.getTitle() === N.nivel);
  if (!nivel || nivel.getType() !== FormApp.ItemType.CHECKBOX) {
    throw new Error("Primero cambia la pregunta \"Nivel\" a \"Casillas\" en el editor del formulario y vuelve a ejecutar.");
  }
  const hoja = SpreadsheetApp.openById(IDS_HOJA).getSheetByName("Respuestas");
  const cab = encabezados(hoja);
  const filas = Math.max(hoja.getLastRow() - 1, 1);
  const casillas = form.getItems().filter((i) => i.getType() === FormApp.ItemType.CHECKBOX).map((i) => i.getTitle());
  const informe = [];
  casillas.forEach((titulo) => {
    const col = cab.indexOf(titulo);
    if (col < 0) return;
    const rango = hoja.getRange(2, col + 1, filas, 1);
    const datos = rango.getValues();
    const rem = REEMPLAZOS_HOJA[titulo] || {}, basura = (BASURA_EN_CASILLAS[titulo] || []).map((x) => x.toLowerCase());
    let cambios = 0;
    datos.forEach((r) => {
      const antes = String(r[0] || "").trim();
      if (!antes) return;
      const vistos = new Set();
      const partes = antes.split(/,\s*/).map((x) => rem[x.trim()] || x.trim()).join(", ").split(/,\s*/)
        .filter((x) => x && !basura.includes(x.toLowerCase()) && !vistos.has(x.toLowerCase()) && vistos.add(x.toLowerCase()));
      const despues = seguro(partes.join(", "));
      if (despues !== antes) { r[0] = despues; cambios++; }
    });
    if (cambios) { rango.setValues(datos); informe.push("- Hoja, " + titulo + ": " + cambios + " celdas ordenadas"); }
  });
  Object.keys(QUITAR_OPCIONES).forEach((titulo) => {
    const it = form.getItems().find((i) => i.getTitle() === titulo && itemDeSeleccion(i));
    if (!it) return;
    const sel = itemDeSeleccion(it);
    const quitar = QUITAR_OPCIONES[titulo];
    const quedan = sel.x.getChoices().map((c) => c.getValue()).filter((v) => !quitar.includes(v));
    sel.x.setChoiceValues(quedan);
    informe.push("- Formulario, " + titulo + ": se quitó " + quitar.join(", "));
  });
  sincronizarOpciones(); // agrega al formulario cualquier opción que use la hoja
  Logger.log("Listo:\n" + (informe.join("\n") || "- La hoja ya estaba ordenada") + "\n(Ver arriba el detalle de opciones del formulario.)");
}

// Después de cada envío: lo que se escribió en "Otro" queda como opción para la próxima vez.
function aprenderOpciones(cab, fila) {
  try {
    const items = FormApp.openById(IDS_FORMULARIO).getItems();
    PREGUNTAS_CON_OTRO.forEach((titulo) => {
      const it = items.find((i) => i.getTitle() === titulo && itemDeSeleccion(i));
      const col = cab.indexOf(titulo);
      if (!it || col < 0) return;
      const sel = itemDeSeleccion(it);
      if (sel.tipo === "lista") return;
      agregarOpciones(sel, titulo, partesDe(fila[col], sel.tipo), true);
    });
  } catch (e) {
    console.error("No se pudieron agregar las opciones nuevas: " + e); // nunca frena el registro de la beca
  }
}


// Aviso a BpH: otra persona registró una beca que queda "Pendiente de revisión" (no se ve en el buscador).
// Aviso a BpH: otra persona registró una beca (o propuso cambios a una) que queda "Pendiente de revisión".
// Si "original" trae la fila de la beca publicada, es una corrección: solo se muestran los cambios propuestos.
function avisarPendiente(hoja, fila, codigo, quien, nota, original) {
  const cab = encabezados(hoja);
  const valores = hoja.getRange(fila, 1, 1, cab.length).getValues()[0];
  let enlace = "";
  try { enlace = enlaceParaCorregir(cab, valores).url; }
  catch (e) { console.error("No se pudo armar el enlace para revisar la beca pendiente: " + e); }
  const revisar = (txt) => enlace ? link(txt, enlace) : "(usa \"Corregir una beca\" en el formulario con el código " + codigo + ")";
  const publicar = enlacePublicar(codigo);
  const talCual = (txt) => publicar ? link(txt, publicar)
    : "en el formulario elige \"" + ACCION_PUBLICAR + "\" con el código " + codigo + " para " + txt;
  const esCorreccion = !!original;
  const texto = esCorreccion
    ? ficha(cab, original, codigo) + "\nCorrección propuesta por: " + (quien || "(sin correo)") + "\n" +
      textoCambios(cab, original, valores, "Cambios propuestos:") + "\n" +
      "La beca publicada no cambia hasta que la apruebes, y conserva su código.\n" +
      "Para aprobarla: " + revisar("revisar y aprobar los cambios") + ", o " + talCual("aprobarla tal cual") + " (sin cambiar nada).\n" +
      "Si no quieres esos cambios, no hagas nada."
    : (nota || "") + ficha(cab, valores, codigo) + "\nRegistrada por: " + (quien || "(sin correo)") + "\n" + detalleCorto(cab, valores) + "\n" +
      "Todavía no aparece en el buscador. Antes de aprobarla, busca su nombre en el buscador para no publicar un duplicado.\n" +
      "Para aprobarla: " + revisar("revisar y publicar") + ", o " + talCual("publicarla tal cual") + " (sin cambiar nada).";
  const asunto = esCorreccion ? "Corrección pendiente de revisión" : "Beca pendiente de revisión";
  anotar(asunto, codigo, texto);
  if (!puedeAvisar(quien)) return; // mucho envío de la misma persona o sin cuota: queda solo en la bitácora
  const url = urlBitacora();
  avisar(CORREOS_AVISO_PENDIENTES.join(","), asunto + ": " + codigo,
    texto + (url ? "\n\n" + link("Ver la bitácora", url) : ""));
}

// Seguridad (spam): como mucho MAX_AVISOS_POR_PERSONA correos al día por quien envía, y nunca si queda poca cuota de Gmail.
const MAX_AVISOS_POR_PERSONA = 5;
function puedeAvisar(quien) {
  try {
    if (MailApp.getRemainingDailyQuota() < 10) return false;
    const clave = "avisos_" + Utilities.formatDate(new Date(), zonaHoja(), "yyyyMMdd") + "_" + String(quien || "anonimo").toLowerCase();
    const props = PropertiesService.getScriptProperties();
    const n = Number(props.getProperty(clave) || 0) + 1;
    props.setProperty(clave, String(n));
    return n <= MAX_AVISOS_POR_PERSONA;
  } catch (err) { console.error("No se pudo revisar el límite de avisos: " + err); return true; }
}

// ---------- Textos cortos para la bitácora y el aviso ----------

// Enlace con texto: en la bitácora y en el correo se ve solo el texto, clicable.
// Seguridad: el enlace se marca con caracteres invisibles que el script quita de todo lo que viene del formulario,
// así nadie puede meter un "[[Aprobar aquí|sitio falso]]" en el nombre de una beca y que se vuelva un enlace real.
const ENL_A = "\uE000", ENL_B = "\uE001", ENL_C = "\uE002";
const RE_ENLACE = /\uE000([^\uE001]+)\uE001([^\uE002\s]+)\uE002/g;
function limpio(t) { return String(t === null || t === undefined ? "" : t).replace(/[\uE000-\uE002]/g, ""); }
function link(texto, url) {
  url = limpio(url).trim();
  return /^https?:\/\//i.test(url) ? ENL_A + limpio(texto) + ENL_B + url + ENL_C : limpio(texto);
}

// "BPASAE001_1 · Becas de pregrado en NYU Abu Dhabi (Emiratos Árabes Unidos)"
function ficha(cab, fila, codigo) {
  const v = (t) => limpio(valorDe(cab, fila, t) || "").trim();
  const pais = v(N.pais) || v(N.paisViejo);
  return codigo + " · " + (v(N.nombre) || "(sin nombre)") + (pais ? " (" + pais + ")" : "");
}

// "Nivel: Pregrado · Cobertura: Completa · Convocatoria" (enlace)
function detalleCorto(cab, fila) {
  const v = (t) => limpio(valorDe(cab, fila, t) || "").trim();
  const partes = [];
  if (v("Nivel")) partes.push("Nivel: " + v("Nivel"));
  if (v("Cobertura")) partes.push("Cobertura: " + v("Cobertura"));
  if (v(N.enlace)) partes.push(link("convocatoria", v(N.enlace)));
  return partes.join(" · ");
}

// Texto de un valor para comparar y mostrar (fechas y horas como en la hoja).
function textoValor(v, zona) {
  if (v instanceof Date) return Utilities.formatDate(v, zona, v.getFullYear() < 1901 ? "HH:mm" : "dd/MM/yyyy");
  return limpio(v).trim();
}

// Lista de lo que cambió entre dos versiones de la beca: "- Idioma: «Inglés» → «Inglés, Árabe»".
function textoCambios(cab, antes, despues, titulo) {
  const zona = zonaHoja();
  const iCorr = cab.indexOf(TITULO_CORREGIR), iCorreo = cab.findIndex((h) => /correo|email/i.test(h));
  const lineas = [];
  cab.forEach((h, i) => {
    if (!h || i === 0 || i === iCorr || i === iCorreo || h === N.verificado || COLUMNAS_PROPIAS.includes(h) || SOLO_FORMULARIO.includes(h)) return;
    const a = textoValor(antes[i], zona), d = textoValor(despues[i], zona);
    if (a === d) return;
    if (!d && (/^Plazo \d|^¿Agregar/.test(h) || h === N.mes || h === N.fechaVieja || h === N.variasVieja || Object.values(U).includes(h))) return; // campos viejos que el formulario ya no llena
    lineas.push("- " + h + ": «" + (a || "vacío") + "» → «" + (d || "vacío") + "»");
  });
  return lineas.length ? titulo + "\n" + lineas.join("\n") : "No cambió ningún dato (solo se marcó como verificada).";
}

// Enlace para volver a corregir la beca que está en esa fila.
function lineaEnlaces(hoja, filaNum) {
  try {
    const cab = encabezados(hoja);
    const fila = hoja.getRange(filaNum, 1, 1, cab.length).getValues()[0];
    const { url, faltan } = enlaceParaCorregir(cab, fila);
    return "\n" + link("Corregirla", url) + (faltan.length ? " (hay que volver a llenar: " + faltan.join(", ") + ")" : "");
  } catch (e) {
    console.error("No se pudo armar el enlace para corregir: " + e);
    return "";
  }
}

// Enlace prellenado: "Publicar o volver a publicar una beca" con el código ya puesto. "" si no se pudo.
function enlacePublicar(codigo) {
  try {
    const form = FormApp.openById(IDS_FORMULARIO);
    const r = form.createResponse();
    let listo = 0;
    form.getItems().forEach((item) => {
      const t = item.getTitle();
      if (t === TITULO_ACCION) { r.withItemResponse(item.asMultipleChoiceItem().createResponse(ACCION_PUBLICAR)); listo++; }
      else if (t === TITULO_CODIGO) { r.withItemResponse(item.asTextItem().createResponse(codigoOriginal(codigo))); listo++; }
    });
    return listo === 2 ? r.toPrefilledUrl() : "";
  } catch (e) {
    console.error("No se pudo armar el enlace para publicar tal cual: " + e);
    return "";
  }
}

// ---------- Bitácora (Google Docs de solo lectura) ----------

// Ejecutar UNA vez a mano: crea la bitácora (o muestra su enlace si ya existe).
function crearBitacora() {
  const doc = bitacora();
  console.log("Bitácora: " + doc.getUrl());
  return doc.getUrl();
}

function urlBitacora() {
  try { return bitacora().getUrl(); } catch (e) { return ""; }
}

// Abre la bitácora guardada en las propiedades del script; si no existe, la crea.
function bitacora() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty("BITACORA_ID");
  if (id) {
    try { return DocumentApp.openById(id); } catch (e) { console.error("No se pudo abrir la bitácora, se crea otra: " + e); }
  }
  const doc = DocumentApp.create(TITULO_BITACORA);
  const cuerpo = doc.getBody();
  const tituloDoc = cuerpo.getChild(0).asParagraph(); // setText no devuelve nada, por eso va en dos líneas
  tituloDoc.setText(TITULO_BITACORA);
  tituloDoc.setHeading(DocumentApp.ParagraphHeading.TITLE);
  cuerpo.appendParagraph("Cada envío del formulario queda anotado aquí, el más reciente arriba. " +
    "Busca con Ctrl+F (en el celular: menú ⋮ > Buscar y reemplazar) por nombre o código de beca. Este documento es solo de lectura.")
    .setHeading(DocumentApp.ParagraphHeading.NORMAL);
  cuerpo.appendParagraph("").appendHorizontalRule();
  doc.saveAndClose();
  compartirBitacora(doc.getId(), true);
  props.setProperty("BITACORA_ID", doc.getId());
  return DocumentApp.openById(doc.getId());
}

function compartirBitacora(id, mover) {
  const archivo = DriveApp.getFileById(id);
  if (mover) {
    try { archivo.moveTo(DriveApp.getFolderById(CARPETA_BITACORA)); } catch (e) { console.error("No se pudo mover la bitácora a la carpeta: " + e); }
  }
  LECTORES_BITACORA.forEach((c) => {
    try { archivo.addViewer(c); } catch (e) { console.error("No se pudo compartir la bitácora con " + c + ": " + e); }
  });
}

// Posición donde empiezan las anotaciones (justo después de la línea del encabezado).
function inicioBitacora(cuerpo) {
  for (let i = 0; i < cuerpo.getNumChildren(); i++) {
    const el = cuerpo.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH && el.asParagraph().findElement(DocumentApp.ElementType.HORIZONTAL_RULE)) return i + 1;
  }
  return cuerpo.getNumChildren();
}

// Marcas de la bitácora: ❌ = pendiente (falta aprobarla), ✅ = resuelta, ⚠️ = no se pudo hacer.
const MARCA_PENDIENTE = "❌", MARCA_LISTA = "✅", MARCA_ERROR = "⚠️";
const TITULOS_PENDIENTES = ["Beca pendiente de revisión", "Corrección pendiente de revisión"];
function marcaDe(titulo) {
  if (TITULOS_PENDIENTES.indexOf(titulo) >= 0) return MARCA_PENDIENTE;
  return /^No se pudo/.test(titulo) ? MARCA_ERROR : MARCA_LISTA;
}

// Anota un envío arriba de todo: título con fecha, el texto completo y los enlaces clicables.
// Si "resuelve" viene ("publicada" o "corregida"), las entradas ❌ de ese código pasan a ✅ con la fecha.
// Nunca frena el registro de la beca: si algo falla, solo queda en el registro de ejecución.
function anotar(titulo, codigo, texto, resuelve) {
  try {
    const doc = bitacora();
    const cuerpo = doc.getBody();
    // Medir el largo cuesta: se revisa una vez cada 25 anotaciones.
    const props = PropertiesService.getScriptProperties();
    const n = Number(props.getProperty("anotaciones") || 0) + 1;
    props.setProperty("anotaciones", String(n % 25));
    if (n % 25 === 0 && cuerpo.getText().length > LIMITE_BITACORA) archivarBitacora(doc);
    let pos = inicioBitacora(cuerpo);
    const fecha = Utilities.formatDate(new Date(), zonaHoja(), "dd/MM/yyyy HH:mm");
    if (resuelve && codigo) marcarResueltas(cuerpo, codigo, resuelve + " el " + fecha);
    cuerpo.insertParagraph(pos++, marcaDe(titulo) + " " + fecha + " · " + titulo + (codigo ? " · " + codigo : ""))
      .setHeading(DocumentApp.ParagraphHeading.HEADING3);
    String(texto).split("\n").forEach((linea) => {
      const { visible, enlaces } = partirEnlaces(linea);
      const p = cuerpo.insertParagraph(pos++, visible).setHeading(DocumentApp.ParagraphHeading.NORMAL);
      if (enlaces.length) enlaces.forEach((x) => p.editAsText().setLinkUrl(x.ini, x.fin, x.url));
      else enlazar(p, visible);
    });
    cuerpo.insertParagraph(pos, "").setHeading(DocumentApp.ParagraphHeading.NORMAL).appendHorizontalRule();
    doc.saveAndClose();
  } catch (e) {
    console.error("No se pudo anotar en la bitácora (" + titulo + " " + codigo + "): " + e);
  }
}

// Partes de un título de la bitácora: "❌ 14/11/2026 16:39 · Corrección pendiente de revisión · BMEUDE003".
function partesTitulo(texto) {
  const m = String(texto).match(/^(❌|✅|⚠️)?\s*(\d\d\/\d\d\/\d{4} \d\d:\d\d) · ([^·]+?)(?: · ([^·]+?))?(?: · (.+))?$/);
  return m ? { marca: m[1] || "", fecha: m[2], titulo: m[3].trim(), codigo: (m[4] || "").trim(), nota: m[5] || "" } : null;
}

// Pasa a ✅ las entradas ❌ (pendientes) de ese código y les agrega "· publicada el …" o "· corregida el …".
function marcarResueltas(cuerpo, codigo, nota) {
  const cod = codigoOriginal(codigo).toUpperCase();
  const titulos = cuerpo.getParagraphs().filter((p) => p.getHeading() === DocumentApp.ParagraphHeading.HEADING3);
  titulos.forEach((p) => {
    const t = partesTitulo(p.getText());
    if (!t || t.marca !== MARCA_PENDIENTE || codigoOriginal(t.codigo).toUpperCase() !== cod) return;
    const texto = p.editAsText();
    texto.deleteText(0, MARCA_PENDIENTE.length - 1);
    texto.insertText(0, MARCA_LISTA);
    texto.appendText(" · " + nota);
  });
}

// Ejecutar UNA vez a mano: pone ❌/✅/⚠️ a las entradas que quedaron sin marca y pasa a ✅
// las pendientes que ya se resolvieron después (publicada o corregida con el mismo código).
function repararBitacora() {
  const doc = bitacora();
  const titulos = doc.getBody().getParagraphs().filter((p) => p.getHeading() === DocumentApp.ParagraphHeading.HEADING3);
  const resueltas = {}; // código -> primera resolución posterior (la bitácora va de nueva a vieja)
  let cambios = 0;
  titulos.slice().reverse().forEach((p) => { // de la más vieja a la más nueva
    const t = partesTitulo(p.getText());
    if (!t) return;
    if (!t.marca) { p.editAsText().insertText(0, marcaDe(t.titulo) + " "); t.marca = marcaDe(t.titulo); cambios++; }
    const cod = codigoOriginal(t.codigo).toUpperCase();
    if (t.titulo === "Beca publicada" || t.titulo === "Beca corregida") {
      (resueltas[cod] = resueltas[cod] || []).forEach((q) => {
        const texto = q.editAsText();
        texto.deleteText(0, MARCA_PENDIENTE.length - 1);
        texto.insertText(0, MARCA_LISTA);
        texto.appendText(" · " + (t.titulo === "Beca publicada" ? "publicada" : "corregida") + " el " + t.fecha);
        cambios++;
      });
      resueltas[cod] = [];
    } else if (t.marca === MARCA_PENDIENTE && !t.nota) {
      (resueltas[cod] = resueltas[cod] || []).push(p);
    }
  });
  doc.saveAndClose();
  Logger.log("Bitácora revisada: " + cambios + " cambios.");
}

// Convierte en enlace cada https://… que haya en el párrafo.
function enlazar(parrafo, linea) {
  const texto = parrafo.editAsText();
  const re = /https?:\/\/\S+/g;
  let m;
  while ((m = re.exec(linea)) !== null) texto.setLinkUrl(m.index, m.index + m[0].length - 1, m[0]);
}

// Separa "[[texto|url]]" en el texto visible y la lista de enlaces (posición dentro de ese texto).
function partirEnlaces(linea) {
  const enlaces = [];
  let visible = "", desde = 0, m;
  const re = new RegExp(RE_ENLACE.source, "g");
  while ((m = re.exec(linea)) !== null) {
    visible += linea.slice(desde, m.index);
    enlaces.push({ ini: visible.length, fin: visible.length + m[1].length - 1, url: m[2] });
    visible += m[1];
    desde = m.index + m[0].length;
  }
  return { visible: visible + linea.slice(desde), enlaces };
}

// La bitácora se llenó: guarda una copia con todo lo anterior y deja la original (mismo enlace) solo con el encabezado.
function archivarBitacora(doc) {
  const fecha = Utilities.formatDate(new Date(), zonaHoja(), "dd/MM/yyyy");
  const copia = DriveApp.getFileById(doc.getId()).makeCopy(TITULO_BITACORA + " (anotaciones hasta el " + fecha + ")");
  compartirBitacora(copia.getId(), true);
  const cuerpo = doc.getBody();
  const desde = inicioBitacora(cuerpo);
  for (let i = cuerpo.getNumChildren() - 1; i >= desde; i--) {
    const el = cuerpo.getChild(i);
    try { el.removeFromParent(); } catch (e) { if (el.getType() === DocumentApp.ElementType.PARAGRAPH) el.asParagraph().clear(); }
  }
  const p = cuerpo.insertParagraph(desde, "Las anotaciones anteriores al " + fecha + " están en: " + copia.getUrl());
  enlazar(p, p.getText());
}
