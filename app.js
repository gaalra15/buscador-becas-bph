// Buscador de becas: lee el CSV publicado, calcula si cada beca está abierta y aplica filtros que se guardan en la URL.
(function () {
  const MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
  const FILTROS = ["nivel", "area", "pais", "cobertura", "idioma", "modalidad"];
  const CAMPOS_LISTA = ["nivel", "area", "idioma", "modalidad"]; // celdas que pueden traer varios valores separados por ; o ,

  const $ = (sel) => document.querySelector(sel);
  const norm = (s) => String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().trim();

  // --- CSV (soporta comillas, comas y saltos de línea dentro de celdas) ---
  function parseCSV(text) {
    const rows = []; let row = []; let cell = ""; let q = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (q) {
        if (c === '"' && text[i + 1] === '"') { cell += '"'; i++; }
        else if (c === '"') q = false;
        else cell += c;
      } else if (c === '"') q = true;
      else if (c === ",") { row.push(cell); cell = ""; }
      else if (c === "\n" || c === "\r") {
        if (c === "\r" && text[i + 1] === "\n") i++;
        row.push(cell); rows.push(row); row = []; cell = "";
      } else cell += c;
    }
    if (cell !== "" || row.length) { row.push(cell); rows.push(row); }
    return rows.filter((r) => r.some((v) => v.trim() !== ""));
  }

  // --- Fechas ---
  function hoy() { const d = new Date(); return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
  function parseFecha(s) {
    s = String(s || "").trim();
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); // dd/mm/aaaa (formato de Google Sheets en español)
    if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
    return null;
  }
  function parseMes(s) {
    const n = norm(s);
    const i = MESES.findIndex((m) => n.startsWith(m.slice(0, 3)));
    if (i >= 0) return i;
    const k = parseInt(n, 10);
    return k >= 1 && k <= 12 ? k - 1 : -1;
  }
  const fmt = (d) => d.toLocaleDateString("es-HN", { day: "numeric", month: "long", year: "numeric" });
  const diasHasta = (d) => Math.round((d - hoy()) / 86400000);

  // Devuelve { abierta, etiqueta, clase, texto, orden }
  // plazos: [{ronda, tipo, fecha:"aaaa-mm-dd", hora, zona, ingreso, estado:"Confirmada"|"Estimada"}] (columna "plazos")
  function leerPlazos(s) {
    try { const l = JSON.parse(s || "[]"); return Array.isArray(l) ? l.filter((p) => parseFecha(p.fecha)) : []; } catch (e) { return []; }
  }

  function calcularEstado(tipo, valor, plazos) {
    const t = norm(tipo);
    const pronto = (d) => diasHasta(d) <= DIAS_CIERRA_PRONTO;

    if (t.includes("todo el ano") || t === "abierta") {
      return { abierta: true, etiqueta: "Abierta todo el año", clase: "abierta", texto: "Sin fecha de cierre", orden: 1e6 };
    }
    if (plazos && plazos.length) {
      const orden = plazos.map((p) => ({ ...p, d: parseFecha(p.fecha) })).sort((a, b) => a.d - b.d);
      const prox = orden.find((p) => diasHasta(p.d) >= 0);
      const detalle = (p) => fmt(p.d) + (norm(p.estado) === "estimada" ? " (fecha estimada)" : "") + (p.tipo ? " · " + p.tipo : "");
      if (prox) return { abierta: true, etiqueta: pronto(prox.d) ? "Cierra pronto" : "Abierta", clase: pronto(prox.d) ? "pronto" : "abierta", texto: "Próximo plazo: " + detalle(prox), orden: diasHasta(prox.d) };
      const ult = orden[orden.length - 1];
      return { abierta: false, etiqueta: "Cerrada", clase: "cerrada", texto: "Último plazo: " + detalle(ult), orden: 2e6 };
    }
    if (t.includes("sin fecha")) {
      return { abierta: false, etiqueta: "Sin fecha publicada", clase: "cerrada", texto: "La convocatoria aún no publica fecha límite", orden: 2.5e6 };
    }
    const esMes = t.includes("mes") || (t.includes("varias") && !/\d/.test(String(valor)) && parseMes(valor) >= 0);
    if (t.includes("varias") && !esMes) {
      const fechas = String(valor).split(/[;\n]/).map(parseFecha).filter(Boolean).sort((a, b) => a - b);
      const prox = fechas.find((d) => diasHasta(d) >= 0);
      if (prox) return { abierta: true, etiqueta: pronto(prox) ? "Cierra pronto" : "Abierta", clase: pronto(prox) ? "pronto" : "abierta", texto: "Próximo cierre: " + fmt(prox), orden: diasHasta(prox) };
      if (fechas.length) return { abierta: false, etiqueta: "Cerrada", clase: "cerrada", texto: "Último cierre: " + fmt(fechas[fechas.length - 1]), orden: 2e6 };
      if (String(valor).trim() && parseMes(valor) < 0) return { abierta: false, etiqueta: "Revisar fechas", clase: "cerrada", texto: String(valor).trim(), orden: 3e6 };
    }
    if (esMes) {
      const mes = parseMes(valor);
      if (mes >= 0) {
        const actual = hoy().getMonth();
        // Supuesto: "solo mes" significa que cierra al final de ese mes, cada año.
        if (mes >= actual) {
          const fin = new Date(hoy().getFullYear(), mes + 1, 0);
          return { abierta: true, etiqueta: pronto(fin) ? "Cierra pronto" : "Abierta", clase: pronto(fin) ? "pronto" : "abierta", texto: "Cierra en " + MESES[mes] + " (día por confirmar)", orden: diasHasta(fin) };
        }
        return { abierta: false, etiqueta: "Cerrada este año", clase: "cerrada", texto: "Suele cerrar en " + MESES[mes], orden: 2e6 };
      }
    }
    const d = parseFecha(valor);
    if (d) {
      if (diasHasta(d) < 0) return { abierta: false, etiqueta: "Cerrada", clase: "cerrada", texto: "Cerró el " + fmt(d), orden: 2e6 };
      return { abierta: true, etiqueta: pronto(d) ? "Cierra pronto" : "Abierta", clase: pronto(d) ? "pronto" : "abierta", texto: "Cierra el " + fmt(d), orden: diasHasta(d) };
    }
    return { abierta: false, etiqueta: "Revisar fecha", clase: "cerrada", texto: "Fecha de cierre no disponible", orden: 3e6 };
  }

  // --- Datos ---
  function aObjetos(rows) {
    const cab = rows[0].map(norm);
    const idx = {};
    for (const [campo, nombre] of Object.entries(COLUMNAS)) idx[campo] = cab.indexOf(norm(nombre));
    const faltan = Object.keys(idx).filter((k) => idx[k] < 0 && ["nombre", "tipoCierre", "cierre", "enlace"].includes(k)); // el resto es opcional
    if (faltan.length) console.warn("Columnas no encontradas en el CSV:", faltan.map((k) => COLUMNAS[k]));
    return rows.slice(1).map((r) => {
      const o = {};
      for (const campo in idx) o[campo] = idx[campo] >= 0 ? (r[idx[campo]] || "").trim() : "";
      return o;
    });
  }
  const aprobada = (v) => ["si", "x", "true", "verdadero", "aprobada", "aprobado", "1"].includes(norm(v));
  // "Pregrado y posgrado" (tercera opción de Nivel) cuenta para los dos filtros.
  const SEPARADOR = { nivel: /[;,]|\s+y\s+/i };
  const mayuscula = (s) => s.charAt(0).toUpperCase() + s.slice(1);
  const valores = (beca, campo) => CAMPOS_LISTA.includes(campo) ? beca[campo].split(SEPARADOR[campo] || /[;,]/).map((s) => mayuscula(s.trim())).filter(Boolean) : [beca[campo]].filter(Boolean);

  let becas = [];
  let pestana = "buscar";

  // --- URL <-> filtros ---
  function leerURL() {
    const p = new URLSearchParams(location.search);
    $("#q").value = p.get("q") || "";
    $("#abiertas").checked = p.get("abiertas") === "1";
    for (const f of FILTROS) {
      const v = p.get(f); const sel = $("#f-" + f);
      if (!v) { sel.value = ""; continue; }
      const op = [...sel.options].find((o) => norm(o.value) === norm(v));
      sel.value = op ? op.value : "";
    }
    if (FILTROS.some((f) => p.get(f))) $("details.filtros").open = true;
  }
  function escribirURL() {
    const p = new URLSearchParams();
    if (pestana === "resumen") p.set("pestana", "resumen");
    if ($("#q").value.trim()) p.set("q", $("#q").value.trim());
    if ($("#abiertas").checked) p.set("abiertas", "1");
    for (const f of FILTROS) { const v = $("#f-" + f).value; if (v) p.set(f, v); }
    const qs = p.toString();
    try { history.replaceState(null, "", location.pathname + (qs ? "?" + qs : "")); } catch (e) { /* en algunos marcos incrustados no se permite */ }
  }

  // --- Pintar ---
  function llenarSelects() {
    for (const f of FILTROS) {
      const set = new Map();
      becas.forEach((b) => valores(b, f).forEach((v) => { if (!set.has(norm(v))) set.set(norm(v), v); }));
      const sel = $("#f-" + f);
      [...set.values()].sort((a, b) => a.localeCompare(b, "es")).forEach((v) => sel.add(new Option(v, v)));
    }
  }

  function tarjeta(b) {
    const el = (tag, clase, texto) => { const e = document.createElement(tag); if (clase) e.className = clase; if (texto != null) e.textContent = texto; return e; };
    const lista = (v) => String(v || "").split(/\s*[;,]\s*/).filter(Boolean);
    const li = el("li", "card");
    if (b.pais) li.append(el("div", "pais", b.pais));
    const cuerpo = el("div", "cuerpo");

    // Lo esencial a la vista: estado, nombre, próxima fecha y etiquetas cortas
    cuerpo.append(el("span", "estado " + b.estado.clase, b.estado.etiqueta), el("h2", null, b.nombre), el("p", "cierre", b.estado.texto));
    const tags = el("ul", "tags");
    [...lista(b.nivel), ...lista(b.tipoPosgrado), b.cobertura ? "Cobertura " + b.cobertura.toLowerCase() : "", ...lista(b.idioma), ...lista(b.modalidad)]
      .filter((t) => t && norm(t) !== "otro").forEach((t) => tags.append(el("li", null, t)));
    if (tags.children.length) cuerpo.append(tags);
    if (b.area) { const a = el("p", "area"); a.append(el("span", null, "Área: "), b.area.replace(/\s*;\s*/g, ", ")); cuerpo.append(a); }

    // El resto, plegado
    const edad = b.edadMin && b.edadMax ? b.edadMin + " a " + b.edadMax + " años" : b.edadMin ? "Desde " + b.edadMin + " años" : b.edadMax ? "Hasta " + b.edadMax + " años" : "";
    const dl = el("dl", "datos");
    [["Enfoque", b.subarea], ["Nacionalidades elegibles", b.nacionalidades], ["Edad", edad],
     ["Duración de la financiación", b.duracion], ["El monto depende de", b.importeSegun], ["Plazas", b.plazas]].forEach(([k, v]) => {
      if (!v) return;
      dl.append(el("dt", null, k), el("dd", null, v.replace(/\s*;\s*/g, ", ")));
    });
    let ul = null;
    if (b.listaPlazos.length > 1 || (b.listaPlazos[0] && (b.listaPlazos[0].ronda || b.listaPlazos[0].hora))) {
      ul = el("ul", "plazos");
      b.listaPlazos.slice().sort((x, y) => parseFecha(x.fecha) - parseFecha(y.fecha)).forEach((p) => {
        const partes = [...new Set([p.ronda, p.tipo].filter(Boolean))].join(" · ");
        const hora = p.hora ? ", " + p.hora + (p.zona ? " (" + p.zona + ")" : "") : "";
        const li2 = el("li", diasHasta(parseFecha(p.fecha)) < 0 ? "pasado" : null);
        if (partes) li2.append(el("strong", null, partes), document.createElement("br"));
        li2.append(fmt(parseFecha(p.fecha)) + hora +
          (p.ingreso && norm(p.ingreso) !== "no especificado" ? " · ingreso " + p.ingreso.toLowerCase() : "") +
          (norm(p.estado) === "estimada" ? " · estimada" : ""));
        ul.append(li2);
      });
    }
    if (dl.children.length || ul) {
      const det = el("details", "mas");
      det.append(el("summary", null, "Ver detalles"));
      if (ul) det.append(el("p", "sub", "Fechas límite"), ul);
      if (dl.children.length) det.append(dl);
      cuerpo.append(det);
    }

    const pie = el("div", "pie-card");
    const vf = parseFecha(b.verificado);
    if (vf) pie.append(el("p", "verificado", "Verificada el " + fmt(vf)));
    if (/^https?:\/\//i.test(b.enlace)) {
      const a = el("a", "ir", "Ver convocatoria oficial"); a.href = b.enlace; a.target = "_blank"; a.rel = "noopener";
      pie.append(a);
    }
    if (b.id) pie.append(el("p", "codigo", "Código: " + b.id)); // para ubicar la beca en la hoja y corregirla por ID
    cuerpo.append(pie);
    li.append(cuerpo);
    return li;
  }

  function aplicar() {
    const q = norm($("#q").value);
    const soloAbiertas = $("#abiertas").checked;
    const sel = Object.fromEntries(FILTROS.map((f) => [f, norm($("#f-" + f).value)]));
    const res = becas.filter((b) => {
      if (soloAbiertas && !b.estado.abierta) return false;
      for (const f of FILTROS) if (sel[f] && !valores(b, f).some((v) => norm(v) === sel[f])) return false;
      if (q && !norm([b.id, b.nombre, b.pais, b.nacionalidades, b.nivel, b.tipoPosgrado, b.area, b.subarea, b.cobertura, b.idioma, b.modalidad].join(" ")).includes(q)) return false;
      return true;
    });
    const lista = $("#lista"); lista.replaceChildren(...res.map(tarjeta));
    $("#vacio").hidden = res.length > 0;
    $("#conteo").textContent = res.length === 1 ? "1 beca" : res.length + " becas";
    const activos = FILTROS.filter((f) => sel[f]).length + (soloAbiertas ? 1 : 0) + (q ? 1 : 0);
    $("#limpiar").hidden = activos === 0;
    escribirURL();
  }

  // --- Pestaña "Resumen de las becas": conteos calculados con las becas publicadas ---
  const CONTEOS = [["nivel", "Por nivel"], ["cobertura", "Por cobertura"], ["area", "Por área"], ["pais", "Por país"], ["idioma", "Por idioma"], ["modalidad", "Por modalidad"]];
  const MAX_FILAS = 12; // países e idiomas pueden ser muchos: se muestran los más frecuentes

  function pintarResumen() {
    const el = (tag, clase, texto) => { const e = document.createElement(tag); if (clase) e.className = clase; if (texto != null) e.textContent = texto; return e; };
    const panel = $("#panel-resumen");
    const n = (f) => becas.filter(f).length;
    const tiles = [
      ["Becas publicadas", becas.length, "En el buscador", ""],
      ["Abiertas", n((b) => b.estado.clase === "abierta"), "Cierran en más de " + DIAS_CIERRA_PRONTO + " días o están abiertas todo el año", "abierta"],
      ["Cierran pronto", n((b) => b.estado.clase === "pronto"), "En los próximos " + DIAS_CIERRA_PRONTO + " días", "pronto"],
      ["Cerradas", n((b) => /^cerrada/i.test(b.estado.etiqueta)), "Ya pasaron todas sus fechas", "cerrada"],
      ["Sin fecha publicada", n((b) => b.estado.etiqueta === "Sin fecha publicada"), "Revisa el enlace de cada ficha", "cerrada"],
    ];
    const revisar = n((b) => /^revisar/i.test(b.estado.etiqueta));
    if (revisar) tiles.push(["Fecha por revisar", revisar, "La fecha no se pudo leer", "cerrada"]);

    const fila = el("div", "stats");
    tiles.forEach(([titulo, valor, ayuda, clase]) => {
      const t = el("div", "stat");
      t.append(el("p", "stat-titulo", titulo), el("p", "stat-num" + (clase ? " " + clase : ""), valor.toLocaleString("es-HN")), el("p", "stat-ayuda", ayuda));
      fila.append(t);
    });

    const grupos = el("div", "grupos");
    CONTEOS.forEach(([campo, titulo]) => {
      const mapa = new Map();
      becas.forEach((b) => [...new Set(valores(b, campo).map(norm))].forEach((k) => {
        const v = valores(b, campo).find((x) => norm(x) === k);
        if (!mapa.has(k)) mapa.set(k, { v, c: 0 });
        mapa.get(k).c++;
      }));
      const todos = [...mapa.values()].sort((a, b) => b.c - a.c || a.v.localeCompare(b.v, "es"));
      if (!todos.length) return;
      const lista = todos.slice(0, MAX_FILAS);
      const max = lista[0].c;
      const g = el("section", "grupo");
      g.append(el("h3", null, titulo));
      const ul = el("ul", "conteo");
      lista.forEach(({ v, c }) => {
        const li = el("li");
        const bt = el("button", "fila");
        bt.type = "button";
        bt.title = "Ver las " + c + " becas: " + v;
        const barra = el("span", "barra-fondo");
        const dato = el("span", "barra-dato");
        dato.style.width = Math.max(2, (c / max) * 100) + "%";
        barra.append(dato);
        bt.append(el("span", "fila-et", v), barra, el("span", "fila-num", c.toLocaleString("es-HN")));
        bt.addEventListener("click", () => verFiltro(campo, v));
        li.append(bt);
        ul.append(li);
      });
      g.append(ul);
      if (todos.length > MAX_FILAS) g.append(el("p", "grupo-nota", "Y " + (todos.length - MAX_FILAS) + " más. Usa el filtro en Buscar becas para verlos todos."));
      grupos.append(g);
    });

    panel.replaceChildren(
      el("p", "resumen-intro", "Toca cualquier fila para ver esas becas. Una beca puede contar en varias filas (por ejemplo, si es de pregrado y posgrado o de varias áreas)."),
      fila, grupos);
  }

  // Desde el resumen: abre Buscar becas con ese único filtro puesto
  function verFiltro(campo, valor) {
    $("#q").value = ""; $("#abiertas").checked = false;
    FILTROS.forEach((f) => ($("#f-" + f).value = ""));
    const op = [...$("#f-" + campo).options].find((o) => norm(o.value) === norm(valor));
    if (op) $("#f-" + campo).value = op.value;
    mostrarPestana("buscar");
    aplicar();
    $("#panel-buscar").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function mostrarPestana(cual, sinURL) {
    pestana = cual === "resumen" ? "resumen" : "buscar";
    ["buscar", "resumen"].forEach((p) => {
      const activa = p === pestana;
      $("#panel-" + p).hidden = !activa;
      $("#tab-" + p).setAttribute("aria-selected", String(activa));
      $("#tab-" + p).tabIndex = activa ? 0 : -1;
    });
    if (!sinURL) escribirURL(); // al cargar no se escribe: los filtros de la URL todavía no se leyeron
  }

  function iniciarPestanas() {
    const tabs = ["buscar", "resumen"];
    tabs.forEach((p, i) => {
      const tab = $("#tab-" + p);
      tab.addEventListener("click", () => mostrarPestana(p));
      tab.addEventListener("keydown", (e) => { // flechas para moverse entre pestañas
        if (e.key !== "ArrowRight" && e.key !== "ArrowLeft") return;
        const otra = tabs[(i + (e.key === "ArrowRight" ? 1 : tabs.length - 1)) % tabs.length];
        mostrarPestana(otra);
        $("#tab-" + otra).focus();
      });
    });
    if (new URLSearchParams(location.search).get("pestana") === "resumen") mostrarPestana("resumen", true);
  }

  async function iniciar() {
    $("#aviso-ejemplo").hidden = !/datos-ejemplo\.csv$/.test(CSV_URL);
    try {
      const resp = await fetch(CSV_URL, { cache: "no-store" });
      if (!resp.ok) throw new Error("HTTP " + resp.status);
      const rows = parseCSV(await resp.text());
      if (rows.length < 2) {
        $("#estado-carga").textContent = "Todavía no hay becas publicadas. Vuelve pronto.";
        $("#panel-resumen .vacio").textContent = "Todavía no hay becas publicadas. Vuelve pronto.";
        return;
      }
      const cab = rows[0].map(norm);
      const faltan = Object.entries(COLUMNAS).filter(([k, n]) => ["nombre", "tipoCierre", "cierre", "enlace"].includes(k) && !cab.includes(norm(n))).map(([, n]) => n);
      if (faltan.length) {
        $("#estado-carga").textContent = "La hoja no tiene estas columnas: " + faltan.join(", ") + ". Revisa los encabezados o COLUMNAS en config.js.";
        return;
      }
      becas = aObjetos(rows)
        .filter((b) => b.nombre && (COLUMNAS.aprobada == null || aprobada(b.aprobada)))
        .map((b) => { const listaPlazos = leerPlazos(b.plazos); return { ...b, listaPlazos, estado: calcularEstado(b.tipoCierre, b.cierre, listaPlazos) }; })
        .sort((a, b) => a.estado.orden - b.estado.orden || a.nombre.localeCompare(b.nombre, "es"));
    } catch (e) {
      console.error(e);
      $("#estado-carga").textContent = "No se pudieron cargar las becas. Revisa tu conexión e intenta de nuevo más tarde.";
      $("#panel-resumen .vacio").textContent = "No se pudieron cargar las becas. Revisa tu conexión e intenta de nuevo más tarde.";
      return;
    }
    $("#estado-carga").hidden = true;
    llenarSelects();
    leerURL();
    pintarResumen();
    ["#q", "#abiertas", ...FILTROS.map((f) => "#f-" + f)].forEach((s) => $(s).addEventListener(s === "#q" ? "input" : "change", aplicar));
    $("#limpiar").addEventListener("click", () => {
      $("#q").value = ""; $("#abiertas").checked = false; FILTROS.forEach((f) => ($("#f-" + f).value = "")); aplicar();
    });
    aplicar();
  }

  // --- Vista celular / web (celular por defecto; index.html la aplica antes de pintar) ---
  function iniciarVista() {
    const botones = document.querySelectorAll(".vista button[data-vista]");
    const marcar = (v) => {
      document.documentElement.dataset.vista = v;
      botones.forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.vista === v)));
    };
    marcar(document.documentElement.dataset.vista || "celular");
    botones.forEach((b) => b.addEventListener("click", () => {
      marcar(b.dataset.vista);
      try { localStorage.setItem("vista", b.dataset.vista); } catch (e) { /* sin almacenamiento: la vista no se recuerda */ }
    }));
  }

  // --- Modo claro / oscuro (sin elección guardada sigue el del sistema) ---
  function iniciarTema() {
    const raiz = document.documentElement;
    const botones = document.querySelectorAll(".tema button");
    const oscuroSistema = window.matchMedia ? window.matchMedia("(prefers-color-scheme: dark)") : null;
    const actual = () => raiz.dataset.theme || (oscuroSistema && oscuroSistema.matches ? "dark" : "light");
    const marcar = () => botones.forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.tema === actual())));
    marcar();
    if (oscuroSistema && oscuroSistema.addEventListener) oscuroSistema.addEventListener("change", marcar);
    botones.forEach((b) => b.addEventListener("click", () => {
      raiz.dataset.theme = b.dataset.tema;
      try { localStorage.setItem("tema", b.dataset.tema); } catch (e) { /* sin almacenamiento: no se recuerda */ }
      marcar();
    }));
  }

  // Expuesto solo para pruebas
  window.__becas = { calcularEstado, parseCSV, leerPlazos };
  document.addEventListener("DOMContentLoaded", () => { iniciarTema(); iniciarVista(); iniciarPestanas(); iniciar(); });
})();
