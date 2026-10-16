// URL del CSV publicado de la pestaña "Publicadas" (Archivo > Compartir > Publicar en la web > CSV).
// Es la ÚNICA línea que hay que cambiar para usar los datos reales:
const CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTRp3SYQuRF13-jq918KdkJnedCK-fFFhuQQ6h2lN5quyO6EAhN_7ME2uPhe2LtHxxzaSWVmK0c-DKo/pub?gid=1027982118&single=true&output=csv";

// Nombres de las columnas tal como aparecen en la primera fila de la hoja.
// Solo cambia el texto de la derecha si en la hoja se llaman distinto.
const COLUMNAS = {
  nombre: "nombre",
  pais: "pais",
  nivel: "nivel",
  area: "area",
  subarea: "subarea",
  cobertura: "cobertura",
  idioma: "idioma",
  modalidad: "modalidad",
  tipoCierre: "tipo_cierre",
  cierre: "cierre",
  enlace: "enlace",
  aprobada: "aprobada",
  // Campos nuevos (opcionales: si la hoja no los tiene, la tarjeta simplemente no los muestra)
  nacionalidades: "nacionalidades",
  tipoPosgrado: "tipo_posgrado",
  edadMin: "edad_min",
  edadMax: "edad_max",
  duracion: "duracion",
  importeSegun: "importe_segun",
  plazas: "plazas",
  plazos: "plazos",
  verificado: "verificado",
};

// Días antes del cierre en que la beca se marca como "Cierra pronto".
const DIAS_CIERRA_PRONTO = 30;
