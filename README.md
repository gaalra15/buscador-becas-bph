# Buscador de becas · Becas para Hondureños

Buscador web de becas para estudiantes hondureños, hecho para la organización Becas para Hondureños (BpH).

## Cómo funciona

```
Google Form → Google Sheets → pestaña "Publicadas" (CSV publicado) → esta página
```

- El equipo de BpH registra y corrige becas solo desde un formulario de Google; un Apps Script asigna un ID a cada beca (BPH-0001…), aprueba lo que envía la cuenta de BpH y avisa por correo.
- La página es estática (HTML, CSS y JavaScript sin dependencias ni compilación) y lee el CSV publicado de la hoja, así que cada beca nueva aparece sola, sin volver a desplegar.
- El estado (abierta, cierra pronto, cerrada) se calcula con las fechas límite; una beca puede tener varias rondas.
- Búsqueda por texto, filtros por nivel, país, área, cobertura, idioma y modalidad, modo claro/oscuro y vista celular/web. Los filtros quedan en la URL para compartir enlaces como `?nivel=Posgrado&abiertas=1`.

## Archivos

| Archivo | Qué hace |
|---|---|
| `index.html` | Estructura de la página |
| `styles.css` | Estilos con la marca de BpH |
| `app.js` | Lee el CSV, calcula estados, filtra y dibuja las tarjetas |
| `config.js` | URL del CSV, nombres de columnas y días para «Cierra pronto» |
| `logo.jpg` | Logo de BpH |

## Desplegar

GitHub Pages (rama `main`, carpeta raíz): cada push a `main` se publica solo en gaalra15.github.io/buscador-becas-bph. El script de Google vive en `apps-script/id-de-beca.gs` y se pega a mano en Apps Script (ver `VERSION_SCRIPT`).

## Probar en local

```
python -m http.server 8000
```

y abrir `http://localhost:8000`.
