# Preparador de Envíos Crea+

Herramienta web estática para validar bases de destinatarios y generar archivos Excel por bloques, listos para que Power Automate los lea desde una tabla.

## ¿Qué hace?

- Lee un Excel grande de destinatarios directamente en el navegador.
- Permite elegir la hoja cuando el archivo tiene varias.
- Mapea columnas de correo, nombre y variables opcionales.
- Valida correos vacíos, formato básico de correo, duplicados y filas incompletas.
- Mantiene trazabilidad con `id_envio` e `id_campaña`.
- Divide los registros válidos en bloques configurables; por defecto, 250 filas.
- Descarga un ZIP con:
  - `bloque_001.xlsx`, `bloque_002.xlsx`, etc.
  - `reporte_errores.xlsx`
  - `resumen_campaña.xlsx`
  - `plantilla_correo.html`

## ¿Qué NO hace?

- No envía correos.
- No usa credenciales.
- No usa Microsoft Graph.
- No tiene backend.
- No se conecta a Power Automate.
- No sube bases, plantillas ni reportes a ningún servidor.

## Uso local

Abre `index.html` en el navegador o publícalo como sitio estático, por ejemplo con GitHub Pages.

1. Escribe el nombre de la campaña.
2. Sube el Excel de destinatarios.
3. Elige la hoja y mapea columnas.
4. Escribe asunto y mensaje HTML.
5. Presiona **Validar base**.
6. Revisa el resumen.
7. Presiona **Generar y descargar ZIP**.

## Estructura de los bloques

Cada `bloque_XXX.xlsx` contiene una hoja llamada `Envios` y una tabla llamada `TablaEnvios` con estas columnas:

- `id_envio`
- `id_campaña`
- `nombre_campaña`
- `bloque`
- `correo`
- `nombre`
- `asunto`
- `mensaje_html`
- `estado_envio`
- `fecha_envio`
- `error_envio`
- `respondio`
- `fecha_respuesta`
- `tipo_respuesta`
- `estado_seguimiento`
- `observacion`

Valores iniciales incluidos:

- `estado_envio = PENDIENTE`
- `respondio = NO`
- `estado_seguimiento = SIN ENVIAR`
- `fecha_envio` vacío
- `error_envio` vacío

## Privacidad

La aplicación usa librerías por CDN (`SheetJS`, `JSZip` y `FileSaver`) y procesa todo localmente en el navegador. Los datos del Excel no salen del equipo de la persona usuaria.

## Desarrollo y pruebas

Validación rápida de sintaxis:

```bash
node --check app.js
```

Para probar funcionalmente, usa un Excel ficticio con columnas como `correo`, `nombre`, `celular`, `dni`, `ciudad`, `institución` y algunos correos duplicados o inválidos.
