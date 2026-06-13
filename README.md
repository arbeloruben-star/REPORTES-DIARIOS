# App Reportabilidad Soldesp - Komatsu

Aplicacion web/PWA para registrar asistencia, actividades, resumen de turno y exportacion Excel.

## Ejecucion en Windows

1. Abrir esta carpeta.
2. Ejecutar `ejecutar_app.bat`.
3. Entrar desde el PC a `http://localhost:5000`.
4. Desde un telefono en la misma red, entrar a `http://IP_DEL_PC:5000`.

La base SQLite `reportabilidad.db` se crea automaticamente con datos iniciales.

## Flujo recomendado en terreno

1. Revisar o cargar asistencia del turno.
2. Entrar a `Registro de actividades`.
3. Usar `Iniciar frente`, `Cerrar frente`, `Marcar no ejecutado` o `Duplicar anterior`.
4. En `Cierre de turno`, revisar todos los registros del dia antes de exportar.
5. Exportar Excel desde `Cierre de turno` o el menu superior.

## PWA

La app incluye `manifest.webmanifest` y `service-worker.js`.
En produccion con HTTPS, el telefono puede instalarla como aplicacion desde el navegador.
La version actual cachea la interfaz y muestra una pantalla offline; el guardado offline con sincronizacion queda como fase siguiente.

## Despliegue en nube

Archivos incluidos:

- `requirements.txt`
- `Procfile`
- `runtime.txt`
- `render.yaml`

Para Render, el `render.yaml` usa un disco persistente en `/var/data` y guarda SQLite en `DATABASE_PATH=/var/data/reportabilidad.db`.
En otros proveedores se puede usar `gunicorn app:app` y configurar `DATABASE_PATH` hacia una carpeta persistente.

## Envio de correos

La app puede enviar la asistencia inicial y el informe final por correo usando una casilla Microsoft 365 / Outlook configurada en variables de entorno.
No guardar usuarios ni contrasenas dentro del codigo.

Variables principales en Render:

- `SMTP_HOST`: `smtp.office365.com`
- `SMTP_PORT`: `587`
- `SMTP_USER`: correo emisor
- `SMTP_PASSWORD`: clave o clave de aplicacion del correo emisor
- `SMTP_FROM`: correo emisor, normalmente igual a `SMTP_USER`
- `MAIL_FROM_NAME`: nombre visible del remitente
- `ASSISTANCE_MAIL_TO`: destinatarios de asistencia, separados por coma o punto y coma
- `REPORT_MAIL_TO`: destinatarios del informe final
- `MAIL_CC`: copias comunes opcionales

## Reglas HH implementadas

- Ejecutado: HH directas.
- Parcial: permite separar HH directas e indirectas; si no se separa, todo queda directo con advertencia.
- Stand-by: HH indirectas o no utilizadas segun causal.
- No ejecutado: HH no utilizadas.
- C07: brecha por capacidad insuficiente / falta de dotacion.
- C02, C03, C04, C05, C06 y C08: brecha por condicion operacional o mandante.
