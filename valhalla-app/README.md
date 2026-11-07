# Valhalla · App de gestión

Web app para celular de **Valhalla (estética vehicular)**: cotizador de tiempo y precio por marca y modelo, agenda del taller, clientes, presupuestos y trabajos realizados. Con perfil **admin** para stock de productos y agenda de contactos.
Usa un **Google Sheet como base de datos** (planilla `Valhalla_Gestion`) a través de Google Apps Script, y se publica gratis en **GitHub Pages**.

Sin configurar nada, la app arranca en **modo demo** con datos de prueba, así se puede mostrar enseguida (usuarios demo: `admin` / `1234` y `operador` / `0000`).

## Usuarios y roles

Los usuarios se cargan en la hoja **Usuarios** (usuario, nombre, PIN, rol).

| Rol | Puede |
|---|---|
| `admin` | Todo, más la pestaña **Admin**: stock (productos, entradas/salidas, reposición) y contactos (alta, edición, importar/exportar) |
| `operador` | Cotizar, agenda, clientes, presupuestos, turnos y trabajos. No ve stock ni contactos |

El PIN se valida en el Apps Script (del lado de Google), no en el navegador, y los operadores no reciben los datos de stock ni de contactos. **Cambiá los PIN de ejemplo antes de usarla en serio.**

## Estructura

```
index.html              → la app (HTML + CSS + JS, sin dependencias)
manifest.webmanifest    → permite "Agregar a pantalla de inicio"
icon.svg                → ícono
apps-script/Code.gs     → API que lee y escribe el Google Sheet
```

## 1. Conectar el Google Sheet (Apps Script)

1. Subí `Valhalla_Gestion.xlsx` a Google Drive y abrilo con **Hojas de cálculo de Google** (Archivo → Guardar como Hojas de cálculo de Google).
2. En la planilla: **Extensiones → Apps Script**.
3. Borrá lo que haya y pegá el contenido de `apps-script/Code.gs`. Guardá.
4. Elegí la función `probar` y tocá **Ejecutar** → autorizá los permisos. En el registro tiene que mostrar cuántos clientes y servicios leyó.
5. **Implementar → Nueva implementación → Tipo: Aplicación web**
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier usuario**
6. Copiá la URL que termina en `/exec`.

> Cada vez que cambies `Code.gs`, hacé **Implementar → Gestionar implementaciones → Editar → Nueva versión** para que la URL use el código nuevo.

## 2. Publicar la app en GitHub Pages

1. Creá un repo (por ejemplo `valhalla-app`) y subí estos archivos.
2. En el repo: **Settings → Pages → Branch: `main` / root → Save**.
3. En un minuto queda en `https://TU-USUARIO.github.io/valhalla-app/`.

## 3. Conectar la app con el Sheet

Dos opciones:

- **Fija:** en `index.html`, completá `CONFIG.API_URL` con la URL `/exec`.
- **Desde la app:** en la pantalla de ingreso, *Ajustes de conexión* → pegar la URL → *Guardar y conectar* (queda guardada en ese celular).

Después se entra con usuario y PIN de la hoja Usuarios. El indicador arriba a la derecha pasa de **DEMO** a **EN VIVO**.

## 4. Instalarla en el celular

Abrí la URL en el celular → menú del navegador → **Agregar a pantalla de inicio**. Se abre como una app, a pantalla completa.

## Cómo calcula

- Días estimados = días base del servicio × factor del vehículo, redondeado a medio día.
- El factor sale del segmento del modelo (Chico 0,8 · Mediano 1 · SUV 1,25 · Pick-up 1,4 · Grande 1,5) o del ajuste propio de la hoja `Marcas_Modelos`.
- La entrega cuenta días hábiles según `Parametros` (días no laborables y feriados).
- Precio = precio base × factor de precio del segmento.

Es la misma lógica que las fórmulas del Sheet, así que la app y la planilla siempre coinciden.

## WhatsApp

- **Botones de WhatsApp (incluidos, gratis):** cotización, "tu auto está listo", contactar clientes/proveedores y **pedido de reposición al proveedor** con la lista de productos bajos. Abren WhatsApp con el mensaje escrito; se envía a mano.
- **Contactos:** WhatsApp no deja leer ni sincronizar contactos desde afuera; usa la agenda del teléfono. Por eso la app **importa** un archivo `.vcf` exportado del celular y **exporta** la agenda de Valhalla como `.vcf` para cargarla en el teléfono (y así aparece en WhatsApp).
- **Envío automático (opcional, a futuro):** con la WhatsApp Business Platform (Cloud API de Meta) el Apps Script podría mandar avisos solos (turno confirmado, auto listo). Requiere cuenta de Meta Business verificada, un número dedicado y plantillas aprobadas; Meta cobra por mensaje de plantilla, y las respuestas dentro de las 24 h de que escribió el cliente son gratis.

## Seguridad

La URL del Apps Script es pública, pero cada pedido exige usuario y PIN válidos de la hoja Usuarios. Usá PIN distintos a los de ejemplo, no compartas la planilla y cambiá el PIN de alguien que deje de trabajar en Valhalla.
