# Portal de certificados QCP

Aplicación web en español para convertir certificados Excel a PDF por lotes, con una imagen de firma y un QR por certificado. Basada en el comportamiento del flujo n8n **«Portal certificados QCP - lotes dinámicos v3»**, con una implementación independiente: no necesita n8n, QuickChart ni el servicio externo `pdf-code-fixer`.

![Vista del portal en desarrollo](docs/portal.png)

La pantalla usa el nombre y el emblema de **Quality Control Perú S.A.C.**, tomados del membrete de la plantilla proporcionada. Incluye una guía de tres pasos, contador de archivos, firma opcional desplegable y presentación adaptada a computadora y celular. La página de consulta comparte esa identidad visual. La ilustración del inicio es decorativa; los certificados se generan desde los archivos cargados.

## Qué hace

- Recibe hasta 20 archivos `.xlsx` por lote desde el navegador, procesados uno a uno. También conserva la entrada PDF de la versión anterior. Cada archivo admite 15 MB y el resultado hasta 100 páginas.
- Exporta únicamente la hoja **Certificado**, con su área de impresión, filas de encabezado y saltos de página. Conserva los resultados guardados de las fórmulas; no recalcula mediciones ni actualiza conexiones de datos. El Excel original no se guarda ni se modifica.
- En la plantilla Excel, coloca un QR de 66 puntos **a la izquierda de «Escanee este QR»**. Conserva la imagen de firma situada junto a «Autorizado y firmador por:» o «Autorizado y firmado por:», y omite los sellos circulares conforme al PDF de referencia. Cada archivo obtiene un identificador y un QR distintos, incluso si repite el mismo código.
- Detecta códigos `QCP-0000-0000`. Si existe un único código distinto, permite reemplazar todas sus apariciones por el código solicitado. Conserva el nombre base y entrega un archivo `.pdf`.
- Si hay varios códigos, los conserva y muestra una advertencia. Si no encuentra texto editable, no inventa ni corrige códigos. No incluye OCR.
- Permite reemplazar la firma del Excel mediante una imagen PNG/JPG autorizada; es opcional y se aplica al lote. La ubica a la izquierda del texto del autor. La firma no se almacena como archivo independiente.
- Si la plantilla Excel no permite ubicar el QR o la firma sin cubrir texto o imágenes, devuelve un error para corregirla. No añade páginas a esa plantilla. Para entradas PDF genéricas se mantiene la búsqueda de espacios libres y, si hace falta, una página adicional de consulta.
- Aplica cifrado AES-256 con apertura sin contraseña y permisos que limitan edición y copia en lectores compatibles. Permite imprimir. Esos permisos no garantizan que otros programas no puedan modificar el PDF.
- Guarda el PDF final y sus metadatos en el servidor. El QR apunta a una página pública de consulta y descarga. La página muestra un hash SHA-256 del archivo.
- Rechaza PDFs con contraseña o campos de firma digital para no invalidarlos silenciosamente.

**La imagen de firma, el QR y el registro no constituyen una firma digital ni acreditan la identidad del emisor.** La sustitución de códigos conserva tamaño/color y aproxima la fuente con fuentes PDF estándar; documentos con tipografías especiales deben revisarse visualmente. Texto girado que requiera corrección se rechaza.

## Ejecutar en desarrollo

Requisitos: Python **3.12**, [uv](https://docs.astral.sh/uv/), **LibreOffice Calc** y fuentes Liberation, DejaVu y Carlito. En Debian/Ubuntu: `apt-get install libreoffice-calc fonts-liberation fonts-dejavu-core fonts-crosextra-carlito`. El Dockerfile instala el conversor y las fuentes automáticamente. Utiliza el checkout existente; no es necesario crear otro worktree.

```sh
cd /workspace/Chamba
sh scripts/install.sh
sh scripts/start.sh
```

En una instalación local fuera de Codex, entra en el directorio clonado de Chamba. El servidor escucha en el puerto 8000. Para cambiarlo, exporta `PORT` y `PUBLIC_BASE_URL` antes de arrancar. En desarrollo el origen por defecto es `http://127.0.0.1:8000`; los QR de ese modo sirven únicamente para pruebas en la máquina del servidor.

El archivo `.env.example` documenta los nombres de variables; **la aplicación no carga archivos `.env` automáticamente**. Exporta las variables o configúralas en tu proveedor. Nunca guardes claves ni PDFs de clientes en Git.

Para probar sin documentos reales, genera un PDF ficticio de dos páginas, súbelo al portal y solicita el código `QCP-2026-0099`:

```sh
.venv/bin/python scripts/create_demo.py --output /tmp/qcp-demo.pdf
```

## Validar

```sh
cd /workspace/Chamba
.venv/bin/python -m pytest -q
```

Las pruebas generan Excel y PDFs ficticios; comprueban la conversión, resultados de fórmulas sin recálculo, exportación de una sola hoja, QR al costado y distinto por archivo, eliminación de sellos, conservación/reemplazo de firma, corrección del código, permisos, almacenamiento y errores. No utilizan certificados de clientes. Las pruebas de Excel requieren LibreOffice.

## Preparar el Excel

Usa un archivo `.xlsx` con una hoja llamada **Certificado**, área de impresión configurada y el texto **Escanee este QR** visible. La ubicación de la firma y los sellos se reconoce en el bloque de autorización de la plantilla QCP proporcionada; otras plantillas necesitan adaptación. Abre, calcula y guarda el archivo en Excel antes de subirlo: se rechazan fórmulas con errores o sin resultados guardados. No se admiten `.xls`, macros, archivos con contraseña u objetos incrustados.

Se usa LibreOffice para convertir una copia temporal. Los fondos de hoja de la plantilla se aplican como una imagen completa debajo del contenido; las alturas de las filas se fijan para respetar los saltos. Los formatos numéricos de la plantilla usan coma decimal y espacio entre millares, como el PDF de referencia; se conservan el formato General, las fechas y los formatos con moneda o región explícita. Algunas tipografías de Microsoft se sustituyen por fuentes disponibles en el servidor y pueden producir pequeñas diferencias visuales. Revisa los primeros PDFs de cada plantilla antes de usarla habitualmente.

Con el servidor arrancado, comprueba además una carga y descarga reales:

```sh
.venv/bin/python scripts/smoke.py
```

Esta comprobación guarda un certificado ficticio marcado SIN VALIDEZ. Usa `--url https://tu-portal.example.com` para comprobar una publicación. La carga y el procesamiento son públicos y no requieren contraseña.

## Publicar para obtener un enlace

El repositorio y la publicación web son cosas distintas: subir este código a GitHub **no inicia un servidor web**. Esta aplicación necesita un servidor Python con disco persistente; GitHub Pages no ejecuta este backend.

Puedes desplegar el `Dockerfile` en un proveedor que admita contenedores y almacenamiento persistente, por ejemplo un servicio web Docker en Render:

1. Conecta este repositorio y selecciona la rama `main` y el runtime Docker.
2. El portal permite cargar y procesar Excel y PDFs sin contraseña, por decisión del usuario. Si configuraste `PORTAL_ACCESS_KEY` en una versión anterior, puedes eliminarla: la aplicación ya no la utiliza. Se mantienen los límites de archivos, páginas y almacenamiento. La conversión se limita a un archivo simultáneo para reducir el uso de memoria.
3. Configura `PUBLIC_BASE_URL` con la dirección **HTTPS** del servicio, sin barra final ni ruta. En Render también se acepta su variable automática `RENDER_EXTERNAL_URL` si no defines `PUBLIC_BASE_URL`.
4. Monta un **disco persistente** en `/data`, escribible por el usuario del contenedor (UID 10001). El proveedor puede cobrar por el servicio o el disco; revisa su precio antes de crear recursos.
5. Usa `/healthz` como comprobación de salud. Abre la página y procesa un PDF ficticio; descarga el resultado, escanea su QR desde otro dispositivo y comprueba que abre la página de consulta.

Los procesos no sobreviven a un reinicio; el proveedor debe arrancar el contenedor nuevamente. Los archivos deben sobrevivir en el disco. Prueba un reinicio y confirma que un certificado anterior sigue disponible antes de utilizar documentos reales. No cambies el dominio después de emitir certificados: los QR existentes conservarán el dominio con el que fueron generados.

En **Render Free**, el almacenamiento local es temporal: un reinicio, suspensión o despliegue puede eliminar los certificados y dejar sus QR sin destino. Sirve para probar el flujo; los certificados que deban seguir disponibles requieren almacenamiento persistente (Google Drive sigue pendiente).

Con **Auto-Deploy** activado para `main`, cada cambio enviado a GitHub inicia un nuevo despliegue en Render. El sitio cambia cuando ese despliegue termina con estado **Live**; la dirección del portal se mantiene.

También puedes usar Docker en una máquina propia:

```sh
docker build -t chamba-qcp .
docker run --rm -p 8000:8000 --mount type=volume,source=qcp-data,target=/data chamba-qcp
```

Ese ejemplo es para pruebas locales. Para publicación usa HTTPS, la URL pública y un proxy frontal. `PUBLIC_BASE_URL` determina los QR; no se confía en encabezados de origen enviados por clientes.

### Variables

| Nombre | Uso |
| --- | --- |
| `PUBLIC_BASE_URL` | Origen HTTPS del portal publicado. Por defecto, origen local de pruebas. |
| `CERTIFICATE_STORAGE_DIR` | Directorio de certificados. Por defecto `data/certificates`; en Docker `/data/certificates`. |
| `MAX_STORAGE_MB` | Cuota total del directorio. Por defecto 1024 MB; no elimina archivos automáticamente. |
| `PORT` | Puerto de escucha. Por defecto 8000. |

Los enlaces usan identificadores aleatorios, pero **cualquier persona que tenga el enlace o el QR puede descargar el certificado**, igual que la opción «cualquiera con el enlace» del flujo original. Haz copias de seguridad del directorio de almacenamiento; borrarlo rompe los enlaces emitidos. Utiliza una única instancia y un único worker con ese disco. Para escalar a varias instancias se necesitaría almacenamiento compartido y una base de datos.

## Diferencias respecto al JSON original

Google Drive queda pendiente por decisión del usuario: esta versión guarda archivos en el servidor y los QR apuntan al portal. El código del servicio `pdf-code-fixer` no estaba incluido; el procesamiento se implementa aquí con LibreOffice y PyMuPDF. No se presupone la existencia de `firmaluis.png`: se conserva la firma de la plantilla o se carga una imagen autorizada en la interfaz.

El QR verifica la disponibilidad de la copia registrada, no la autenticidad del contenido. El código original no incluía la interfaz ni la implementación del servicio PDF; esta aplicación aporta ambos componentes con las limitaciones documentadas arriba.
