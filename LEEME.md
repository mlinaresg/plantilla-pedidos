# Complemento "Plantilla reenvío pedidos"

Botón en la cinta de redacción de Outlook (Clásico, Nuevo Outlook y Web)
que, al pulsarlo, **añade `pedidosclientes@plexus.es` como destinatario si no
estaba ya puesto**, e inserta arriba del cuerpo:

```
NOMBRE CLIENTE:
DESCRIPCIÓN DEL PROYECTO:
FECHA DE FACTURACIÓN APROXIMADA:
DIRECCIÓN DE ENTREGA: (si no se indica nada, la dirección de entrega será Santiago)
```

Flujo pensado para comerciales reenviando pedidos: **Reenviar > pulsar el
botón**, y ya está — no hace falta escribir ni buscar el destinatario a
mano.

## 1. Configuración actual: SOLO PC local, sin subir a ningún sitio

Este `manifest.xml` ya está apuntando a `https://localhost:3000/`, así que
no hace falta editarlo. Solo tienes que servir la carpeta en tu propio PC:

1. Instala Node.js si no lo tienes.
2. En una terminal: `npm install -g office-addin-dev-certs http-server`
   y luego `npx office-addin-dev-certs install` (instala y da de confianza
   un certificado HTTPS local; solo hace falta una vez).
3. Desde dentro de esta carpeta, ejecuta:
   ```
   http-server . -S -C "%USERPROFILE%\.office-addin-dev-certs\localhost.crt" -K "%USERPROFILE%\.office-addin-dev-certs\localhost.key" -p 3000
   ```
   y deja esa ventana abierta mientras quieras usar el botón. (O usa el
   arranque automático de la sección 6 para no tener que hacerlo a mano.)

**Límites de este modo:** solo funciona en Outlook Clásico/Nuevo de este PC,
y en Outlook Web si lo abres desde un navegador de este mismo PC. No
funciona desde el móvil ni desde otro equipo.

**Cuando decidas alojarlo en algún sitio (GitHub Pages u otro):** sube estos
mismos archivos y luego reemplaza las apariciones de `https://localhost:3000`
en `manifest.xml` por la URL real — es el único cambio necesario.

## 2. Instalar el complemento en Outlook (una vez por cuenta)

**En Outlook Web o Nuevo Outlook:**
Configuración (engranaje) > Correo > Personalizar acciones > Complementos, o
bien Inicio > Obtener complementos > Mis complementos > Agregar un
complemento personalizado > Agregar desde archivo, y selecciona tu
`manifest.xml`.

**En Outlook clásico:**
Archivo > Administrar complementos (te abre el mismo panel de la web) —o
directamente en la cinta: Inicio > Obtener complementos > Mis complementos >
Agregar complemento personalizado > Agregar desde archivo.

Una vez instalado en cualquiera de los dos, queda disponible también en el
otro y en el móvil, porque se guarda asociado a tu buzón.

## 3. Probarlo

1. Abre o reenvía un correo cualquiera (sin poner ningún destinatario).
2. En la cinta de redacción, pulsa el botón **"Insertar plantilla"**
   (grupo "Pedidos clientes").
3. Deberías ver `pedidosclientes@plexus.es` añadido solo en "Para", y los
   4 campos insertados arriba del cuerpo.
4. Prueba también con `pedidosclientes@plexus.es` ya puesto de antemano:
   no debería duplicarse, solo insertar los campos.

## 4. Personalizar

- **Cambiar el correo objetivo:** en `commands.js`, primera línea con
  `CORREO_OBJETIVO`.
- **Cambiar los campos o el texto:** funciones `plantillaHtml()` y
  `plantillaTexto()` en `commands.js` (mantén ambas en sincronía si editas
  una).
- **Cambiar el nombre/icono del botón:** en `manifest.xml`, dentro de
  `bt:ShortStrings` (`Button.Label`) y `bt:Images`.

## 5. Nota conocida (no es un fallo de este código)

En Nuevo Outlook y Outlook Web, tras insertar la plantilla el cursor salta al
final del cuerpo en vez de quedarse junto al texto insertado. Es una
limitación documentada de la API de Office (no permite posicionar el cursor
tras `prependAsync`), no algo que se pueda corregir desde el complemento.

## 6. Arranque automático al iniciar sesión

Si no quieres tener que abrir la terminal y lanzar `http-server` cada vez,
hay dos archivos que lo automatizan:

- **`start-servidor.ps1`** — hace las mismas comprobaciones que hemos ido
  viendo (que exista la carpeta, los archivos, el certificado, que
  `http-server` esté instalado, y que no haya ya un servidor escuchando en
  el puerto 3000) y, si todo está bien, arranca el servidor **oculto**
  (sin ventana visible). Todo queda registrado en `servidor.log`, dentro de
  esta misma carpeta, por si algún día el botón de Outlook no responde y
  hay que ver qué pasó.
- **`registrar-tarea.ps1`** — script de instalación, se ejecuta **una sola
  vez**: crea una tarea en el Programador de tareas de Windows que lanza
  `start-servidor.ps1` automáticamente cada vez que inicias sesión. No
  requiere permisos de administrador.

Ambos asumen que la carpeta está en `C:\QGPL\Outlook\plantilla-addin`; si la
mueves de sitio, edita la ruta al principio de los dos archivos.

**Para activarlo, en PowerShell:**
```
cd C:\QGPL\Outlook\plantilla-addin
.\registrar-tarea.ps1
```

**Para probarlo ya, sin reiniciar sesión:**
```
Start-ScheduledTask -TaskName "Plantilla pedidos - servidor local"
```
Luego revisa `servidor.log` para confirmar que dice "arrancado correctamente",
y prueba el botón en Outlook.

**Para quitarlo más adelante**, si algún día ya no lo necesitas:
```
Unregister-ScheduledTask -TaskName "Plantilla pedidos - servidor local"
```

### Si el Programador de tareas te da "Acceso denegado"

Algunas políticas de empresa lo bloquean para usuarios estándar. Alternativa
que casi nunca está restringida: un acceso directo en la carpeta de Inicio
de Windows, usando `arrancar-oculto.vbs` + `instalar-inicio-automatico.ps1`.

```
cd C:\QGPL\Outlook\plantilla-addin
Unblock-File .\arrancar-oculto.vbs
Unblock-File .\instalar-inicio-automatico.ps1
.\instalar-inicio-automatico.ps1
```

Para probarlo ya, sin reiniciar sesión:
```
wscript.exe "C:\QGPL\Outlook\plantilla-addin\arrancar-oculto.vbs"
```
y revisa `servidor.log` igual que antes.

Para quitarlo más adelante: borra el archivo `Plantilla pedidos - servidor
local.lnk` de la carpeta de Inicio (escribe `shell:startup` en el Explorador
de archivos para llegar directo ahí).

## 7. Camino a futuro: detección automática

Este complemento ya está preparado para, el día que interese, añadirle
detección automática (que salte solo al añadir el destinatario, sin pulsar
el botón) mediante el evento `OnMessageRecipientsChanged`. Esa parte requiere
que un administrador de Microsoft 365 de Plexus lo despliegue de forma
centralizada desde el panel de administración — el botón manual seguirá
funcionando igual mientras tanto.
