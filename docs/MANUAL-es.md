# MEFISTOFELE DeckWhisperer – Manual (Español)

**Remote notes & control for PowerPoint**
[English](MANUAL-en.md) · [Italiano](MANUAL-it.md) · [Français](MANUAL-fr.md) · [Español](MANUAL-es.md) · [العربية](MANUAL-ar.md) · [中文](MANUAL-zh.md)

En este manual, la **página de gestión** es la que se abre en el ordenador donde se ejecuta PowerPoint; la **página remota** es la que ves en el teléfono.

MEFISTOFELE DeckWhisperer lee las **notas del orador** de PowerPoint mientras la presentación está en marcha y las muestra en tiempo real
en una página web que abres en el teléfono (o en cualquier dispositivo de la misma red). Desde el teléfono también puedes avanzar y retroceder las diapositivas,
saltar a una diapositiva y ver la diapositiva actual y la siguiente.

## Requisitos
- **Windows 10/11** con **PowerPoint** de escritorio, o un **Mac** (macOS 12 o posterior) con **PowerPoint para Mac**.
- El ordenador y el teléfono en la **misma red** (Wi-Fi o cable).

## Inicio rápido
**Windows**
1. Descarga `mefistofele.exe` desde la página de [Releases](https://github.com/battacharlie/PPT-Notes/releases) (última versión).
2. Haz doble clic. Si Windows muestra «Windows protegió su PC»: **Más información → Ejecutar de todas formas**. En el primer inicio, permite el acceso a las **redes privadas**.

**Mac**
1. Descarga `mefistofele-mac.zip` desde la página de [Releases](https://github.com/battacharlie/PPT-Notes/releases) y descomprímelo.
2. Abre **Terminal**, escribe `chmod +x ` (con el espacio), arrastra el archivo `mefistofele` a la ventana y pulsa Intro. Después repite sin `chmod +x ` para ejecutarlo.
   Si macOS lo bloquea («no se puede abrir porque no se puede verificar al desarrollador»): ejecuta `xattr -d com.apple.quarantine ` seguido del archivo, o **Ajustes del Sistema → Privacidad y seguridad → Abrir igualmente**.
3. La primera vez, macOS pregunta si **Terminal puede controlar Microsoft PowerPoint**: pulsa **OK** (se puede cambiar luego en **Ajustes del Sistema → Privacidad y seguridad → Automatización**).

**Después (en ambos sistemas)**
- La **página de gestión** se abre en el navegador con un código QR. Escanéalo con el teléfono.
- En PowerPoint inicia la presentación (**Presentación con diapositivas → Desde el principio**): las notas aparecen en el teléfono y siguen las diapositivas.

El programa se ejecuta en una ventana negra (Terminal en Mac): **ciérrala para detener el programa**.

## La página de gestión (en el ordenador)
Dirección: `http://localhost:8080/start` (se abre sola). El botón **Guía** arriba a la derecha abre este manual en el idioma elegido.

**Idioma.** Elige el idioma en el menú **Idioma**: English (predeterminado), Italiano, Français, Español, العربية o 中文.
Se aplica a la página de gestión, a la página remota en **todos los teléfonos** (siguen la página de gestión) y a la guía. La elección se guarda.

**Adaptador de red del QR.** Si el ordenador tiene varios adaptadores de red (por ejemplo Wi-Fi y red de producción), elige en el menú
**Adaptador de red del QR** el que usan los teléfonos: el QR y la dirección cambian al instante. *Automático* deja que elija el programa.

Después elige **qué muestran los teléfonos**. Los elementos, por orden:

| Casilla | Efecto en el teléfono |
|---|---|
| Barra de herramientas | El teléfono muestra la barra de botones de colores, con la franja y el candado. **Si está desactivada**, el teléfono solo muestra los elementos elegidos abajo y **no se puede cambiar nada** |
| Estilo del texto | Fuente, tamaño, colores, interlineado, espaciado de letras |
| Diapositiva actual | Imagen de la diapositiva que ve el público |
| Diapositiva siguiente | Imagen de la próxima diapositiva |
| Ir a la diapositiva | Casilla numérica y botón **Ir** para saltar a una diapositiva |
| Navegación | Botones grandes **Siguiente** y **Anterior** |
| Bloquear pantalla remota | Quita **todos los controles** del teléfono: solo quedan las notas, que únicamente se pueden **desplazar** si superan la pantalla, más las vistas previas elegidas. Los comandos dejan de aceptarse |

También puedes definir un **PIN** para los controles: el teléfono lo pide la primera vez que pulsas Siguiente/Anterior o Ir.
Tus elecciones se **guardan** (`mefistofele.json` junto al programa) y llegan **al instante** a los teléfonos conectados.
La página de gestión se abre **solo en el ordenador** que ejecuta el programa.

## La página remota (en el teléfono)
Cada función tiene su color. Botón **relleno** = activado, botón **solo contorno** = desactivado.

| Color | Botón | Qué hace |
|---|---|---|
| Morado | Estilo del texto | Muestra/oculta los controles del texto |
| Verde | Navegación | Muestra/oculta Siguiente / Anterior e Ir a la diapositiva |
| Naranja | Diapositiva actual | Muestra/oculta la vista previa de la diapositiva actual |
| Azul | Diapositiva siguiente | Muestra/oculta la vista previa de la siguiente |
| Turquesa | Pantalla completa | Entra/sale de pantalla completa (no disponible en iPhone) |
| Rojo | Bloquear pantalla | Bloquea el teléfono contra toques accidentales |
| Gris | QR / Info | QR para conectar otro teléfono / versión, guía y créditos |
| Amarillo | Ocultar barra | Cierra la barra |

**Tamaño del texto**: escribe el número en la casilla, o usa **a−** (más pequeño) y **A+** (más grande).

**Franja superior.** Al cerrar la barra queda una franja con cuatro botones: **Barra** (reabre la barra),
**Pantalla** (pantalla completa), **Navegación** (muestra/oculta Siguiente, Anterior e Ir) y **Bloquear**.

**Bloqueo de pantalla.** Con el candado la pantalla ignora cualquier toque (las notas siguen actualizándose).
Para **desbloquear, mantén pulsado 5 segundos el botón rojo de arriba**: una barra se llena y la pantalla se libera. Si sueltas antes, sigue bloqueada.

**El teléfono no entra en reposo** mientras la página está abierta. El primer toque en la página activa la protección;
si el teléfono se apaga igualmente, comprueba que el *ahorro de batería* no esté activo.

Abajo: botones grandes **Anterior** y **Siguiente**, más la casilla **Ir** para saltar a un número de diapositiva.
Si hay varias presentaciones abiertas, elige la que seguir en el menú **Presentación** (no disponible con *Bloquear pantalla remota*).
Cada teléfono recuerda sus elecciones (estilo del texto, qué mostrar). El idioma es el elegido en la página de gestión.

## Opciones de línea de comandos
Windows: `mefistofele.exe --port 8080 --pin 1234 --no-control --no-browser --demo`
Mac: `./mefistofele --port 8080 --pin 1234 --no-control --no-browser --demo`

- `--port` puerto del servidor (8080 por defecto)
- `--pin` PIN para los controles
- `--no-control` desactiva por completo los controles del teléfono
- `--no-browser` no abre la página de gestión en el ordenador
- `--demo` prueba sin PowerPoint

## Solución de problemas
- **El teléfono no se conecta**: ¿misma red? Si el ordenador tiene varios adaptadores, elige el correcto en la página de gestión. Firewall de Windows: permite las *redes privadas* (y las públicas si tu Wi-Fi se considera pública). En Mac permite las conexiones entrantes si se solicita.
- **No aparecen notas**: la presentación debe estar **en marcha** (pantalla completa), no solo abierta.
- **Mac: no pasa nada**: revisa **Ajustes del Sistema → Privacidad y seguridad → Automatización** y comprueba que Terminal pueda controlar Microsoft PowerPoint.
- **Vistas previas**: en Mac dependen de que PowerPoint permita la exportación; si fallan, el motivo aparece en lugar de la imagen y las notas y controles siguen funcionando.
- **No hay vistas previas (Windows)**: abre `http://localhost:8080/diagnostica`, pulsa el botón de prueba y copia todo.
- Los errores también se escriben en `mefistofele.log` junto al programa (o en la carpeta temporal).

---
Created and developed by [www.andreabattaglioli.it](https://www.andreabattaglioli.it) ·
Powered by [MyWebTvLive.com](https://www.mywebtvlive.com) · Made in Bolognina
