# Convenios y finiquitos

App de Streamlit para Recursos Humanos. Hace lo mismo que la macro de Excel de convenios: se capturan los
datos del colaborador y se descargan en PDF el convenio de finiquito (sin días, con 15, 30 o 45 días, o las
cuatro opciones en un solo archivo), la carta de renuncia y la constancia laboral.

## Archivos

| Archivo | Para qué sirve |
| --- | --- |
| `app.py` | La pantalla. Solo acomoda los campos y botones. |
| `ui.py` | Componentes visuales, iconos vectoriales y barra de estado Noir Liquid Aurora. |
| `documentos.py` | Cálculos, textos legales, validación, PDF y Excel. No depende de Streamlit. |
| `assets/` | Hoja de estilos (`styles.css`), membrete (`fondo.jpg`), logotipo (`logo.png`) y sello (`sello.png`). |
| `requirements.txt` | Librerías, fijadas en la versión con la que se probó. |
| `.streamlit/config.toml` | Colores de la app. |
| `.streamlit/secrets.toml.example` | Ejemplo de secretos: clave de acceso y datos de quien firma. |

## Probarla en tu equipo (VS Code)

Necesita Python 3.11 o más nuevo. Abre la carpeta en VS Code y, en la terminal:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

En macOS o Linux cambia la segunda línea por `source .venv/bin/activate`. Si PowerShell no deja activar el
entorno, usa `.venv\Scripts\activate.bat` desde `cmd`. La app abre en `http://localhost:8501`. Sin archivo
de secretos no pide clave.

## Subirla a Streamlit Community Cloud

1. Sube la carpeta a un repositorio de GitHub **privado**. El proyecto lleva el membrete y el sello de la
   empresa. `.streamlit/secrets.toml` no se sube (está en `.gitignore`).
2. En [share.streamlit.io](https://share.streamlit.io) elige **Create app** y luego el repositorio, la rama
   y `app.py` como archivo principal. Para publicar desde un repositorio necesitas ser administrador de él.
3. Antes de publicar abre **Advanced settings**: elige Python 3.12 (también sirven 3.11 y 3.13) y pega
   ahí el contenido de tu `secrets.toml`. La versión de Python ya no se puede cambiar después sin borrar y
   volver a crear la app.
4. **Deploy**. Cada `git push` a esa rama actualiza la app.
5. Con el botón **Share** de la app invitas por correo a quienes la van a usar.

Guías oficiales: [publicar](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy),
[compartir](https://docs.streamlit.io/deploy/streamlit-community-cloud/share-your-app),
[límites del servicio](https://docs.streamlit.io/deploy/streamlit-community-cloud/status).

## Quién puede entrar

Hay dos candados y se pueden usar juntos:

- **App privada.** Una app publicada desde un repositorio privado nace privada: solo entran las personas
  que invites por correo. Community Cloud permite una sola app privada por cuenta (workspace).
- **Clave de acceso.** Si en los secretos existe `clave`, la app la pide antes de mostrar cualquier cosa.
  Es una clave compartida y sencilla: sirve de segundo candado, no sustituye a la app privada.

Si la app queda pública y sin clave, cualquiera con el enlace puede generar documentos con el membrete y el
sello.

## Secretos

Copia `.streamlit/secrets.toml.example` como `.streamlit/secrets.toml` (en tu equipo) o pega su contenido en
los secretos de la app (en la nube).

```toml
clave = "la-clave-que-usara-rh"

[firmante]                 # opcional: datos de quien firma la constancia, ya llenos al abrir
nombre = "LIC. NOMBRE Y APELLIDOS"
area = "RECURSOS HUMANOS"
telefono = "000 000 0000"
correo = "correo@empresa.mx"
direccion = "Calle y número\nColonia y código postal\nCiudad, estado"
firma = "assets/firma.png"   # imagen de la firma dentro del proyecto
```

Sin `[firmante]`, quien usa la app escribe esos datos y sube la imagen de la firma en la sección «Quién
firma la constancia»; duran lo que dure la sesión. Guardar la firma en `assets/` implica subirla al
repositorio: hazlo solo si es privado.

## Cómo se usa

1. Llenar las secciones 1 a 3 (y los descuentos, si hay). Las fechas se escriben como `20/02/2026`; también
   se aceptan `20-2-26` y `20022026`. Los montos aceptan `7274.01` o `$7,274.01`.
2. En la barra lateral elegir los días de convenio. Ahí se ve el total de cada opción.
3. Pulsar **Ver y descargar…** del documento. Se abre una ventana con el PDF a la vista y el botón
   **Descargar PDF**.
4. Si falta o está mal escrito algún dato, la barra lateral dice cuál y el documento no se puede abrir
   hasta corregirlo.
5. «Historial de esta sesión» guarda cada convenio abierto, permite volver a cargarlo en el formulario y
   descargarlo en Excel con las columnas de la hoja Historico (cuatro renglones por persona).

## Datos

- La app no guarda lo capturado en ningún archivo ni base de datos. Vive en la memoria de la sesión y se
  pierde al cerrar o recargar la página; por eso el historial hay que descargarlo en Excel antes de salir.
- Lo capturado (nombre, RFC, número de seguro social, sueldo) sí viaja al servidor donde corre la app, que
  es donde se arman los PDF. Community Cloud hospeda todas las apps en Estados Unidos. Si eso no es
  aceptable para datos de personal, la misma app corre igual en un servidor interno con
  `streamlit run app.py`.
- La hora del convenio, cuando se deja vacía, es la del momento según el estado elegido (Sinaloa y Sonora,
  por ejemplo, van una hora atrás del centro); no depende del reloj del servidor. La pantalla muestra qué
  hora va a salir. En municipios con horario distinto al de su estado, como los de la frontera norte,
  escribe la hora en el campo.

## Dónde cambiar cosas

Todo está en `documentos.py`:

- **Textos de los documentos**: bloque «Textos de los documentos». La redacción es la del Excel, palabra por
  palabra, incluidas dos erratas que conviene que Jurídico decida si se corrigen: «TRANFERENCIA BANCARIA»
  (en `tercera`) y «OTRAS PRESTACIONES EXCENTAS» (en `PERCEPCIONES`).
- **Razones sociales**: lista `EMPRESAS`. En pantalla también se puede escribir una que no esté.
- **Percepciones y descuentos**: listas `PERCEPCIONES` y `DESCUENTOS`. Un renglón que agregues ahí sale
  en la pantalla y en el convenio sin tocar nada más; para que también salga en el Excel hay que agregarlo
  en `filas_historial`.
- **Reglas de captura**: `validar` (qué pide cada documento) y `mal_escritos` (avisos junto a cada campo).
- **Acomodo de cada hoja**: funciones `_hoja_convenio`, `renuncia` y `constancia`. Las medidas van en puntos
  sobre hoja carta (612 × 792). El convenio reduce la letra solo cuando hace falta para caber en una hoja.
- **Membrete, logotipo y sello**: sustituye las imágenes de `assets/` conservando el nombre.
- **Colores**: `.streamlit/config.toml`.

## Igual que en el Excel, y lo que cambia

- Con días de convenio se suma sueldo ÷ 30 × días a «Otras prestaciones exentas». Sin días el pago dice
  transferencia bancaria; con 15, 30 o 45 días, efectivo.
- El importe con letra sale igual que en la macro («UN MIL», sin acentos).
- Cada descuento se imprime con el nombre del renglón donde se capturó, igual en las cuatro opciones.
- «Descanso y festivo laborado» se imprime con su nombre y sí se suma al total.
- El RFC y el número de seguro social solo se piden para la constancia.
- La firma de la constancia no viene dentro de la app: se sube en cada sesión o se configura en los secretos.

## Versiones probadas

Streamlit 1.67.0, reportlab 5.0.1, Pillow 12.3.0 y openpyxl 3.1.5, con Python 3.11, 3.12 y 3.13.
`requirements.txt` fija esas versiones; para actualizar una, cambia el número y vuelve a probar los tres
documentos.
