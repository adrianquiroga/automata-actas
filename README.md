# Sistema de Automatización de Actas Institucionales (DataEng)

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/adrianquiroga/automata-actas/blob/main/automata_actas_colab.ipynb)

Este sistema permite automatizar de forma inteligente la generación de actas institucionales a partir de la transcripción de reuniones, utilizando Inteligencia Artificial (Gemini 2.5 Flash) y plantillas en Microsoft Word (`.docx`).

---

## ⚡ Modos de Uso Disponibles

El proyecto cuenta con tres modos de funcionamiento adaptados a las necesidades de los usuarios:

### Opción 1: Ejecución Directa mediante Enlace de GitHub (La más fácil)
*La secretaria simplemente hace clic en un enlace para abrir el cuaderno directamente desde GitHub en Google Colab.*

1. **Abrir el enlace:**
   - Haz clic en el botón azul **"Open In Colab"** que se encuentra arriba o ingresa a este enlace:
     [Abrir en Google Colab](https://colab.research.google.com/github/adrianquiroga/automata-actas/blob/main/automata_actas_colab.ipynb)
2. **Guardar una copia (Recomendado):**
   - Una vez abierto, ve a **Archivo -> Guardar una copia en Drive** para guardarlo en tu cuenta personal de Google Drive.
3. **Configurar y ejecutar:**
   - **Forma recomendada (Segura):** Haz clic en el icono de llave (🔑) en el menú lateral izquierdo de Google Colab, agrega un secreto con el nombre `GEMINI_API_KEY`, pega tu clave en el valor y activa el interruptor de acceso al cuaderno.
   - **Forma interactiva:** Si no usas secretos, al presionar **Play** (triángulo a la izquierda) se te solicitará la API key de forma oculta en una casilla segura de contraseña.
   - Selecciona el **`NIVEL_DETALLE`** a la derecha y presiona **Play**.
   - Sube los archivos cuando el programa lo solicite.

### Opción 2: Subiendo el archivo `.ipynb` a Google Drive
*Útil si quieres tener el archivo del cuaderno directamente guardado y organizado en tus carpetas de Google Drive.*

1. **Subir el cuaderno:**
   - Sube el archivo `automata_actas_colab.ipynb` de este repositorio a tu Google Drive (por ejemplo, en la carpeta *Colab Notebooks*).
   - Se mostrará con el icono naranja de Colab (`CO`).
2. **Abrir y ejecutar:**
   - Haz doble clic sobre el archivo en Drive para abrirlo en Google Colab.
   - Configura de forma segura tu **`GEMINI_API_KEY`** (en los Secretos de Colab) o introdúcela en la casilla interactiva al ejecutar, selecciona el **`NIVEL_DETALLE`** y ejecuta presionando el botón de **Play**.

*Para desarrolladores o usuarios que prefieren trabajar de forma local utilizando una interfaz de escritorio nativa.*

#### Requisitos previos:
* Python 3.8 o superior instalado en el sistema.

#### Instalación y ejecución:
1. Clona o descarga este repositorio en tu computador.
2. Instala las dependencias ejecutando en tu consola:
   ```bash
   pip install google-genai python-docx
   ```
3. Ejecuta el script de escritorio:
   ```bash
   python automata_actas.py
   ```
4. Se abrirá una ventana de escritorio donde podrás rellenar los datos de API Key, elegir la plantilla y la transcripción mediante exploradores de archivos nativos de Windows/Mac.

---

## 📋 Estructura de la Plantilla de Word (`.docx`)

El motor del programa busca palabras especiales encerradas entre signos menores y mayores, y las reemplaza con el análisis del contenido de la reunión provisto por la Inteligencia Artificial.

### Marcadores sugeridos en la plantilla:
* **Generales:** `<FechaActual>`, `<TemaReunion>`, `<HoraInicio>`, etc.
* **Orden del día (1 al 5):** `<<Agenda1>>` hasta `<<Agenda5>>`.
* **Desarrollo/Comentarios:** `<<Comentario1>>` hasta `<<Comentario5>>` (se redactan automáticamente en prosa formal institucional según el nivel de detalle configurado).
* **Compromisos/Acciones pendientes:** `<<Accion1>>`, `<<Responsable1>>`, `<<Fecha1>>`, `<<Tipo1>>` (el campo tipo clasifica la acción automáticamente como *Informativa*, *Decisoria* o *Estratégica*).
* **Compromisos varios:** `<<ResponsablesVarios>>`, `<<AccionVarios>>`, `<<TipoVarios>>`, `<<FechaVarios>>`.

---

## 🎙️ Formatos de Transcripción Admitidos
* **Texto Plano (`.txt`) [Recomendado]:** Es el formato más rápido, ligero y estable. Puedes descargar o exportar la transcripción de tu videoconferencia (Teams, Zoom, Meet, etc.) y guardarla como texto plano.
* **Documentos de Word (`.docx`)**
* **Subtítulos WebVTT (`.vtt`)**