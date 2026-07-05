# Sistema de Automatización de Actas Institucionales (DataEng)

Este sistema permite automatizar de forma inteligente la generación de actas institucionales a partir de la transcripción de reuniones, utilizando Inteligencia Artificial (Gemini 2.5 Flash) y plantillas en Microsoft Word (`.docx`).

---

## ⚡ Modos de Uso Disponibles

El proyecto cuenta con dos modos de funcionamiento adaptados a las necesidades de los usuarios:

### Opción 1: Ejecución en la Nube con Google Colab (Recomendada)
*Ideal para secretarios/as y usuarios sin conocimientos técnicos, ya que no requiere instalar nada en la computadora.*

#### Pasos para utilizarlo:
1. **Obtener la Clave de API de Gemini (Gratuita):**
   - Regístrate en [Google AI Studio](https://aistudio.google.com/) con tu cuenta de Gmail.
   - Haz clic en **"Get API Key"** y luego en **"Create API Key"** para obtener tu clave.
2. **Subir el cuaderno a Google Drive:**
   - Sube el archivo `automata_actas_colab.ipynb` de este repositorio a tu Google Drive (por ejemplo, en la carpeta *Colab Notebooks*).
   - Se mostrará con el icono naranja de Colab (`CO`).
3. **Ejecutar el programa:**
   - Haz doble clic sobre el archivo en Drive para abrirlo en Google Colab.
   - Pega tu clave de API en el campo **`API_KEY`** a la derecha de la celda de ejecución.
   - Selecciona el **`NIVEL_DETALLE`** ("Resumido", "Medio" o "Detallado").
   - Presiona el botón de **Play** (triángulo a la izquierda).
4. **Cargar archivos y descargar resultado:**
   - El script instalará automáticamente las dependencias en la nube.
   - Haz clic en **"Elegir archivos"** para subir tu **Plantilla Word** (`.docx`) y tu **Transcripción** (`.txt`, `.docx` o `.vtt`).
   - Al finalizar, el acta ya completada se descargará de forma automática en tu carpeta local de **Descargas**.

---

### Opción 2: Ejecución Local con Interfaz Gráfica (GUI)
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