# Sistema de Automatización de Actas Institucionales

Este proyecto permite automatizar la generación de actas institucionales a partir de la transcripción de reuniones procesadas en formato JSON, integrando una interfaz gráfica interactiva para la selección de archivos y un motor de análisis contextual para la clasificación automática de compromisos.

## 🚀 Características

- **Carga de Datos Flexible:** Interfaz gráfica nativa con `tkinter` para seleccionar cualquier archivo `.json` sin importar su nombre de origen (`datos.json`, `gemini_...`, etc.).
- **Análisis Contextual (Backend):** Clasificación automática e inyección dinámica de los campos `<<Tipo1>>` a `<<Tipo5>>` y `<<TipoVarios>>` (*Informativo*, *Decisorio* o *Estratégico*) analizando los verbos rectores en la prosa de los comentarios de la reunión.
- **Compatibilidad con Plantillas Word:** Mapeo directo y seguro de marcadores en texto plano y tablas dentro de la plantilla institucional `Formato_F-2-2-16_Place.docx`.
- **Estructura Limpia:** Preservación de datos puros en el archivo de intercambio de datos (JSON), separando la capa de información pura de la lógica de negocio y etiquetas visuales.

## 📦 Requisitos e Instalación

1. Asegúrate de tener instalado **Python 3.8** o superior en tu sistema.
2. Instala las dependencias requeridas ejecutando el siguiente comando en tu terminal:

```bash
pip install -r requirements.txt