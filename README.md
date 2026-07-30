```
 ███████╗ ██████╗███████╗       █████╗ ██████╗ ██╗
 ██╔════╝██╔════╝██╔════╝      ██╔══██╗██╔══██╗██║
 █████╗  ██║     █████╗ █████╗ ███████║██████╔╝██║
 ██╔══╝  ██║     ██╔══╝ ╚════╝ ██╔══██║██╔═══╝ ██║
 ██║     ╚██████╗██║           ██║  ██║██║     ██║
 ╚═╝      ╚═════╝╚═╝           ╚═╝  ╚═╝╚═╝     ╚═╝
```

<div align="center">

# Análisis FCF

**Clasificación inteligente de prensa para la Federación Colombiana de Fútbol**

[![Streamlit](https://img.shields.io/badge/Streamlit-App-FF4B4B?logo=streamlit&logoColor=white)](https://fcf-api.streamlit.app/)
[![Python](https://img.shields.io/badge/Python-3.10+-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![OpenAI](https://img.shields.io/badge/OpenAI-GPT--4.1--nano-412991?logo=openai&logoColor=white)](https://openai.com/)
[![License](https://img.shields.io/badge/License-Privado-lightgrey)]()

**[🔗 Ver aplicación en vivo](https://fcf-api.streamlit.app/)**

</div>

---

## 📋 Descripción

**Análisis FCF** es una aplicación de inteligencia de medios construida en Streamlit que automatiza la clasificación editorial de dossiers de prensa relacionados con la **Federación Colombiana de Fútbol**. La herramienta procesa archivos Excel con cientos de noticias y, mediante embeddings semánticos y modelos de lenguaje de OpenAI, asigna de forma consistente:

- **Impacto** (tono: Positivo / Negativo / Neutro)
- **Tema** (Institucional, Torneos - Copas - Ligas, Selecciones, Gestión, Jugadores, Entorno)
- **Subtema** (etiqueta específica del hecho noticioso)
- **Vocero** (detección de menciones a voceros oficiales)

Todo esto respetando reglas de negocio propias del fútbol colombiano: identificación de partidos de eliminatorias, partidos del Mundial, rivales específicos por selección, fotos institucionales y noticias duplicadas.

---

## ✨ Características principales

- 🧠 **Clasificación por lotes con IA** — Agrupa noticias similares (por título, cuerpo y similitud semántica de embeddings) y las clasifica una sola vez por grupo, garantizando consistencia editorial.
- 🔗 **Detección de duplicados** — Combina coincidencia exacta de título, similitud difusa (`SequenceMatcher`) y comparación de URLs para evitar doble conteo.
- ⚽ **Reglas de dominio futbolístico** — Detecta automáticamente partidos ante selecciones rivales, partidos de eliminatorias y partidos del Mundial 2026 mediante expresiones regulares contextuales.
- 📊 **Integración con configuración externa (VLOOKUP)** — Cruza `NOMBRE DE MEDIO → REGIÓN` contra un archivo `Configuracion.xlsx` del repositorio.
- 🔗 **Preservación de hipervínculos** — Lee y reescribe los hipervínculos de columnas `LINK`/`WEB` directamente desde el XLSX original usando `openpyxl`.
- 📥 **Exportación a Excel enriquecido** — Genera un archivo `.xlsx` con encabezados estilizados, columnas autoajustadas y links funcionales.
- 🎨 **Interfaz moderna** — UI personalizada con CSS propio, métricas en tiempo real y barra de progreso por etapas.

---

## 🛠️ Stack tecnológico

| Componente          | Tecnología                          |
|---------------------|--------------------------------------|
| Framework web        | Streamlit                           |
| Procesamiento de datos | pandas, NumPy                     |
| Embeddings           | OpenAI `text-embedding-3-small`     |
| Clasificación        | OpenAI `gpt-4.1-nano`               |
| Similaridad semántica | scikit-learn (`cosine_similarity`) |
| Manejo de Excel       | openpyxl                            |
| Normalización de texto | unidecode, re (regex)              |

---

## 🚀 Uso

1. Ingresa a la aplicación: **[fcf-api.streamlit.app](https://fcf-api.streamlit.app/)**
2. Carga el dossier de prensa en formato `.xlsx` con las columnas requeridas:
   `ID`, `FECHA`, `HORA`, `TIPO DE MEDIO`, `NOMBRE DE MEDIO`, `REGION`, `SECCIÓN`, `TÍTULO`, `RESUMEN`
3. La app aplicará automáticamente el cruce de región desde `Configuracion.xlsx`.
4. Presiona **"Analizar FCF"** para iniciar la clasificación por IA.
5. Descarga el archivo `Analisis_FCF.xlsx` con las columnas `Impacto`, `TEMA`, `SUBTEMA` y `VOCERO` completadas.

---

## ⚙️ Configuración local (desarrollo)

```bash
git clone <url-del-repositorio>
cd analisis-fcf
pip install -r requirements.txt
```

Crea un archivo `.streamlit/secrets.toml`:

```toml
OPENAI_API_KEY = "tu-api-key-aqui"
```

Ejecuta la aplicación:

```bash
streamlit run app.py
```

> Asegúrate de incluir `Configuracion.xlsx` en la raíz del proyecto con las hojas de mapeo de medios y regiones.

---

## 📁 Estructura esperada del archivo de entrada

| Columna         | Descripción                              |
|------------------|-------------------------------------------|
| TÍTULO           | Titular de la noticia                     |
| RESUMEN          | Cuerpo o resumen del artículo             |
| NOMBRE DE MEDIO  | Medio de comunicación                     |
| REGION           | Región (se completa automáticamente)      |
| LINK / WEB       | Hipervínculos a la nota original          |

---

## 👤 Autoría

Desarrollado por **Johnathan Cortés** para **GlobalNews Group**, como parte de sus soluciones de inteligencia de medios y monitoreo de noticias para la Federación Colombiana de Fútbol.

---

<div align="center">

*Análisis FCF · GlobalNews Group*

</div>
