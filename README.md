```markdown
  ███████╗ ██████╗  ███████╗         █████╗  ██████╗  ██╗
  ██╔════╝██╔════╝  ██╔════╝        ██╔══██╗ ██╔══██╗ ██║
  █████╗  ██║       █████╗   ───    ███████║ ██████╔╝ ██║
  ██╔══╝  ██║       ██╔══╝   ───    ██╔══██║ ██╔═══╝  ██║
  ██║     ╚██████╗  ██║             ██║  ██║ ██║      ██║
  ╚═╝      ╚═════╝  ╚═╝             ╚═╝  ╚═╝ ╚═╝      ╚═╝

> **Sistema Inteligente de Monitoreo, Clasificación Reputacional e Inteligencia de Medios para la Federación Colombiana de Fútbol (FCF)**

[![Live App](https://img.shields.io/badge/Streamlit_App-Online-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)](https://fcf-api.streamlit.app/)
[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![OpenAI](https://img.shields.io/badge/OpenAI-API-412991?style=for-the-badge&logo=openai&logoColor=white)](https://openai.com)

---

## 🚀 App en Vivo

🔗 **Acceso a la plataforma:** [https://fcf-api.streamlit.app/](https://fcf-api.streamlit.app/)

---

## 📋 Tabla de Contenidos

- [Descripción General](#-descripción-general)
- [Características Principales](#-características-principales)
- [Flujo de Procesamiento](#-flujo-de-procesamiento)
- [Requisitos e Instalación](#-requisitos-e-instalación)
- [Configuración del Sistema](#-configuración-del-sistema)
- [Estructura del Proyecto](#-estructura-del-proyecto)
- [Taxonomía y Criterios de Clasificación](#-taxonomía-y-criterios-de-clasificación)
- [Formatos de Entrada y Salida](#-formatos-de-entrada-y-salida)

---

## ⚽ Descripción General

**FCF-API** (Análisis FCF) es un sistema avanzado de inteligencia de medios y análisis reputacional diseñado exclusivamente para procesar, limpiar y clasificar información de prensa de la **Federación Colombiana de Fútbol (FCF)**.

Utilizando modelos de Lenguaje y Embeddings de **OpenAI** (`gpt-4.1-nano-2025-04-14` y `text-embedding-3-small`), la herramienta analiza grandes volúmenes de noticias contenidas en archivos Excel, deduplica registros, agrupa publicaciones por similitud semántica y clasifica automáticamente el **Impacto**, **TEMA**, **SUBTEMA** y **VOCERO**.

---

## ✨ Características Principales

* 🔍 **Desduplicación Multinivel**:
  * Filtrado por enlace exacto (`WEB` / `LINK`).
  * Coincidencia exacta de títulos y resúmenes.
  * Comparación difusa (Fuzzy Matching) basada en ratio de similitud (`SequenceMatcher`).

* 🧠 **Agrupamiento Semántico (Vector Clustering)**:
  * Generación de embeddings vectoriales a través de OpenAI `text-embedding-3-small`.
  * Algoritmo de conjuntos disjuntos **DSU (Disjoint Set Union)** con cálculo de similitud del coseno para agrupar notas con la misma temática sin repetir llamadas innecesarias al LLM.

* 🏷️ **Clasificación Automática con Inteligencia Artificial**:
  * **Impacto**: Positivo, Negativo, Neutro o Duplicada.
  * **TEMA**: Institucional, Torneos - Copas - Ligas, Selecciones, Gestión, Jugadores, Entorno.
  * **SUBTEMA**: Inferencia de etiquetas específicas (ej. *Partidos del Mundial*, *Partido Eliminatorias*, *Partido ante [Rival]*).
  * **VOCERO**: Detección inteligente de voceros (ej. *Ramón Jesurun*).

* 🗺️ **Enriquecimiento Regional Automático**:
  * Búsqueda dinámica y mapeo automatizado `NOMBRE DE MEDIO` $\rightarrow$ `REGION` consultando el diccionario de `Configuracion.xlsx`.

* 🎨 **Interfaz de Usuario (UX/UI)**:
  * Tema corporativo FCF (Azul `#173b7a`, Dorado `#d9a441`, Rojo `#c8202f`).
  * Indicadores dinámicos de progreso y métricas de procesamiento en tiempo real.

* 📊 **Exportación Profesional a Excel**:
  * Formateo de estilos, ancho automático de columnas y conservación bidireccional de hipervínculos web (`openpyxl`).

---

## 🔄 Flujo de Procesamiento

```text
┌─────────────────────────────────────────────────────────────────┐
│                    1. DOSSIER ENTRADA (.XLSX)                   │
│         Carga de datos + Extracción de links incrustados        │
└────────────────────────────────┬────────────────────────────────┘
                                 │
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                 2. MAPEADO REGIONAL DINÁMICO                    │
│        Lookup automap Medio -> Región (Configuracion.xlsx)      │
└────────────────────────────────┬────────────────────────────────┘
                                 │
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                   3. FILTRADO Y DEDUPLICACIÓN                   │
│       Detección de duplicados por URL, Título exacto y Fuzzy    │
└────────────────────────────────┬────────────────────────────────┘
                                 │
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                 4. EMBEDDINGS & DSU CLUSTERING                  │
│   Vectorización semántica (OpenAI) + Agrupamiento Disjoint-Set   │
└────────────────────────────────┬────────────────────────────────┘
                                 │
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                  5. CLASIFICACIÓN LLM (OPENAI)                  │
│       Inferencia de Impacto, Tema, Subtema y Vocero por Grupo   │
└────────────────────────────────┬────────────────────────────────┘
                                 │
                                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                   6. EXPORTACIÓN DE RESULTADOS                  │
│          Visualización en Streamlit + Descargar Excel (.XLSX)   │
└─────────────────────────────────────────────────────────────────┘
```

---

## 🛠️ Requisitos e Instalación

### Pre-requisitos
* **Python 3.10** o superior.
* API Key de OpenAI habilitada.

### Pasos para Ejecución Local

1. **Clonar el repositorio:**
   ```bash
   git clone https://github.com/tu-usuario/FCF-API.git
   cd FCF-API
   ```

2. **Crear y activar el entorno virtual:**
   ```bash
   # En macOS/Linux:
   python3 -m venv venv
   source venv/bin/activate

   # En Windows:
   python -m venv venv
   venv\Scripts\activate
   ```

3. **Instalar las dependencias:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Ejecutar la aplicación:**
   ```bash
   streamlit run app.py
   ```

---

## ⚙️ Configuración del Sistema

### 1. Variables Secretas (`.streamlit/secrets.toml`)

Crea la carpeta `.streamlit` en la raíz del proyecto y dentro añade el archivo `secrets.toml`:

```toml
OPENAI_API_KEY = "sk-proj-tu-api-key-de-openai-aqui"
```

### 2. Archivo de Configuración de Regiones (`Configuracion.xlsx`)

El sistema busca automáticamente un archivo `Configuracion.xlsx` en la raíz para asociar medios con su región geográfica. Debe tener una estructura como la siguiente:

| NOMBRE DE MEDIO | REGION |
| :--- | :--- |
| El Tiempo | Nacional |
| El Espectador | Nacional |
| El Heraldo | Barranquilla |
| Telemedellín | Antioquia |

---

## 📁 Estructura del Proyecto

```text
FCF-API/
├── .streamlit/
│   └── secrets.toml         # Credenciales y claves de API (Streamlit Secrets)
├── app.py                   # Código principal de la aplicación Streamlit e integración IA
├── Configuracion.xlsx       # Matriz de consulta Medio -> Región
├── requirements.txt         # Lista de dependencias del proyecto
└── README.md                # Documentación general del repositorio
```

---

## 🏷️ Taxonomía y Criterios de Clasificación

### **1. Categorías de Tema (`TEMA`)**
* **Institucional**: Declaraciones oficiales, comunicados, temas corporativos o de marca FCF.
* **Torneos - Copas - Ligas**: Liga BetPlay, Copa Colombia, Torneos CONMEBOL, FIFA, etc.
* **Selecciones**: Selección Colombia Masculina, Femenina, Sub-20, Sub-17, Mayores.
* **Gestión**: Decisiones directivas, logística, infraestructura, contratos y convenios.
* **Jugadores**: Noticias individuales centrándose en desempeño o novedades de futbolistas.
* **Entorno**: Noticias contextuales del fútbol nacional e internacional sin acción directa FCF.

### **2. Tono / Impacto (`Impacto`)**
* 🟢 **Positivo**: Resalta logros, victorias, buenas prácticas o reconocimientos.
* 🔴 **Negativo**: Contiene críticas, señalamientos, controversias o afectación reputacional.
* ⚪ **Neutro**: Reportes informativos genéricos sin juicio de valor implícito.
* ⬛ **Duplicada**: Filas identificadas como repetidas de otra noticia ya procesada.

### **3. Casos Específicos & Normalizaciones**
* **Uso de Foto FCF**: `Institucional / Foto` (Si el texto indica *"Foto tomada de FCF"*).
* **Mención de Logo**: `Institucional / Logo` (Si la noticia no menciona a la FCF explícitamente pero la incluye).
* **Partido de Selección**: Si se detecta un encuentro, el subtema se normaliza dinámicamente:
  * *Partidos del Mundial*
  * *Partido Eliminatorias*
  * *Partido ante [Selección Rival]*

---

## 📊 Formatos de Entrada y Salida

### **Columnas Requeridas en el Dossier de Entrada (.xlsx)**
* `ID`
* `FECHA`
* `HORA`
* `TIPO DE MEDIO`
* `NOMBRE DE MEDIO`
* `REGION`
* `SECCIÓN`
* `TÍTULO`
* `RESUMEN`
* `LINK` / `WEB` *(Opcional pero recomendado)*

### **Columnas Agregadas en la Salida (.xlsx)**
* `Impacto` *(Positivo, Negativo, Neutro, Duplicada)*
* `TEMA` *(Categoría asignada)*
* `SUBTEMA` *(Etiqueta detallada)*
* `VOCERO` *(Detección del portavoz)*

---

<p align="center">
  <sub>Desarrollado para la <b>Federación Colombiana de Fútbol (FCF)</b> | App: <a href="https://fcf-api.streamlit.app/">fcf-api.streamlit.app</a></sub>
</p>
```
