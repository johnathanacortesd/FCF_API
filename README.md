```text
  ███████╗ ██████╗  ███████╗         █████╗  ██████╗  ██╗
  ██╔════╝██╔════╝  ██╔════╝        ██╔══██╗ ██╔══██╗ ██║
  █████╗  ██║       █████╗   ───    ███████║ ██████╔╝ ██║
  ██╔══╝  ██║       ██╔══╝   ───    ██╔══██║ ██╔═══╝  ██║
  ██║     ╚██████╗  ██║             ██║  ██║ ██║      ██║
  ╚═╝      ╚═════╝  ╚═╝             ╚═╝  ╚═╝ ╚═╝      ╚═╝
```

> **Sistema Inteligente de Monitoreo, Clasificación Reputacional e Inteligencia de Medios para la Federación Colombiana de Fútbol (FCF)**

[![Live App](https://img.shields.io/badge/Streamlit_App-Online-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)](https://fcf-api.streamlit.app/)
[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![OpenAI](https://img.shields.io/badge/OpenAI-API-412991?style=for-the-badge&logo=openai&logoColor=white)](https://openai.com)

---

## 🚀 App en Vivo

🔗 **Acceso a la plataforma web:** [https://fcf-api.streamlit.app/](https://fcf-api.streamlit.app/)

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

**FCF-API** es un sistema inteligente de análisis reputacional y procesamiento masivo de prensa desarrollado para la **Federación Colombiana de Fútbol (FCF)**.

Impulsado por modelos de Lenguaje y Embeddings de **OpenAI** (`gpt-4.1-nano-2025-04-14` y `text-embedding-3-small`), el sistema recibe dossieres de noticias en formato Excel, normaliza medios regionales, detecta registros duplicados, agrupa publicaciones por similitud semántica y clasifica automáticamente cada impacto en **Impacto (Tono)**, **TEMA**, **SUBTEMA** y **VOCERO**.

---

## ✨ Características Principales

* 🔍 **Desduplicación Inteligente**:
  * Filtrado automático por URLs (`WEB` / `LINK`).
  * Detección de títulos idénticos y coincidencia difusa (Fuzzy Matching).

* 🧠 **Clustering Semántico (Vector Embeddings)**:
  * Vectorización de resúmenes y títulos mediante `text-embedding-3-small`.
  * Agrupación eficiente usando **Disjoint Set Union (DSU)** y similitud del coseno para evitar consultas repetitivas al modelo LLM.

* 🏷️ **Clasificación Automatizada por IA**:
  * **Impacto**: Positivo, Negativo, Neutro o Duplicada.
  * **TEMA**: Institucional, Torneos - Copas - Ligas, Selecciones, Gestión, Jugadores, Entorno.
  * **SUBTEMA**: Asignación contextual específica (ej. *Partidos del Mundial*, *Partido Eliminatorias*, *Partido ante [Rival]*).
  * **VOCERO**: Detección de voceros institucionales (ej. *Ramón Jesurun*).

* 🗺️ **Lookup de Regiones**:
  * Cruce automático `NOMBRE DE MEDIO` $\rightarrow$ `REGION` consultando el diccionario de `Configuracion.xlsx`.

* 🎨 **Interfaz de Usuario (UX/UI)**:
  * Estilo corporativo FCF (Azul `#173b7a`, Dorado `#d9a441`, Rojo `#c8202f`).
  * Indicadores visuales de avance, tiempos de respuesta y métricas de procesamiento.

* 📊 **Exportación Profesional**:
  * Descarga directa en `.xlsx` formateado, incluyendo preservación de enlaces incrustados.

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
* API Key de OpenAI activa.

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
   pip install streamlit pandas numpy openpyxl openai scikit-learn unidecode
   ```

4. **Ejecutar la aplicación:**
   ```bash
   streamlit run app.py
   ```

---

## ⚙️ Configuración del Sistema

### 1. Variables Secretas (`.streamlit/secrets.toml`)

Crea la carpeta `.streamlit` en la raíz del proyecto y añade el archivo `secrets.toml`:

```toml
OPENAI_API_KEY = "sk-proj-tu-api-key-de-openai-aqui"
```

### 2. Configuración de Regiones (`Configuracion.xlsx`)

El sistema consulta localmente el archivo `Configuracion.xlsx` para asociar medios con su región. Debe contar con la siguiente estructura básica:

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
│   └── secrets.toml         # Claves privadas (OpenAI API Key)
├── app.py                   # Script principal de Streamlit y lógica de IA
├── Configuracion.xlsx       # Matriz de consulta Medio -> Región
├── README.md                # Documentación principal del repositorio
└── requirements.txt         # Dependencias del proyecto
```

---

## 🏷️ Taxonomía y Criterios de Clasificación

### **1. Categorías de Tema (`TEMA`)**
* **Institucional**: Comunicados oficiales, pronunciamientos directivos o uso de marca FCF.
* **Torneos - Copas - Ligas**: Liga BetPlay, Copa Colombia, Torneos CONMEBOL, FIFA, etc.
* **Selecciones**: Selecciones Colombia (Masculina, Femenina, Sub-20, Sub-17, Mayores).
* **Gestión**: Decisiones directivas, logística, acuerdos comerciales, infraestructura.
* **Jugadores**: Contenido enfocado en el desempeño o novedades individuales de futbolistas.
* **Entorno**: Temas del fútbol nacional e internacional sin relación directa con la FCF.

### **2. Tono / Impacto (`Impacto`)**
* 🟢 **Positivo**: Enfoque favorable, reconocimientos, gestión destacada o triunfos.
* 🔴 **Negativo**: Críticas, controversias, cuestionamientos o impacto reputacional adverso.
* ⚪ **Neutro**: Notas puramente informativas sin sesgo implícito.
* ⬛ **Duplicada**: Filas identificadas como copias de otro registro ya analizado.

### **3. Casos Específicos & Normalizaciones**
* **Fotografía FCF**: `Institucional / Foto` (Si el texto indica *"Foto tomada de FCF"*).
* **Uso de Logo**: `Institucional / Logo` (Si la noticia incluye el logo sin mención explícita en texto).
* **Partidos de Selección**: Normalización automática a `Partidos del Mundial`, `Partido Eliminatorias` o `Partido ante [Rival]`.

---

## 📊 Formatos de Entrada y Salida

### **Columnas Entrada (.xlsx)**
* `ID`
* `FECHA`
* `HORA`
* `TIPO DE MEDIO`
* `NOMBRE DE MEDIO`
* `REGION`
* `SECCIÓN`
* `TÍTULO`
* `RESUMEN`
* `LINK` / `WEB` *(Opcional)*

### **Columnas Asignadas en Salida (.xlsx)**
* `Impacto` *(Positivo, Negativo, Neutro, Duplicada)*
* `TEMA` *(Categoría principal)*
* `SUBTEMA` *(Etiqueta contextual)*
* `VOCERO` *(Portavoz detectado)*

---

<p align="center">
  <sub>Desarrollado por Johnathan Cortés para GlobalNews Group y el cliente <b>Federación Colombiana de Fútbol (FCF)</b> | Web App: <a href="https://fcf-api.streamlit.app/">fcf-api.streamlit.app</a></sub>
</p>
