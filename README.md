```markdown
 ███████╗██████╗ ███████╗     █████╗ ██████╗ ██╗
 ██╔════╝██╔════╝██╔════╝    ██╔══██╗██╔══██╗██║
 █████╗  ██║     █████╗  ─── ███████║██████╔╝██║
 ██╔══╝  ██║     ██╔══╝  ─── ██╔══██║██╔═══╝ ██║
 ██║     ╚█████╗ ██║         ██║  ██║██║     ██║
 ╚═╝      ╚════╝ ╚═╝         ╚═╝  ╚═╝╚═╝     ╚═╝

> **Sistema Inteligente de Monitoreo, Clasificación Reputacional e Inteligencia de Medios para la Federación Colombiana de Fútbol (FCF)**

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg?logo=python&logoColor=white)](https://python.org)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.30%2B-FF4B4B.svg?logo=streamlit&logoColor=white)](https://streamlit.io)
[![OpenAI](https://img.shields.io/badge/OpenAI-API-412991.svg?logo=openai&logoColor=white)](https://openai.com)
[![Status](https://img.shields.io/badge/Status-Production-brightgreen.svg)]()

---

## 📋 Tabla de Contenidos

1. [Descripción General](#-descripción-general)
2. [Características Principales](#-características-principales)
3. [Flujo de Procesamiento](#-flujo-de-procesamiento)
4. [Requisitos e Instalación](#-requisitos-e-instalación)
5. [Configuración de Entorno](#-configuración-de-entorno)
6. [Estructura del Proyecto](#-estructura-del-proyecto)
7. [Taxonomía y Criterios de Clasificación](#-taxonomía-y-criterios-de-clasificación)
8. [Exportación e Informes](#-exportación-e-informes)

---

## ⚽ Descripción General

**FCF-API / Análisis FCF** es una aplicación Web interactiva desarrollada en Streamlit que automatiza el análisis de prensa y monitoreo de reputación mediática para la **Federación Colombiana de Fútbol**. 

A través de modelos de IA de **OpenAI** (`gpt-4.1-nano-2025-04-14` y `text-embedding-3-small`), el sistema procesa dossieres masivos en formato Excel, elimina duplicados, agrupa noticias semánticamente similares y clasifica automáticamente cada impacto en categorías, temas, subtemas específicos y voceros.

---

## ✨ Características Principales

* 🔍 **Detección Automática de Duplicados**: Filtrado inteligente por URLs (`WEB`), coincidencia exacta de texto y algoritmos de similitud difusa (Fuzzy Matching).
* 🧠 **Agrupamiento Semántico (Vector Embeddings)**: Generación de embeddings vectoriales combinados con **Disjoint Set Union (DSU)** para clustering de noticias sobre el mismo hecho noticioso.
* 🏷️ **Clasificación Automática con LLM**: Categorización precisa del **Impacto** *(Positivo, Negativo, Neutro)*, **TEMA** e inferencia contextual de **SUBTEMA**.
* 🗺️ **Mapeo Automático de Regiones**: Enriquecimiento del campo `REGION` mediante lookup dinámico contra `Configuracion.xlsx`.
* 🗣️ **Detección de Voceros**: Identificación inteligente de menciones a voceros clave (ej. *Ramón Jesurun*).
* 🎨 **Interfaz de Alta Experiencia (UX/UI)**: Diseño personalizado con la identidad visual corporativa de la FCF (Azul, Dorado, Rojo) y métricas dinámicas de procesamiento.
* 📊 **Exportación Avanzada a Excel**: Generación de reportes formateados con celdas estilizadas y conservación intacta de hipervínculos (`LINK` / `WEB`).

---

## 🔄 Flujo de Procesamiento

```
┌──────────────────────────────────────────────────────────────────────────────────┐
│                                 ENTRADA DE DATOS                                 │
│                     Carga de Dossier XLSX + Extracción de Links                  │
└────────────────────────────────────────┬─────────────────────────────────────────┘
                                         │
                                         ▼
┌──────────────────────────────────────────────────────────────────────────────────┐
│                           MAPEADO DINÁMICO DE REGIONES                           │
│                 Cruce con diccionario local (Configuracion.xlsx)                 │
└────────────────────────────────────────┬─────────────────────────────────────────┘
                                         │
                                         ▼
┌──────────────────────────────────────────────────────────────────────────────────┐
│                              DETECCIÓN DE DUPLICADOS                             │
│                  Coincidencia por URLs + Similitud de Títulos                   │
└────────────────────────────────────────┬─────────────────────────────────────────┘
                                         │
                                         ▼
┌──────────────────────────────────────────────────────────────────────────────────┐
│                            AGRUPAMIENTO SEMÁNTICO                                │
│       OpenAI Text Embeddings (3-small) + DSU (Disjoint Set Union) + Cosine Sim   │
└────────────────────────────────────────┬─────────────────────────────────────────┘
                                         │
                                         ▼
┌──────────────────────────────────────────────────────────────────────────────────┐
│                        CLASIFICACIÓN INTELIGENTE (LLM)                           │
│           Evaluación de Tono / Tema / Subtema / Vocero (GPT-4.1 Nano)            │
└────────────────────────────────────────┬─────────────────────────────────────────┘
                                         │
                                         ▼
┌──────────────────────────────────────────────────────────────────────────────────┐
│                                SALIDA DE REPORTES                                │
│                   Visualización en Dashboard + Descarga Excel                    │
└──────────────────────────────────────────────────────────────────────────────────┘
```

---

## 🛠️ Requisitos e Instalación

### Requisitos Previos

* Python **3.10** o superior.
* Clave de API de OpenAI (`OPENAI_API_KEY`).

### Pasos de Instalación

1. **Clonar el repositorio:**
   ```bash
   git clone https://github.com/tu-usuario/FCF-API.git
   cd FCF-API
   ```

2. **Crear y activar un entorno virtual:**
   ```bash
   python -m venv venv
   # En Windows:
   venv\Scripts\activate
   # En macOS/Linux:
   source venv/bin/activate
   ```

3. **Instalar dependencias:**
   ```bash
   pip install streamlit pandas numpy openpyxl openai scikit-learn unidecode
   ```

---

## ⚙️ Configuración de Entorno

### 1. Clave de API de OpenAI (`secrets.toml`)

Crea la carpeta `.streamlit` en la raíz del proyecto y añade el archivo `secrets.toml`:

```
.
├── .streamlit/
│   └── secrets.toml
├── app.py
└── Configuracion.xlsx
```

Contenido de `.streamlit/secrets.toml`:
```toml
OPENAI_API_KEY = "sk-proj-tu-api-key-aqui..."
```

### 2. Archivo de Configuración Regional (`Configuracion.xlsx`)

El archivo opcional `Configuracion.xlsx` permite mapear automáticamente medios a sus regiones correspondientes. Debe contener al menos una hoja con las columnas:

| NOMBRE DE MEDIO | REGION |
| :--- | :--- |
| El Tiempo | Nacional |
| El Espectador | Nacional |
| El Heraldo | Barranquilla |

---

## 📁 Estructura del Proyecto

```
FCF-API/
├── .streamlit/
│   └── secrets.toml         # Configuración de credenciales locales
├── app.py                   # Aplicación principal de Streamlit y lógica de IA
├── Configuracion.xlsx       # Matriz de mapeo Medio -> Región
├── README.md                # Documentación del proyecto
└── requirements.txt         # Dependencias del proyecto
```

---

## 🏷️ Taxonomía y Criterios de Clasificación

### **Temas Preestablecidos (`TEMA`)**
* 🏛️ **Institucional**
* 🏆 **Torneos - Copas - Ligas**
* 🇨🇴 **Selecciones**
* 💼 **Gestión**
* ⚽ **Jugadores**
* 🌐 **Entorno**

### **Tono e Impacto (`Impacto`)**
* 🟢 **Positivo**: Enfoque favorable, logros, gestión destacada o reconocimientos.
* 🔴 **Negativo**: Críticas, cuestionamientos, controversias o afectación reputacional.
* ⚪ **Neutro**: Notas informativas puras sin sesgo evaluativo implícito.
* ⬛ **Duplicada**: Registros repetidos filtrados por el sistema.

### **Casos Especiales Normalizados**
* **Menciones indirectas / Usos de Marca**: Fotografía de FCF o pie de foto $\rightarrow$ `Institucional / Foto`.
* **Presencia de Logo sin texto FCF**: $\rightarrow$ `Institucional / Logo`.
* **Partidos de Selección**: Normalización automática a `Partidos del Mundial`, `Partido Eliminatorias` o `Partido ante [Rival]`.

---

## 🚀 Ejecución de la Aplicación

Para iniciar el servidor local de Streamlit, ejecuta:

```bash
streamlit run app.py
```

Accede desde tu navegador en: `http://localhost:8501`

---

## 📊 Exportación e Informes

El reporte final exportable (`Analisis_FCF.xlsx`) incluye:
* Formato visual estilizado corporativo (Cabecera azul FCF con texto en blanco).
* Ancho de columnas auto-ajustado dinámicamente.
* Enlaces funcionales interactivos (conservando hipervínculos originales de las celdas `LINK` y `WEB`).

---

<p align="center">
  <sub>Desarrollado para la <b>Federación Colombiana de Fútbol (FCF)</b></sub>
</p>
```
