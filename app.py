# ======================================
# Importaciones
# ======================================
import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment
from collections import defaultdict, Counter
from difflib import SequenceMatcher
from copy import deepcopy
import datetime
import io
import openai
import re
import time
from unidecode import unidecode
import numpy as np
from sklearn.metrics.pairwise import cosine_similarity
from sklearn.cluster import AgglomerativeClustering
import json
import asyncio
import hashlib
from typing import List, Dict, Tuple, Optional, Any
from pathlib import Path

# ======================================
# Configuración general
# ======================================
st.set_page_config(
    page_title="Análisis de Noticias FCF · Realizado por Johnathan Cortés",
    page_icon="⚽",
    layout="wide",
    initial_sidebar_state="collapsed"
)

OPENAI_MODEL_EMBEDDING     = "text-embedding-3-small"
OPENAI_MODEL_CLASIFICACION = "gpt-4.1-nano-2025-04-14"

# Concurrencia controlada para evitar Rate Limits (429)
CONCURRENT_REQUESTS          = 3
SIMILARITY_THRESHOLD_GROUP   = 0.82
SIMILARITY_THRESHOLD_TITULOS = 0.93

PRICE_INPUT_1M     = 0.05
PRICE_OUTPUT_1M    = 0.40
PRICE_EMBEDDING_1M = 0.02

if 'tokens_input' not in st.session_state: st.session_state['tokens_input']     = 0
if 'tokens_output' not in st.session_state: st.session_state['tokens_output']    = 0
if 'tokens_embedding' not in st.session_state: st.session_state['tokens_embedding'] = 0

STOPWORDS_ES = set("""
a ante bajo cabe con contra de desde durante en entre hacia hasta mediante
para por segun sin so sobre tras y o u e la el los las un una unos unas lo
al del se su sus le les mi mis tu tus nuestro nuestros vuestra vuestras este
esta estos estas ese esa esos esas aquel aquella aquellos aquellas que cual
cuales quien quienes cuyo cuya cuyos cuyas como cuando donde cual es son fue
fueron era eran sera seran seria serian he ha han habia han hay hubo habra
habria estoy esta estan estaba estaban estamos estan estar estare estaria
estuvieron estarian estuvo asi ya mas menos tan tanto cada muy todo toda todos
todas ser haber hacer tener poder deber ir dar ver saber querer llegar pasar
encontrar creer decir poner salir volver seguir llevar sentir cambiar contra
""".split())

# ======================================
# Adaptadores de Compatibilidad OpenAI (v0.x y v1.x)
# ======================================
_client = None
_async_client = None

def get_openai_clients():
    """Garantiza la instanciación única de clientes para optimizar conexiones asíncronas."""
    global _client, _async_client
    if hasattr(openai, "OpenAI"):
        if _client is None:
            _client = openai.OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
        if _async_client is None:
            _async_client = openai.AsyncOpenAI(api_key=st.secrets["OPENAI_API_KEY"])
    return _client, _async_client

def sync_embedding_create(input_texts: List[str], model: str) -> Tuple[List[List[float]], int]:
    """Generación sincrónica de embeddings adaptable a la versión de la SDK."""
    client, _ = get_openai_clients()
    if client is not None:
        resp = client.embeddings.create(input=input_texts, model=model)
        embeddings = [d.embedding for d in resp.data]
        total_tokens = resp.usage.total_tokens
        return embeddings, total_tokens
    else:
        resp = openai.Embedding.create(input=input_texts, model=model)
        embeddings = [d["embedding"] for d in resp["data"]]
        total_tokens = resp["usage"]["total_tokens"]
        return embeddings, total_tokens

async def async_chat_completion(model: str, messages: list, max_tokens: int, temperature: float, response_format: dict = None) -> Tuple[str, int, int]:
    """Generación asíncrona de completions adaptable a la versión de la SDK."""
    _, async_client = get_openai_clients()
    if async_client is not None:
        kw = {
            "model": model,
            "messages": messages,
            "max_tokens": max_tokens,
            "temperature": temperature,
        }
        if response_format:
            kw["response_format"] = response_format
        resp = await async_client.chat.completions.create(**kw)
        content = resp.choices[0].message.content
        prompt_tokens = resp.usage.prompt_tokens
        completion_tokens = resp.usage.completion_tokens
        return content, prompt_tokens, completion_tokens
    else:
        kw = {
            "model": model,
            "messages": messages,
            "max_tokens": max_tokens,
            "temperature": temperature,
        }
        if response_format:
            kw["response_format"] = response_format
        resp = await openai.ChatCompletion.acreate(**kw)
        content = resp.choices[0].message.content
        prompt_tokens = resp["usage"]["prompt_tokens"]
        completion_tokens = resp["usage"]["completion_tokens"]
        return content, prompt_tokens, completion_tokens

# ======================================
# Caché Global de Embeddings
# ======================================
class EmbeddingCache:
    def __init__(self):
        self._cache: Dict[str, List[float]] = {}
        self._hits = 0
        self._misses = 0

    def _key(self, text):
        return hashlib.md5(text[:2000].encode('utf-8', errors='ignore')).hexdigest()

    def get(self, text):
        k = self._key(text)
        if k in self._cache:
            self._hits += 1
            return self._cache[k]
        self._misses += 1
        return None

    def put(self, text, emb):
        self._cache[self._key(text)] = emb

    def get_many(self, textos):
        results = [None] * len(textos)
        missing = []
        for i, t in enumerate(textos):
            c = self.get(t)
            if c is not None:
                results[i] = c
            else:
                missing.append(i)
        return results, missing

    def stats(self):
        total = self._hits + self._misses
        rate = (self._hits / total * 100) if total > 0 else 0
        return f"Cache: {self._hits} hits, {self._misses} misses ({rate:.0f}%)"

    def clear(self):
        self._cache.clear()
        self._hits = 0
        self._misses = 0

if '_emb_cache' not in st.session_state:
    st.session_state['_emb_cache'] = EmbeddingCache()

def get_embedding_cache():
    return st.session_state['_emb_cache']

# ======================================
# CSS Estilizado
# ======================================
def load_custom_css():
    st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500;700&family=Google+Sans+Text:wght@400;500;700&family=Roboto+Mono:wght@400;500&display=swap');
:root {
    --bg:#f8f9fa;--s1:#ffffff;--s2:#f1f3f4;--border:#dadce0;
    --text:#202124;--text2:#3c4043;--text3:#5f6368;
    --accent:#059669;--accent-bg:#ecfdf5;--accent-bdr:#a7f3d0;
    --r2:12px;--shadow-sm:0 1px 2px rgba(60,64,67,0.1),0 1px 3px rgba(60,64,67,0.08);
}
html,body,[data-testid="stApp"]{
    background:var(--bg)!important;color:var(--text)!important;
    font-family:'Google Sans Text',sans-serif;
}
.app-header{background:var(--s1);border:1px solid var(--border);border-radius:var(--r2);padding:1.2rem;margin-bottom:1rem;display:flex;align-items:center;gap:1rem;box-shadow:var(--shadow-sm);}
.app-header-icon{width:40px;height:40px;background:linear-gradient(135deg,#059669,#047857);border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:1.2rem;color:white;flex-shrink:0;}
.app-header-title{font-family:'Google Sans',sans-serif;font-size:1.25rem;font-weight:700;color:var(--text);}
.app-header-badge{background:var(--accent-bg);border:1px solid var(--accent-bdr);color:#047857;font-family:'Roboto Mono',monospace;font-size:0.65rem;font-weight:500;padding:0.25rem 0.75rem;border-radius:100px;text-transform:uppercase;}
.metrics-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:0.6rem;margin:0.8rem 0}
.metric-card{background:var(--s1);border:1px solid var(--border);border-radius:var(--r2);padding:0.8rem;text-align:center;box-shadow:var(--shadow-sm);}
.metric-val{font-family:'Google Sans',sans-serif;font-size:1.5rem;font-weight:700;line-height:1;margin-bottom:0.3rem;}
.metric-lbl{font-family:'Roboto Mono',monospace;font-size:0.62rem;color:var(--text3);text-transform:uppercase;}
.sec-label{font-family:'Google Sans',sans-serif;font-size:0.75rem;font-weight:700;color:var(--text2);letter-spacing:0.08em;text-transform:uppercase;padding-bottom:0.3rem;border-bottom:2px solid var(--s2);margin:1rem 0 0.5rem;}
</style>
""", unsafe_allow_html=True)

# ======================================
# Utilidades de Datos y Configuración
# ======================================
def check_password():
    if st.session_state.get("password_correct", False):
        return True
    _, col, _ = st.columns([1, 2, 1])
    with col:
        with st.form("pw"):
            st.markdown("<h3 style='text-align:center;'>Sistema de Análisis FCF</h3>", unsafe_allow_html=True)
            pw = st.text_input("Contraseña", type="password")
            if st.form_submit_button("Ingresar", use_container_width=True):
                if pw == st.secrets.get("APP_PASSWORD", "INVALID"):
                    st.session_state["password_correct"] = True
                    st.rerun()
                else:
                    st.error("Contraseña incorrecta")
    return False

def load_local_config():
    paths = [Path("Configuracion.xlsx"), Path("configuracion.xlsx"), Path("Config.xlsx"), Path("config.xlsx")]
    for p in paths:
        if p.exists(): return p
    return None

def load_config_maps(config_path):
    try:
        config_sheets = pd.read_excel(config_path, sheet_name=None, engine='openpyxl')
        region_map = {}
        if 'Regiones' in config_sheets:
            region_map = pd.Series(
                config_sheets['Regiones'].iloc[:, 1].values,
                index=config_sheets['Regiones'].iloc[:, 0].astype(str).str.lower().str.strip()
            ).dropna().to_dict()
            
        internet_map = {}
        if 'Internet' in config_sheets:
            internet_map = pd.Series(
                config_sheets['Internet'].iloc[:, 1].values,
                index=config_sheets['Internet'].iloc[:, 0].astype(str).str.lower().str.strip()
            ).dropna().to_dict()
            
        return region_map, internet_map
    except Exception as e:
        st.warning(f"Error al cargar la configuración de mapeos locales: {e}")
        return {}, {}

# Sistema de Reintentos Exponencial y Detección Especializada de Error de Cuota (429)
def call_with_retries(fn, *a, **kw):
    d = 3
    for att in range(5):
        try: 
            return fn(*a, **kw)
        except Exception as e:
            err_str = str(e).lower()
            if "rate limit" in err_str or "429" in err_str:
                time.sleep(d * 3)
                d *= 2
            else:
                if att == 4: raise e
                time.sleep(d)
                d *= 2

async def acall_with_retries(fn, *a, **kw):
    d = 3
    for att in range(5):
        try: 
            return await fn(*a, **kw)
        except Exception as e:
            err_str = str(e).lower()
            if "rate limit" in err_str or "429" in err_str:
                await asyncio.sleep(d * 3)
                d *= 2
            else:
                if att == 4: raise e
                await asyncio.sleep(d)
                d *= 2

def normalize_title_for_comparison(title):
    if not isinstance(title, str): return ""
    cleaned = re.sub(r"\s+[\|–—-]\s+[^\|–—-]+$", "", title).strip()
    if ":" in cleaned:
        parts = cleaned.split(":", 1)
        if len(parts[1].strip()) >= 10: cleaned = parts[1].strip()
    return re.sub(r"\W+", " ", cleaned).lower().strip()

def normalizar_tipo_medio_fcf(tipo_raw):
    if not isinstance(tipo_raw, str): return str(tipo_raw)
    t = unidecode(tipo_raw.strip().lower())
    tipo_map = {
        'aire': 'Televisión', 'cable': 'Televisión', 'television': 'Televisión', 'televisión': 'Televisión',
        'am': 'Radio', 'fm': 'Radio', 'radio': 'Radio',
        'online': 'Internet', 'internet': 'Internet',
        'diario': 'Prensa', 'prensa': 'Prensa',
        'revista': 'Revistas', 'revistas': 'Revistas'
    }
    return tipo_map.get(t, tipo_raw.strip().title())

def buscar_vocero(resumen):
    if not isinstance(resumen, str) or not resumen.strip(): return ""
    res_norm = unidecode(resumen.lower())
    targets = ["ramon jesurun franco", "ramon jesurun", "jesurun"]
    if any(t in res_norm for t in targets):
        return "Ramón Jesurun"
    return ""

def texto_para_embedding(titulo, resumen, max_len=1800):
    t = str(titulo or "").strip()
    r = str(resumen or "").strip()
    return f"{t}. {r}"[:max_len]

# Helpers de mapeo robusto de cabeceras (Ignora espacios o acentos accidentales)
def get_row_value(row: dict, key_to_find: str) -> Any:
    norm_target = unidecode(key_to_find.strip().lower())
    for k, v in row.items():
        if unidecode(str(k).strip().lower()) == norm_target:
            return v
    return None

def set_row_value(row: dict, key_to_set: str, value: Any):
    norm_target = unidecode(key_to_set.strip().lower())
    for k in list(row.keys()):
        if unidecode(str(k).strip().lower()) == norm_target:
            row[k] = value
            return
    row[key_to_set] = value

def clean_json_string(s: str) -> str:
    s = s.strip()
    if s.startswith("```"):
        s = re.sub(r"^```(?:json)?\s*", "", s, flags=re.IGNORECASE)
        s = re.sub(r"\s*```$", "", s)
    return s.strip()

# Generador de Subtemas Dinámicos Lógicos de Respaldo (Evita "Varios" y "Sin clasificar")
def obtener_subtema_fallback(titulo: str, tema: str) -> str:
    words = [w.strip() for w in re.findall(r'[A-Za-zÀ-ÿ\d]+', str(titulo)) if len(w) > 4]
    
    if len(words) >= 2:
        phrase = f"Actualidad sobre {words[0].lower()} y {words[1].lower()}"
        return phrase[0].upper() + phrase[1:]
    elif len(words) == 1:
        phrase = f"Noticias relacionadas con {words[0].lower()}"
        return phrase[0].upper() + phrase[1:]
        
    defaults = {
        "Institucional": "Gestión Administrativa de FCF",
        "Torneos - Copas - Ligas": "Competición y Actividad de Clubes",
        "Selecciones": "Actualidad de Selecciones Nacionales",
        "Gestión": "Planificación y Proyectos Técnicos",
        "Jugadores": "Rendimiento Deportivo de Jugadores",
        "Entorno": "Relaciones del Fútbol Profesional"
    }
    return defaults.get(tema, "Información General de FCF")

def string_norm_label(s):
    if not s: return ""
    s = unidecode(s.lower())
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    return " ".join(t for t in s.split() if t not in STOPWORDS_ES)

# ======================================
# Consolidación de Subtemas (Deduplicación)
# ======================================
def dedup_labels(etiquetas: List[str], umbral: float = 0.82) -> List[str]:
    unique = list(dict.fromkeys(etiquetas))
    if len(unique) <= 1:
        return etiquetas
    normed = [string_norm_label(u) for u in unique]
    n = len(unique)
    parent = list(range(n))

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb: parent[rb] = ra

    for i in range(n):
        if not normed[i]: continue
        for j in range(i + 1, n):
            if not normed[j] or find(i) == find(j): continue
            if SequenceMatcher(None, normed[i], normed[j]).ratio() >= umbral:
                union(i, j)
                    
    le = get_embeddings_batch(unique)
    vp = [(i, le[i]) for i in range(n) if le[i] is not None]
    if len(vp) >= 2:
        vi, vv = zip(*vp)
        sm = cosine_similarity(np.array(vv))
        for pi in range(len(vi)):
            for pj in range(pi + 1, len(vi)):
                if sm[pi][pj] >= umbral:
                    if find(vi[pi]) != find(vi[pj]):
                        union(vi[pi], vi[pj])

    freq = Counter(etiquetas)
    grupos = defaultdict(list)
    for i in range(n):
        grupos[find(i)].append(i)
    canon = {}
    for root, members in grupos.items():
        cands = [unique[m] for m in members]
        canon[root] = max(cands, key=lambda c: (freq[c], len(c)))
    lm = {unique[i]: canon[find(i)] for i in range(n)}
    return [lm.get(e, e) for e in etiquetas]

def consolidar_sinonimos_llm(subtemas_unicos: List[str]) -> Dict[str, str]:
    if len(subtemas_unicos) <= 1:
        return {s: s for s in subtemas_unicos}
        
    prompt = (
        "Eres un analista de datos experto en taxonomías de fútbol.\n"
        "Tienes la siguiente lista de subtemas periodísticos generados para noticias de la FCF:\n"
        f"{', '.join(subtemas_unicos)}\n\n"
        "Tu tarea consiste en identificar aquellos subtemas que se refieren exactamente al MISMO EVENTO, PARTIDO o CONCEPTO "
        "y unificarlos bajo un único nombre representativo, claro y conciso.\n\n"
        "REGLAS:\n"
        "1. Agrupa variantes de partidos, análisis previos, rivales, etc. del mismo encuentro en un único subtema estructurado como: 'Partido Colombia vs [Rival]'\n"
        "   Ejemplo: 'Análisis de Portugal', 'Partido contra Portugal', 'Análisis previo Colombia-Portugal' deben unificarse todos en: 'Partido Colombia vs Portugal'.\n"
        "2. No fusiones conceptos que pertenezcan a eventos o categorías claramente distintas.\n"
        "3. Devuelve EXCLUSIVAMENTE un objeto JSON donde las claves sean los subtemas originales y los valores sean los subtemas unificados.\n\n"
        "Esquema JSON esperado:\n"
        '{"Subtema Original 1": "Subtema Unificado", "Subtema Original 2": "Subtema Unificado"}'
    )
    try:
        content, _, _ = call_with_retries(
            async_chat_completion,
            model=OPENAI_MODEL_CLASIFICACION,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=1500,
            temperature=0.0,
            response_format={"type": "json_object"}
        )
        cleaned = clean_json_string(content)
        return json.loads(cleaned)
    except Exception as e:
        return {s: s for s in subtemas_unicos}

# ======================================
# Estructura DSU para Agrupaciones
# ======================================
class DSU:
    def __init__(self, n):
        self.p = list(range(n))
        self.rank = [0] * n

    def find(self, i):
        path = []
        while self.p[i] != i:
            path.append(i)
            i = self.p[i]
        for node in path: self.p[node] = i
        return i

    def union(self, i, j):
        ri, rj = self.find(i), self.find(j)
        if ri == rj: return
        if self.rank[ri] < self.rank[rj]: ri, rj = rj, ri
        self.p[rj] = ri
        if self.rank[ri] == self.rank[rj]: self.rank[ri] += 1

    def grupos(self, n):
        c = defaultdict(list)
        for i in range(n): c[self.find(i)].append(i)
        return dict(c)

# ======================================
# Llamadas de Embeddings y Clasificación
# ======================================
async def clasificar_fcf_llm(titulo: str, resumen: str, sem: asyncio.Semaphore) -> dict:
    async with sem:
        prompt = (
            f"Eres un analista experto en reputación de medios especializado en la Federación Colombiana de Fútbol (FCF).\n"
            f"Tu tarea consiste en analizar la noticia suministrada y clasificar su TEMA, SUBTEMA e IMPACTO (Tono) con criterios estrictos y lógicos.\n\n"
            f"--- NOTICIA ---\n"
            f"Título: {titulo}\n"
            f"Resumen: {resumen}\n\n"
            f"--- REGLAS DE CLASIFICACIÓN ---\n"
            f"1. **TEMA**: Debe ser estrictamente uno de los siguientes 6 valores (elige el que mejor aplique):\n"
            f"   - Institucional (asuntos de la FCF como organización, finanzas, patrocinadores, asambleas de la FCF, decisiones ejecutivas)\n"
            f"   - Torneos - Copas - Ligas (partidos locales de clubes, copas nacionales/internacionales, boletería, arbitraje, etc.)\n"
            f"   - Selecciones (noticias, partidos, entrenamientos o convocatorias de la Selección Colombia masculina o femenina de cualquier categoría)\n"
            f"   - Gestión (capacitaciones de FCF, licencias de técnicos, de estadios, programas de talento o desarrollo de la federación)\n"
            f"   - Jugadores (noticias enfocadas en el rendimiento, transferencias o actualidad de futbolistas individuales: Luis Díaz, James, etc.)\n"
            f"   - Entorno (noticias de clubes o ligas, aniversarios, relaciones de la FCF con el gobierno, e incidentes del fútbol nacional)\n\n"
            f"2. **SUBTEMA**:\n"
            f"   - Crea o asocia un subtema específico para la noticia.\n"
            f"   - Debe ser una frase nominal concreta de 2 a 4 palabras, sin verbos conjugados, sin marcas comerciales y con ortografía correcta.\n"
            f"   - PROHIBICIÓN ESTRICTA: NUNCA uses la palabra 'Varios', 'Otros', 'Sin clasificar', 'Error' o términos placeholders. Si no encuentras un subtema exacto, genera una frase nominal descriptiva basada en el hecho central.\n\n"
            f"3. **IMPACTO**: Califica el tono reputacional hacia la FCF:\n"
            f"   - 'Positivo': Si la noticia resalta explícitamente una gestión, logro, o anuncio exitoso directo de la FCF.\n"
            f"   - 'Negativo': Si contiene críticas directas, cuestionamientos públicos, multas, fallas organizativas, quejas o comentarios desfavorables hacia la FCF.\n"
            f"   - 'Neutro': Información de partidos, crónicas ordinarias de resultados, fichajes, convocatorias regulares o mención puramente periodística donde no se exalta ni se critica a la FCF.\n\n"
            f"Responde estrictamente en formato JSON con el siguiente esquema:\n"
            f'{{"tema": "...", "subtema": "...", "impacto": "Positivo|Negativo|Neutro"}}'
        )

        try:
            content, prompt_tokens, completion_tokens = await acall_with_retries(
                async_chat_completion,
                model=OPENAI_MODEL_CLASIFICACION,
                messages=[{"role": "user", "content": prompt}],
                max_tokens=150,
                temperature=0.0,
                response_format={"type": "json_object"}
            )
            st.session_state['tokens_input'] += prompt_tokens
            st.session_state['tokens_output'] += completion_tokens
            
            cleaned_content = clean_json_string(content)
            res = json.loads(cleaned_content)
            
            tema = str(res.get("tema", "")).strip()
            subtema = str(res.get("subtema", "")).strip()
            impacto = str(res.get("impacto", "")).strip().title()

            if tema not in ["Institucional", "Torneos - Copas - Ligas", "Selecciones", "Gestión", "Jugadores", "Entorno"]:
                tema = "Institucional"
                
            sub_lower = subtema.lower()
            if not subtema or any(x in sub_lower for x in ["varios", "otros", "no aplica", "n/a", "error", "sin subtema"]):
                subtema = obtener_subtema_fallback(titulo, tema)

            if impacto not in ["Positivo", "Negativo", "Neutro"]:
                impacto = "Neutro"
                
            return {"tema": tema, "subtema": subtema, "impacto": impacto}
        except Exception as e:
            tema_fallback = "Institucional"
            sub_fallback = obtener_subtema_fallback(titulo, tema_fallback)
            return {"tema": tema_fallback, "subtema": sub_fallback, "impacto": "Neutro"}

# ======================================
# Proceso de Agrupamiento y Análisis
# ======================================
async def procesar_analisis_fcf(rows, headers, region_map, internet_map, pbar):
    n = len(rows)
    pbar.progress(0.05, "Mapeando regiones y buscando vocería...")
    
    # 1. Mapeos locales rigurosos
    for row in rows:
        res_val = get_row_value(row, "RESUMEN") or ""
        set_row_value(row, "VOCERO", buscar_vocero(str(res_val)))

        # Determinar Región antes de alterar NOMBRE DE MEDIO
        nm_val = get_row_value(row, "NOMBRE DE MEDIO") or ""
        nm_clean = str(nm_val).strip().lower()
        if nm_clean in region_map:
            set_row_value(row, "REGION", region_map[nm_clean])
        else:
            current_reg = get_row_value(row, "REGION")
            if not current_reg:
                set_row_value(row, "REGION", "N/A")

        # Re-mapeo y normalización de Tipo de Medio
        tm_val = get_row_value(row, "TIPO DE MEDIO") or ""
        tipo_norm = normalizar_tipo_medio_fcf(str(tm_val))
        set_row_value(row, "TIPO DE MEDIO", tipo_norm)

        # Mapeo de nombres de Internet solo para tipo Online / Internet
        if str(tm_val).strip().lower() in ["online", "internet"] or tipo_norm == "Internet":
            if nm_clean in internet_map:
                set_row_value(row, "NOMBRE DE MEDIO", internet_map[nm_clean])

    # 2. Agrupamiento Lógico de Noticias Similares
    pbar.progress(0.15, "Agrupando noticias similares por títulos...")
    dsu = DSU(n)
    titulos = [str(get_row_value(r, "TÍTULO") or get_row_value(r, "TITULO") or "").strip() for r in rows]
    norm_titles = [normalize_title_for_comparison(t) for t in titulos]

    for i in range(n):
        if not norm_titles[i]: continue
        for j in range(i + 1, n):
            if not norm_titles[j] or dsu.find(i) == dsu.find(j): continue
            if SequenceMatcher(None, norm_titles[i], norm_titles[j]).ratio() >= SIMILARITY_THRESHOLD_TITULOS:
                dsu.union(i, j)

    pbar.progress(0.30, "Analizando semántica de texto...")
    textos = [
        texto_para_embedding(
            str(get_row_value(r, "TÍTULO") or get_row_value(r, "TITULO") or ""),
            str(get_row_value(r, "RESUMEN") or "")
        )
        for r in rows
    ]
    embs = get_embeddings_batch(textos)
    validos = [(i, e) for i, e in enumerate(embs) if e is not None]

    if len(validos) >= 2:
        idxs, M = zip(*validos)
        labels = AgglomerativeClustering(
            n_clusters=None,
            distance_threshold=1 - SIMILARITY_THRESHOLD_GROUP,
            metric="cosine",
            linkage="complete"
        ).fit(np.array(M)).labels_
        
        g = defaultdict(list)
        for k, lbl in enumerate(labels):
            g[lbl].append(idxs[k])
            
        for cl in g.values():
            for j in cl[1:]:
                dsu.union(cl[0], j)

    # 3. Procesamiento en Paralelo de Clasificación Inicial con IA
    grupos = dsu.grupos(n)
    total_grupos = len(grupos)
    pbar.progress(0.45, f"Clasificando {total_grupos} grupos de noticias con {OPENAI_MODEL_CLASIFICACION}...")

    sem = asyncio.Semaphore(CONCURRENT_REQUESTS)
    tasks = []
    cids = list(grupos.keys())

    for cid in cids:
        rep_idx = grupos[cid][0]
        rep_title = str(get_row_value(rows[rep_idx], "TÍTULO") or get_row_value(rows[rep_idx], "TITULO") or "")
        rep_resumen = str(get_row_value(rows[rep_idx], "RESUMEN") or "")
        tasks.append(clasificar_fcf_llm(rep_title, rep_resumen, sem))

    resultados = await asyncio.gather(*tasks)
    rpg = {cids[k]: resultados[k] for k in range(len(cids))}

    # Asignación inicial de clasificaciones a los registros
    for cid, idxs in grupos.items():
        evaluacion = rpg.get(cid, {"tema": "Institucional", "subtema": "Actualidad de FCF", "impacto": "Neutro"})
        for i in idxs:
            set_row_value(rows[i], "TEMA", evaluacion["tema"])
            set_row_value(rows[i], "SUBTEMA", evaluacion["subtema"])
            set_row_value(rows[i], "Impacto", evaluacion["impacto"])

    # 4. Pipeline de Consolidación Semántica de Subtemas
    pbar.progress(0.85, "Consolidando subtemas similares para consistencia...")
    subtemas_iniciales = [get_row_value(r, "SUBTEMA") for r in rows if get_row_value(r, "SUBTEMA")]
    
    # Capa 1 y 2: Deduplicación léxico-semántica local
    subtemas_dedup = dedup_labels(subtemas_iniciales, umbral=0.82)
    mapa_dedup = dict(zip(subtemas_iniciales, subtemas_dedup))
    
    for r in rows:
        sub_act = get_row_value(r, "SUBTEMA")
        if sub_act in mapa_dedup:
            set_row_value(r, "SUBTEMA", mapa_dedup[sub_act])
            
    # Capa 3: Agrupación avanzada de sinónimos de partidos/eventos por IA
    subtemas_unificados_unicos = list(set(mapa_dedup.values()))
    mapa_sinonimos = consolidar_sinonimos_llm(subtemas_unificados_unicos)
    
    for r in rows:
        sub_act = get_row_value(r, "SUBTEMA")
        if sub_act in mapa_sinonimos:
            set_row_value(r, "SUBTEMA", mapa_sinonimos[sub_act])

    # 5. Capa 4: Homogeneización de Tema e Impacto para subtemas consolidados
    sub_to_data = defaultdict(list)
    for r in rows:
        sub = get_row_value(r, "SUBTEMA")
        tema = get_row_value(r, "TEMA")
        imp = get_row_value(r, "Impacto")
        sub_to_data[sub].append((tema, imp))
        
    sub_to_best = {}
    for sub, items in sub_to_data.items():
        temas = [it[0] for it in items if it[0]]
        impactos = [it[1] for it in items if it[1]]
        best_tema = Counter(temas).most_common(1)[0][0] if temas else "Institucional"
        best_imp = Counter(impactos).most_common(1)[0][0] if impactos else "Neutro"
        sub_to_best[sub] = (best_tema, best_imp)
        
    for r in rows:
        sub = get_row_value(r, "SUBTEMA")
        if sub in sub_to_best:
            set_row_value(r, "TEMA", sub_to_best[sub][0])
            set_row_value(r, "Impacto", sub_to_best[sub][1])

    pbar.progress(1.0, "Análisis de clasificación y consolidación finalizado.")
    return rows

# ======================================
# Generador de Excel de Salida
# ======================================
def generate_fcf_output(rows, headers):
    wb = Workbook()
    ws = wb.active
    ws.title = "Analisis FCF"

    ws.append(headers)
    font_header = Font(bold=True)
    for idx in range(1, len(headers) + 1):
        ws.cell(row=1, column=idx).font = font_header

    font_hyperlink = Font(color="059669", underline="single")
    align_left = Alignment(horizontal='left')

    for r_idx, row in enumerate(rows, start=2):
        out_row = []
        links = {}
        for c_idx, h in enumerate(headers, start=1):
            val = row.get(h)
            cv = None
            if isinstance(val, dict) and "url" in val:
                cv = val.get("value", "Link")
                if val.get("url"): links[c_idx] = val["url"]
            else:
                cv = val
            out_row.append(cv)
        ws.append(out_row)

        for c_idx, url in links.items():
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.hyperlink = url
            cell.font = font_hyperlink
            cell.alignment = align_left

        for c_idx, h in enumerate(headers, start=1):
            if str(h).strip().lower() == "fecha":
                cell = ws.cell(row=r_idx, column=c_idx)
                if isinstance(cell.value, (datetime.datetime, datetime.date)):
                    cell.number_format = 'DD/MM/YYYY'
                elif isinstance(cell.value, pd.Timestamp):
                    cell.value = cell.value.to_pydatetime()
                    cell.number_format = 'DD/MM/YYYY'

    for idx, h in enumerate(headers, start=1):
        letter = ws.cell(row=1, column=idx).column_letter
        h_clean = str(h).strip().lower()
        if "título" in h_clean or "titulo" in h_clean or "resumen" in h_clean:
            ws.column_dimensions[letter].width = 50
        elif "link" in h_clean:
            ws.column_dimensions[letter].width = 15
        else:
            ws.column_dimensions[letter].width = 18

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ======================================
# Orquestador del Flujo de Trabajo
# ======================================
async def run_fcf_pipeline(df_file):
    get_embedding_cache().clear()
    st.session_state.update({'tokens_input': 0, 'tokens_output': 0, 'tokens_embedding': 0})
    t0 = time.time()

    try:
        openai.api_key = st.secrets["OPENAI_API_KEY"]
    except Exception as e:
        st.error("Error: la llave OPENAI_API_KEY no se encuentra definida en st.secrets.")
        st.stop()

    with st.status("Cargando y procesando archivos de entrada...", expanded=True) as status_step:
        config_path = load_local_config()
        if not config_path:
            st.error("❌ No se encontró el archivo 'Configuracion.xlsx' en el repositorio.")
            st.stop()
            
        region_map, internet_map = load_config_maps(config_path)

        wb_in = load_workbook(df_file, data_only=True)
        sheet = wb_in.active

        # Extraer cabeceras reales
        headers = [str(cell.value) for cell in sheet[1] if cell.value is not None]
        
        # Mapeo de filas del Excel en diccionarios
        rows = []
        for row in sheet.iter_rows(min_row=2):
            if all(c.value is None for c in row): continue
            row_data = {}
            for i, h in enumerate(headers):
                if i < len(row):
                    cell = row[i]
                    val = cell.value
                    url = cell.hyperlink.target if (cell.hyperlink and cell.hyperlink.target) else None
                    if url:
                        row_data[h] = {"value": val or "Link", "url": url}
                    else:
                        row_data[h] = val
            rows.append(row_data)

        status_step.update(label="✓ Archivos cargados correctamente.", state="complete")

    # Validar columnas de entrada críticas
    norm_headers = [h.strip().lower() for h in headers]
    for r_col in ["título", "resumen"]:
        if not any(r_col in nh for nh in norm_headers):
            st.error(f"Columna requerida ausente en el archivo de entrada: {r_col.upper()}")
            st.stop()

    with st.status("Analizando contenido con Inteligencia Artificial...", expanded=True) as status_step:
        pbar = st.progress(0.0)
        rows_processed = await procesar_analisis_fcf(rows, headers, region_map, internet_map, pbar)
        status_step.update(label="✓ Clasificación de noticias completada.", state="complete")

    ci = (st.session_state['tokens_input']     / 1e6) * PRICE_INPUT_1M
    co = (st.session_state['tokens_output']    / 1e6) * PRICE_OUTPUT_1M
    ce = (st.session_state['tokens_embedding'] / 1e6) * PRICE_EMBEDDING_1M

    st.session_state["output_data"] = generate_fcf_output(rows_processed, headers)
    st.session_state["output_filename"] = f"Informe_IA_FCF_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    st.session_state["processing_complete"] = True
    st.session_state.update({
        "total_rows": len(rows),
        "process_duration": f"{time.time() - t0:.0f}s",
        "process_cost": f"${ci + co + ce:.4f} USD",
        "cache_stats": get_embedding_cache().stats()
    })

# ======================================
# Función Main (Interfaz Streamlit)
# ======================================
def main():
    load_custom_css()
    if not check_password(): return

    st.markdown("""
    <div class="app-header">
        <div class="app-header-icon">⚽</div>
        <div class="app-header-text">
            <div class="app-header-title">Análisis de Noticias FCF · Procesamiento de Reputación</div>
            <div style="font-size:0.75rem; color:#5f6368;">Clasificación de noticias con el modelo gpt-4.1-nano-2025-04-14</div>
        </div>
        <div class="app-header-badge">FCF IA</div>
    </div>""", unsafe_allow_html=True)

    if not st.session_state.get("processing_complete", False):
        st.markdown('<div class="sec-label">Subir Dossier para Análisis</div>', unsafe_allow_html=True)
        st.markdown("Asegúrate de que tu archivo cuente con las columnas obligatorias: **TÍTULO** y **RESUMEN**.")
        f1 = st.file_uploader("Dossier (.xlsx)", type=["xlsx"], label_visibility="collapsed")

        if st.button("▶ Iniciar Análisis Reputacional", use_container_width=True, type="primary"):
            if not f1:
                st.error("Por favor, sube un archivo antes de continuar.")
            else:
                asyncio.run(run_fcf_pipeline(f1))
                st.rerun()
    else:
        st.markdown(
            '<div style="background-color:var(--accent-bg); border:1px solid var(--accent-bdr); border-radius:12px; padding:1rem; margin-bottom:1rem;">'
            '<h4 style="color:#047857; margin:0;">✓ Procesamiento completado</h4>'
            '<p style="margin:0.2rem 0 0 0; font-size:0.85rem;">Las noticias han sido evaluadas y agrupadas lógicamente.</p>'
            '</div>',
            unsafe_allow_html=True
        )

        st.markdown('<div class="sec-label">Métricas de Ejecución</div>', unsafe_allow_html=True)
        st.markdown(f"""
        <div class="metrics-grid">
          <div class="metric-card"><div class="metric-val">{st.session_state.total_rows}</div><div class="metric-lbl">Total Noticias</div></div>
          <div class="metric-card"><div class="metric-val" style="color:#059669;">6</div><div class="metric-lbl">Temas Posibles</div></div>
          <div class="metric-card"><div class="metric-val" style="color:#1a73e8;">{st.session_state.process_duration}</div><div class="metric-lbl">Tiempo de Cómputo</div></div>
          <div class="metric-card"><div class="metric-val" style="color:#d97706;">{st.session_state.process_cost}</div><div class="metric-lbl">Costo Consumo</div></div>
        </div>""", unsafe_allow_html=True)
        
        st.caption(f"📊 Detalle del cache: {st.session_state.cache_stats}")

        st.markdown('<div class="sec-label">Descargar Resultado</div>', unsafe_allow_html=True)
        st.download_button(
            "⬇ Descargar Informe de Análisis FCF",
            data=st.session_state.output_data,
            file_name=st.session_state.output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
        
        if st.button("Procesar un nuevo dossier", use_container_width=True):
            pwd = st.session_state.get("password_correct")
            st.session_state.clear()
            st.session_state.password_correct = pwd
            st.rerun()

    st.markdown(
        '<div style="text-align:center; font-family:\'Roboto Mono\',monospace; font-size:0.65rem; color:#9aa0a6; margin-top:2rem; padding-top:1rem; border-top:1px solid #e8eaed;">'
        'Análisis de Noticias FCF v2.0 · Johnathan Cortés ©'
        '</div>',
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
