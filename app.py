import io
import json
import re
import time
from collections import defaultdict
from difflib import SequenceMatcher
from pathlib import Path

import numpy as np
import openai
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from sklearn.metrics.pairwise import cosine_similarity
from unidecode import unidecode


st.set_page_config(
    page_title="Analisis FCF",
    page_icon="FCF",
    layout="wide",
    initial_sidebar_state="collapsed",
)

OPENAI_MODEL_EMBEDDING = "text-embedding-3-small"
OPENAI_MODEL_CLASIFICACION = "gpt-4.1-nano-2025-04-14"

TEMAS_FCF = [
    "Institucional",
    "Torneos - Copas - Ligas",
    "Selecciones",
    "Gestión",
    "Jugadores",
    "Entorno",
]

TONOS_FCF = ["Positivo", "Negativo", "Neutro"]

INPUT_REQUIRED_COLUMNS = [
    "ID",
    "FECHA",
    "HORA",
    "TIPO DE MEDIO",
    "NOMBRE DE MEDIO",
    "REGION",
    "SECCIÓN",
    "TÍTULO",
    "RESUMEN",
]

OUTPUT_TONO_COL = "Impacto"
OUTPUT_TEMA_COL = "TEMA"
OUTPUT_SUBTEMA_COL = "SUBTEMA"
OUTPUT_VOCERO_COL = "VOCERO"

SIMILARIDAD_TITULO = 0.93
SIMILARIDAD_DUPLICADO_TITULO = 0.97
SIMILARIDAD_EMBEDDING = 0.90
BATCH_EMBEDDINGS = 80

SELECCIONES = [
    "Argentina", "Bolivia", "Brasil", "Chile", "Costa Rica", "Ecuador",
    "Estados Unidos", "Guatemala", "Honduras", "Mexico", "México", "Panama",
    "Panamá", "Paraguay", "Peru", "Perú", "Uruguay", "Venezuela", "España",
    "Inglaterra", "Francia", "Alemania", "Italia", "Portugal", "Marruecos",
    "Japon", "Japón", "Corea", "Senegal", "Nigeria", "Camerun", "Camerún",
    "Australia", "Nueva Zelanda", "Canada", "Canadá", "Qatar",
]

SELECCION_ALIASES = {
    "Argentina": ["argentina", "albiceleste"],
    "Bolivia": ["bolivia"],
    "Brasil": ["brasil", "brazil", "canarinha"],
    "Chile": ["chile"],
    "Costa Rica": ["costa rica", "ticos", "tica"],
    "Ecuador": ["ecuador", "tri"],
    "Estados Unidos": ["estados unidos", "eeuu", "ee uu", "usa", "usmnt"],
    "Guatemala": ["guatemala"],
    "Honduras": ["honduras"],
    "Mexico": ["mexico", "seleccion mexicana", "tri mexicano"],
    "Panama": ["panama", "canaleros"],
    "Paraguay": ["paraguay", "albirroja"],
    "Peru": ["peru", "blanquirroja"],
    "Uruguay": ["uruguay", "charrua", "celeste"],
    "Venezuela": ["venezuela", "vinotinto"],
    "España": ["espana", "españa"],
    "Inglaterra": ["inglaterra", "england"],
    "Francia": ["francia", "france"],
    "Alemania": ["alemania", "germany"],
    "Italia": ["italia", "italy"],
    "Portugal": ["portugal"],
    "Marruecos": ["marruecos", "morocco"],
    "Japon": ["japon", "japón", "japan"],
    "Corea": ["corea", "corea del sur", "south korea"],
    "Senegal": ["senegal"],
    "Nigeria": ["nigeria"],
    "Camerun": ["camerun", "camerún", "cameroon"],
    "Australia": ["australia"],
    "Nueva Zelanda": ["nueva zelanda", "new zealand"],
    "Canada": ["canada", "canadá"],
    "Qatar": ["qatar", "catar"],
}

COLOMBIA_ALIASES = [
    "colombia", "seleccion colombia", "seleccion colombiana",
    "seleccion de colombia", "tricolor", "la tricolor",
]


class DSU:
    def __init__(self, n):
        self.parent = list(range(n))
        self.rank = [0] * n

    def find(self, x):
        while self.parent[x] != x:
            self.parent[x] = self.parent[self.parent[x]]
            x = self.parent[x]
        return x

    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra == rb:
            return
        if self.rank[ra] < self.rank[rb]:
            ra, rb = rb, ra
        self.parent[rb] = ra
        if self.rank[ra] == self.rank[rb]:
            self.rank[ra] += 1

    def groups(self):
        out = defaultdict(list)
        for i in range(len(self.parent)):
            out[self.find(i)].append(i)
        return list(out.values())


def normalize_text(value):
    text = "" if pd.isna(value) else str(value)
    text = unidecode(text.lower())
    text = re.sub(r"https?://\S+|www\.\S+", " ", text)
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    return " ".join(text.split())


def get_column(df, candidates):
    by_norm = {normalize_text(col): col for col in df.columns}
    for candidate in candidates:
        key = normalize_text(candidate)
        if key in by_norm:
            return by_norm[key]
    return None


def load_custom_css():
    st.markdown(
        """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
:root{
  --fcf-blue:#173b7a;
  --fcf-gold:#d9a441;
  --fcf-red:#c8202f;
  --surface:#ffffff;
  --muted:#667085;
  --border:#d9e2ef;
  --soft:#f5f8fc;
}
.stApp{background:linear-gradient(180deg,#f7faff 0%,#ffffff 36%);font-family:'Inter',sans-serif;}
.block-container{padding-top:1.3rem;max-width:1220px;}
.fcf-hero{border:1px solid var(--border);border-radius:8px;background:var(--surface);padding:1rem 1.1rem;margin-bottom:.9rem;box-shadow:0 8px 24px rgba(23,59,122,.08);}
.fcf-kicker{font-size:.72rem;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:var(--fcf-blue);margin-bottom:.25rem;}
.fcf-title{font-size:1.8rem;line-height:1.1;font-weight:800;color:#12233f;margin:0;}
.fcf-sub{color:var(--muted);font-size:.9rem;margin-top:.35rem;}
.sec-label{font-size:.72rem;font-weight:800;color:var(--fcf-blue);letter-spacing:.08em;text-transform:uppercase;margin:.8rem 0 .4rem;}
.upload-zone{display:grid;grid-template-columns:1fr 1fr;gap:.7rem;margin:.4rem 0 .8rem;}
.upload-zone-card{background:#fff;border:1.5px dashed var(--border);border-radius:8px;padding:.75rem .85rem;}
.upload-zone-title{font-size:.88rem;font-weight:800;color:#12233f;}
.upload-zone-desc{font-size:.75rem;color:var(--muted);margin-top:.2rem;}
.metrics-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:.7rem;margin:.75rem 0;}
.metric-card{background:#fff;border:1px solid var(--border);border-radius:8px;padding:.8rem .9rem;box-shadow:0 6px 18px rgba(23,59,122,.06);}
.metric-val{font-size:1.45rem;font-weight:800;line-height:1;color:#12233f;}
.metric-lbl{font-size:.68rem;font-weight:800;letter-spacing:.07em;text-transform:uppercase;color:var(--muted);margin-top:.28rem;}
.success-banner{background:#effaf4;border:1px solid #bbebcd;border-left:4px solid #16a34a;border-radius:8px;padding:.8rem 1rem;margin:.5rem 0 .8rem;color:#14532d;font-weight:700;}
div.stButton > button:first-child{background:var(--fcf-blue);border-color:var(--fcf-blue);border-radius:8px;font-weight:800;}
div.stDownloadButton > button:first-child{background:var(--fcf-red);border-color:var(--fcf-red);border-radius:8px;font-weight:800;color:white;}
@media(max-width:800px){.upload-zone,.metrics-grid{grid-template-columns:1fr}.fcf-title{font-size:1.35rem}}
</style>
""",
        unsafe_allow_html=True,
    )


def extract_embedded_links(xlsx_bytes, columns=("LINK", "WEB")):
    workbook = load_workbook(io.BytesIO(xlsx_bytes), data_only=False)
    worksheet = workbook.active
    headers = [cell.value for cell in worksheet[1]]
    header_positions = {normalize_text(value): idx + 1 for idx, value in enumerate(headers)}
    links = {}

    for column_name in columns:
        column_position = header_positions.get(normalize_text(column_name))
        if not column_position:
            continue

        for excel_row in range(2, worksheet.max_row + 1):
            cell = worksheet.cell(row=excel_row, column=column_position)
            target = None
            if cell.hyperlink:
                target = cell.hyperlink.target or cell.hyperlink.location
            elif isinstance(cell.value, str) and cell.value.strip().lower().startswith(("http://", "https://")):
                target = cell.value.strip()

            if target:
                row_index = excel_row - 2
                links.setdefault(row_index, {})[column_name] = target

    return links


def normalize_url(value):
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip()
    if not text:
        return ""
    return text.rstrip("/").lower()


def get_row_link(row_index, col_name, df, hyperlinks):
    if hyperlinks and row_index in hyperlinks and col_name in hyperlinks[row_index]:
        return hyperlinks[row_index][col_name]
    actual_col = get_column(df, [col_name])
    if actual_col:
        return df.at[row_index, actual_col]
    return ""


def load_local_config():
    paths_to_try = [
        Path("Configuracion.xlsx"),
        Path("configuracion.xlsx"),
        Path(__file__).parent / "Configuracion.xlsx",
        Path(__file__).parent / "configuracion.xlsx",
        Path(__file__).parent.parent / "Configuracion.xlsx",
        Path(__file__).parent.parent / "configuracion.xlsx",
        Path(__file__).parent.parent / "Grill-API-main" / "Configuracion.xlsx",
        Path(__file__).parent.parent / "Grill-API-main" / "configuracion.xlsx",
    ]
    for path in paths_to_try:
        if path.exists():
            return path
    return None


def load_config_sheets(config_source):
    return pd.read_excel(config_source, sheet_name=None, engine="openpyxl")


def load_region_lookup(config_source):
    sheets = load_config_sheets(config_source)
    for sheet_name, sheet_df in sheets.items():
        if sheet_df.empty:
            continue

        medio_col = get_column(sheet_df, ["NOMBRE DE MEDIO", "MEDIO", "Nombre de Medio"])
        region_col = get_column(sheet_df, ["REGION", "Región"])
        if medio_col and region_col:
            source = sheet_df[[medio_col, region_col]].dropna(subset=[medio_col])
        elif normalize_text(sheet_name) == "regiones" and sheet_df.shape[1] >= 2:
            source = sheet_df.iloc[:, [0, 1]].dropna(subset=[sheet_df.columns[0]])
            medio_col = source.columns[0]
            region_col = source.columns[1]
        else:
            continue

        lookup = {}
        for _, row in source.iterrows():
            key = normalize_text(row[medio_col])
            value = row[region_col]
            if key and not pd.isna(value):
                lookup[key] = str(value).strip()
        if lookup:
            return lookup, sheet_name

    return {}, None


def apply_region_lookup(df, medio_col, region_col, config_source=None, config_label=None):
    config_path = config_source or load_local_config()
    if not config_path:
        return df, "No se encontro Configuracion.xlsx; se conserva REGION del archivo cargado."

    lookup, sheet_name = load_region_lookup(config_path)
    if not lookup:
        label = config_label or getattr(config_path, "name", "Configuracion.xlsx")
        return df, f"No se encontraron columnas NOMBRE DE MEDIO y REGION en {label}."

    result = df.copy()
    mapped = result[medio_col].apply(lambda value: lookup.get(normalize_text(value)))
    result[region_col] = mapped.fillna(result[region_col])
    found = int(mapped.notna().sum())
    label = config_label or getattr(config_path, "name", "Configuracion.xlsx")
    missing = len(result) - found
    return result, f"Buscarv NOMBRE DE MEDIO -> REGION aplicado desde {label} ({sheet_name}): {found}/{len(result)} medios encontrados; {missing} sin coincidencia."


def build_analysis_text(title, summary):
    title = "" if pd.isna(title) else str(title).strip()
    summary = "" if pd.isna(summary) else str(summary).strip()
    return f"TITULO: {title}\nRESUMEN: {summary}".strip()


def detect_duplicate_rows(df, medio_col, title_col, web_col=None, hyperlinks=None):
    duplicates = set()

    by_media = defaultdict(list)
    for idx, row in df.iterrows():
        media_key = normalize_text(row.get(medio_col, ""))
        title_key = normalize_text(row.get(title_col, ""))
        if media_key and title_key:
            by_media[media_key].append((idx, title_key))

    for rows in by_media.values():
        for pos, (idx, title_key) in enumerate(rows):
            if idx in duplicates:
                continue
            for other_idx, other_title in rows[pos + 1:]:
                if other_idx in duplicates:
                    continue
                if title_key == other_title:
                    duplicates.add(other_idx)
                    continue
                if SequenceMatcher(None, title_key, other_title).ratio() >= SIMILARIDAD_DUPLICADO_TITULO:
                    duplicates.add(other_idx)

    if web_col or hyperlinks:
        seen_web = {}
        for idx in df.index:
            web_link = normalize_url(get_row_link(idx, "WEB", df, hyperlinks))
            if not web_link:
                continue
            if web_link in seen_web:
                duplicates.add(idx)
            else:
                seen_web[web_link] = idx

    return duplicates


def is_fcf_photo_summary(summary):
    text = normalize_text(summary)
    if not text:
        return False
    return bool(re.search(r"\bfoto\s*(de\s*)?fcf\b", text))


def title_case_label(text):
    clean = " ".join(str(text).strip().split())
    if not clean:
        return clean
    small = {"de", "del", "la", "el", "los", "las", "ante", "y"}
    words = []
    for pos, word in enumerate(clean.split()):
        low = word.lower()
        if pos > 0 and low in small:
            words.append(low)
        else:
            words.append(word[:1].upper() + word[1:].lower())
    return " ".join(words)


def has_any_phrase(norm, phrases):
    return any(re.search(rf"\b{re.escape(normalize_text(phrase))}\b", norm) for phrase in phrases)


def is_partido_context(norm):
    match_words = r"\b(partido|encuentro|juego|duelo|fecha|fixture|calendario|debut|enfrenta|enfrentara|enfrentará|enfrento|enfrentó|vs|contra|ante)\b"
    tournament_words = r"\b(mundial|copa del mundo|mundial 2026|eliminatoria|eliminatorias|clasificatorio|clasificatorias|seleccion|selecciones)\b"
    colombia_context = has_any_phrase(norm, COLOMBIA_ALIASES)
    return bool(re.search(match_words, norm)) and (bool(re.search(tournament_words, norm)) or colombia_context)


def label_for_selection_alias(alias_norm):
    for label, aliases in SELECCION_ALIASES.items():
        if alias_norm in {normalize_text(alias) for alias in aliases}:
            return title_case_label(label)
    return title_case_label(alias_norm)


def find_rival_selection(norm):
    colombia_pattern = r"(?:colombia|seleccion colombia|seleccion colombiana|seleccion de colombia|tricolor|la tricolor)"
    action_pattern = r"(?:enfrenta|enfrentara|enfrento|jugara|jugo|juega|mide|medira|choca|chocara|disputa|disputara)"

    for label, aliases in SELECCION_ALIASES.items():
        for alias in aliases:
            alias_norm = normalize_text(alias)
            rival_pattern = re.escape(alias_norm)
            patterns = [
                rf"\b(?:ante|contra)\s+(?:la\s+seleccion\s+de\s+|seleccion\s+de\s+|seleccion\s+)?{rival_pattern}\b",
                rf"\bfrente\s+a\s+(?:la\s+seleccion\s+de\s+|seleccion\s+de\s+|seleccion\s+)?{rival_pattern}\b",
                rf"\bvs\.?\s+{rival_pattern}\b",
                rf"\b{rival_pattern}\s+vs\.?\s+{colombia_pattern}\b",
                rf"\b{colombia_pattern}\s+vs\.?\s+{rival_pattern}\b",
                rf"\b{colombia_pattern}\s+[-–]\s+{rival_pattern}\b",
                rf"\b{rival_pattern}\s+[-–]\s+{colombia_pattern}\b",
                rf"\b{colombia_pattern}\s+(?:se\s+)?{action_pattern}\s+(?:ante\s+|contra\s+|frente\s+a\s+|con\s+|a\s+)?{rival_pattern}\b",
                rf"\b{rival_pattern}\s+(?:se\s+)?{action_pattern}\s+(?:ante\s+|contra\s+|frente\s+a\s+|con\s+|a\s+)?{colombia_pattern}\b",
            ]
            if any(re.search(pattern, norm) for pattern in patterns):
                return title_case_label(label)

    return None


def detect_partido_subtema(text):
    norm = normalize_text(text)
    if not is_partido_context(norm):
        return None

    rival = find_rival_selection(norm)
    if rival:
        return f"Partido ante {rival}"

    if re.search(r"\b(copa del mundo|copa mundial|copa mundo|mundial 2026|mundial)\b", norm):
        return "Partido Mundial 2026"
    if re.search(r"\b(eliminatoria|eliminatorias|clasificatorio|clasificatorias)\b", norm):
        return "Partido Eliminatorias"
    return None


def normalize_subtema_fcf(text, tema, subtema):
    partido = detect_partido_subtema(text)
    if partido:
        return "Selecciones", partido
    generic = normalize_text(subtema)
    if generic in {
        "partido mundial", "partido copa del mundo", "partido mundial 2026",
        "partido de mundial", "partido del mundial", "juego mundial",
        "encuentro mundial", "partido copa mundo", "partido copa mundial",
        "partido de copa del mundo", "partido del copa del mundo",
    }:
        return "Selecciones", "Partido Mundial 2026"
    if generic in {"partido eliminatorias", "partido eliminatoria", "juego eliminatorias"}:
        return "Selecciones", "Partido Eliminatorias"
    return tema, subtema


def apply_partido_normalization_to_result(result, row_idx, text):
    current_tema = result.at[row_idx, OUTPUT_TEMA_COL]
    current_subtema = result.at[row_idx, OUTPUT_SUBTEMA_COL]
    tema, subtema = normalize_subtema_fcf(text, current_tema, current_subtema)
    result.at[row_idx, OUTPUT_TEMA_COL] = tema
    result.at[row_idx, OUTPUT_SUBTEMA_COL] = subtema


def choose_group_partido_subtema(row_texts):
    detected = [detect_partido_subtema(text) for text in row_texts]
    detected = [item for item in detected if item]
    if not detected:
        return None

    rival_labels = sorted({item for item in detected if item.startswith("Partido ante ")})
    if len(rival_labels) == 1:
        return rival_labels[0]
    if len(rival_labels) > 1:
        return None
    if any(item == "Partido Mundial 2026" for item in detected):
        return "Partido Mundial 2026"
    if any(item == "Partido Eliminatorias" for item in detected):
        return "Partido Eliminatorias"
    return None


def detect_vocero(summary):
    text = normalize_text(summary)
    patterns = [
        "ramon jesurun franco",
        "ramon jesurun",
        "jesurun",
    ]
    return "Ramón Jesurun" if any(p in text for p in patterns) else "Sin vocero"


def representative_index(group, texts):
    if len(group) == 1:
        return group[0]
    lengths = [(len(texts[i]), i) for i in group]
    return sorted(lengths, reverse=True)[0][1]


def group_similar_news(df, title_col, summary_col, progress=None):
    n = len(df)
    dsu = DSU(n)
    titles = df[title_col].fillna("").astype(str).tolist()
    summaries = df[summary_col].fillna("").astype(str).tolist()
    texts = [build_analysis_text(t, s) for t, s in zip(titles, summaries)]

    exact_buckets = defaultdict(list)
    for i, (title, summary) in enumerate(zip(titles, summaries)):
        key = normalize_text(f"{title} {summary}")[:500]
        if key:
            exact_buckets[key].append(i)
    for bucket in exact_buckets.values():
        for item in bucket[1:]:
            dsu.union(bucket[0], item)

    norm_titles = [normalize_text(title) for title in titles]
    for i in range(n):
        if progress and n:
            progress.progress(min(0.25, 0.05 + 0.20 * (i + 1) / n), "Agrupando titulos similares...")
        if not norm_titles[i]:
            continue
        for j in range(i + 1, n):
            if dsu.find(i) == dsu.find(j) or not norm_titles[j]:
                continue
            if SequenceMatcher(None, norm_titles[i], norm_titles[j]).ratio() >= SIMILARIDAD_TITULO:
                dsu.union(i, j)

    embeddings = get_embeddings(texts, progress)
    if embeddings is not None and len(embeddings) == n:
        matrix = np.array(embeddings)
        sim = cosine_similarity(matrix)
        for i in range(n):
            if progress and n:
                progress.progress(min(0.55, 0.30 + 0.25 * (i + 1) / n), "Agrupando noticias por similitud semantica...")
            for j in range(i + 1, n):
                if dsu.find(i) != dsu.find(j) and sim[i][j] >= SIMILARIDAD_EMBEDDING:
                    dsu.union(i, j)

    return dsu.groups(), texts


def get_embeddings(texts, progress=None):
    try:
        vectors = []
        clean_texts = [text[:2500] for text in texts]
        for start in range(0, len(clean_texts), BATCH_EMBEDDINGS):
            batch = clean_texts[start:start + BATCH_EMBEDDINGS]
            response = openai.Embedding.create(input=batch, model=OPENAI_MODEL_EMBEDDING)
            vectors.extend([item["embedding"] for item in response["data"]])
            if progress and clean_texts:
                progress.progress(
                    min(0.30, 0.25 + 0.05 * len(vectors) / len(clean_texts)),
                    "Calculando similitud semantica...",
                )
        return vectors
    except Exception as exc:
        st.warning(f"No fue posible calcular embeddings. Se usara agrupacion por texto y titulo. Detalle: {exc}")
        return None


def classify_group(text):
    prompt = f"""
Eres un analista de reputacion y prensa de la Federacion Colombiana de Futbol (FCF).

Clasifica la noticia usando exactamente estas opciones:

TONO:
- Positivo: hay exaltacion, logro, defensa favorable, reconocimiento o gestion positiva atribuida a la FCF.
- Negativo: hay critica, acusacion, cuestionamiento, comentario negativo o afectacion reputacional contra la FCF.
- Neutro: la FCF no se menciona, o se menciona sin exaltacion ni critica.

TEMA:
- Institucional
- Torneos - Copas - Ligas
- Selecciones
- Gestión
- Jugadores
- Entorno

SUBTEMA:
- Debe ser una etiqueta breve y especifica basada en el hecho central.
- No generalices. Noticias iguales o muy similares deben poder compartir el mismo subtema.
- Si la noticia es sobre un partido de Colombia en Mundial, Copa del Mundo o Mundial 2026, evita separar artificialmente "Partido Mundial", "Partido Copa del Mundo" y "Partido Mundial 2026".
- Cuando el rival aparezca en el texto, usa el formato "Partido ante [seleccion]" para agrupar con precision. Ejemplos: "Partido ante Argentina", "Partido ante Brasil".
- Si no aparece rival claro, usa "Partido Mundial 2026" para Mundial, Copa del Mundo, Copa Mundo o Mundial 2026.
- Usa nombres concretos cuando ayuden: "Designacion arbitral", "Convocatoria Seleccion Colombia", "Foto", "Comunicado FCF", etc.

TEXTO:
{text[:3500]}

Responde unicamente JSON valido:
{{"tono":"Positivo|Negativo|Neutro","tema":"uno de los temas permitidos","subtema":"etiqueta breve"}}
""".strip()

    response = openai.ChatCompletion.create(
        model=OPENAI_MODEL_CLASIFICACION,
        messages=[{"role": "user", "content": prompt}],
        temperature=0,
        max_tokens=120,
        response_format={"type": "json_object"},
    )
    raw = response["choices"][0]["message"]["content"]
    data = json.loads(raw)
    tono = str(data.get("tono", "Neutro")).strip().title()
    tema = str(data.get("tema", "Entorno")).strip()
    subtema = str(data.get("subtema", "Sin subtema")).strip()

    if tono not in TONOS_FCF:
        tono = "Neutro"
    if tema not in TEMAS_FCF:
        tema = "Entorno"
    if not subtema:
        subtema = "Sin subtema"
    tema, subtema = normalize_subtema_fcf(text, tema, subtema)

    return {OUTPUT_TONO_COL: tono, OUTPUT_TEMA_COL: tema, OUTPUT_SUBTEMA_COL: subtema[:80]}


def process_dataframe(df, title_col, summary_col, medio_col, web_col, hyperlinks, progress):
    result = df.copy()
    result[OUTPUT_VOCERO_COL] = result[summary_col].apply(detect_vocero)
    result[OUTPUT_TONO_COL] = ""
    result[OUTPUT_TEMA_COL] = ""
    result[OUTPUT_SUBTEMA_COL] = ""

    duplicate_rows = detect_duplicate_rows(result, medio_col, title_col, web_col, hyperlinks)
    for row_idx in duplicate_rows:
        result.at[row_idx, OUTPUT_TONO_COL] = "Duplicada"
        result.at[row_idx, OUTPUT_TEMA_COL] = "-"
        result.at[row_idx, OUTPUT_SUBTEMA_COL] = "-"
        result.at[row_idx, OUTPUT_VOCERO_COL] = "-"

    active_indices = [idx for idx in result.index if idx not in duplicate_rows]
    if not active_indices:
        progress.progress(1.0, "Completado")
        return result, len(duplicate_rows)

    active_df = result.loc[active_indices].reset_index(drop=True)

    groups, texts = group_similar_news(active_df, title_col, summary_col, progress)
    total_groups = len(groups)

    for pos, group in enumerate(groups, start=1):
        progress.progress(
            min(0.95, 0.55 + 0.40 * pos / max(total_groups, 1)),
            f"Clasificando grupo {pos}/{total_groups}...",
        )

        photo_rows = [
            i for i in group
            if is_fcf_photo_summary(active_df.at[i, summary_col])
        ]

        normal_rows = [i for i in group if i not in photo_rows]
        if normal_rows:
            rep = representative_index(normal_rows, texts)
            group_partido_subtema = choose_group_partido_subtema([texts[i] for i in normal_rows])
            try:
                classification = classify_group(texts[rep])
            except Exception as exc:
                st.warning(f"No se pudo clasificar un grupo. Se marco como Neutro/Entorno. Detalle: {exc}")
                classification = {
                    OUTPUT_TONO_COL: "Neutro",
                    OUTPUT_TEMA_COL: "Entorno",
                    OUTPUT_SUBTEMA_COL: "Sin subtema",
                }
            for row_idx in normal_rows:
                original_idx = active_indices[row_idx]
                result.at[original_idx, OUTPUT_TONO_COL] = classification[OUTPUT_TONO_COL]
                result.at[original_idx, OUTPUT_TEMA_COL] = classification[OUTPUT_TEMA_COL]
                result.at[original_idx, OUTPUT_SUBTEMA_COL] = classification[OUTPUT_SUBTEMA_COL]
                if group_partido_subtema:
                    result.at[original_idx, OUTPUT_TEMA_COL] = "Selecciones"
                    result.at[original_idx, OUTPUT_SUBTEMA_COL] = group_partido_subtema
                else:
                    apply_partido_normalization_to_result(result, original_idx, texts[row_idx])

        for row_idx in photo_rows:
            original_idx = active_indices[row_idx]
            result.at[original_idx, OUTPUT_TONO_COL] = "Neutro"
            result.at[original_idx, OUTPUT_TEMA_COL] = "Institucional"
            result.at[original_idx, OUTPUT_SUBTEMA_COL] = "Foto"

    progress.progress(1.0, "Completado")
    return result, len(duplicate_rows)


def dataframe_to_excel(df, hyperlinks=None):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Analisis FCF")
        ws = writer.book["Analisis FCF"]
        header_positions = {normalize_text(cell.value): cell.column for cell in ws[1]}
        hyperlink_font = Font(color="0563C1", underline="single")

        for row_index, row_links in (hyperlinks or {}).items():
            excel_row = row_index + 2
            for column_name in ("LINK", "WEB"):
                target = row_links.get(column_name)
                column_position = header_positions.get(normalize_text(column_name))
                if target and column_position:
                    cell = ws.cell(row=excel_row, column=column_position)
                    cell.value = "Link"
                    cell.hyperlink = target
                    cell.font = hyperlink_font

        header_fill = PatternFill("solid", fgColor="1F4E79")
        header_font = Font(color="FFFFFF", bold=True)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
        for column in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in column)
            ws.column_dimensions[column[0].column_letter].width = min(max(max_len + 2, 14), 55)
    return buffer.getvalue()


def main():
    load_custom_css()
    st.markdown(
        """
        <div class="fcf-hero">
          <div class="fcf-kicker">Federacion Colombiana de Futbol</div>
          <h1 class="fcf-title">Analisis FCF</h1>
          <div class="fcf-sub">Clasificacion de Impacto, TEMA, SUBTEMA y VOCERO para noticias sobre la FCF.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.sidebar:
        st.subheader("Modelo")
        st.write(OPENAI_MODEL_CLASIFICACION)
        st.subheader("Temas prestablecidos")
        for tema in TEMAS_FCF:
            st.write(f"- {tema}")

    st.markdown('<div class="sec-label">Archivos de entrada</div>', unsafe_allow_html=True)
    st.markdown(
        """
        <div class="upload-zone">
          <div class="upload-zone-card">
            <div class="upload-zone-title">Dossier FCF</div>
            <div class="upload-zone-desc">XLSX con NOMBRE DE MEDIO, REGION, TÍTULO, RESUMEN, LINK y WEB.</div>
          </div>
          <div class="upload-zone-card">
            <div class="upload-zone-title">Configuracion local</div>
            <div class="upload-zone-desc">Se toma automaticamente de Configuracion.xlsx dentro del repo.</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    uploaded = st.file_uploader("Dossier FCF", type=["xlsx"], label_visibility="collapsed")
    if not uploaded:
        st.info("El archivo debe incluir el formato FCF con NOMBRE DE MEDIO, REGION, TÍTULO y RESUMEN.")
        return

    try:
        xlsx_bytes = uploaded.getvalue()
        hyperlinks = extract_embedded_links(xlsx_bytes)
        df = pd.read_excel(io.BytesIO(xlsx_bytes), engine="openpyxl")
    except Exception as exc:
        st.error(f"No se pudo leer el XLSX: {exc}")
        return

    title_col = get_column(df, ["TÍTULO", "Título", "Titulo"])
    summary_col = get_column(df, ["RESUMEN", "Resumen", "Resumen - Aclaracion", "Resumen - Aclaración"])
    medio_col = get_column(df, ["NOMBRE DE MEDIO", "Nombre de Medio", "Medio"])
    region_col = get_column(df, ["REGION", "Región"])
    web_col = get_column(df, ["WEB", "Web"])

    if not title_col or not summary_col or not medio_col or not region_col:
        st.error("No encontre las columnas requeridas: NOMBRE DE MEDIO, REGION, TÍTULO y RESUMEN.")
        st.write("Columnas detectadas:", list(df.columns))
        return

    missing_expected = [
        col for col in INPUT_REQUIRED_COLUMNS
        if get_column(df, [col]) is None
    ]
    if missing_expected:
        st.warning("Columnas del formato FCF no detectadas: " + ", ".join(missing_expected))

    df, lookup_message = apply_region_lookup(df, medio_col, region_col)
    st.info(lookup_message)
    st.markdown(
        f"""
        <div class="metrics-grid">
          <div class="metric-card"><div class="metric-val">{len(df)}</div><div class="metric-lbl">Filas</div></div>
          <div class="metric-card"><div class="metric-val">{df[medio_col].nunique(dropna=True)}</div><div class="metric-lbl">Medios</div></div>
          <div class="metric-card"><div class="metric-val">{df[region_col].nunique(dropna=True)}</div><div class="metric-lbl">Regiones</div></div>
          <div class="metric-card"><div class="metric-val">{len(hyperlinks)}</div><div class="metric-lbl">Filas con links</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.success(f"Archivo cargado. Texto a analizar: {title_col} + {summary_col}.")
    st.dataframe(df.head(10), use_container_width=True)

    if st.button("Analizar FCF", type="primary", use_container_width=True):
        try:
            openai.api_key = st.secrets["OPENAI_API_KEY"]
        except Exception:
            st.error("OPENAI_API_KEY no encontrada en Streamlit secrets.")
            return

        progress = st.progress(0, "Iniciando...")
        start = time.time()
        with st.spinner("Procesando archivo..."):
            output, duplicate_count = process_dataframe(
                df,
                title_col,
                summary_col,
                medio_col,
                web_col,
                hyperlinks,
                progress,
            )
        elapsed = time.time() - start

        st.session_state["fcf_output"] = output
        st.session_state["fcf_hyperlinks"] = hyperlinks
        st.session_state["fcf_duplicates"] = duplicate_count
        st.session_state["fcf_elapsed"] = elapsed

    if "fcf_output" in st.session_state:
        output = st.session_state["fcf_output"]
        st.markdown('<div class="success-banner">Analisis completado. Listo para descargar.</div>', unsafe_allow_html=True)
        st.markdown(
            f"""
            <div class="metrics-grid">
              <div class="metric-card"><div class="metric-val">{len(output)}</div><div class="metric-lbl">Total</div></div>
              <div class="metric-card"><div class="metric-val">{len(output) - st.session_state.get('fcf_duplicates', 0)}</div><div class="metric-lbl">Unicas</div></div>
              <div class="metric-card"><div class="metric-val">{st.session_state.get('fcf_duplicates', 0)}</div><div class="metric-lbl">Duplicadas</div></div>
              <div class="metric-card"><div class="metric-val">{st.session_state.get('fcf_elapsed', 0):.0f}s</div><div class="metric-lbl">Tiempo</div></div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.dataframe(output.head(50), use_container_width=True)
        st.download_button(
            "Descargar XLSX clasificado",
            data=dataframe_to_excel(output, st.session_state.get("fcf_hyperlinks")),
            file_name="Analisis_FCF.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )


if __name__ == "__main__":
    main()
