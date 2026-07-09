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
SIMILARIDAD_EMBEDDING = 0.90
BATCH_EMBEDDINGS = 80


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


def load_region_lookup(config_path):
    sheets = pd.read_excel(config_path, sheet_name=None, engine="openpyxl")
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


def apply_region_lookup(df, medio_col, region_col):
    config_path = load_local_config()
    if not config_path:
        return df, "No se encontro Configuracion.xlsx; se conserva REGION del archivo cargado."

    lookup, sheet_name = load_region_lookup(config_path)
    if not lookup:
        return df, f"No se encontraron columnas NOMBRE DE MEDIO y REGION en {config_path.name}."

    result = df.copy()
    mapped = result[medio_col].apply(lambda value: lookup.get(normalize_text(value)))
    result[region_col] = mapped.fillna(result[region_col])
    found = int(mapped.notna().sum())
    return result, f"Buscarv aplicado desde {config_path.name} ({sheet_name}): {found}/{len(result)} regiones actualizadas."


def build_analysis_text(title, summary):
    title = "" if pd.isna(title) else str(title).strip()
    summary = "" if pd.isna(summary) else str(summary).strip()
    return f"TITULO: {title}\nRESUMEN: {summary}".strip()


def is_fcf_photo_summary(summary):
    text = normalize_text(summary)
    if not text:
        return False
    return bool(re.search(r"\bfoto\s*(de\s*)?fcf\b", text))


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

    return {OUTPUT_TONO_COL: tono, OUTPUT_TEMA_COL: tema, OUTPUT_SUBTEMA_COL: subtema[:80]}


def process_dataframe(df, title_col, summary_col, progress):
    result = df.copy()
    result[OUTPUT_VOCERO_COL] = result[summary_col].apply(detect_vocero)
    result[OUTPUT_TONO_COL] = ""
    result[OUTPUT_TEMA_COL] = ""
    result[OUTPUT_SUBTEMA_COL] = ""

    groups, texts = group_similar_news(result, title_col, summary_col, progress)
    total_groups = len(groups)

    for pos, group in enumerate(groups, start=1):
        progress.progress(
            min(0.95, 0.55 + 0.40 * pos / max(total_groups, 1)),
            f"Clasificando grupo {pos}/{total_groups}...",
        )

        photo_rows = [
            i for i in group
            if is_fcf_photo_summary(result.at[i, summary_col])
        ]

        normal_rows = [i for i in group if i not in photo_rows]
        if normal_rows:
            rep = representative_index(normal_rows, texts)
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
                result.at[row_idx, OUTPUT_TONO_COL] = classification[OUTPUT_TONO_COL]
                result.at[row_idx, OUTPUT_TEMA_COL] = classification[OUTPUT_TEMA_COL]
                result.at[row_idx, OUTPUT_SUBTEMA_COL] = classification[OUTPUT_SUBTEMA_COL]

        for row_idx in photo_rows:
            result.at[row_idx, OUTPUT_TONO_COL] = "Neutro"
            result.at[row_idx, OUTPUT_TEMA_COL] = "Institucional"
            result.at[row_idx, OUTPUT_SUBTEMA_COL] = "Foto"

    progress.progress(1.0, "Completado")
    return result


def dataframe_to_excel(df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Analisis FCF")
        ws = writer.book["Analisis FCF"]
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
    st.title("Analisis FCF")
    st.caption("Clasificacion de Impacto, TEMA, SUBTEMA y VOCERO para noticias sobre la Federacion Colombiana de Futbol.")

    with st.sidebar:
        st.subheader("Modelo")
        st.write(OPENAI_MODEL_CLASIFICACION)
        st.subheader("Temas prestablecidos")
        for tema in TEMAS_FCF:
            st.write(f"- {tema}")

    uploaded = st.file_uploader("Sube un archivo XLSX", type=["xlsx"])
    if not uploaded:
        st.info("El archivo debe incluir el formato FCF con NOMBRE DE MEDIO, REGION, TÍTULO y RESUMEN.")
        return

    try:
        df = pd.read_excel(uploaded, engine="openpyxl")
    except Exception as exc:
        st.error(f"No se pudo leer el XLSX: {exc}")
        return

    title_col = get_column(df, ["TÍTULO", "Título", "Titulo"])
    summary_col = get_column(df, ["RESUMEN", "Resumen", "Resumen - Aclaracion", "Resumen - Aclaración"])
    medio_col = get_column(df, ["NOMBRE DE MEDIO", "Nombre de Medio", "Medio"])
    region_col = get_column(df, ["REGION", "Región"])

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
    st.success(
        f"Archivo cargado: {len(df)} filas. Medio: {medio_col}. Region: {region_col}. "
        f"Titulo: {title_col}. Resumen: {summary_col}."
    )
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
            output = process_dataframe(df, title_col, summary_col, progress)
        elapsed = time.time() - start

        st.session_state["fcf_output"] = output
        st.session_state["fcf_elapsed"] = elapsed

    if "fcf_output" in st.session_state:
        output = st.session_state["fcf_output"]
        st.subheader("Resultado")
        st.caption(f"Tiempo de procesamiento: {st.session_state.get('fcf_elapsed', 0):.0f}s")
        st.dataframe(output.head(50), use_container_width=True)
        st.download_button(
            "Descargar XLSX clasificado",
            data=dataframe_to_excel(output),
            file_name="Analisis_FCF.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )


if __name__ == "__main__":
    main()
