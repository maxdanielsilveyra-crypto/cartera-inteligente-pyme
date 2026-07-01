import io
import re
import unicodedata
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Cartera Inteligente PyME",
    page_icon="🎯",
    layout="wide"
)

REQUIRED_COLUMNS = [
    "CR",
    "CLIENTE",
    "ACTIVIDAD ECONÓMICA",
    "CTA PYME",
    "TPV",
    "FLOTILLA"
]

def normalize_text(value):
    if pd.isna(value):
        return ""
    text = str(value).strip()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return text.lower()

def normalize_yes_no(value):
    text = normalize_text(value)
    if text in ["si", "sí", "s", "1", "true", "x", "contratado", "con"]:
        return "Sí"
    if text in ["no", "n", "0", "false", "", "nan", "sin"]:
        return "No"
    return str(value).strip()

def load_excel(file):
    df = pd.read_excel(file)
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        st.error(f"Faltan estas columnas en el archivo: {', '.join(missing)}")
        st.stop()

    df = df[REQUIRED_COLUMNS].copy()
    df["CR"] = df["CR"].astype(str).str.strip()
    df["CLIENTE"] = df["CLIENTE"].astype(str).str.strip()
    df["ACTIVIDAD ECONÓMICA"] = df["ACTIVIDAD ECONÓMICA"].astype(str).str.strip()

    for col in ["CTA PYME", "TPV", "FLOTILLA"]:
        df[col] = df[col].apply(normalize_yes_no)

    return df

def apply_filters(df, cr_values, actividades, cta, tpv, flotilla):
    out = df.copy()

    if cr_values:
        out = out[out["CR"].isin(cr_values)]

    if actividades:
        out = out[out["ACTIVIDAD ECONÓMICA"].isin(actividades)]

    if cta != "Todos":
        out = out[out["CTA PYME"] == cta]

    if tpv != "Todos":
        out = out[out["TPV"] == tpv]

    if flotilla != "Todos":
        out = out[out["FLOTILLA"] == flotilla]

    return out

def parse_ai_prompt(prompt, df):
    text = normalize_text(prompt)
    result = {
        "cr": [],
        "actividad": [],
        "cta": "Todos",
        "tpv": "Todos",
        "flotilla": "Todos"
    }

    cr_matches = re.findall(r"\b\d{3,5}\b", text)
    available_cr = set(df["CR"].astype(str))
    result["cr"] = [cr for cr in cr_matches if cr in available_cr]

    if "sin tpv" in text or "no tpv" in text or "sin terminal" in text:
        result["tpv"] = "No"
    elif "con tpv" in text or "tengan tpv" in text or "tiene tpv" in text:
        result["tpv"] = "Sí"

    if "sin cuenta" in text or "sin cta" in text or "no cuenta pyme" in text or "sin pyme" in text:
        result["cta"] = "No"
    elif "con cuenta" in text or "con cta" in text or "cuenta pyme" in text or "cta pyme" in text:
        result["cta"] = "Sí"

    if "sin flotilla" in text or "no flotilla" in text:
        result["flotilla"] = "No"
    elif "con flotilla" in text or "flotilla" in text or "transporte" in text or "transportistas" in text:
        result["flotilla"] = "Sí"

    actividad_norm = {
        normalize_text(a): a
        for a in sorted(df["ACTIVIDAD ECONÓMICA"].dropna().unique())
    }

    keywords = {
        "restaurante": ["restaurante", "restaurantes"],
        "compraventa de calzado": ["calzado", "zapatos"],
        "usuarios menores comercio": ["comercio", "comercios"],
        "servicios profesionales": ["servicios profesionales", "profesionales"],
        "transporte": ["transporte", "transportistas", "carga", "flotillas"],
        "fotografia": ["fotografia", "fotografía"],
        "turismo": ["turismo", "agencia de turismo"],
        "contaduria": ["contaduria", "contaduría", "contadores"]
    }

    selected = []
    for key, words in keywords.items():
        if any(w in text for w in words):
            for norm_name, original in actividad_norm.items():
                if key in norm_name:
                    selected.append(original)

    result["actividad"] = sorted(set(selected))
    return result

def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Cartera objetivo")
    output.seek(0)
    return output

st.title("🎯 Cartera Inteligente PyME")
st.caption("Construye carteras objetivo filtrando por CR, actividad económica y productos contratados.")

with st.sidebar:
    st.header("1. Carga tu archivo")
    file = st.file_uploader("Sube el Excel de clientes", type=["xlsx"])

    st.markdown("---")
    st.caption("No subas el Excel a GitHub si contiene datos internos. Cárgalo aquí desde la app.")

if file is None:
    st.info("Sube tu archivo Excel para comenzar.")
    st.stop()

df = load_excel(file)

st.subheader("¿Qué cartera deseas construir hoy?")
prompt = st.text_input(
    "Describe tu segmento",
    placeholder="Ejemplo: restaurantes del CR 432 con Cuenta PyME y sin TPV"
)

parsed = None
if prompt:
    parsed = parse_ai_prompt(prompt, df)
    st.success("Interpreté tu solicitud y sugerí filtros. Puedes ajustarlos manualmente abajo.")

with st.expander("🎛️ Filtros de segmentación", expanded=True):
    col1, col2, col3, col4, col5 = st.columns(5)

    cr_options = sorted(df["CR"].dropna().unique())
    actividad_options = sorted(df["ACTIVIDAD ECONÓMICA"].dropna().unique())

    default_cr = parsed["cr"] if parsed else []
    default_actividad = parsed["actividad"] if parsed else []
    default_cta = parsed["cta"] if parsed else "Todos"
    default_tpv = parsed["tpv"] if parsed else "Todos"
    default_flotilla = parsed["flotilla"] if parsed else "Todos"

    with col1:
        cr_values = st.multiselect("CR", cr_options, default=default_cr)
    with col2:
        actividades = st.multiselect("Actividad económica", actividad_options, default=default_actividad)
    with col3:
        cta = st.selectbox("Cuenta PyME", ["Todos", "Sí", "No"], index=["Todos", "Sí", "No"].index(default_cta))
    with col4:
        tpv = st.selectbox("TPV", ["Todos", "Sí", "No"], index=["Todos", "Sí", "No"].index(default_tpv))
    with col5:
        flotilla = st.selectbox("Flotilla", ["Todos", "Sí", "No"], index=["Todos", "Sí", "No"].index(default_flotilla))

filtered = apply_filters(df, cr_values, actividades, cta, tpv, flotilla)

st.markdown("### Resultado de la segmentación")

k1, k2, k3, k4 = st.columns(4)
k1.metric("Clientes encontrados", f"{len(filtered):,}")
k2.metric("Con CTA PyME", f"{(filtered['CTA PYME'] == 'Sí').sum():,}")
k3.metric("Con TPV", f"{(filtered['TPV'] == 'Sí').sum():,}")
k4.metric("Con Flotilla", f"{(filtered['FLOTILLA'] == 'Sí').sum():,}")

st.dataframe(
    filtered,
    use_container_width=True,
    hide_index=True
)

excel_bytes = to_excel_bytes(filtered)
st.download_button(
    "⬇️ Descargar cartera filtrada en Excel",
    data=excel_bytes,
    file_name="cartera_objetivo.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

with st.expander("📊 Resumen rápido del segmento"):
    c1, c2 = st.columns(2)
    with c1:
        st.write("Clientes por CR")
        st.bar_chart(filtered["CR"].value_counts().head(15))
    with c2:
        st.write("Top actividades")
        st.bar_chart(filtered["ACTIVIDAD ECONÓMICA"].value_counts().head(15))