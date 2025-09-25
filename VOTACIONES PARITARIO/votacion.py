# -*- coding: utf-8 -*-
# pip install streamlit pandas openpyxl
import streamlit as st
import pandas as pd
from datetime import datetime
from pathlib import Path

# ===================== CONFIG =====================
DATA_DIR = Path(".")
FILE_CAND  = DATA_DIR / "CANDIDATOS.xlsx"
FILE_TOK   = DATA_DIR / "TOKEN.xlsx"
FILE_VOTOS = DATA_DIR / "VOTOS.xlsx"
MAX_SEL = 6

APP_TITLE = "Votación representantes Comité Paritario de Higiene y Seguridad Nivel Central"
st.set_page_config(page_title=APP_TITLE, page_icon="🗳️", layout="wide")
st.title(f"🗳️ {APP_TITLE}")
st.caption("Votación anónima. Se registra solo la selección realizada; el código de acceso queda marcado como usado.")

# ===================== SESSION STATE =====================
if "auth" not in st.session_state:
    st.session_state.auth = False
if "done" not in st.session_state:
    st.session_state.done = False
if "token_in" not in st.session_state:
    st.session_state.token_in = ""

# ===================== HELPERS =====================
def pick_col(df: pd.DataFrame, names):
    low = {c.strip().lower(): c for c in df.columns}
    for n in names:
        if n.strip().lower() in low:
            return low[n.strip().lower()]
    return None

def norm_bool_series(s: pd.Series):
    return s.astype(str).str.strip().str.lower().isin(["true","1","yes","si","sí","x"])

def ensure_cols(df: pd.DataFrame, needed: dict):
    for c, default in needed.items():
        if c not in df.columns:
            df[c] = default
    return df

@st.cache_data
def load_candidatos():
    df = pd.read_excel(FILE_CAND, dtype=str)

    col_label = pick_col(df, [
        "NOMRES A UTILIZAR  PARA VOTAR",
        "NOMBRES A UTILIZAR  PARA VOTAR",
        "NOMBRES A UTILIZAR PARA VOTAR",
        "NOMBRE A UTILIZAR PARA VOTAR",
        "NOMBRES"
    ])
    if col_label is None:
        c_nom = pick_col(df, ["NOMBRES"])
        c_ap1 = pick_col(df, ["PRIMER APELLIDO"])
        c_ap2 = pick_col(df, ["SEGUNDO APELLIDO"])
        df["__label_tmp__"] = (
            df.get(c_nom, "").fillna("") + " " +
            df.get(c_ap1, "").fillna("") + " " +
            df.get(c_ap2, "").fillna("")
        ).str.replace(r"\s+", " ", regex=True).str.strip()
        col_label = "__label_tmp__"

    col_div = pick_col(df, ["Dependencia/División", "Dependencia/Division", "Dependencia", "División", "Division"])

    c_run = pick_col(df, ["RUN (sin puntos)", "RUN"])
    c_dv  = pick_col(df, ["DV"])
    if c_run and c_dv:
        df["__id__"] = (df[c_run].astype(str).str.replace(".", "", regex=False).str.strip()
                        + "-" + df[c_dv].astype(str).str.strip())
    elif c_run:
        df["__id__"] = df[c_run].astype(str).str.replace(".", "", regex=False).str.strip()
    else:
        df["__id__"] = (df.index + 1).astype(str)

    col_vol = pick_col(df, ["Voluntario","VOLUNTARIO"])
    df["__vol__"] = norm_bool_series(df[col_vol].fillna("")) if col_vol else False

    df["__div__"]   = df[col_div].astype(str).str.strip() if col_div else ""
    base_label      = df[col_label].astype(str).str.strip()
    df["__label__"] = base_label.where(df["__div__"] == "", base_label + " — (" + df["__div__"] + ")")
    df.loc[df["__vol__"] == True, "__label__"] = df.loc[df["__vol__"] == True, "__label__"] + "  [Voluntario]"

    df = df[df["__label__"].str.len() > 0].copy()
    df = df.sort_values(by=["__vol__", "__label__"], ascending=[False, True])
    return df[["__id__", "__label__", "__div__", "__vol__"]].reset_index(drop=True)

@st.cache_data
def load_tokens():
    df = pd.read_excel(FILE_TOK, dtype=str)
    c_tok  = pick_col(df, ["TOKEN", "token", "Token", "código", "codigo", "Código de acceso", "Codigo de acceso"])
    c_used = pick_col(df, ["Usado", "usado", "USED"])
    c_fuso = pick_col(df, ["FechaUso", "FechaU", "fechauso", "fecha u", "fecha u."])
    c_mail = pick_col(df, ["Correo electrónico institucional", "Correo electronico institucional", "correo"])
    if c_tok is None:
        raise ValueError("TOKEN.xlsx debe tener una columna con el código (por ejemplo 'TOKEN').")
    rename = {c_tok: "token"}
    if c_used: rename[c_used] = "Usado"
    if c_fuso: rename[c_fuso] = "FechaUso"
    if c_mail: rename[c_mail] = "correo"
    df = df.rename(columns=rename)
    if "correo" in df.columns:
        df["correo"] = df["correo"].astype(str).str.strip().str.lower()
    else:
        df["correo"] = ""
    df["token"] = (df["token"].astype(str).str.strip().str.upper()
                   .str.replace(" ", "", regex=False).str.replace("-", "", regex=False))
    df = ensure_cols(df, {"Usado":"", "FechaUso":""})
    return df

def save_tokens(df):
    cols = ["token","correo","Usado","FechaUso"]
    for c in cols:
        if c not in df.columns: df[c] = ""
    df[cols].to_excel(FILE_TOK, index=False)

def load_votes():
    if FILE_VOTOS.exists():
        return pd.read_excel(FILE_VOTOS, dtype=str)
    return pd.DataFrame(columns=["Token","SeleccionadosIDs","SeleccionadosNombres","Cantidad","Fecha"])

def save_votes(df):
    df.to_excel(FILE_VOTOS, index=False)

# ===================== PANTALLA 0: YA VOTÓ =====================
if st.session_state.done:
    st.success("¡Gracias por votar! ✅ Tu respuesta fue registrada.")
    st.caption("Puedes cerrar esta ventana. Si necesitas asistencia, contacta al equipo organizador.")
    st.stop()

# ===================== PANTALLA 1: LOGIN CON CÓDIGO =====================
if not st.session_state.auth:
    st.subheader("Acceso")
    st.caption("Tu **código de acceso** te llegó por correo. Es personal y de **un solo uso**.")

    with st.form("login_form", clear_on_submit=False):
        raw = st.text_input("Ingresa tu **código de acceso**", type="password",
                            help="Puedes copiar/pegar; se ignorarán espacios y guiones.")
        submit = st.form_submit_button("Continuar", type="primary")

    if not submit:
        st.stop()

    token_in = raw.upper().replace(" ", "").replace("-", "").strip()
    if not token_in:
        st.error("Debes ingresar tu código de acceso.")
        st.stop()

    try:
        tok = load_tokens()
    except Exception as e:
        st.error(f"No se pudo cargar TOKEN.xlsx: {e}")
        st.stop()

    row = tok.loc[tok["token"] == token_in]
    if row.empty:
        st.error("Código de acceso inválido.")
        st.stop()
    if str(row.iloc[0].get("Usado","")).strip().lower() in ["true","1","yes","si","sí","x","used"]:
        st.error("Este código de acceso ya fue usado.")
        st.stop()

    # OK → guardar en sesión y pasar a la papeleta
    st.session_state.token_in = token_in
    st.session_state.auth = True
    st.success("Código de acceso válido ✅")
    st.rerun()   # 👈 ahora sí funciona

# ===================== PANTALLA 2: PAPELETA =====================
token_in = st.session_state.token_in

cand = load_candidatos()

st.subheader("Papeleta")
col1, col2, col3 = st.columns([2, 1.2, 1])
with col1:
    filtro = st.text_input("Buscar por nombre").strip().lower()
with col2:
    divisiones = ["Todas"] + sorted([d for d in cand["__div__"].unique().tolist() if d])
    sel_div = st.selectbox("División", divisiones, index=0)
with col3:
    solo_vol = st.checkbox("Mostrar sólo voluntarios", value=False)

base = cand.copy()
if filtro:
    base = base[base["__label__"].str.lower().str.contains(filtro)]
if sel_div != "Todas":
    base = base[base["__div__"] == sel_div]
if solo_vol:
    base = base[base["__vol__"] == True]

st.write(f"**Coincidencias:** {len(base)}")

df_view = base[["__label__", "__id__"]].rename(columns={"__label__": "Candidato", "__id__": "ID"})
df_view["Elegir"] = False

edited = st.data_editor(
    df_view,
    hide_index=True,
    column_config={
        "Elegir": st.column_config.CheckboxColumn(
            label="Elegir",
            help=f"Marca hasta {MAX_SEL} candidatos",
            default=False
        ),
        "Candidato": st.column_config.TextColumn(width="large"),
        "ID": st.column_config.TextColumn(width="medium"),
    },
    disabled=["Candidato", "ID"],
    use_container_width=True,
    height=420,
)

sel_df = edited[edited["Elegir"] == True]
seleccion_ids = sel_df["ID"].tolist()
label_clean = sel_df["Candidato"].tolist()

st.markdown(f"**Seleccionados:** {len(seleccion_ids)} / {MAX_SEL}")
if len(seleccion_ids) > MAX_SEL:
    st.error(f"Seleccionaste {len(seleccion_ids)}. El máximo es {MAX_SEL}. Desmarca algunos antes de continuar.")
    st.stop()

enviar = st.button("Enviar voto", type="primary", disabled=(len(seleccion_ids)==0))

if enviar:
    votos = load_votes()
    ahora = datetime.now().isoformat(timespec="seconds")
    nuevo = pd.DataFrame([{
        "Token": token_in,
        "SeleccionadosIDs": ";".join(seleccion_ids),
        "SeleccionadosNombres": "; ".join(label_clean),
        "Cantidad": str(len(seleccion_ids)),
        "Fecha": ahora,
    }])
    votos = pd.concat([votos, nuevo], ignore_index=True)
    save_votes(votos)

    tok = load_tokens()
    tok.loc[tok["token"] == token_in, "Usado"] = "True"
    tok.loc[tok["token"] == token_in, "FechaUso"] = ahora
    save_tokens(tok)

    st.session_state.done = True
    st.success("¡Voto registrado correctamente! ✅")
    st.balloons()
    st.rerun()
