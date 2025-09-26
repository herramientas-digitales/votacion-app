# -*- coding: utf-8 -*-
# pip install streamlit pandas openpyxl msal requests
import streamlit as st
import pandas as pd
from datetime import datetime, timezone
from pathlib import Path
import unicodedata, json, requests
import msal

# ===================== CONFIG =====================
DATA_DIR = Path(".")  # CANDIDATOS.xlsx desde el repo
FILE_CAND  = DATA_DIR / "CANDIDATOS.xlsx"
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
if "selected_ids" not in st.session_state:
    st.session_state.selected_ids = set()

# ===================== UTILIDADES =====================
def _strip_accents(s: str) -> str:
    if s is None: return ""
    s = unicodedata.normalize("NFD", str(s))
    return "".join(ch for ch in s if unicodedata.category(ch) != "Mn")

def _normalize_token(s: str) -> str:
    return str(s or "").upper().replace(" ", "").replace("-", "").strip()

def pick_col(df: pd.DataFrame, names):
    low = {c.strip().lower(): c for c in df.columns}
    for n in names:
        if n.strip().lower() in low:
            return low[n.strip().lower()]
    return None

# ===================== MICROSOFT GRAPH =====================
# Config esperada en Streamlit Secrets:
# TENANT_ID, CLIENT_ID, CLIENT_SECRET
# GRAPH_SITE_ID, GRAPH_LIST_TOKENS_ID, GRAPH_LIST_VOTOS_ID
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_BASE  = "https://graph.microsoft.com/v1.0"

@st.cache_resource(show_spinner=False)
def _graph_token():
    app = msal.ConfidentialClientApplication(
        client_id=st.secrets["CLIENT_ID"],
        authority=f"https://login.microsoftonline.com/{st.secrets['TENANT_ID']}",
        client_credential=st.secrets["CLIENT_SECRET"],
    )
    result = app.acquire_token_silent(GRAPH_SCOPE, account=None)
    if not result:
        result = app.acquire_token_for_client(scopes=GRAPH_SCOPE)
    if "access_token" not in result:
        raise RuntimeError(f"Graph auth failed: {result}")
    return result["access_token"]

def _gheaders():
    return {"Authorization": f"Bearer {_graph_token()}","Content-Type":"application/json"}

def _sp_list_items(list_id: str, params: dict):
    url = f"{GRAPH_BASE}/sites/{st.secrets['GRAPH_SITE_ID']}/lists/{list_id}/items"
    r = requests.get(url, headers=_gheaders(), params=params, timeout=30)
    r.raise_for_status()
    return r.json()

def _sp_get_item(list_id: str, item_id: str):
    url = f"{GRAPH_BASE}/sites/{st.secrets['GRAPH_SITE_ID']}/lists/{list_id}/items/{item_id}"
    r = requests.get(url, headers=_gheaders(), timeout=30)
    r.raise_for_status()
    return r.json()

def _sp_patch_item_fields(list_id: str, item_id: str, fields: dict):
    url = f"{GRAPH_BASE}/sites/{st.secrets['GRAPH_SITE_ID']}/lists/{list_id}/items/{item_id}/fields"
    r = requests.patch(url, headers=_gheaders(), data=json.dumps(fields), timeout=30)
    r.raise_for_status()
    return True

def _sp_add_item(list_id: str, fields: dict):
    url = f"{GRAPH_BASE}/sites/{st.secrets['GRAPH_SITE_ID']}/lists/{list_id}/items"
    body = {"fields": fields}
    r = requests.post(url, headers=_gheaders(), data=json.dumps(body), timeout=30)
    r.raise_for_status()
    return r.json()

# ------- Lógica de tokens/votos sobre Microsoft Lists -------
def token_get_item(token_norm: str):
    # Filtra por Title == token (usa columna "Title" para guardar el token)
    # Si tu lista usa otro campo para token, cambia fields/Title por fields/<NombreCampo>
    params = {
        "$filter": f"fields/Title eq '{token_norm}'",
        "$select": "id,fields",
        "$top": "1",
    }
    data = _sp_list_items(st.secrets["GRAPH_LIST_TOKENS_ID"], params)
    items = data.get("value", [])
    return items[0] if items else None

def token_is_used(token_norm: str) -> bool:
    item = token_get_item(token_norm)
    if not item:  # token no existe => inválido
        return False
    fields = item.get("fields", {})
    val = str(fields.get("Usado", "")).strip().lower()
    return val in ["true","1","yes","si","sí","x","used"]

def mark_token_used(token_norm: str, fecha_iso: str):
    item = token_get_item(token_norm)
    if not item:
        raise ValueError("Token no encontrado en la lista TOKEN.")
    item_id = item["id"]
    return _sp_patch_item_fields(
        st.secrets["GRAPH_LIST_TOKENS_ID"],
        item_id,
        {"Usado": True, "FechaUso": fecha_iso}
    )

def append_vote(token_norm: str, seleccion_ids, seleccion_nombres, cantidad: int, fecha_iso: str):
    fields = {
        # Puedes renombrar estos campos en tu lista VOTOS; ajusta los nombres aquí
        "Token": token_norm,
        "SeleccionadosIDs": ";".join(seleccion_ids),
        "SeleccionadosNombres": "; ".join(seleccion_nombres),
        "Cantidad": cantidad,
        "Fecha": fecha_iso,
    }
    return _sp_add_item(st.secrets["GRAPH_LIST_VOTOS_ID"], fields)

# ===================== CANDIDATOS desde Excel =====================
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

    # Voluntario (columna "Voluntario" o última)
    col_vol = pick_col(df, ["Voluntario","VOLUNTARIO"])
    if col_vol is None:
        col_vol = df.columns[-1]
    t = df[col_vol].astype(str).map(_strip_accents).str.strip().str.lower()
    df["__vol__"] = t.isin(["true","1","yes","si","x","sí","sí","si"])

    df["__div__"]   = df[col_div].astype(str).str.strip() if col_div else ""
    base_label      = df[col_label].astype(str).str.strip()
    df["__label__"] = base_label.where(df["__div__"] == "", base_label + " — (" + df["__div__"] + ")")
    df.loc[df["__vol__"] == True, "__label__"] = df.loc[df["__vol__"] == True, "__label__"] + "  [Voluntario]"

    df = df[df["__label__"].str.len() > 0].copy()
    df = df.sort_values(by=["__vol__", "__label__"], ascending=[False, True])
    return df[["__id__", "__label__", "__div__", "__vol__"]].reset_index(drop=True)

# ===================== PANTALLAS =====================
# 0: ya votó
if st.session_state.done:
    st.success("¡Gracias por votar! ✅ Tu respuesta fue registrada.")
    st.caption("Puedes cerrar esta ventana. Si necesitas asistencia, contacta al equipo organizador.")
    st.stop()

# 1: login
if not st.session_state.auth:
    st.subheader("Acceso")
    st.caption("Tu **código de acceso** te llegó por correo. Es personal y de **un solo uso**.")

    with st.form("login_form", clear_on_submit=False):
        raw = st.text_input("Ingresa tu **código de acceso**", type="password",
                            help="Puedes copiar/pegar; se ignorarán espacios y guiones.")
        submit = st.form_submit_button("Continuar", type="primary")

    if not submit:
        st.stop()

    token_in = _normalize_token(raw)
    if not token_in:
        st.error("Debes ingresar tu código de acceso.")
        st.stop()

    item = token_get_item(token_in)
    if not item:
        st.error("Código de acceso inválido.")
        st.stop()
    if token_is_used(token_in):
        st.error("Este código de acceso ya fue usado.")
        st.stop()

    st.session_state.token_in = token_in
    st.session_state.auth = True
    st.success("Código de acceso válido ✅")
    st.rerun()

# 2: papeleta
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

label_map = dict(zip(cand["__id__"].astype(str), cand["__label__"]))
df_view = base[["__label__", "__id__"]].rename(columns={"__label__":"Candidato","__id__":"ID"}).copy()
df_view["ID"] = df_view["ID"].astype(str)
df_view["Elegir"] = df_view["ID"].apply(lambda x: x in st.session_state.selected_ids)

edited = st.data_editor(
    df_view,
    key="grid_candidatos",
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

edited["ID"] = edited["ID"].astype(str)
antes = set(st.session_state.selected_ids)
ids_en_vista = set(edited["ID"].tolist())
marcados_en_vista = set(edited.loc[edited["Elegir"] == True, "ID"].tolist())
st.session_state.selected_ids -= (ids_en_vista - marcados_en_vista)
st.session_state.selected_ids |= marcados_en_vista
st.session_state.selected_ids &= set(cand["__id__"].astype(str).tolist())
if st.session_state.selected_ids != antes:
    st.rerun()

seleccion_ids = list(st.session_state.selected_ids)
label_clean   = [label_map.get(x, x) for x in seleccion_ids]

st.markdown(f"**Seleccionados:** {len(seleccion_ids)} / {MAX_SEL}")
if len(seleccion_ids) > MAX_SEL:
    st.error(f"Seleccionaste {len(seleccion_ids)}. El máximo es {MAX_SEL}. Desmarca algunos antes de continuar.")
    st.stop()

enviar = st.button("Enviar voto", type="primary", disabled=(len(seleccion_ids)==0))

if enviar:
    # doble verificación anti-replay
    if token_is_used(token_in):
        st.error("Este código de acceso ya fue usado.")
        st.stop()

    ahora = datetime.now(timezone.utc).isoformat(timespec="seconds")

    # 1) guardar voto
    append_vote(token_in, seleccion_ids, label_clean, len(seleccion_ids), ahora)
    # 2) marcar token usado
    mark_token_used(token_in, ahora)

    st.session_state.done = True
    st.success("¡Voto registrado correctamente! ✅")
    st.balloons()
    st.rerun()

