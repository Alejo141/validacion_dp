"""
Aplicación Streamlit para determinar usuarios a facturar
DISPOWER S.A.S. E.S.P. - Sistema ZNI/SISFV
"""

import streamlit as st
import pandas as pd
import io
from typing import Optional

# ─── Configuración de página ──────────────────────────────────────────────────

st.set_page_config(
    page_title="Validador · DISPOWER",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── Estilos ──────────────────────────────────────────────────────────────────

st.markdown("""
<style>
  /* Paleta */
  :root {
    --solar-gold:   #F5A623;
    --deep-navy:    #0D2540;
    --mid-slate:    #1E3A5F;
    --light-slate:  #2E547A;
    --off-white:    #F0F4F8;
    --green-ok:     #2ECC71;
    --red-no:       #E74C3C;
    --text-main:    #E8EEF4;
    --text-muted:   #8FA8C0;
  }

  /* Fondo global */
  .stApp { background-color: var(--deep-navy); color: var(--text-main); }

  /* Sidebar */
  [data-testid="stSidebar"] {
    background-color: var(--mid-slate);
    border-right: 2px solid var(--solar-gold);
  }
  [data-testid="stSidebar"] * { color: var(--text-main) !important; }

  /* Métricas */
  [data-testid="stMetric"] {
    background: var(--mid-slate);
    border-radius: 8px;
    padding: 1rem 1.2rem;
    border-left: 3px solid var(--solar-gold);
  }
  [data-testid="stMetricLabel"] { color: var(--text-muted) !important; font-size: .8rem; }
  [data-testid="stMetricValue"] { color: var(--solar-gold) !important; font-size: 1.8rem; }

  /* Encabezado */
  .app-header {
    background: linear-gradient(135deg, var(--mid-slate) 0%, var(--light-slate) 100%);
    border-bottom: 3px solid var(--solar-gold);
    padding: 1.2rem 1.8rem;
    border-radius: 10px;
    margin-bottom: 1.5rem;
  }
  .app-header h1 { color: var(--solar-gold); margin: 0; font-size: 1.6rem; letter-spacing: .5px; }
  .app-header p  { color: var(--text-muted); margin: .3rem 0 0; font-size: .85rem; }

  /* Tarjetas de estado */
  .card-no  { background: rgba(231,76,60,.15);  border:1px solid var(--red-no);   border-radius:8px; padding:.6rem 1rem; }
  .card-si  { background: rgba(46,204,113,.15); border:1px solid var(--green-ok); border-radius:8px; padding:.6rem 1rem; }
  .card-no span, .card-si span { font-weight:700; font-size:.95rem; }

  /* Tabla */
  [data-testid="stDataFrame"] { border: 1px solid var(--light-slate); border-radius: 8px; }

  /* Divider */
  hr { border-color: var(--light-slate); }

  /* Botón descarga */
  .stDownloadButton > button {
    background-color: var(--solar-gold) !important;
    color: var(--deep-navy) !important;
    font-weight: 700;
    border: none;
    border-radius: 6px;
  }
</style>
""", unsafe_allow_html=True)

# ─── Cabecera ─────────────────────────────────────────────────────────────────

st.markdown("""
<div class="app-header">
  <h1>⚡ Motor de Facturación · ZNI / SISFV</h1>
  <p>Automatización de base para facturación — DISPOWER S.A.S. E.S.P.</p>
</div>
""", unsafe_allow_html=True)

# ═════════════════════════════════════════════════════════════════════════════
# FUNCIONES DE PROCESAMIENTO
# ═════════════════════════════════════════════════════════════════════════════

def leer_archivo(archivo) -> Optional[pd.DataFrame]:
    """Lee un archivo Excel o CSV cargado por el usuario."""
    try:
        nombre = archivo.name.lower()
        if nombre.endswith(".csv"):
            return pd.read_csv(archivo, dtype=str)
        elif nombre.endswith((".xlsx", ".xls")):
            return pd.read_excel(archivo, dtype=str)
        else:
            st.error(f"Formato no soportado: {archivo.name}")
            return None
    except Exception as e:
        st.error(f"Error al leer {archivo.name}: {e}")
        return None


def limpiar_nui(serie: pd.Series) -> pd.Series:
    """Convierte NUI a texto limpio sin espacios ni nulos."""
    return serie.astype(str).str.strip().str.upper().replace("NAN", pd.NA)


def detectar_duplicados(df: pd.DataFrame, campo: str, nombre_base: str) -> int:
    """Informa duplicados en un campo dado y retorna la cantidad."""
    total = df[campo].duplicated().sum()
    if total > 0:
        st.warning(f"⚠️ **{nombre_base}**: {total} NUI duplicados detectados — se consolidará por NUI.")
    return int(total)


def consolidar_sac(df_sac: pd.DataFrame) -> pd.DataFrame:
    """
    Consolida la base SAC por NUI aplicando la lógica de prioridad:
    1. Si existe algún ticket Abierto con Submenu3 = 'Bloquea facturación' → BLOQUEA
    2. Si existe algún ticket Abierto con Submenu3 = 'Descuento comercial'  → DESCUENTO
    3. Si no hay tickets abiertos relevantes → SIN_NOVEDAD_SAC
    """
    # Normalizar columnas clave
    for col in ["NUI", "Submenu3"]:
        if col in df_sac.columns:
            df_sac[col] = df_sac[col].astype(str).str.strip()

    # Normalizar campos de estado y semáforo
    col_estado   = detectar_columna(df_sac, ["estado", "estado del ticket", "estadoticket"])
    col_semaforo = detectar_columna(df_sac, ["semaforo", "semáforo"])
    col_submenu  = detectar_columna(df_sac, ["submenu3", "sub menu 3", "submenu_3"])

    if not col_estado:
        st.error("❌ No se encontró columna de Estado en la base SAC.")
        return pd.DataFrame()

    df_sac["_estado"]   = df_sac[col_estado].str.strip().str.lower()
    df_sac["_submenu3"] = df_sac[col_submenu].str.strip().str.lower() if col_submenu else ""

    ESTADOS_ABIERTOS = ["abierto", "open", "activo"]
    BLOQUEA          = "bloquea facturación"
    DESCUENTO        = "descuento comercial"

    resultados = []
    for nui, grupo in df_sac.groupby("NUI"):
        abiertos = grupo[grupo["_estado"].isin(ESTADOS_ABIERTOS)]

        bloquea   = abiertos[abiertos["_submenu3"].str.contains(BLOQUEA, na=False)]
        descuento = abiertos[abiertos["_submenu3"].str.contains(DESCUENTO, na=False)]

        if not bloquea.empty:
            resultados.append({"NUI": nui, "_sac_decision": "BLOQUEA"})
        elif not descuento.empty:
            resultados.append({"NUI": nui, "_sac_decision": "DESCUENTO"})
        elif not abiertos.empty:
            # Tickets abiertos sin Submenu3 relevante → no bloquea facturación
            resultados.append({"NUI": nui, "_sac_decision": "ABIERTO_SIN_REGLA"})
        else:
            resultados.append({"NUI": nui, "_sac_decision": "SIN_NOVEDAD_SAC"})

    return pd.DataFrame(resultados)


def consolidar_hurtos(df_hurtos: pd.DataFrame) -> pd.Series:
    """Retorna un conjunto de NUI únicos presentes en la base de hurtos."""
    col_nui = detectar_columna(df_hurtos, ["nui"])
    if not col_nui:
        st.error("❌ No se encontró columna NUI en la base de Hurtos.")
        return pd.Index([])
    return set(df_hurtos[col_nui].astype(str).str.strip().str.upper().dropna().unique())


def detectar_columna(df: pd.DataFrame, variantes: list) -> Optional[str]:
    """Busca una columna en el DataFrame usando una lista de nombres posibles (case-insensitive)."""
    cols_lower = {c.lower().strip(): c for c in df.columns}
    for v in variantes:
        if v.lower() in cols_lower:
            return cols_lower[v.lower()]
    return None


def aplicar_reglas(
    df_usuarios: pd.DataFrame,
    df_sac_consolidado: pd.DataFrame,
    nui_hurtos: set,
) -> pd.DataFrame:
    """
    Aplica las 4 reglas de negocio en orden de prioridad por NUI.
    Retorna tabla con columnas: NUI, Estado de Facturación, Motivo, Fuente de decisión.
    """
    col_nui_usr = detectar_columna(df_usuarios, ["nui"])
    if not col_nui_usr:
        st.error("❌ No se encontró columna NUI en la base de Usuarios.")
        return pd.DataFrame()

    df_usuarios["NUI"] = limpiar_nui(df_usuarios[col_nui_usr])
    df_usuarios = df_usuarios.dropna(subset=["NUI"])
    df_usuarios = df_usuarios.drop_duplicates(subset=["NUI"])

    # Merge con SAC consolidado
    df = df_usuarios[["NUI"]].merge(df_sac_consolidado, on="NUI", how="left")
    df["_sac_decision"] = df["_sac_decision"].fillna("SIN_REGISTRO_SAC")

    filas = []
    for _, row in df.iterrows():
        nui      = row["NUI"]
        decision = row["_sac_decision"]

        # ── Regla 1: No facturar — bloquea facturación ──
        if decision == "BLOQUEA":
            filas.append({
                "NUI": nui,
                "Estado de Facturación": "No facturar",
                "Motivo": "Ticket abierto - Bloquea facturación",
                "Fuente de decisión": "SAC",
            })

        # ── Regla 2: Sí facturar — descuento comercial (prioridad sobre hurtos) ──
        elif decision == "DESCUENTO":
            filas.append({
                "NUI": nui,
                "Estado de Facturación": "Sí facturar",
                "Motivo": "Ticket abierto - Descuento comercial",
                "Fuente de decisión": "SAC",
            })

        # ── Reglas 3 y 4: Sin ticket SAC bloqueante → revisar hurtos ──
        else:
            if nui in nui_hurtos:
                # Regla 3: No facturar — hurto
                filas.append({
                    "NUI": nui,
                    "Estado de Facturación": "No facturar",
                    "Motivo": "Usuario reportado en hurtos",
                    "Fuente de decisión": "Hurtos",
                })
            else:
                # Regla 4: Facturación normal
                filas.append({
                    "NUI": nui,
                    "Estado de Facturación": "Sí facturar",
                    "Motivo": "Sin novedades",
                    "Fuente de decisión": "Sin coincidencias",
                })

    return pd.DataFrame(filas)


def exportar_excel(df_resultado: pd.DataFrame) -> bytes:
    """Genera el archivo Excel con dos hojas: NO_FACTURAR y FACTURAR."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        no_facturar = df_resultado[df_resultado["Estado de Facturación"] == "No facturar"].reset_index(drop=True)
        facturar    = df_resultado[df_resultado["Estado de Facturación"] == "Sí facturar"].reset_index(drop=True)

        no_facturar.to_excel(writer, sheet_name="NO_FACTURAR", index=False)
        facturar.to_excel(writer, sheet_name="FACTURAR", index=False)

        # Autoajuste de columnas
        for sheet_name, df_sheet in [("NO_FACTURAR", no_facturar), ("FACTURAR", facturar)]:
            ws = writer.sheets[sheet_name]
            for col in ws.columns:
                max_len = max((len(str(cell.value or "")) for cell in col), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)

    return buffer.getvalue()


# ═════════════════════════════════════════════════════════════════════════════
# SIDEBAR — CARGA DE ARCHIVOS
# ═════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## 📂 Cargar bases de datos")
    st.markdown("---")

    st.markdown("#### 👥 Base de Usuarios")
    archivo_usuarios = st.file_uploader("NUI únicos de usuarios", type=["xlsx", "xls", "csv"], key="usuarios")

    st.markdown("#### 🎫 Base SAC")
    archivo_sac = st.file_uploader("Tickets de atención al cliente", type=["xlsx", "xls", "csv"], key="sac")

    st.markdown("#### 🔒 Base Hurtos")
    archivo_hurtos = st.file_uploader("NUI reportados por hurto", type=["xlsx", "xls", "csv"], key="hurtos")

    st.markdown("---")
    procesar = st.button("⚡ Procesar Facturación", use_container_width=True, type="primary")

# ═════════════════════════════════════════════════════════════════════════════
# LÓGICA PRINCIPAL
# ═════════════════════════════════════════════════════════════════════════════

if procesar:
    if not archivo_usuarios:
        st.error("❌ Debes cargar al menos la **Base de Usuarios** para continuar.")
        st.stop()

    with st.spinner("Procesando bases de datos…"):

        # Lectura
        df_usuarios = leer_archivo(archivo_usuarios)
        df_sac      = leer_archivo(archivo_sac) if archivo_sac else None
        df_hurtos   = leer_archivo(archivo_hurtos) if archivo_hurtos else None

        if df_usuarios is None:
            st.error("Error crítico al leer la base de usuarios.")
            st.stop()

        # Validaciones y consolidación SAC
        df_sac_consolidado = pd.DataFrame(columns=["NUI", "_sac_decision"])
        if df_sac is not None:
            col_nui_sac = detectar_columna(df_sac, ["nui"])
            if col_nui_sac:
                df_sac["NUI"] = limpiar_nui(df_sac[col_nui_sac])
                df_sac = df_sac.dropna(subset=["NUI"])
                detectar_duplicados(df_sac, "NUI", "SAC")
                df_sac_consolidado = consolidar_sac(df_sac)
            else:
                st.warning("⚠️ Base SAC sin columna NUI reconocible — se omitirá.")

        # Consolidación Hurtos
        nui_hurtos = set()
        if df_hurtos is not None:
            col_nui_h = detectar_columna(df_hurtos, ["nui"])
            if col_nui_h:
                df_hurtos["NUI"] = limpiar_nui(df_hurtos[col_nui_h])
                df_hurtos = df_hurtos.dropna(subset=["NUI"])
                detectar_duplicados(df_hurtos, "NUI", "Hurtos")
                nui_hurtos = consolidar_hurtos(df_hurtos)
            else:
                st.warning("⚠️ Base Hurtos sin columna NUI reconocible — se omitirá.")

        # Aplicar reglas
        df_resultado = aplicar_reglas(df_usuarios, df_sac_consolidado, nui_hurtos)

        if df_resultado.empty:
            st.error("No se pudo generar el resultado. Revisa el formato de los archivos.")
            st.stop()

    # Guardar en sesión
    st.session_state["df_resultado"] = df_resultado
    st.success("✅ Procesamiento completado.")

# ─── Resultados ───────────────────────────────────────────────────────────────

if "df_resultado" in st.session_state:
    df_resultado = st.session_state["df_resultado"]

    total        = len(df_resultado)
    facturar_n   = (df_resultado["Estado de Facturación"] == "Sí facturar").sum()
    no_facturar_n= (df_resultado["Estado de Facturación"] == "No facturar").sum()
    pct_si       = facturar_n / total * 100 if total else 0
    pct_no       = no_facturar_n / total * 100 if total else 0

    # Métricas
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total usuarios",      f"{total:,}")
    c2.metric("✅ Sí facturar",       f"{facturar_n:,}")
    c3.metric("🚫 No facturar",       f"{no_facturar_n:,}")
    c4.metric("% Facturados",         f"{pct_si:.1f}%")
    c5.metric("% No facturados",      f"{pct_no:.1f}%")

    st.markdown("---")

    # Filtros y buscador
    col_f1, col_f2 = st.columns([2, 3])
    with col_f1:
        filtro_estado = st.selectbox(
            "Filtrar por estado",
            ["Todos", "Sí facturar", "No facturar"],
        )
    with col_f2:
        buscar_nui = st.text_input("🔍 Buscar por NUI", placeholder="Ingresa un NUI…")

    df_vista = df_resultado.copy()
    if filtro_estado != "Todos":
        df_vista = df_vista[df_vista["Estado de Facturación"] == filtro_estado]
    if buscar_nui.strip():
        df_vista = df_vista[df_vista["NUI"].str.contains(buscar_nui.strip().upper(), na=False)]

    # Colorear la columna de estado
    def colorear_estado(val):
        if val == "Sí facturar":
            return "background-color: rgba(46,204,113,.25); color: #2ECC71; font-weight:700"
        elif val == "No facturar":
            return "background-color: rgba(231,76,60,.25); color: #E74C3C; font-weight:700"
        return ""

    st.markdown(f"**{len(df_vista):,} registros** mostrados")
    st.dataframe(
        df_vista.style.applymap(colorear_estado, subset=["Estado de Facturación"]),
        use_container_width=True,
        height=420,
    )

    # Descarga
    st.markdown("---")
    excel_bytes = exportar_excel(df_resultado)
    st.download_button(
        label="⬇️ Descargar Resultado_Facturacion.xlsx",
        data=excel_bytes,
        file_name="Resultado_Facturacion.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

else:
    # Estado vacío
    st.markdown("""
    <div style="text-align:center; padding:4rem 2rem; color:#8FA8C0;">
        <div style="font-size:3.5rem;">⚡</div>
        <h3 style="color:#F5A623; margin-top:1rem;">Listo para procesar</h3>
        <p>Carga las bases de datos en el panel izquierdo y presiona <strong>Procesar Facturación</strong>.</p>
        <p style="font-size:.8rem; margin-top:1.5rem;">
            Base de Usuarios (requerida) · Base SAC (opcional) · Base Hurtos (opcional)
        </p>
    </div>
    """, unsafe_allow_html=True)
