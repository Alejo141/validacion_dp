"""
Motor de Facturación — DISPOWER S.A.S. E.S.P.
Sistema ZNI / SISFV

Estructura esperada de archivos
────────────────────────────────
• db_usuarios.xlsx
    Columna requerida: NUI

• db_sac.xlsx  (o cualquier nombre — es la base SAC con tickets)
    Columnas requeridas: NUI, Semaforo, SubMenu3
    Semaforo: Crítico | Moderado | Leve = ticket ABIERTO
    Semaforo: Cerrado                   = ticket CERRADO

• db_hurtos_sac.xlsx  (base de hurtos)
    Columnas requeridas: NUI, SubMenu3
    Solo se consideran hurtos los NUI con SubMenu3 = BLOQUEA FACTURACION
"""

import io
import unicodedata
import streamlit as st
import pandas as pd

# ─── Configuración ────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Motor de Facturación · DISPOWER",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
  :root {
    --solar-gold:  #F5A623; --deep-navy: #0D2540;
    --mid-slate:   #1E3A5F; --light-slate: #2E547A;
    --green-ok:    #2ECC71; --red-no: #E74C3C;
    --text-main:   #E8EEF4; --text-muted: #8FA8C0;
  }
  .stApp { background-color: var(--deep-navy); color: var(--text-main); }
  [data-testid="stSidebar"] { background-color: var(--mid-slate); border-right: 2px solid var(--solar-gold); }
  [data-testid="stSidebar"] * { color: var(--text-main) !important; }
  [data-testid="stMetric"] { background: var(--mid-slate); border-radius: 8px; padding: 1rem 1.2rem; border-left: 3px solid var(--solar-gold); }
  [data-testid="stMetricLabel"] { color: var(--text-muted) !important; font-size: .8rem; }
  [data-testid="stMetricValue"] { color: var(--solar-gold) !important; font-size: 1.8rem; }
  .app-header { background: linear-gradient(135deg,var(--mid-slate) 0%,var(--light-slate) 100%); border-bottom: 3px solid var(--solar-gold); padding: 1.2rem 1.8rem; border-radius: 10px; margin-bottom: 1.5rem; }
  .app-header h1 { color: var(--solar-gold); margin:0; font-size:1.6rem; }
  .app-header p  { color: var(--text-muted); margin:.3rem 0 0; font-size:.85rem; }
  [data-testid="stDataFrame"] { border: 1px solid var(--light-slate); border-radius: 8px; }
  hr { border-color: var(--light-slate); }
  .stDownloadButton > button { background-color: var(--solar-gold) !important; color: var(--deep-navy) !important; font-weight: 700; border: none; border-radius: 6px; }
  .warn-box { background: rgba(245,166,35,.12); border: 1px solid var(--solar-gold); border-radius: 8px; padding: .7rem 1rem; margin-bottom:.5rem; font-size:.85rem; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="app-header">
  <h1>⚡ Motor de Facturación · ZNI / SISFV</h1>
  <p>Clasificación automática de usuarios para facturación — DISPOWER S.A.S. E.S.P.</p>
</div>
""", unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════════════════════
# CONSTANTES
# ═════════════════════════════════════════════════════════════════════════════

COL_NUI      = "NUI"
COL_SEMAFORO = "Semaforo"
COL_SUBMENU3 = "SubMenu3"

# Semáforos que indican ticket ABIERTO (comparación sin tildes + mayúsculas)
SEMAFOROS_ABIERTOS = {"CRITICO", "MODERADO", "LEVE"}

# Valores de SubMenu3 (sin tildes + mayúsculas)
SUBMENU_BLOQUEA   = "BLOQUEA FACTURACION"
SUBMENU_DESCUENTO = "DESCUENTO COMERCIAL"


# ═════════════════════════════════════════════════════════════════════════════
# FUNCIONES DE UTILIDAD
# ═════════════════════════════════════════════════════════════════════════════

def quitar_tildes(texto: str) -> str:
    """Elimina diacríticos. 'Crítico' → 'Critico'"""
    return "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )


def normalizar_nui(serie: pd.Series) -> pd.Series:
    """NUI → texto MAYÚSCULAS sin espacios; NaN para nulos."""
    return serie.astype(str).str.strip().str.upper().replace("NAN", pd.NA)


def normalizar_texto(serie: pd.Series) -> pd.Series:
    """Texto → sin tildes + MAYÚSCULAS. 'Crítico' → 'CRITICO'"""
    return (
        serie.fillna("").astype(str)
        .str.strip()
        .apply(quitar_tildes)
        .str.upper()
    )


def leer_excel(archivo) -> pd.DataFrame | None:
    """Lee un archivo Excel; retorna None y muestra error si falla."""
    try:
        return pd.read_excel(archivo, dtype=str)
    except Exception as e:
        st.error(f"❌ Error al leer **{archivo.name}**: {e}")
        return None


def validar_columnas(df: pd.DataFrame, requeridas: list, nombre: str) -> bool:
    """Verifica que el DataFrame tenga las columnas esperadas."""
    faltantes = [c for c in requeridas if c not in df.columns]
    if faltantes:
        st.error(
            f"❌ **{nombre}** — columnas faltantes: `{faltantes}`\n\n"
            f"Columnas encontradas: `{list(df.columns)}`"
        )
        return False
    return True


# ═════════════════════════════════════════════════════════════════════════════
# VALIDACIONES Y REPORTES DE CALIDAD
# ═════════════════════════════════════════════════════════════════════════════

def reportar_calidad_hurtos(df_h: pd.DataFrame) -> dict:
    """
    Analiza y reporta la calidad de la base de hurtos.
    Retorna diccionario con métricas de calidad.
    """
    total_registros = len(df_h)
    nui_unicos      = df_h[COL_NUI].nunique()
    dup_internos    = df_h[df_h.duplicated(subset=[COL_NUI], keep=False)][COL_NUI].nunique()

    # SubMenu3: solo BLOQUEA son hurtos efectivos
    bloquea_nui  = set(df_h[df_h[COL_SUBMENU3] == SUBMENU_BLOQUEA][COL_NUI].dropna())
    nobloquea_nui= set(df_h[df_h[COL_SUBMENU3] != SUBMENU_BLOQUEA][COL_NUI].dropna())
    # NUI con SOLO NO BLOQUEA (nunca aparecen como BLOQUEA)
    solo_nobloquea = nobloquea_nui - bloquea_nui

    return {
        "total_registros": total_registros,
        "nui_unicos": nui_unicos,
        "dup_internos": dup_internos,
        "bloquea_nui": bloquea_nui,
        "solo_nobloquea": solo_nobloquea,
    }


def reportar_calidad_sac(df_s: pd.DataFrame, nui_hurtos_efectivos: set) -> dict:
    """
    Analiza calidad del SAC y calcula duplicados cruce hurtos∩SAC_abiertos.
    """
    sac_ab     = df_s[df_s[COL_SEMAFORO].isin(SEMAFOROS_ABIERTOS)]
    nui_sac_ab = set(sac_ab[COL_NUI].dropna())

    dup_cruce = nui_hurtos_efectivos & nui_sac_ab  # resueltos por SAC, no por hurtos

    return {
        "total_registros": len(df_s),
        "nui_unicos": df_s[COL_NUI].nunique(),
        "tickets_abiertos": len(sac_ab),
        "nui_abiertos": len(nui_sac_ab),
        "dup_cruce": dup_cruce,
        "sac_ab": sac_ab,
    }


# ═════════════════════════════════════════════════════════════════════════════
# LÓGICA DE NEGOCIO
# ═════════════════════════════════════════════════════════════════════════════

def consolidar_sac(df_sac: pd.DataFrame) -> pd.DataFrame:
    """
    Consolida la base SAC por NUI.

    Jerarquía de decisión por NUI (orden de prioridad):
      BLOQUEA          → ticket abierto + SubMenu3 = BLOQUEA FACTURACION
      DESCUENTO        → ticket abierto + SubMenu3 = DESCUENTO COMERCIAL
      ABIERTO_SIN_REGLA→ ticket abierto + otro SubMenu3
      SIN_NOVEDAD_SAC  → sin tickets abiertos

    Estado ABIERTO = Semáforo en {Crítico, Moderado, Leve}
    """
    df = df_sac.copy()
    df[COL_NUI]      = normalizar_nui(df[COL_NUI])
    df[COL_SEMAFORO] = normalizar_texto(df[COL_SEMAFORO])
    df[COL_SUBMENU3] = normalizar_texto(df[COL_SUBMENU3])
    df = df.dropna(subset=[COL_NUI])

    filas = []
    for nui, grupo in df.groupby(COL_NUI):
        abiertos = grupo[grupo[COL_SEMAFORO].isin(SEMAFOROS_ABIERTOS)]
        if not abiertos.empty:
            if (abiertos[COL_SUBMENU3] == SUBMENU_BLOQUEA).any():
                dec = "BLOQUEA"
            elif (abiertos[COL_SUBMENU3] == SUBMENU_DESCUENTO).any():
                dec = "DESCUENTO"
            else:
                dec = "ABIERTO_SIN_REGLA"
        else:
            dec = "SIN_NOVEDAD_SAC"
        filas.append({"NUI": nui, "_dec": dec})

    return pd.DataFrame(filas)


def obtener_nui_hurtos(df_hurtos: pd.DataFrame) -> set:
    """
    Retorna el conjunto de NUI efectivos de hurtos.
    Solo se consideran hurtos los NUI con SubMenu3 = BLOQUEA FACTURACION.
    NUI con solo NO BLOQUEA FACTURACION se excluyen.
    """
    df = df_hurtos.copy()
    df[COL_NUI]      = normalizar_nui(df[COL_NUI])
    df[COL_SUBMENU3] = normalizar_texto(df[COL_SUBMENU3])
    df = df.dropna(subset=[COL_NUI])

    # Solo NUI que tienen al menos un registro BLOQUEA FACTURACION
    return set(df[df[COL_SUBMENU3] == SUBMENU_BLOQUEA][COL_NUI].dropna().unique())


def aplicar_reglas(
    df_usuarios: pd.DataFrame,
    df_sac_consolidado: pd.DataFrame,
    nui_hurtos: set,
    nui_sac_abiertos: set,
) -> pd.DataFrame:
    """
    Aplica las 4 reglas de facturación a cada NUI de la base de usuarios.

    Regla 1 — No facturar : SAC abierto + BLOQUEA FACTURACION
    Regla 2 — Sí facturar : SAC abierto + DESCUENTO COMERCIAL  (prioridad sobre hurtos)
    Regla 3 — No facturar : sin ticket SAC abierto + aparece en hurtos
    Regla 4 — Sí facturar : sin ticket SAC abierto + no aparece en hurtos

    Nota: ABIERTO_SIN_REGLA (ticket abierto con otro SubMenu3) no bloquea
    la consulta de hurtos — se revisa igualmente la base de hurtos.
    SAC solo tiene prioridad ABSOLUTA cuando la decisión es BLOQUEA o DESCUENTO.
    """
    # Normalizar usuarios
    df_u = df_usuarios[[COL_NUI]].copy()
    df_u[COL_NUI] = normalizar_nui(df_u[COL_NUI])
    df_u = df_u.dropna(subset=[COL_NUI]).drop_duplicates(subset=[COL_NUI])

    # Unir con decisión SAC
    df = df_u.merge(df_sac_consolidado, on=COL_NUI, how="left")
    df["_dec"] = df["_dec"].fillna("SIN_REGISTRO_SAC")

    filas = []
    for _, row in df.iterrows():
        nui = row[COL_NUI]
        dec = row["_dec"]

        if dec == "BLOQUEA":
            # Regla 1 — SAC bloquea facturación
            filas.append({
                "NUI": nui,
                "Estado de Facturación": "No facturar",
                "Motivo": "Ticket abierto - Bloquea facturación",
                "Fuente de decisión": "SAC",
            })
        elif dec == "DESCUENTO":
            # Regla 2 — SAC descuento comercial (prioridad sobre hurtos)
            filas.append({
                "NUI": nui,
                "Estado de Facturación": "Sí facturar",
                "Motivo": "Ticket abierto - Descuento comercial",
                "Fuente de decisión": "SAC",
            })
        else:
            # Para ABIERTO_SIN_REGLA, SIN_NOVEDAD_SAC y SIN_REGISTRO_SAC
            # → SAC no define la facturación → consultar base de hurtos
            if nui in nui_hurtos:
                # Regla 3 — No facturar por hurto
                filas.append({
                    "NUI": nui,
                    "Estado de Facturación": "No facturar",
                    "Motivo": "Usuario reportado en hurtos",
                    "Fuente de decisión": "Hurtos",
                })
            else:
                # Regla 4 — Facturación normal
                filas.append({
                    "NUI": nui,
                    "Estado de Facturación": "Sí facturar",
                    "Motivo": "Sin novedades",
                    "Fuente de decisión": "Sin coincidencias",
                })

    return pd.DataFrame(filas)


def exportar_excel(df_resultado: pd.DataFrame) -> bytes:
    """Genera Excel con hoja NO_FACTURAR y hoja FACTURAR."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        no_fac = df_resultado[
            df_resultado["Estado de Facturación"] == "No facturar"
        ].reset_index(drop=True)
        si_fac = df_resultado[
            df_resultado["Estado de Facturación"] == "Sí facturar"
        ].reset_index(drop=True)

        no_fac.to_excel(writer, sheet_name="NO_FACTURAR", index=False)
        si_fac.to_excel(writer, sheet_name="FACTURAR",    index=False)

        for sname, dfs in [("NO_FACTURAR", no_fac), ("FACTURAR", si_fac)]:
            ws = writer.sheets[sname]
            for col in ws.columns:
                mx = max((len(str(c.value or "")) for c in col), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(mx + 4, 60)

    return buffer.getvalue()


# ═════════════════════════════════════════════════════════════════════════════
# SIDEBAR — CARGA DE ARCHIVOS
# ═════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## 📂 Cargar bases de datos")
    st.markdown("---")

    st.markdown("#### 👥 Base de Usuarios")
    st.caption("Columna requerida: **NUI**")
    archivo_usuarios = st.file_uploader(
        "db_usuarios.xlsx", type=["xlsx","xls"], key="usuarios", label_visibility="collapsed"
    )

    st.markdown("#### 🎫 Base SAC")
    st.caption("Columnas: **NUI · Semaforo · SubMenu3**")
    archivo_sac = st.file_uploader(
        "db_sac.xlsx", type=["xlsx","xls"], key="sac", label_visibility="collapsed"
    )

    st.markdown("#### 🔒 Base Hurtos")
    st.caption("Columnas: **NUI · SubMenu3**")
    archivo_hurtos = st.file_uploader(
        "db_hurtos_sac.xlsx", type=["xlsx","xls"], key="hurtos", label_visibility="collapsed"
    )

    st.markdown("---")
    procesar = st.button("⚡ Procesar Facturación", use_container_width=True, type="primary")

    with st.expander("📋 Estructura esperada"):
        st.markdown("""
**db_usuarios.xlsx**
`NUI` — identificador único del usuario

**db_sac.xlsx**
`NUI` · `Semaforo` · `SubMenu3`
Semáforo abierto: *Crítico / Moderado / Leve*
SubMenu3 bloqueante: *BLOQUEA FACTURACION*
SubMenu3 descuento: *DESCUENTO COMERCIAL*

**db_hurtos_sac.xlsx**
`NUI` · `SubMenu3`
Solo NUI con *BLOQUEA FACTURACION* se consideran hurtos efectivos
        """)


# ═════════════════════════════════════════════════════════════════════════════
# PROCESAMIENTO
# ═════════════════════════════════════════════════════════════════════════════

if procesar:
    if not archivo_usuarios:
        st.error("❌ Carga la **Base de Usuarios** para continuar.")
        st.stop()

    with st.spinner("Procesando bases de datos…"):

        # ── Lectura ──────────────────────────────────────────────────────────
        df_u = leer_excel(archivo_usuarios)
        df_s = leer_excel(archivo_sac)    if archivo_sac    else None
        df_h = leer_excel(archivo_hurtos) if archivo_hurtos else None

        if df_u is None: st.stop()

        if not validar_columnas(df_u, [COL_NUI], "Base de Usuarios"):
            st.stop()

        df_u[COL_NUI] = normalizar_nui(df_u[COL_NUI])
        df_u = df_u.dropna(subset=[COL_NUI]).drop_duplicates(subset=[COL_NUI])

        # ── Hurtos ───────────────────────────────────────────────────────────
        nui_hurtos      = set()
        cal_hurtos      = {}
        nui_sac_abiertos= set()

        if df_h is not None:
            if validar_columnas(df_h, [COL_NUI, COL_SUBMENU3], "Base Hurtos"):
                df_h[COL_NUI]      = normalizar_nui(df_h[COL_NUI])
                df_h[COL_SUBMENU3] = normalizar_texto(df_h[COL_SUBMENU3])
                df_h = df_h.dropna(subset=[COL_NUI])
                cal_hurtos = reportar_calidad_hurtos(df_h)
                nui_hurtos = cal_hurtos["bloquea_nui"]

        # ── SAC ───────────────────────────────────────────────────────────────
        df_sac_c = pd.DataFrame(columns=[COL_NUI, "_dec"])
        cal_sac  = {}

        if df_s is not None:
            if validar_columnas(df_s, [COL_NUI, COL_SEMAFORO, COL_SUBMENU3], "Base SAC"):
                df_s[COL_NUI]      = normalizar_nui(df_s[COL_NUI])
                df_s[COL_SEMAFORO] = normalizar_texto(df_s[COL_SEMAFORO])
                df_s[COL_SUBMENU3] = normalizar_texto(df_s[COL_SUBMENU3])
                df_s = df_s.dropna(subset=[COL_NUI])
                cal_sac  = reportar_calidad_sac(df_s, nui_hurtos)
                nui_sac_abiertos = set(cal_sac["sac_ab"][COL_NUI].dropna())
                df_sac_c = consolidar_sac(df_s)

        # ── Aplicar reglas ────────────────────────────────────────────────────
        df_resultado = aplicar_reglas(df_u, df_sac_c, nui_hurtos, nui_sac_abiertos)

        if df_resultado.empty:
            st.error("No se pudo generar resultado. Revisa los archivos.")
            st.stop()

    st.session_state["df_resultado"] = df_resultado
    st.session_state["cal_hurtos"]   = cal_hurtos
    st.session_state["cal_sac"]      = cal_sac
    st.success("✅ Procesamiento completado.")


# ═════════════════════════════════════════════════════════════════════════════
# VISUALIZACIÓN
# ═════════════════════════════════════════════════════════════════════════════

if "df_resultado" in st.session_state:
    df_res    = st.session_state["df_resultado"]
    cal_h     = st.session_state.get("cal_hurtos", {})
    cal_s     = st.session_state.get("cal_sac", {})

    total  = len(df_res)
    si_n   = (df_res["Estado de Facturación"] == "Sí facturar").sum()
    no_n   = (df_res["Estado de Facturación"] == "No facturar").sum()
    pct_si = si_n / total * 100 if total else 0
    pct_no = no_n / total * 100 if total else 0

    # ── Métricas principales ─────────────────────────────────────────────────
    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("👥 Total usuarios",  f"{total:,}")
    c2.metric("✅ Sí facturar",      f"{si_n:,}")
    c3.metric("🚫 No facturar",      f"{no_n:,}")
    c4.metric("% Facturados",       f"{pct_si:.1f}%")
    c5.metric("% No facturados",    f"{pct_no:.1f}%")

    # ── Alertas de calidad de datos ──────────────────────────────────────────
    st.markdown("---")
    alertas = []

    if cal_h:
        dup_int   = cal_h.get("dup_internos", 0)
        solo_nb   = cal_h.get("solo_nobloquea", set())
        nui_hurtos_ef = cal_h.get("bloquea_nui", set())

        if dup_int:
            alertas.append(
                f"⚠️ **Hurtos** — {dup_int} NUI duplicados internamente "
                f"(aparecen más de una vez en la base de hurtos). Se consolidó por NUI."
            )
        if solo_nb:
            alertas.append(
                f"ℹ️ **Hurtos** — {len(solo_nb)} NUI tienen solo `NO BLOQUEA FACTURACION` "
                f"en SubMenu3 y fueron **excluidos** como hurtos efectivos."
            )

    if cal_s:
        dup_cruce = cal_s.get("dup_cruce", set())
        if dup_cruce:
            alertas.append(
                f"🔀 **Cruce SAC × Hurtos** — {len(dup_cruce)} NUI aparecen en la base de hurtos "
                f"Y tienen ticket SAC abierto. **SAC tiene prioridad** — estos se resolvieron por SAC."
            )

    if alertas:
        with st.expander(f"⚠️ Alertas de calidad de datos ({len(alertas)})", expanded=True):
            for a in alertas:
                st.markdown(f'<div class="warn-box">{a}</div>', unsafe_allow_html=True)

    # ── Detalle por motivo ───────────────────────────────────────────────────
    with st.expander("📊 Detalle por motivo de decisión"):
        resumen = (
            df_res.groupby(["Estado de Facturación","Motivo","Fuente de decisión"])
            .size().reset_index(name="Cantidad")
            .sort_values("Cantidad", ascending=False)
        )
        st.dataframe(resumen, use_container_width=True, hide_index=True)

    st.markdown("---")

    # ── Filtros y tabla ──────────────────────────────────────────────────────
    cf1, cf2 = st.columns([2,3])
    with cf1:
        filtro = st.selectbox("Filtrar por estado", ["Todos","Sí facturar","No facturar"])
    with cf2:
        buscar = st.text_input("🔍 Buscar por NUI", placeholder="Ingresa un NUI…")

    df_vista = df_res.copy()
    if filtro != "Todos":
        df_vista = df_vista[df_vista["Estado de Facturación"] == filtro]
    if buscar.strip():
        df_vista = df_vista[df_vista["NUI"].str.contains(buscar.strip().upper(), na=False)]

    def color_estado(val):
        if val == "Sí facturar":
            return "background-color:rgba(46,204,113,.25);color:#2ECC71;font-weight:700"
        elif val == "No facturar":
            return "background-color:rgba(231,76,60,.25);color:#E74C3C;font-weight:700"
        return ""

    st.markdown(f"**{len(df_vista):,} registros** mostrados")
    st.dataframe(
        df_vista.style.map(color_estado, subset=["Estado de Facturación"]),
        use_container_width=True,
        height=440,
    )

    # ── Descarga ─────────────────────────────────────────────────────────────
    st.markdown("---")
    excel_bytes = exportar_excel(df_res)
    st.download_button(
        label="⬇️ Descargar Resultado_Facturacion.xlsx",
        data=excel_bytes,
        file_name="Resultado_Facturacion.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

else:
    st.markdown("""
    <div style="text-align:center;padding:4rem 2rem;color:#8FA8C0;">
        <div style="font-size:3.5rem;">⚡</div>
        <h3 style="color:#F5A623;margin-top:1rem;">Listo para procesar</h3>
        <p>Carga las bases en el panel izquierdo y presiona <strong>Procesar Facturación</strong>.</p>
        <p style="font-size:.8rem;margin-top:1.5rem;">
            Base de Usuarios (requerida) · Base SAC · Base Hurtos
        </p>
    </div>
    """, unsafe_allow_html=True)
