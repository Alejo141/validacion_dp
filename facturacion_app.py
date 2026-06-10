"""
Motor de Facturación — DISPOWER S.A.S. E.S.P.
Sistema ZNI / SISFV

Estructura de archivos esperada
────────────────────────────────
• db_usuarios.xlsx   → columna NUI (una fila por usuario)
• db_sac.xlsx        → columnas NUI, Semaforo, SubMenu3
                       Semaforo: Crítico | Moderado | Leve = Abierto
                       Semaforo: Cerrado = Cerrado
• db_hurtos_sac.xlsx → columnas NUI, Semaforo, SubMenu3
                       (todos los registros corresponden a hurtos)
"""

import io
import streamlit as st
import pandas as pd

# ─── Configuración de página ──────────────────────────────────────────────────

st.set_page_config(
    page_title="Motor de Facturación · DISPOWER",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── Estilos ──────────────────────────────────────────────────────────────────

st.markdown("""
<style>
  :root {
    --solar-gold:  #F5A623;
    --deep-navy:   #0D2540;
    --mid-slate:   #1E3A5F;
    --light-slate: #2E547A;
    --green-ok:    #2ECC71;
    --red-no:      #E74C3C;
    --text-main:   #E8EEF4;
    --text-muted:  #8FA8C0;
  }
  .stApp { background-color: var(--deep-navy); color: var(--text-main); }

  [data-testid="stSidebar"] {
    background-color: var(--mid-slate);
    border-right: 2px solid var(--solar-gold);
  }
  [data-testid="stSidebar"] * { color: var(--text-main) !important; }

  [data-testid="stMetric"] {
    background: var(--mid-slate);
    border-radius: 8px;
    padding: 1rem 1.2rem;
    border-left: 3px solid var(--solar-gold);
  }
  [data-testid="stMetricLabel"] { color: var(--text-muted) !important; font-size: .8rem; }
  [data-testid="stMetricValue"] { color: var(--solar-gold) !important; font-size: 1.8rem; }

  .app-header {
    background: linear-gradient(135deg, var(--mid-slate) 0%, var(--light-slate) 100%);
    border-bottom: 3px solid var(--solar-gold);
    padding: 1.2rem 1.8rem;
    border-radius: 10px;
    margin-bottom: 1.5rem;
  }
  .app-header h1 { color: var(--solar-gold); margin:0; font-size:1.6rem; letter-spacing:.5px; }
  .app-header p  { color: var(--text-muted); margin:.3rem 0 0; font-size:.85rem; }

  [data-testid="stDataFrame"] { border: 1px solid var(--light-slate); border-radius: 8px; }
  hr { border-color: var(--light-slate); }

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
  <p>Clasificación automática de usuarios para facturación — DISPOWER S.A.S. E.S.P.</p>
</div>
""", unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════════════════════
# CONSTANTES — nombres de columnas según estructura real de los archivos
# ═════════════════════════════════════════════════════════════════════════════

COL_NUI       = "NUI"
COL_SEMAFORO  = "Semaforo"
COL_SUBMENU3  = "SubMenu3"

# Valores de Semáforo que indican ticket ABIERTO
SEMAFOROS_ABIERTOS = {"CRITICO", "MODERADO", "LEVE"}

# Valores relevantes de SubMenu3 (en mayúsculas para comparación)
SUBMENU_BLOQUEA   = "BLOQUEA FACTURACION"
SUBMENU_DESCUENTO = "DESCUENTO COMERCIAL"


# ═════════════════════════════════════════════════════════════════════════════
# FUNCIONES DE UTILIDAD
# ═════════════════════════════════════════════════════════════════════════════

def leer_excel(archivo) -> pd.DataFrame | None:
    """Lee un archivo Excel cargado por el usuario y retorna DataFrame."""
    try:
        return pd.read_excel(archivo, dtype=str)
    except Exception as e:
        st.error(f"❌ Error al leer **{archivo.name}**: {e}")
        return None


def normalizar_nui(serie: pd.Series) -> pd.Series:
    """Convierte NUI a texto en mayúsculas sin espacios; marca nulos."""
    return serie.astype(str).str.strip().str.upper().replace("NAN", pd.NA)


def quitar_tildes(texto: str) -> str:
    """Elimina tildes y diacríticos. Ej: 'Crítico' → 'Critico'"""
    import unicodedata
    return "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )


def normalizar_texto(serie: pd.Series) -> pd.Series:
    """Normaliza texto: strip + sin tildes + mayúsculas.
    
    'Crítico' → 'CRITICO' | 'crítico' → 'CRITICO'
    Evita fallos por tildes en los valores del Excel.
    """
    return (
        serie.fillna("")
        .astype(str)
        .str.strip()
        .apply(quitar_tildes)
        .str.upper()
    )


def reportar_duplicados(df: pd.DataFrame, nombre: str) -> None:
    """Advierte en pantalla si hay NUI duplicados en una base."""
    n_dup = df[COL_NUI].duplicated().sum()
    if n_dup:
        st.warning(f"⚠️ **{nombre}**: {n_dup:,} registros con NUI duplicado — se consolidará por NUI.")


def validar_columnas(df: pd.DataFrame, requeridas: list[str], nombre: str) -> bool:
    """Verifica que el DataFrame contenga las columnas esperadas."""
    faltantes = [c for c in requeridas if c not in df.columns]
    if faltantes:
        st.error(f"❌ **{nombre}** no tiene las columnas requeridas: {faltantes}\n\n"
                 f"Columnas encontradas: {list(df.columns)}")
        return False
    return True


# ═════════════════════════════════════════════════════════════════════════════
# LÓGICA DE NEGOCIO
# ═════════════════════════════════════════════════════════════════════════════

def consolidar_sac(df_sac: pd.DataFrame) -> pd.DataFrame:
    """
    Consolida la base SAC por NUI.

    Lógica de estado:
      • Semáforo en {Crítico, Moderado, Leve} → ticket ABIERTO
      • Semáforo = Cerrado                     → ticket CERRADO

    Decisión por NUI (prioridad descendente):
      1. BLOQUEA   — ticket abierto + SubMenu3 = BLOQUEA FACTURACION
      2. DESCUENTO — ticket abierto + SubMenu3 = DESCUENTO COMERCIAL
      3. ABIERTO_OTRO — ticket abierto sin SubMenu3 relevante
      4. SIN_NOVEDAD_SAC — solo tickets cerrados o sin registros
    """
    df = df_sac.copy()
    df[COL_NUI]      = normalizar_nui(df[COL_NUI])
    df[COL_SEMAFORO] = normalizar_texto(df[COL_SEMAFORO])
    df[COL_SUBMENU3] = normalizar_texto(df[COL_SUBMENU3])
    df = df.dropna(subset=[COL_NUI])

    resultados = []
    for nui, grupo in df.groupby(COL_NUI):
        abiertos = grupo[grupo[COL_SEMAFORO].isin(SEMAFOROS_ABIERTOS)]

        if not abiertos.empty:
            if (abiertos[COL_SUBMENU3] == SUBMENU_BLOQUEA).any():
                decision = "BLOQUEA"
            elif (abiertos[COL_SUBMENU3] == SUBMENU_DESCUENTO).any():
                decision = "DESCUENTO"
            else:
                decision = "ABIERTO_OTRO"
        else:
            decision = "SIN_NOVEDAD_SAC"

        resultados.append({"NUI": nui, "_sac_decision": decision})

    return pd.DataFrame(resultados)


def obtener_nui_hurtos(df_hurtos: pd.DataFrame) -> set:
    """Retorna conjunto de NUI únicos presentes en la base de hurtos."""
    df = df_hurtos.copy()
    df[COL_NUI] = normalizar_nui(df[COL_NUI])
    return set(df[COL_NUI].dropna().unique())


def aplicar_reglas(
    df_usuarios: pd.DataFrame,
    df_sac_consolidado: pd.DataFrame,
    nui_hurtos: set,
) -> pd.DataFrame:
    """
    Aplica las 4 reglas de negocio a cada NUI de la base de usuarios.

    Regla 1 — No facturar: ticket abierto + BLOQUEA FACTURACION (SAC)
    Regla 2 — Sí facturar: ticket abierto + DESCUENTO COMERCIAL (SAC, prioridad sobre hurtos)
    Regla 3 — No facturar: sin ticket abierto + aparece en hurtos
    Regla 4 — Sí facturar: sin novedad en ninguna base
    """
    # Normalizar NUI de usuarios
    df_u = df_usuarios[[COL_NUI]].copy()
    df_u[COL_NUI] = normalizar_nui(df_u[COL_NUI])
    df_u = df_u.dropna(subset=[COL_NUI]).drop_duplicates(subset=[COL_NUI])

    # Unir con decisión SAC
    df = df_u.merge(df_sac_consolidado, on=COL_NUI, how="left")
    df["_sac_decision"] = df["_sac_decision"].fillna("SIN_REGISTRO_SAC")

    filas = []
    for _, row in df.iterrows():
        nui      = row[COL_NUI]
        decision = row["_sac_decision"]

        if decision == "BLOQUEA":
            # Regla 1 — No facturar
            filas.append({
                "NUI": nui,
                "Estado de Facturación": "No facturar",
                "Motivo": "Ticket abierto - Bloquea facturación",
                "Fuente de decisión": "SAC",
            })
        elif decision == "DESCUENTO":
            # Regla 2 — Sí facturar (gana sobre hurtos)
            filas.append({
                "NUI": nui,
                "Estado de Facturación": "Sí facturar",
                "Motivo": "Ticket abierto - Descuento comercial",
                "Fuente de decisión": "SAC",
            })
        elif nui in nui_hurtos:
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
    """Genera archivo Excel con dos hojas: NO_FACTURAR y FACTURAR."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        no_facturar = df_resultado[
            df_resultado["Estado de Facturación"] == "No facturar"
        ].reset_index(drop=True)
        facturar = df_resultado[
            df_resultado["Estado de Facturación"] == "Sí facturar"
        ].reset_index(drop=True)

        no_facturar.to_excel(writer, sheet_name="NO_FACTURAR", index=False)
        facturar.to_excel(writer, sheet_name="FACTURAR", index=False)

        # Autoajuste de ancho de columnas
        for sheet_name, df_sheet in [("NO_FACTURAR", no_facturar), ("FACTURAR", facturar)]:
            ws = writer.sheets[sheet_name]
            for col in ws.columns:
                max_len = max(
                    (len(str(cell.value or "")) for cell in col), default=10
                )
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)

    return buffer.getvalue()


# ═════════════════════════════════════════════════════════════════════════════
# SIDEBAR — CARGA DE ARCHIVOS
# ═════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## 📂 Cargar bases de datos")
    st.markdown("---")

    st.markdown("#### 👥 Base de Usuarios")
    st.caption("Archivo: `db_usuarios.xlsx` — columna requerida: **NUI**")
    archivo_usuarios = st.file_uploader(
        "Selecciona archivo", type=["xlsx", "xls"], key="usuarios"
    )

    st.markdown("#### 🎫 Base SAC")
    st.caption("Archivo: `db_sac.xlsx` — columnas: **NUI, Semaforo, SubMenu3**")
    archivo_sac = st.file_uploader(
        "Selecciona archivo", type=["xlsx", "xls"], key="sac"
    )

    st.markdown("#### 🔒 Base Hurtos")
    st.caption("Archivo: `db_hurtos_sac.xlsx` — columna requerida: **NUI**")
    archivo_hurtos = st.file_uploader(
        "Selecciona archivo", type=["xlsx", "xls"], key="hurtos"
    )

    st.markdown("---")
    procesar = st.button(
        "⚡ Procesar Facturación", use_container_width=True, type="primary"
    )

    # Leyenda de estructura esperada
    with st.expander("📋 Estructura esperada"):
        st.markdown("""
**db_usuarios.xlsx**
- `NUI` — identificador único del usuario

**db_sac.xlsx**
- `NUI` — identificador del usuario
- `Semaforo` — `Crítico` / `Moderado` / `Leve` = abierto; `Cerrado` = cerrado
- `SubMenu3` — `BLOQUEA FACTURACION` / `DESCUENTO COMERCIAL` / otros

**db_hurtos_sac.xlsx**
- `NUI` — usuarios con caso de hurto
        """)


# ═════════════════════════════════════════════════════════════════════════════
# PROCESAMIENTO PRINCIPAL
# ═════════════════════════════════════════════════════════════════════════════

if procesar:
    if not archivo_usuarios:
        st.error("❌ Debes cargar la **Base de Usuarios** para continuar.")
        st.stop()

    with st.spinner("Procesando bases de datos…"):

        # ── Lectura ──────────────────────────────────────────────────────────
        df_usuarios = leer_excel(archivo_usuarios)
        df_sac      = leer_excel(archivo_sac)      if archivo_sac    else None
        df_hurtos   = leer_excel(archivo_hurtos)   if archivo_hurtos else None

        if df_usuarios is None:
            st.stop()

        # ── Validar columnas requeridas ───────────────────────────────────────
        if not validar_columnas(df_usuarios, [COL_NUI], "Base de Usuarios"):
            st.stop()

        # ── Consolidar SAC ────────────────────────────────────────────────────
        df_sac_consolidado = pd.DataFrame(columns=["NUI", "_sac_decision"])
        if df_sac is not None:
            if validar_columnas(df_sac, [COL_NUI, COL_SEMAFORO, COL_SUBMENU3], "Base SAC"):
                reportar_duplicados(df_sac, "SAC")
                df_sac_consolidado = consolidar_sac(df_sac)
            else:
                st.warning("⚠️ Base SAC omitida por columnas faltantes.")

        # ── Consolidar Hurtos ─────────────────────────────────────────────────
        nui_hurtos: set = set()
        if df_hurtos is not None:
            if validar_columnas(df_hurtos, [COL_NUI], "Base Hurtos"):
                reportar_duplicados(df_hurtos, "Hurtos")
                nui_hurtos = obtener_nui_hurtos(df_hurtos)
            else:
                st.warning("⚠️ Base Hurtos omitida por columnas faltantes.")

        # ── Aplicar reglas ────────────────────────────────────────────────────
        df_resultado = aplicar_reglas(df_usuarios, df_sac_consolidado, nui_hurtos)

        if df_resultado.empty:
            st.error("No se pudo generar el resultado. Revisa el formato de los archivos.")
            st.stop()

    st.session_state["df_resultado"] = df_resultado
    st.success("✅ Procesamiento completado.")


# ═════════════════════════════════════════════════════════════════════════════
# VISUALIZACIÓN DE RESULTADOS
# ═════════════════════════════════════════════════════════════════════════════

if "df_resultado" in st.session_state:
    df_resultado = st.session_state["df_resultado"]

    total         = len(df_resultado)
    facturar_n    = (df_resultado["Estado de Facturación"] == "Sí facturar").sum()
    no_facturar_n = (df_resultado["Estado de Facturación"] == "No facturar").sum()
    pct_si        = facturar_n    / total * 100 if total else 0
    pct_no        = no_facturar_n / total * 100 if total else 0

    # Métricas resumen
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total usuarios",  f"{total:,}")
    c2.metric("✅ Sí facturar",   f"{facturar_n:,}")
    c3.metric("🚫 No facturar",   f"{no_facturar_n:,}")
    c4.metric("% Facturados",    f"{pct_si:.1f}%")
    c5.metric("% No facturados", f"{pct_no:.1f}%")

    # Detalle por fuente de decisión
    st.markdown("---")
    with st.expander("📊 Detalle por motivo"):
        resumen = (
            df_resultado.groupby(["Estado de Facturación", "Motivo", "Fuente de decisión"])
            .size()
            .reset_index(name="Cantidad")
            .sort_values("Cantidad", ascending=False)
        )
        st.dataframe(resumen, use_container_width=True, hide_index=True)

    st.markdown("---")

    # Filtros
    col_f1, col_f2 = st.columns([2, 3])
    with col_f1:
        filtro_estado = st.selectbox(
            "Filtrar por estado",
            ["Todos", "Sí facturar", "No facturar"],
        )
    with col_f2:
        buscar_nui = st.text_input(
            "🔍 Buscar por NUI", placeholder="Ingresa un NUI…"
        )

    df_vista = df_resultado.copy()
    if filtro_estado != "Todos":
        df_vista = df_vista[df_vista["Estado de Facturación"] == filtro_estado]
    if buscar_nui.strip():
        df_vista = df_vista[
            df_vista["NUI"].str.contains(buscar_nui.strip().upper(), na=False)
        ]

    def colorear_estado(val):
        if val == "Sí facturar":
            return "background-color: rgba(46,204,113,.25); color:#2ECC71; font-weight:700"
        elif val == "No facturar":
            return "background-color: rgba(231,76,60,.25); color:#E74C3C; font-weight:700"
        return ""

    st.markdown(f"**{len(df_vista):,} registros** mostrados")
    st.dataframe(
        df_vista.style.map(colorear_estado, subset=["Estado de Facturación"]),
        use_container_width=True,
        height=440,
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
    # Estado inicial — sin datos
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
