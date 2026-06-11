"""
Motor de Facturación — DISPOWER S.A.S. E.S.P.
Sistema ZNI / SISFV

Estructura esperada de archivos
────────────────────────────────
• db_usuarios.xlsx      → columna: NUI
• db_sac.xlsx           → columnas: NUI, Semaforo, SubMenu3, FechaCreacion, FechaCierre
• db_hurtos_sac.xlsx    → columnas: NUI, Semaforo, SubMenu3, FechaCreacion, FechaCierre

Lógica de período (año/mes seleccionado)
────────────────────────────────────────
Para cada NUI se evalúa el estado del ticket al período seleccionado:

  ABIERTO todo el mes (Semáforo abierto sin FechaCierre en el mes):
    → No facturar / Días facturables = 0

  CERRADO dentro del mes (FechaCierre cae en el mes seleccionado):
    → Facturación proporcional: días_facturables = días desde cierre hasta fin de mes
    → Factor = días_facturables / total_días_mes

  CERRADO antes del mes:
    → El bloqueo ya no aplica ese mes → Sí facturar / Factor = 1.0

  Sin ticket relevante:
    → Sí facturar / Factor = 1.0
"""

import io
import calendar
import unicodedata
from datetime import date

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
    --amber:       #F39C12;
    --text-main:   #E8EEF4; --text-muted:  #8FA8C0;
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
  .warn-box  { background: rgba(245,166,35,.12); border: 1px solid var(--solar-gold); border-radius: 8px; padding: .7rem 1rem; margin-bottom:.5rem; font-size:.85rem; }
  .period-box { background: rgba(46,204,113,.10); border: 1px solid var(--green-ok); border-radius: 8px; padding: .8rem 1.2rem; margin-bottom:1rem; }
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

COL_NUI          = "NUI"
COL_SEMAFORO     = "Semaforo"
COL_SUBMENU3     = "SubMenu3"
COL_FECHA_CIERRE = "FechaCierre"
COL_FECHA_CREACION = "FechaCreacion"
DATE_FORMAT      = "%d-%m-%Y"

SEMAFOROS_ABIERTOS = {"CRITICO", "MODERADO", "LEVE"}
SUBMENU_BLOQUEA    = "BLOQUEA FACTURACION"
SUBMENU_DESCUENTO  = "DESCUENTO COMERCIAL"

MESES_ES = {
    1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril",
    5:"Mayo",  6:"Junio",   7:"Julio", 8:"Agosto",
    9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"
}


# ═════════════════════════════════════════════════════════════════════════════
# UTILIDADES
# ═════════════════════════════════════════════════════════════════════════════

def quitar_tildes(texto: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", texto)
                   if unicodedata.category(c) != "Mn")

def normalizar_nui(serie: pd.Series) -> pd.Series:
    return serie.astype(str).str.strip().str.upper().replace("NAN", pd.NA)

def normalizar_texto(serie: pd.Series) -> pd.Series:
    return serie.fillna("").astype(str).str.strip().apply(quitar_tildes).str.upper()

def parsear_fechas(serie: pd.Series) -> pd.Series:
    """Parsea fechas en formato DD-MM-YYYY; retorna NaT si falla."""
    return pd.to_datetime(serie, format=DATE_FORMAT, errors="coerce")

def leer_excel(archivo) -> pd.DataFrame | None:
    try:
        return pd.read_excel(archivo, dtype=str)
    except Exception as e:
        st.error(f"❌ Error al leer **{archivo.name}**: {e}")
        return None

def validar_columnas(df: pd.DataFrame, requeridas: list, nombre: str) -> bool:
    faltantes = [c for c in requeridas if c not in df.columns]
    if faltantes:
        st.error(f"❌ **{nombre}** — columnas faltantes: `{faltantes}`\n\n"
                 f"Columnas encontradas: `{list(df.columns)}`")
        return False
    return True

def ultimo_dia_mes(anio: int, mes: int) -> pd.Timestamp:
    """Retorna el último día del mes como Timestamp."""
    ultimo = calendar.monthrange(anio, mes)[1]
    return pd.Timestamp(date(anio, mes, ultimo))

def primer_dia_mes(anio: int, mes: int) -> pd.Timestamp:
    return pd.Timestamp(date(anio, mes, 1))


# ═════════════════════════════════════════════════════════════════════════════
# LÓGICA DE PERÍODO Y PRORRATEO
# ═════════════════════════════════════════════════════════════════════════════

def calcular_factor_prorratea(fecha_cierre: pd.Timestamp,
                               inicio_mes: pd.Timestamp,
                               fin_mes: pd.Timestamp) -> float:
    """
    Calcula el factor de facturación proporcional para un ticket que se cerró
    dentro del mes de análisis.

    El período NO facturado es desde inicio_mes hasta fecha_cierre (inclusive).
    El período SÍ facturable es desde fecha_cierre+1 hasta fin_mes.

    Factor = días_facturables / total_días_mes
    Ejemplo: cierre el día 10 de un mes de 30 días
             días_facturables = 30 - 10 = 20
             factor = 20/30 = 0.6667
    """
    total_dias = (fin_mes - inicio_mes).days + 1
    dias_bloqueados = (fecha_cierre - inicio_mes).days + 1  # días desde inicio hasta cierre inclusive
    dias_facturables = total_dias - dias_bloqueados
    if dias_facturables < 0:
        dias_facturables = 0
    return round(dias_facturables / total_dias, 6), int(dias_facturables), int(total_dias)


def decision_ticket_en_periodo(grupo: pd.DataFrame,
                                inicio_mes: pd.Timestamp,
                                fin_mes: pd.Timestamp) -> dict:
    """
    Dado el grupo de tickets de un NUI (ya normalizados), determina la decisión
    para el período mes/año seleccionado.

    Devuelve dict con:
      _dec         : BLOQUEA | BLOQUEA_PARCIAL | DESCUENTO | ABIERTO_SIN_REGLA |
                     SIN_NOVEDAD_SAC | SIN_REGISTRO_SAC
      _factor      : float 0..1  (1.0 = factura completo, 0.0 = no factura nada)
      _dias_fact   : días facturables en el mes
      _dias_total  : días totales del mes
      _fecha_cierre: fecha de cierre si aplica prorrateo
    """
    # ── Filtrar tickets relevantes para el período ────────────────────────────
    # Un ticket es relevante si:
    #   a) Está ABIERTO (semáforo abierto) — vigente al analizar
    #   b) Estaba CERRADO pero se cerró DENTRO del mes → prorrateo
    # Tickets cerrados ANTES del mes → ya no bloquean en este período

    dias_mes   = (fin_mes - inicio_mes).days + 1

    abiertos = grupo[grupo[COL_SEMAFORO].isin(SEMAFOROS_ABIERTOS)]

    # Tickets cerrados cuya fecha de cierre cae dentro del mes
    cerrados = grupo[grupo[COL_SEMAFORO] == "CERRADO"]
    cerrados_en_mes = cerrados[
        (cerrados["_fecha_cierre_dt"] >= inicio_mes) &
        (cerrados["_fecha_cierre_dt"] <= fin_mes)
    ]

    # ── Prioridad 1: ticket ABIERTO con BLOQUEA → No facturar (mes completo) ──
    if (abiertos[COL_SUBMENU3] == SUBMENU_BLOQUEA).any():
        return {"_dec": "BLOQUEA", "_factor": 0.0,
                "_dias_fact": 0, "_dias_total": dias_mes, "_fecha_cierre": None}

    # ── Prioridad 2: ticket ABIERTO con DESCUENTO → Sí facturar ──────────────
    if (abiertos[COL_SUBMENU3] == SUBMENU_DESCUENTO).any():
        return {"_dec": "DESCUENTO", "_factor": 1.0,
                "_dias_fact": dias_mes, "_dias_total": dias_mes, "_fecha_cierre": None}

    # ── Prioridad 3: ticket CERRADO en el mes con BLOQUEA → prorrateo ────────
    cerr_bloquea = cerrados_en_mes[cerrados_en_mes[COL_SUBMENU3] == SUBMENU_BLOQUEA]
    if not cerr_bloquea.empty:
        # Si hay varios, tomar el de cierre más tardío (más días bloqueados)
        fecha_cierre = cerr_bloquea["_fecha_cierre_dt"].max()
        factor, dias_fact, dias_tot = calcular_factor_prorratea(
            fecha_cierre, inicio_mes, fin_mes
        )
        return {"_dec": "BLOQUEA_PARCIAL", "_factor": factor,
                "_dias_fact": dias_fact, "_dias_total": dias_tot,
                "_fecha_cierre": fecha_cierre}

    # ── Prioridad 4: ticket CERRADO en el mes con DESCUENTO → prorrateo ──────
    cerr_descuento = cerrados_en_mes[cerrados_en_mes[COL_SUBMENU3] == SUBMENU_DESCUENTO]
    if not cerr_descuento.empty:
        fecha_cierre = cerr_descuento["_fecha_cierre_dt"].max()
        factor, dias_fact, dias_tot = calcular_factor_prorratea(
            fecha_cierre, inicio_mes, fin_mes
        )
        return {"_dec": "DESCUENTO_PARCIAL", "_factor": factor,
                "_dias_fact": dias_fact, "_dias_total": dias_tot,
                "_fecha_cierre": fecha_cierre}

    # ── Abierto sin regla bloqueante ──────────────────────────────────────────
    if not abiertos.empty:
        return {"_dec": "ABIERTO_SIN_REGLA", "_factor": 1.0,
                "_dias_fact": dias_mes, "_dias_total": dias_mes, "_fecha_cierre": None}

    # ── Sin novedad en SAC para este período ──────────────────────────────────
    return {"_dec": "SIN_NOVEDAD_SAC", "_factor": 1.0,
            "_dias_fact": dias_mes, "_dias_total": dias_mes, "_fecha_cierre": None}


def consolidar_sac_con_periodo(df_sac: pd.DataFrame,
                                inicio_mes: pd.Timestamp,
                                fin_mes: pd.Timestamp) -> pd.DataFrame:
    """
    Consolida SAC por NUI teniendo en cuenta el período seleccionado.
    """
    df = df_sac.copy()
    df[COL_NUI]       = normalizar_nui(df[COL_NUI])
    df[COL_SEMAFORO]  = normalizar_texto(df[COL_SEMAFORO])
    df[COL_SUBMENU3]  = normalizar_texto(df[COL_SUBMENU3])
    df["_fecha_cierre_dt"] = parsear_fechas(df[COL_FECHA_CIERRE])
    df = df.dropna(subset=[COL_NUI])

    filas = []
    for nui, grupo in df.groupby(COL_NUI):
        dec_info = decision_ticket_en_periodo(grupo, inicio_mes, fin_mes)
        dec_info["NUI"] = nui
        filas.append(dec_info)

    return pd.DataFrame(filas)


def obtener_hurtos_con_periodo(df_hurtos: pd.DataFrame,
                                inicio_mes: pd.Timestamp,
                                fin_mes: pd.Timestamp) -> dict:
    """
    Para hurtos, retorna un dict NUI → info de prorrateo.
    Todos los NUI en hurtos son casos de hurto.
    Si el ticket de hurto se cerró en el mes → prorrateo.
    Si estaba abierto todo el mes → bloqueo completo.
    Si se cerró antes del mes → ya no aplica (no está en hurtos activos).
    """
    dias_mes = (fin_mes - inicio_mes).days + 1
    df = df_hurtos.copy()
    df[COL_NUI]            = normalizar_nui(df[COL_NUI])
    df[COL_SEMAFORO]       = normalizar_texto(df[COL_SEMAFORO])
    df["_fecha_cierre_dt"] = parsear_fechas(df[COL_FECHA_CIERRE])
    df = df.dropna(subset=[COL_NUI])

    resultado = {}
    for nui, grupo in df.groupby(COL_NUI):
        # Verificar si el cierre ocurrió dentro del mes
        cerrados_en_mes = grupo[
            (grupo["_fecha_cierre_dt"] >= inicio_mes) &
            (grupo["_fecha_cierre_dt"] <= fin_mes)
        ]
        if not cerrados_en_mes.empty:
            # Prorrateo: tomar el cierre más tardío del mes
            fecha_cierre = cerrados_en_mes["_fecha_cierre_dt"].max()
            factor, dias_fact, dias_tot = calcular_factor_prorratea(
                fecha_cierre, inicio_mes, fin_mes
            )
            resultado[nui] = {
                "hurto_factor": factor,
                "hurto_dias_fact": dias_fact,
                "hurto_dias_total": dias_tot,
                "hurto_fecha_cierre": fecha_cierre,
                "hurto_tipo": "PARCIAL",
            }
        else:
            # Sin cierre en el mes: bloqueo completo
            resultado[nui] = {
                "hurto_factor": 0.0,
                "hurto_dias_fact": 0,
                "hurto_dias_total": dias_mes,
                "hurto_fecha_cierre": None,
                "hurto_tipo": "COMPLETO",
            }

    return resultado


def aplicar_reglas_con_periodo(df_usuarios: pd.DataFrame,
                                df_sac_consolidado: pd.DataFrame,
                                hurtos_info: dict,
                                dias_mes: int) -> pd.DataFrame:
    """
    Aplica las 4 reglas de facturación considerando el período y el prorrateo.

    Columnas resultado:
      NUI · Estado de Facturación · Motivo · Fuente de decisión ·
      Factor · Días Facturables · Días del Mes · Fecha Cierre Bloqueo
    """
    df_u = df_usuarios[[COL_NUI]].copy()
    df_u[COL_NUI] = normalizar_nui(df_u[COL_NUI])
    df_u = df_u.dropna(subset=[COL_NUI]).drop_duplicates(subset=[COL_NUI])

    df = df_u.merge(df_sac_consolidado, on=COL_NUI, how="left")
    df["_dec"]          = df["_dec"].fillna("SIN_REGISTRO_SAC")
    df["_factor"]       = df["_factor"].fillna(1.0)
    df["_dias_fact"]    = df["_dias_fact"].fillna(dias_mes)
    df["_dias_total"]   = df["_dias_total"].fillna(dias_mes)
    df["_fecha_cierre"] = df["_fecha_cierre"].where(df["_fecha_cierre"].notna(), None)

    filas = []
    for _, row in df.iterrows():
        nui  = row[COL_NUI]
        dec  = row["_dec"]
        fact = float(row["_factor"])
        df_  = int(row["_dias_fact"])
        dt_  = int(row["_dias_total"])
        fc_  = row["_fecha_cierre"]
        fc_str = fc_.strftime("%d/%m/%Y") if pd.notna(fc_) and fc_ is not None else "—"

        if dec == "BLOQUEA":
            filas.append({
                "NUI": nui,
                "Estado de Facturación": "No facturar",
                "Motivo": "Ticket abierto - Bloquea facturación",
                "Fuente de decisión": "SAC",
                "Factor": 0.0,
                "Días Facturables": 0,
                "Días del Mes": dt_,
                "Fecha Cierre Bloqueo": "—",
            })

        elif dec == "BLOQUEA_PARCIAL":
            estado = "Sí facturar" if fact > 0 else "No facturar"
            filas.append({
                "NUI": nui,
                "Estado de Facturación": estado,
                "Motivo": f"Ticket cerrado en el mes - Bloquea facturación (prorrateo)",
                "Fuente de decisión": "SAC",
                "Factor": fact,
                "Días Facturables": df_,
                "Días del Mes": dt_,
                "Fecha Cierre Bloqueo": fc_str,
            })

        elif dec == "DESCUENTO":
            filas.append({
                "NUI": nui,
                "Estado de Facturación": "Sí facturar",
                "Motivo": "Ticket abierto - Descuento comercial",
                "Fuente de decisión": "SAC",
                "Factor": 1.0,
                "Días Facturables": dt_,
                "Días del Mes": dt_,
                "Fecha Cierre Bloqueo": "—",
            })

        elif dec == "DESCUENTO_PARCIAL":
            filas.append({
                "NUI": nui,
                "Estado de Facturación": "Sí facturar",
                "Motivo": "Ticket cerrado en el mes - Descuento comercial (prorrateo)",
                "Fuente de decisión": "SAC",
                "Factor": fact,
                "Días Facturables": df_,
                "Días del Mes": dt_,
                "Fecha Cierre Bloqueo": fc_str,
            })

        else:
            # Sin bloqueo SAC → revisar hurtos
            if nui in hurtos_info:
                h = hurtos_info[nui]
                hf    = h["hurto_factor"]
                hdf   = h["hurto_dias_fact"]
                hdt   = h["hurto_dias_total"]
                hfc   = h["hurto_fecha_cierre"]
                hfc_s = hfc.strftime("%d/%m/%Y") if hfc is not None else "—"
                tipo  = h["hurto_tipo"]

                if tipo == "PARCIAL":
                    estado = "Sí facturar" if hf > 0 else "No facturar"
                    motivo = "Usuario en hurtos - Ticket cerrado en el mes (prorrateo)"
                else:
                    estado = "No facturar"
                    motivo = "Usuario reportado en hurtos"

                filas.append({
                    "NUI": nui,
                    "Estado de Facturación": estado,
                    "Motivo": motivo,
                    "Fuente de decisión": "Hurtos",
                    "Factor": hf,
                    "Días Facturables": hdf,
                    "Días del Mes": hdt,
                    "Fecha Cierre Bloqueo": hfc_s,
                })
            else:
                filas.append({
                    "NUI": nui,
                    "Estado de Facturación": "Sí facturar",
                    "Motivo": "Sin novedades",
                    "Fuente de decisión": "Sin coincidencias",
                    "Factor": 1.0,
                    "Días Facturables": dias_mes,
                    "Días del Mes": dias_mes,
                    "Fecha Cierre Bloqueo": "—",
                })

    return pd.DataFrame(filas)


# ═════════════════════════════════════════════════════════════════════════════
# EXPORTACIÓN
# ═════════════════════════════════════════════════════════════════════════════

def exportar_excel(df_resultado: pd.DataFrame, anio: int, mes: int) -> bytes:
    """Genera Excel con hoja NO_FACTURAR, FACTURAR y PRORRATEO."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        no_fac   = df_resultado[df_resultado["Estado de Facturación"] == "No facturar"].reset_index(drop=True)
        si_fac   = df_resultado[df_resultado["Estado de Facturación"] == "Sí facturar"].reset_index(drop=True)
        prorrateo= df_resultado[df_resultado["Factor"].between(0.001, 0.999)].reset_index(drop=True)

        no_fac.to_excel(writer,   sheet_name="NO_FACTURAR", index=False)
        si_fac.to_excel(writer,   sheet_name="FACTURAR",    index=False)
        if not prorrateo.empty:
            prorrateo.to_excel(writer, sheet_name="PRORRATEO", index=False)

        for sname, dfs in [("NO_FACTURAR", no_fac), ("FACTURAR", si_fac)] + \
                          ([("PRORRATEO", prorrateo)] if not prorrateo.empty else []):
            ws = writer.sheets[sname]
            for col in ws.columns:
                mx = max((len(str(c.value or "")) for c in col), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(mx + 4, 50)

    return buffer.getvalue()


# ═════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## 📅 Período de análisis")
    st.markdown("---")

    anio_actual = date.today().year
    mes_actual  = date.today().month

    col_a, col_m = st.columns(2)
    with col_a:
        anio_sel = st.selectbox("Año", list(range(2023, anio_actual + 2)), index=list(range(2023, anio_actual + 2)).index(anio_actual))
    with col_m:
        mes_sel = st.selectbox("Mes", list(range(1, 13)), index=mes_actual - 1,
                               format_func=lambda x: MESES_ES[x])

    inicio_mes = primer_dia_mes(anio_sel, mes_sel)
    fin_mes    = ultimo_dia_mes(anio_sel, mes_sel)
    dias_mes   = (fin_mes - inicio_mes).days + 1

    st.markdown(f"""
    <div style="background:rgba(245,166,35,.1);border:1px solid #F5A623;border-radius:6px;padding:.6rem 1rem;margin-top:.5rem;font-size:.82rem;">
    📆 <b>{MESES_ES[mes_sel]} {anio_sel}</b><br>
    Corte: <b>{fin_mes.strftime('%d/%m/%Y')}</b> · <b>{dias_mes} días</b>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("## 📂 Cargar bases de datos")

    st.markdown("#### 👥 Base de Usuarios")
    st.caption("Columna requerida: **NUI**")
    archivo_usuarios = st.file_uploader("usuarios", type=["xlsx","xls"],
                                        key="usuarios", label_visibility="collapsed")

    st.markdown("#### 🎫 Base SAC")
    st.caption("Columnas: **NUI · Semaforo · SubMenu3 · FechaCierre**")
    archivo_sac = st.file_uploader("sac", type=["xlsx","xls"],
                                   key="sac", label_visibility="collapsed")

    st.markdown("#### 🔒 Base Hurtos")
    st.caption("Columnas: **NUI · Semaforo · FechaCierre**")
    archivo_hurtos = st.file_uploader("hurtos", type=["xlsx","xls"],
                                      key="hurtos", label_visibility="collapsed")

    st.markdown("---")
    procesar = st.button("⚡ Procesar Facturación", use_container_width=True, type="primary")

    with st.expander("📋 Lógica de período"):
        st.markdown(f"""
**Ticket ABIERTO todo el mes**
→ No factura (bloqueo completo)

**Ticket cerrado DENTRO del mes**
→ Factura proporcional:
`Factor = días_desde_cierre / {dias_mes}`

**Ticket cerrado ANTES del mes**
→ Factura completo (ya no bloquea)

**Sin novedad en SAC/Hurtos**
→ Factura completo
        """)


# ═════════════════════════════════════════════════════════════════════════════
# PROCESAMIENTO
# ═════════════════════════════════════════════════════════════════════════════

if procesar:
    if not archivo_usuarios:
        st.error("❌ Carga la **Base de Usuarios** para continuar.")
        st.stop()

    with st.spinner(f"Procesando {MESES_ES[mes_sel]} {anio_sel}…"):

        df_u = leer_excel(archivo_usuarios)
        df_s = leer_excel(archivo_sac)    if archivo_sac    else None
        df_h = leer_excel(archivo_hurtos) if archivo_hurtos else None

        if df_u is None: st.stop()
        if not validar_columnas(df_u, [COL_NUI], "Base de Usuarios"): st.stop()

        df_u[COL_NUI] = normalizar_nui(df_u[COL_NUI])
        df_u = df_u.dropna(subset=[COL_NUI]).drop_duplicates(subset=[COL_NUI])

        # ── SAC ───────────────────────────────────────────────────────────────
        df_sac_c  = pd.DataFrame(columns=[COL_NUI, "_dec", "_factor",
                                           "_dias_fact", "_dias_total", "_fecha_cierre"])
        alertas   = []

        if df_s is not None:
            cols_req_sac = [COL_NUI, COL_SEMAFORO, COL_SUBMENU3, COL_FECHA_CIERRE]
            if validar_columnas(df_s, cols_req_sac, "Base SAC"):
                df_s[COL_NUI] = normalizar_nui(df_s[COL_NUI])
                df_s = df_s.dropna(subset=[COL_NUI])
                df_sac_c = consolidar_sac_con_periodo(df_s, inicio_mes, fin_mes)

                dup_sac = df_s[df_s.duplicated(subset=[COL_NUI], keep=False)][COL_NUI].nunique()
                if dup_sac:
                    alertas.append(f"⚠️ **SAC** — {dup_sac} NUI con múltiples tickets. Se aplicó lógica de prioridad por NUI.")

        # ── Hurtos ───────────────────────────────────────────────────────────
        hurtos_info = {}

        if df_h is not None:
            cols_req_h = [COL_NUI, COL_FECHA_CIERRE]
            if validar_columnas(df_h, cols_req_h, "Base Hurtos"):
                df_h[COL_NUI] = normalizar_nui(df_h[COL_NUI])
                df_h = df_h.dropna(subset=[COL_NUI])
                hurtos_info = obtener_hurtos_con_periodo(df_h, inicio_mes, fin_mes)

                dup_h = df_h[df_h.duplicated(subset=[COL_NUI], keep=False)][COL_NUI].nunique()
                if dup_h:
                    alertas.append(f"⚠️ **Hurtos** — {dup_h} NUI duplicados internamente. Se consolidó por NUI.")

        # ── Cruce SAC × Hurtos ────────────────────────────────────────────────
        nui_hurtos    = set(hurtos_info.keys())
        nui_sac_bloq  = set(df_sac_c[df_sac_c["_dec"].isin(["BLOQUEA","BLOQUEA_PARCIAL"])]["NUI"])
        dup_cruce     = nui_hurtos & nui_sac_bloq
        if dup_cruce:
            alertas.append(f"🔀 **Cruce SAC × Hurtos** — {len(dup_cruce)} NUI en ambas bases. SAC tiene prioridad.")

        # ── Aplicar reglas ────────────────────────────────────────────────────
        df_resultado = aplicar_reglas_con_periodo(df_u, df_sac_c, hurtos_info, dias_mes)

        if df_resultado.empty:
            st.error("No se pudo generar resultado.")
            st.stop()

    st.session_state.update({
        "df_resultado": df_resultado,
        "alertas": alertas,
        "anio_sel": anio_sel,
        "mes_sel": mes_sel,
        "inicio_mes": inicio_mes,
        "fin_mes": fin_mes,
        "dias_mes": dias_mes,
    })
    st.success(f"✅ Procesamiento de **{MESES_ES[mes_sel]} {anio_sel}** completado.")


# ═════════════════════════════════════════════════════════════════════════════
# VISUALIZACIÓN
# ═════════════════════════════════════════════════════════════════════════════

if "df_resultado" in st.session_state:
    df_res   = st.session_state["df_resultado"]
    alertas  = st.session_state.get("alertas", [])
    anio_vis = st.session_state.get("anio_sel", anio_sel)
    mes_vis  = st.session_state.get("mes_sel",  mes_sel)
    dias_vis = st.session_state.get("dias_mes", dias_mes)
    fin_vis  = st.session_state.get("fin_mes",  fin_mes)

    # Banner de período
    st.markdown(f"""
    <div class="period-box">
      📅 Análisis: <strong>{MESES_ES[mes_vis]} {anio_vis}</strong> &nbsp;·&nbsp;
      Corte: <strong>{fin_vis.strftime('%d/%m/%Y')}</strong> &nbsp;·&nbsp;
      <strong>{dias_vis} días</strong> en el mes
    </div>
    """, unsafe_allow_html=True)

    total   = len(df_res)
    no_n    = (df_res["Estado de Facturación"] == "No facturar").sum()
    si_n    = (df_res["Estado de Facturación"] == "Sí facturar").sum()
    prorr_n = df_res["Factor"].between(0.001, 0.999).sum()
    pct_si  = si_n  / total * 100 if total else 0
    pct_no  = no_n  / total * 100 if total else 0

    c1,c2,c3,c4,c5,c6 = st.columns(6)
    c1.metric("👥 Total",          f"{total:,}")
    c2.metric("✅ Sí facturar",     f"{si_n:,}")
    c3.metric("🚫 No facturar",     f"{no_n:,}")
    c4.metric("📊 Prorrateo",       f"{prorr_n:,}")
    c5.metric("% Facturados",      f"{pct_si:.1f}%")
    c6.metric("% No facturados",   f"{pct_no:.1f}%")

    # Alertas
    if alertas:
        st.markdown("---")
        with st.expander(f"⚠️ Alertas de calidad ({len(alertas)})", expanded=True):
            for a in alertas:
                st.markdown(f'<div class="warn-box">{a}</div>', unsafe_allow_html=True)

    # Detalle por motivo
    st.markdown("---")
    with st.expander("📊 Detalle por motivo de decisión"):
        resumen = (
            df_res.groupby(["Estado de Facturación","Motivo","Fuente de decisión"])
            .size().reset_index(name="Cantidad")
            .sort_values("Cantidad", ascending=False)
        )
        st.dataframe(resumen, use_container_width=True, hide_index=True)

    # Filtros y tabla
    st.markdown("---")
    cf1, cf2, cf3 = st.columns([2, 2, 3])
    with cf1:
        filtro_estado = st.selectbox("Estado", ["Todos","Sí facturar","No facturar"])
    with cf2:
        filtro_factor = st.selectbox("Tipo",   ["Todos","Factor completo (1.0)","Prorrateo (<1.0)","Sin facturación (0.0)"])
    with cf3:
        buscar = st.text_input("🔍 Buscar NUI", placeholder="Ingresa un NUI…")

    df_vista = df_res.copy()
    if filtro_estado != "Todos":
        df_vista = df_vista[df_vista["Estado de Facturación"] == filtro_estado]
    if filtro_factor == "Factor completo (1.0)":
        df_vista = df_vista[df_vista["Factor"] == 1.0]
    elif filtro_factor == "Prorrateo (<1.0)":
        df_vista = df_vista[df_vista["Factor"].between(0.001, 0.999)]
    elif filtro_factor == "Sin facturación (0.0)":
        df_vista = df_vista[df_vista["Factor"] == 0.0]
    if buscar.strip():
        df_vista = df_vista[df_vista["NUI"].str.contains(buscar.strip().upper(), na=False)]

    def color_fila(row):
        if row["Factor"] == 0.0:
            return ["background-color:rgba(231,76,60,.2)"]*len(row)
        elif row["Factor"] < 1.0:
            return ["background-color:rgba(243,156,18,.15)"]*len(row)
        return [""]*len(row)

    def color_estado(val):
        if val == "Sí facturar":
            return "color:#2ECC71;font-weight:700"
        elif val == "No facturar":
            return "color:#E74C3C;font-weight:700"
        return ""

    st.markdown(f"**{len(df_vista):,} registros** mostrados")
    st.dataframe(
        df_vista.style
            .apply(color_fila, axis=1)
            .map(color_estado, subset=["Estado de Facturación"]),
        use_container_width=True,
        height=440,
        column_config={
            "Factor": st.column_config.NumberColumn("Factor", format="%.4f"),
        }
    )

    # Descarga
    st.markdown("---")
    excel_bytes = exportar_excel(df_res, anio_vis, mes_vis)
    st.download_button(
        label=f"⬇️ Descargar Resultado_{MESES_ES[mes_vis]}_{anio_vis}.xlsx",
        data=excel_bytes,
        file_name=f"Resultado_Facturacion_{MESES_ES[mes_vis]}_{anio_vis}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

else:
    st.markdown("""
    <div style="text-align:center;padding:4rem 2rem;color:#8FA8C0;">
        <div style="font-size:3.5rem;">⚡</div>
        <h3 style="color:#F5A623;margin-top:1rem;">Listo para procesar</h3>
        <p>Selecciona el <strong>año y mes</strong>, carga las bases y presiona <strong>Procesar Facturación</strong>.</p>
    </div>
    """, unsafe_allow_html=True)
