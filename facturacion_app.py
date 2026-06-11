"""
Motor de Facturación — DISPOWER S.A.S. E.S.P.  ·  Sistema ZNI / SISFV

Lógica de período (FechaCierre como fuente de verdad)
──────────────────────────────────────────────────────
Para cada ticket se calcula la fecha efectiva de cierre:
  • Semáforo ABIERTO (Crítico/Moderado/Leve) → vigente hasta fin_mes
  • Semáforo CERRADO → vigente hasta FechaCierre

Un ticket afecta el período si fc_efectiva >= ini_mes.

Clasificación por NUI:
  BLOQUEA          → ticket activo en el mes con SubMenu3=BLOQUEA FACTURACION
                     y fc_efectiva >= fin_mes  → factor 0 (no factura)
  BLOQUEA_PARCIAL  → igual pero fc_efectiva dentro del mes → prorrateo
  DESCUENTO        → ticket activo con DESCUENTO COMERCIAL → factor 1
  SIN_NOVEDAD_SAC  → sin tickets activos en el período

Prorrateo: factor = (días_mes - días_bloqueados) / días_mes
  días_bloqueados = días desde ini_mes hasta FechaCierre (inclusive)
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
    page_icon="⚡", layout="wide", initial_sidebar_state="expanded",
)

st.markdown("""
<style>
  :root {
    --solar-gold:#F5A623; --deep-navy:#0D2540; --mid-slate:#1E3A5F;
    --light-slate:#2E547A; --green-ok:#2ECC71; --red-no:#E74C3C;
    --amber:#F39C12; --text-main:#E8EEF4; --text-muted:#8FA8C0;
  }
  .stApp{background-color:var(--deep-navy);color:var(--text-main);}
  [data-testid="stSidebar"]{background-color:var(--mid-slate);border-right:2px solid var(--solar-gold);}
  [data-testid="stSidebar"] *{color:var(--text-main)!important;}
  [data-testid="stMetric"]{background:var(--mid-slate);border-radius:8px;padding:1rem 1.2rem;border-left:3px solid var(--solar-gold);}
  [data-testid="stMetricLabel"]{color:var(--text-muted)!important;font-size:.8rem;}
  [data-testid="stMetricValue"]{color:var(--solar-gold)!important;font-size:1.8rem;}
  .app-header{background:linear-gradient(135deg,var(--mid-slate) 0%,var(--light-slate) 100%);border-bottom:3px solid var(--solar-gold);padding:1.2rem 1.8rem;border-radius:10px;margin-bottom:1.5rem;}
  .app-header h1{color:var(--solar-gold);margin:0;font-size:1.6rem;}
  .app-header p{color:var(--text-muted);margin:.3rem 0 0;font-size:.85rem;}
  [data-testid="stDataFrame"]{border:1px solid var(--light-slate);border-radius:8px;}
  hr{border-color:var(--light-slate);}
  .stDownloadButton>button{background-color:var(--solar-gold)!important;color:var(--deep-navy)!important;font-weight:700;border:none;border-radius:6px;}
  .warn-box{background:rgba(245,166,35,.12);border:1px solid var(--solar-gold);border-radius:8px;padding:.7rem 1rem;margin-bottom:.5rem;font-size:.85rem;}
  .period-box{background:rgba(46,204,113,.10);border:1px solid var(--green-ok);border-radius:8px;padding:.8rem 1.2rem;margin-bottom:1rem;}
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
COL_FC       = "FechaCierre"
DATE_FMT     = "%d-%m-%Y"

SEMAFOROS_ABIERTOS = {"CRITICO", "MODERADO", "LEVE"}
SUBMENU_BLOQUEA    = "BLOQUEA FACTURACION"
SUBMENU_DESCUENTO  = "DESCUENTO COMERCIAL"
SUBMENU_NO_BLOQUEA = "NO BLOQUEA FACTURACION"
COL_CONCAT         = "Concatenado"
REPOSICION_KEYWORD = "REPOSICI"   # cubre Reposición / REPOSICION / reposicion

MESES_ES = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
            7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}


# ═════════════════════════════════════════════════════════════════════════════
# UTILIDADES
# ═════════════════════════════════════════════════════════════════════════════

def quitar_tildes(t: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", t) if unicodedata.category(c) != "Mn")

def normalizar_nui(s: pd.Series) -> pd.Series:
    return s.astype(str).str.strip().str.upper().replace("NAN", pd.NA)

def normalizar_texto(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.strip().apply(quitar_tildes).str.upper()

def parsear_fechas(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, format=DATE_FMT, errors="coerce")

def leer_excel(archivo) -> pd.DataFrame | None:
    try:
        return pd.read_excel(archivo, dtype=str)
    except Exception as e:
        st.error(f"❌ Error al leer **{archivo.name}**: {e}")
        return None

def validar_columnas(df: pd.DataFrame, cols: list, nombre: str) -> bool:
    faltantes = [c for c in cols if c not in df.columns]
    if faltantes:
        st.error(f"❌ **{nombre}** — columnas faltantes: `{faltantes}`\n\n"
                 f"Columnas encontradas: `{list(df.columns)}`")
        return False
    return True

def fin_mes_ts(anio: int, mes: int) -> pd.Timestamp:
    return pd.Timestamp(date(anio, mes, calendar.monthrange(anio, mes)[1]))

def ini_mes_ts(anio: int, mes: int) -> pd.Timestamp:
    return pd.Timestamp(date(anio, mes, 1))


# ═════════════════════════════════════════════════════════════════════════════
# NÚCLEO: clasificación por período usando FechaCierre
# ═════════════════════════════════════════════════════════════════════════════

def fc_efectiva(semaforo: str, fecha_cierre, fin: pd.Timestamp) -> pd.Timestamp:
    """
    Fecha efectiva de vigencia del ticket dentro del período.
    - Ticket abierto (semáforo activo) o sin FechaCierre → fin del mes
    - Ticket cerrado → su FechaCierre real
    """
    if semaforo in SEMAFOROS_ABIERTOS or pd.isna(fecha_cierre):
        return fin
    return fecha_cierre


def calcular_prorrateo(fc: pd.Timestamp, ini: pd.Timestamp,
                        fin: pd.Timestamp, dias_mes: int) -> tuple:
    """
    Calcula factor proporcional cuando el ticket cerró DENTRO del mes.
    Días bloqueados = inicio del mes hasta FechaCierre inclusive.
    Días facturables = días restantes hasta fin del mes.
    """
    dias_bloqueados  = (fc - ini).days + 1
    dias_facturables = dias_mes - dias_bloqueados
    if dias_facturables < 0:
        dias_facturables = 0
    factor = round(dias_facturables / dias_mes, 6)
    return factor, int(dias_facturables), int(dias_mes)


def clasificar_nui_sac(grupo: pd.DataFrame, ini: pd.Timestamp,
                        fin: pd.Timestamp, dias_mes: int,
                        col_creacion: str = "FechaCreacion") -> dict:
    """
    Determina la decisión SAC para un NUI en el período dado.

        Generan prorrateo:
      1. SubMenu3=BLOQUEA FACTURACION + cerrado en el mes
      2. SubMenu3=NO BLOQUEA FACTURACION + cerrado en el mes
         + campo Concatenado contiene la palabra "Reposición"

    Lógica por SubMenu3:
      BLOQUEA FACTURACION  → bloquea o prorrateo según FechaCierre
      DESCUENTO COMERCIAL  → sí facturar (completo o prorrateo)
      NO BLOQUEA + REPOSICION en Concatenado → prorrateo si cerrado en mes
      Cualquier otro       → se ignora (no afecta facturación)
    """
    activos_bloquea   = []  # lista de (fc_inicio_bloqueo, fc_fin_bloqueo)
    activos_descuento = []
    activos_repos     = []   # NO BLOQUEA + REPOSICION en Concatenado

    for _, r in grupo.iterrows():
        sub    = r[COL_SUBMENU3]
        concat = str(r.get(COL_CONCAT, "")).upper()
        sem    = r[COL_SEMAFORO]
        fce    = fc_efectiva(sem, r["_fc"], fin)   # cuándo TERMINA el bloqueo

        if fce < ini:
            continue  # terminó antes del mes, no afecta

        if sub == SUBMENU_BLOQUEA:
            # Determinar cuándo INICIA el bloqueo dentro del mes:
            # Si el ticket estaba abierto antes del mes → inicia en ini_mes
            # Si el ticket se creó dentro del mes → inicia en FechaCreacion
            fcreac = parsear_fechas(pd.Series([r.get(col_creacion, "")])).iloc[0]
            if pd.notna(fcreac) and fcreac > ini:
                # Ticket creado dentro del mes
                inicio_bloqueo = max(fcreac, ini)
            else:
                # Ticket existía antes del mes → bloquea desde el primer día
                inicio_bloqueo = ini
            activos_bloquea.append((inicio_bloqueo, fce))

        elif sub == SUBMENU_DESCUENTO:
            activos_descuento.append(fce)

        elif sub == SUBMENU_NO_BLOQUEA:
            # Solo aplica prorrateo si: cerrado DENTRO del mes Y Concatenado
            # contiene la palabra "Reposición"
            if (fce < fin                               # cerrado dentro del mes
                    and REPOSICION_KEYWORD in concat):  # contiene Reposición
                activos_repos.append(fce)
            # Caso contrario (abierto, o sin Reposición): se ignora

    # ── Prioridad 1: BLOQUEA ────────────────────────────────────────────────
    if activos_bloquea:
        # Calcular días TOTALES bloqueados en el mes considerando
        # el inicio más temprano y el fin más tardío de todos los tickets BLOQUEA.
        # Si cualquier ticket empezó antes o en ini_mes y sigue abierto → bloqueo completo.
        inicio_min = min(t[0] for t in activos_bloquea)  # inicio de bloqueo más temprano
        fc_max     = max(t[1] for t in activos_bloquea)  # fin de bloqueo más tardío

        if inicio_min <= ini and fc_max >= fin:
            # Bloqueado todo el mes
            return {"_dec":"BLOQUEA", "_factor":0.0,
                    "_dias_fact":0, "_dias_total":dias_mes, "_fc_display":None}
        else:
            # Calcular días efectivamente bloqueados en el mes
            inicio_bloq_efectivo = max(inicio_min, ini)
            fin_bloq_efectivo    = min(fc_max, fin)
            dias_bloqueados      = (fin_bloq_efectivo - inicio_bloq_efectivo).days + 1
            dias_fact            = dias_mes - dias_bloqueados
            if dias_fact < 0: dias_fact = 0
            factor = round(dias_fact / dias_mes, 6)
            return {"_dec":"BLOQUEA_PARCIAL", "_factor":factor,
                    "_dias_fact":int(dias_fact), "_dias_total":dias_mes,
                    "_fc_display":fin_bloq_efectivo}

    # ── Prioridad 2: DESCUENTO ──────────────────────────────────────────────
    if activos_descuento:
        fc_max = max(activos_descuento)
        if fc_max >= fin:
            return {"_dec":"DESCUENTO", "_factor":1.0,
                    "_dias_fact":dias_mes, "_dias_total":dias_mes, "_fc_display":None}
        else:
            f, df_, dt_ = calcular_prorrateo(fc_max, ini, fin, dias_mes)
            return {"_dec":"DESCUENTO_PARCIAL", "_factor":f,
                    "_dias_fact":df_, "_dias_total":dt_, "_fc_display":fc_max}

    # ── Prioridad 3: NO BLOQUEA + REPOSICION cerrado en el mes → prorrateo ─
    if activos_repos:
        # Tomar el cierre más tardío (mayor días bloqueados)
        fc_max = max(activos_repos)
        f, df_, dt_ = calcular_prorrateo(fc_max, ini, fin, dias_mes)
        return {"_dec":"REPOSICION_PARCIAL", "_factor":f,
                "_dias_fact":df_, "_dias_total":dt_, "_fc_display":fc_max}

    # ── Sin tickets relevantes activos → factura completo ──────────────────
    return {"_dec":"SIN_NOVEDAD_SAC", "_factor":1.0,
            "_dias_fact":dias_mes, "_dias_total":dias_mes, "_fc_display":None}


def clasificar_nui_hurtos(grupo: pd.DataFrame, ini: pd.Timestamp,
                           fin: pd.Timestamp, dias_mes: int) -> dict:
    """
    Determina la situación de hurto de un NUI en el período.
    Misma lógica de FechaCierre: si el hurto estaba activo en el mes → bloquea/prorrateo.
    """
    fc_activos = []
    for _, r in grupo.iterrows():
        fce = fc_efectiva("CERRADO", r["_fc"], fin)
        if fce >= ini:
            fc_activos.append(fce)

    if not fc_activos:
        # Todos los hurtos cerraron antes del mes → no aplica
        return {"_factor":1.0, "_dias_fact":dias_mes, "_tipo":"YA_CERRADO", "_fc":None}

    fc_max = max(fc_activos)
    if fc_max >= fin:
        return {"_factor":0.0, "_dias_fact":0, "_tipo":"COMPLETO", "_fc":None}
    else:
        f, df_, _ = calcular_prorrateo(fc_max, ini, fin, dias_mes)
        return {"_factor":f, "_dias_fact":df_, "_tipo":"PARCIAL", "_fc":fc_max}


# ═════════════════════════════════════════════════════════════════════════════
# PROCESAMIENTO DE BASES
# ═════════════════════════════════════════════════════════════════════════════

def consolidar_sac(df_sac: pd.DataFrame, ini: pd.Timestamp,
                   fin: pd.Timestamp, dias_mes: int) -> pd.DataFrame:
    df = df_sac.copy()
    df[COL_NUI]      = normalizar_nui(df[COL_NUI])
    df[COL_SEMAFORO] = normalizar_texto(df[COL_SEMAFORO])
    df[COL_SUBMENU3] = normalizar_texto(df[COL_SUBMENU3])
    df["_fc"]        = parsear_fechas(df[COL_FC])
    # Normalizar Concatenado si existe; si no, columna vacía
    if COL_CONCAT in df.columns:
        df[COL_CONCAT] = df[COL_CONCAT].fillna("").astype(str).str.upper()
    else:
        df[COL_CONCAT] = ""
    df = df.dropna(subset=[COL_NUI])

    filas = []
    for nui, grupo in df.groupby(COL_NUI):
        info = clasificar_nui_sac(grupo, ini, fin, dias_mes,
                                  col_creacion=COL_FECHA_CREACION)
        info["NUI"] = nui
        filas.append(info)
    return pd.DataFrame(filas)


def consolidar_hurtos(df_h: pd.DataFrame, ini: pd.Timestamp,
                      fin: pd.Timestamp, dias_mes: int) -> dict:
    df = df_h.copy()
    df[COL_NUI] = normalizar_nui(df[COL_NUI])
    df["_fc"]   = parsear_fechas(df[COL_FC])
    df = df.dropna(subset=[COL_NUI])

    resultado = {}
    for nui, grupo in df.groupby(COL_NUI):
        resultado[nui] = clasificar_nui_hurtos(grupo, ini, fin, dias_mes)
    return resultado


def aplicar_reglas(df_usuarios: pd.DataFrame, df_sac_c: pd.DataFrame,
                   hurtos_info: dict, dias_mes: int) -> pd.DataFrame:
    """
    Aplica las 4 reglas de facturación a cada NUI.
    SAC tiene prioridad sobre hurtos. Solo BLOQUEA FACTURACION genera prorrateo.
    """
    df_u = df_usuarios[[COL_NUI]].copy()
    df_u[COL_NUI] = normalizar_nui(df_u[COL_NUI])
    df_u = df_u.dropna(subset=[COL_NUI]).drop_duplicates(subset=[COL_NUI])

    df = df_u.merge(df_sac_c, on=COL_NUI, how="left")
    df["_dec"]       = df["_dec"].fillna("SIN_REGISTRO_SAC")
    df["_factor"]    = df["_factor"].fillna(1.0)
    df["_dias_fact"] = df["_dias_fact"].fillna(dias_mes)
    df["_dias_total"]= df["_dias_total"].fillna(dias_mes)
    df["_fc_display"]= df["_fc_display"].where(df["_fc_display"].notna(), None)

    filas = []
    for _, row in df.iterrows():
        nui  = row[COL_NUI]
        dec  = row["_dec"]
        fac  = float(row["_factor"])
        df_  = int(row["_dias_fact"])
        dt_  = int(row["_dias_total"])
        fcd  = row["_fc_display"]
        fcs  = fcd.strftime("%d/%m/%Y") if pd.notna(fcd) and fcd is not None else "—"

        # ── Regla 1: SAC BLOQUEA completo ────────────────────────────────────
        if dec == "BLOQUEA":
            filas.append({"NUI":nui,"Estado de Facturación":"No facturar",
                "Motivo":"Ticket abierto - Bloquea facturación",
                "Fuente de decisión":"SAC","Factor":0.0,
                "Días Facturables":0,"Días del Mes":dt_,"Fecha Cierre Bloqueo":"—"})

        # ── Regla 1P: SAC BLOQUEA parcial (prorrateo) ────────────────────────
        elif dec == "BLOQUEA_PARCIAL":
            estado = "Sí facturar" if fac > 0 else "No facturar"
            filas.append({"NUI":nui,"Estado de Facturación":estado,
                "Motivo":"Ticket cerrado en el mes - Bloquea facturación (prorrateo)",
                "Fuente de decisión":"SAC","Factor":fac,
                "Días Facturables":df_,"Días del Mes":dt_,"Fecha Cierre Bloqueo":fcs})

        # ── Regla 2: SAC DESCUENTO ───────────────────────────────────────────
        elif dec in ("DESCUENTO","DESCUENTO_PARCIAL"):
            motivo = ("Ticket abierto - Descuento comercial" if dec=="DESCUENTO"
                      else "Ticket cerrado en el mes - Descuento comercial (prorrateo)")
            filas.append({"NUI":nui,"Estado de Facturación":"Sí facturar",
                "Motivo":motivo,"Fuente de decisión":"SAC","Factor":fac,
                "Días Facturables":df_,"Días del Mes":dt_,"Fecha Cierre Bloqueo":fcs})

        # ── Regla 2R: NO BLOQUEA + REPOSICION cerrado en mes → prorrateo ────
        elif dec == "REPOSICION_PARCIAL":
            estado = "Sí facturar" if fac > 0 else "No facturar"
            filas.append({"NUI":nui,"Estado de Facturación":estado,
                "Motivo":"Ticket cerrado en el mes - No bloquea / Reposición (prorrateo)",
                "Fuente de decisión":"SAC","Factor":fac,
                "Días Facturables":df_,"Días del Mes":dt_,"Fecha Cierre Bloqueo":fcs})

        # ── Reglas 3 y 4: sin decisión SAC → revisar hurtos ──────────────────
        else:
            if nui in hurtos_info:
                h = hurtos_info[nui]
                if h["_tipo"] == "YA_CERRADO":
                    # Hurto cerrado antes del mes → no aplica
                    filas.append({"NUI":nui,"Estado de Facturación":"Sí facturar",
                        "Motivo":"Sin novedades","Fuente de decisión":"Sin coincidencias",
                        "Factor":1.0,"Días Facturables":dt_,"Días del Mes":dt_,
                        "Fecha Cierre Bloqueo":"—"})
                elif h["_tipo"] == "COMPLETO":
                    filas.append({"NUI":nui,"Estado de Facturación":"No facturar",
                        "Motivo":"Usuario reportado en hurtos",
                        "Fuente de decisión":"Hurtos","Factor":0.0,
                        "Días Facturables":0,"Días del Mes":dt_,"Fecha Cierre Bloqueo":"—"})
                else:
                    hf  = h["_factor"]; hdf = h["_dias_fact"]
                    hfc = h["_fc"]
                    hfs = hfc.strftime("%d/%m/%Y") if hfc is not None else "—"
                    est = "Sí facturar" if hf > 0 else "No facturar"
                    filas.append({"NUI":nui,"Estado de Facturación":est,
                        "Motivo":"Usuario en hurtos - Ticket cerrado en el mes (prorrateo)",
                        "Fuente de decisión":"Hurtos","Factor":hf,
                        "Días Facturables":hdf,"Días del Mes":dt_,"Fecha Cierre Bloqueo":hfs})
            else:
                filas.append({"NUI":nui,"Estado de Facturación":"Sí facturar",
                    "Motivo":"Sin novedades","Fuente de decisión":"Sin coincidencias",
                    "Factor":1.0,"Días Facturables":dt_,"Días del Mes":dt_,
                    "Fecha Cierre Bloqueo":"—"})

    return pd.DataFrame(filas)


# ═════════════════════════════════════════════════════════════════════════════
# EXPORTACIÓN
# ═════════════════════════════════════════════════════════════════════════════

def autoajustar_columnas(ws) -> None:
    """Ajusta el ancho de columnas y aplica estilo de encabezado."""
    from openpyxl.styles import Font, PatternFill, Alignment
    header_fill = PatternFill("solid", fgColor="0D2540")
    header_font = Font(bold=True, color="F5A623")
    for col in ws.columns:
        mx = max((len(str(cell.value or "")) for cell in col), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(mx + 4, 55)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")


def exportar_excel(df: pd.DataFrame, anio: int, mes: int) -> bytes:
    """
    Genera Excel con 4 hojas:
      NO_FACTURAR : factor = 0.0
      FACTURAR    : factor = 1.0  (sin prorrateo)
      PRORRATEO   : 0.0 < factor < 1.0  (todos los casos proporcionales)
      RESUMEN     : conteos y descripción del período
    """
    from openpyxl.styles import Font, PatternFill, Alignment
    import calendar as _cal

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:

        no_fac = df[df["Factor"] == 0.0].reset_index(drop=True)
        si_fac = df[df["Factor"] == 1.0].reset_index(drop=True)
        prorr  = df[df["Factor"].between(0.001, 0.999)].reset_index(drop=True)

        # Ordenar prorrateo: por Factor ascendente (menor primero = más días bloqueados)
        prorr = prorr.sort_values("Factor").reset_index(drop=True)

        # ── Hoja NO_FACTURAR ─────────────────────────────────────────────────
        no_fac.to_excel(w, sheet_name="NO_FACTURAR", index=False)
        autoajustar_columnas(w.sheets["NO_FACTURAR"])

        # ── Hoja FACTURAR ────────────────────────────────────────────────────
        si_fac.to_excel(w, sheet_name="FACTURAR", index=False)
        autoajustar_columnas(w.sheets["FACTURAR"])

        # ── Hoja PRORRATEO ───────────────────────────────────────────────────
        if not prorr.empty:
            prorr.to_excel(w, sheet_name="PRORRATEO", index=False)
            autoajustar_columnas(w.sheets["PRORRATEO"])
        else:
            # Crear hoja vacía con encabezados si no hay casos
            pd.DataFrame(columns=df.columns).to_excel(
                w, sheet_name="PRORRATEO", index=False
            )
            autoajustar_columnas(w.sheets["PRORRATEO"])

        # ── Hoja RESUMEN ─────────────────────────────────────────────────────
        dias_mes = _cal.monthrange(anio, mes)[1]
        MESES = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
                 7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}
        resumen_data = {
            "Concepto": [
                "Período de análisis",
                "Fecha de corte",
                "Días del mes",
                "",
                "Total usuarios",
                "No facturar (factor 0.0)",
                "Sí facturar (factor 1.0)",
                "Prorrateo (factor parcial)",
                "",
                "% No facturar",
                "% Sí facturar",
                "% Prorrateo",
            ],
            "Valor": [
                f"{MESES[mes]} {anio}",
                f"{dias_mes:02d}/{mes:02d}/{anio}",
                dias_mes,
                "",
                len(df),
                len(no_fac),
                len(si_fac),
                len(prorr),
                "",
                f"{len(no_fac)/len(df)*100:.1f}%" if len(df) else "0%",
                f"{len(si_fac)/len(df)*100:.1f}%" if len(df) else "0%",
                f"{len(prorr)/len(df)*100:.1f}%" if len(df) else "0%",
            ],
        }
        if not prorr.empty:
            motivos = prorr.groupby("Motivo").size().reset_index(name="Cantidad")
            resumen_data["Concepto"] += ["", "--- Detalle prorrateo por motivo ---"] +                                          motivos["Motivo"].tolist()
            resumen_data["Valor"]    += ["", ""] + motivos["Cantidad"].tolist()

        df_res = pd.DataFrame(resumen_data)
        df_res.to_excel(w, sheet_name="RESUMEN", index=False)
        autoajustar_columnas(w.sheets["RESUMEN"])

    return buf.getvalue()


# ═════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## 📅 Período de análisis")
    st.markdown("---")

    hoy = date.today()
    col_a, col_m = st.columns(2)
    with col_a:
        anios = list(range(2023, hoy.year + 2))
        anio_sel = st.selectbox("Año", anios, index=anios.index(hoy.year))
    with col_m:
        mes_sel = st.selectbox("Mes", list(range(1,13)), index=hoy.month-1,
                               format_func=lambda x: MESES_ES[x])

    ini_mes  = ini_mes_ts(anio_sel, mes_sel)
    fin_mes  = fin_mes_ts(anio_sel, mes_sel)
    dias_mes = (fin_mes - ini_mes).days + 1

    st.markdown(f"""
    <div style="background:rgba(245,166,35,.1);border:1px solid #F5A623;
    border-radius:6px;padding:.6rem 1rem;margin-top:.5rem;font-size:.82rem;">
    📆 <b>{MESES_ES[mes_sel]} {anio_sel}</b><br>
    Corte: <b>{fin_mes.strftime('%d/%m/%Y')}</b> · <b>{dias_mes} días</b>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("## 📂 Bases de datos")

    st.markdown("#### 👥 Base de Usuarios")
    st.caption("Columna: **NUI**")
    archivo_usuarios = st.file_uploader("usuarios", type=["xlsx","xls"],
                                        key="usuarios", label_visibility="collapsed")
    st.markdown("#### 🎫 Base SAC")
    st.caption("Columnas: **NUI · Semaforo · SubMenu3 · FechaCierre · Concatenado**")
    archivo_sac = st.file_uploader("sac", type=["xlsx","xls"],
                                   key="sac", label_visibility="collapsed")
    st.markdown("#### 🔒 Base Hurtos")
    st.caption("Columnas: **NUI · FechaCierre**")
    archivo_hurtos = st.file_uploader("hurtos", type=["xlsx","xls"],
                                      key="hurtos", label_visibility="collapsed")

    st.markdown("---")
    procesar = st.button("⚡ Procesar Facturación", use_container_width=True, type="primary")

    with st.expander("📋 Lógica de FechaCierre"):
        st.markdown(f"""
**FechaCierre como fuente de verdad**

Para cada ticket se calcula cuándo dejó de estar activo:
- Semáforo abierto → activo hasta fin del mes
- Semáforo cerrado → activo hasta su `FechaCierre`

**Clasificación en el período:**

| Situación | Resultado |
|---|---|
| Activo todo el mes + BLOQUEA | No facturar |
| Cerrado en el mes + BLOQUEA | Prorrateo |
| Activo + DESCUENTO | Sí facturar |
| Cerrado en el mes + NO BLOQUEA + Reposición | Prorrateo |
| Cerrado antes del mes | Sí facturar |
| En hurtos activo todo el mes | No facturar |
| En hurtos cerrado en el mes | Prorrateo |
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

        alertas = []

        # SAC
        df_sac_c = pd.DataFrame(columns=[COL_NUI,"_dec","_factor","_dias_fact","_dias_total","_fc_display"])
        if df_s is not None:
            if validar_columnas(df_s, [COL_NUI,COL_SEMAFORO,COL_SUBMENU3,COL_FC], "Base SAC"):
                df_s[COL_NUI] = normalizar_nui(df_s[COL_NUI])
                df_s = df_s.dropna(subset=[COL_NUI])
                df_sac_c = consolidar_sac(df_s, ini_mes, fin_mes, dias_mes)
                dup = df_s[df_s.duplicated(subset=[COL_NUI],keep=False)][COL_NUI].nunique()
                if dup:
                    alertas.append(f"⚠️ **SAC** — {dup} NUI con múltiples tickets. Se aplicó lógica de prioridad.")

        # Hurtos
        hurtos_info = {}
        if df_h is not None:
            if validar_columnas(df_h, [COL_NUI,COL_FC], "Base Hurtos"):
                df_h[COL_NUI] = normalizar_nui(df_h[COL_NUI])
                df_h = df_h.dropna(subset=[COL_NUI])
                hurtos_info = consolidar_hurtos(df_h, ini_mes, fin_mes, dias_mes)
                dup = df_h[df_h.duplicated(subset=[COL_NUI],keep=False)][COL_NUI].nunique()
                if dup:
                    alertas.append(f"⚠️ **Hurtos** — {dup} NUI duplicados. Se consolidó por NUI.")

        # Cruce SAC × Hurtos
        nui_sac_bloq = set(df_sac_c[df_sac_c["_dec"].isin(["BLOQUEA","BLOQUEA_PARCIAL"])]["NUI"])
        nui_hurtos   = set(hurtos_info.keys())
        cruce = nui_hurtos & nui_sac_bloq
        if cruce:
            alertas.append(f"🔀 **Cruce SAC × Hurtos** — {len(cruce)} NUI en ambas bases. SAC tiene prioridad.")

        df_resultado = aplicar_reglas(df_u, df_sac_c, hurtos_info, dias_mes)
        if df_resultado.empty:
            st.error("No se pudo generar resultado."); st.stop()

    st.session_state.update({
        "df_resultado": df_resultado, "alertas": alertas,
        "anio_sel": anio_sel, "mes_sel": mes_sel,
        "ini_mes": ini_mes, "fin_mes": fin_mes, "dias_mes": dias_mes,
    })
    st.success(f"✅ **{MESES_ES[mes_sel]} {anio_sel}** procesado.")


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

    st.markdown(f"""
    <div class="period-box">
      📅 Análisis: <strong>{MESES_ES[mes_vis]} {anio_vis}</strong> &nbsp;·&nbsp;
      Corte: <strong>{fin_vis.strftime('%d/%m/%Y')}</strong> &nbsp;·&nbsp;
      <strong>{dias_vis} días</strong> en el mes
    </div>
    """, unsafe_allow_html=True)

    total   = len(df_res)
    no_n    = (df_res["Estado de Facturación"]=="No facturar").sum()
    si_n    = (df_res["Estado de Facturación"]=="Sí facturar").sum()
    prorr_n = df_res["Factor"].between(0.001, 0.999).sum()
    pct_si  = si_n/total*100 if total else 0
    pct_no  = no_n/total*100 if total else 0

    c1,c2,c3,c4,c5,c6 = st.columns(6)
    c1.metric("👥 Total",       f"{total:,}")
    c2.metric("✅ Sí facturar",  f"{si_n:,}")
    c3.metric("🚫 No facturar",  f"{no_n:,}")
    c4.metric("📊 Prorrateo",    f"{prorr_n:,}")
    c5.metric("% Facturados",   f"{pct_si:.1f}%")
    c6.metric("% No facturados",f"{pct_no:.1f}%")

    if alertas:
        st.markdown("---")
        with st.expander(f"⚠️ Alertas de calidad ({len(alertas)})", expanded=True):
            for a in alertas:
                st.markdown(f'<div class="warn-box">{a}</div>', unsafe_allow_html=True)

    st.markdown("---")
    with st.expander("📊 Detalle por motivo de decisión"):
        res = (df_res.groupby(["Estado de Facturación","Motivo","Fuente de decisión"])
               .size().reset_index(name="Cantidad")
               .sort_values("Cantidad", ascending=False))
        st.dataframe(res, use_container_width=True, hide_index=True)

    st.markdown("---")
    cf1,cf2,cf3 = st.columns([2,2,3])
    with cf1:
        f_est = st.selectbox("Estado",["Todos","Sí facturar","No facturar"])
    with cf2:
        f_fac = st.selectbox("Tipo",["Todos","Factor completo (1.0)","Prorrateo (<1.0)","Sin facturación (0.0)"])
    with cf3:
        buscar = st.text_input("🔍 Buscar NUI", placeholder="Ingresa un NUI…")

    dv = df_res.copy()
    if f_est != "Todos": dv = dv[dv["Estado de Facturación"]==f_est]
    if f_fac == "Factor completo (1.0)":   dv = dv[dv["Factor"]==1.0]
    elif f_fac == "Prorrateo (<1.0)":       dv = dv[dv["Factor"].between(0.001,0.999)]
    elif f_fac == "Sin facturación (0.0)":  dv = dv[dv["Factor"]==0.0]
    if buscar.strip(): dv = dv[dv["NUI"].str.contains(buscar.strip().upper(), na=False)]

    def color_fila(row):
        if row["Factor"] == 0.0:   return ["background-color:rgba(231,76,60,.2)"]*len(row)
        elif row["Factor"] < 1.0:  return ["background-color:rgba(243,156,18,.15)"]*len(row)
        return [""]*len(row)

    def color_estado(val):
        if val=="Sí facturar": return "color:#2ECC71;font-weight:700"
        if val=="No facturar": return "color:#E74C3C;font-weight:700"
        return ""

    st.markdown(f"**{len(dv):,} registros** mostrados")
    st.dataframe(
        dv.style.apply(color_fila, axis=1).map(color_estado, subset=["Estado de Facturación"]),
        use_container_width=True, height=440,
        column_config={"Factor": st.column_config.NumberColumn("Factor", format="%.4f")},
    )

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
        <p>Selecciona el <strong>año y mes</strong>, carga las bases y presiona
        <strong>Procesar Facturación</strong>.</p>
    </div>
    """, unsafe_allow_html=True)
