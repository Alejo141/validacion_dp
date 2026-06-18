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
COL_SUBMENU1 = "SubMenu1"
COL_SUBMENU2 = "SubMenu2"
COL_SECCIONAL = "NombreSeccionales"
COL_ID_TICKET = "Id_Tickets"
COL_CANAL     = "canal"
COL_NOMBRE    = "Nombre_Completo"
COL_CEDULA    = "Cedula"
COL_FC       = "FechaCierre"
COL_FECHA_CREACION = "FechaCreacion"
EXTRA_COLS_TICKET = [COL_ID_TICKET, COL_CANAL, COL_FECHA_CREACION, COL_NOMBRE, COL_CEDULA]
DATE_FMT     = "%d-%m-%Y"

SEMAFOROS_ABIERTOS = {"CRITICO", "MODERADO", "LEVE"}
SUBMENU_BLOQUEA    = "BLOQUEA FACTURACION"
SUBMENU_DESCUENTO  = "DESCUENTO COMERCIAL"
SUBMENU_NO_BLOQUEA = "NO BLOQUEA FACTURACION"
COL_CONCAT         = "Concatenado"
REPOSICION_KEYWORD = "REPOSICI"   # cubre Reposición / REPOSICION / reposicion
DANO_KEYWORD        = "DANO"       # SubMenu1 normalizado (sin tilde) para "DAÑO"
SUBMENU1_HURTO     = "HURTO SOLUCION"  # identifica tickets de hurto en SAC

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

def extraer_datos_ticket(row) -> dict:
    """Extrae los campos de identificación del ticket SAC para reportes."""
    return {c: str(row.get(c, "")) for c in EXTRA_COLS_TICKET}


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
    activos_bloquea   = []  # lista de (fc_inicio_bloqueo, fc_fin_bloqueo, submenu2)
    activos_descuento = []  # lista de (fc_fin, submenu2)
    activos_repos     = []  # lista de (fc_fin, submenu2) NO BLOQUEA + REPOSICION
    activos_dano      = []  # lista de (fc_inicio, fc_fin, submenu2, ticket_data) NO BLOQUEA + DAÑO (técnico)

    for _, r in grupo.iterrows():
        sub    = r[COL_SUBMENU3]
        sub1   = quitar_tildes(str(r.get(COL_SUBMENU1, "")).strip().upper())
        sub2   = str(r.get(COL_SUBMENU2, "")).strip()
        concat = str(r.get(COL_CONCAT, "")).upper()
        sem    = r[COL_SEMAFORO]

        # Ignorar tickets creados DESPUÉS del último día del mes analizado
        fcreac_raw = r.get(col_creacion, None)
        if fcreac_raw is not None:
            fc_creac = parsear_fechas(pd.Series([fcreac_raw])).iloc[0]
            if pd.notna(fc_creac) and fc_creac > fin:
                continue  # ticket futuro, no aplica al período

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
            activos_bloquea.append((inicio_bloqueo, fce, sub2, extraer_datos_ticket(r)))

        elif sub == SUBMENU_DESCUENTO:
            activos_descuento.append((fce, sub2, extraer_datos_ticket(r)))

        elif sub == SUBMENU_NO_BLOQUEA:
            # Caso A: contiene "Reposición" en Concatenado → prorrateo si cerrado en mes
            if fce < fin and REPOSICION_KEYWORD in concat:
                activos_repos.append((fce, sub2, extraer_datos_ticket(r)))
            # Caso B: SubMenu1 = DAÑO (soporte técnico) cerrado dentro del mes
            # → puede haber estado abierto desde meses anteriores; prorratea
            # igual que BLOQUEA: desde ini_mes (o FechaCreacion si es del mes) hasta FechaCierre
            elif fce < fin and sub1 == DANO_KEYWORD:
                fcreac = parsear_fechas(pd.Series([r.get(col_creacion, "")])).iloc[0]
                if pd.notna(fcreac) and fcreac > ini:
                    inicio_dano = max(fcreac, ini)
                else:
                    inicio_dano = ini
                activos_dano.append((inicio_dano, fce, sub2, extraer_datos_ticket(r)))
            # Caso contrario (abierto, o sin Reposición/Daño): se ignora

    # ── Prioridad 1: BLOQUEA ────────────────────────────────────────────────
    if activos_bloquea:
        # Calcular días TOTALES bloqueados en el mes considerando
        # el inicio más temprano y el fin más tardío de todos los tickets BLOQUEA.
        # Si cualquier ticket empezó antes o en ini_mes y sigue abierto → bloqueo completo.
        inicio_min = min(t[0] for t in activos_bloquea)  # inicio de bloqueo más temprano
        fc_max     = max(t[1] for t in activos_bloquea)  # fin de bloqueo más tardío
        # SubMenu2 del ticket con el cierre más tardío (el más representativo)
        submenu2    = next(t[2] for t in activos_bloquea if t[1] == fc_max)
        ticket_data = next(t[3] for t in activos_bloquea if t[1] == fc_max)

        if inicio_min <= ini and fc_max >= fin:
            # Bloqueado todo el mes
            return {"_dec":"BLOQUEA", "_factor":0.0,
                    "_dias_fact":0, "_dias_total":dias_mes, "_fc_display":None,
                    "_submenu2":submenu2, "_ticket_data":ticket_data}
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
                    "_fc_display":fin_bloq_efectivo, "_submenu2":submenu2,
                    "_ticket_data":ticket_data}

    # ── Prioridad 2: DESCUENTO ──────────────────────────────────────────────
    if activos_descuento:
        fc_max      = max(t[0] for t in activos_descuento)
        submenu2    = next(t[1] for t in activos_descuento if t[0] == fc_max)
        ticket_data = next(t[2] for t in activos_descuento if t[0] == fc_max)
        if fc_max >= fin:
            return {"_dec":"DESCUENTO", "_factor":1.0,
                    "_dias_fact":dias_mes, "_dias_total":dias_mes, "_fc_display":None,
                    "_submenu2":submenu2, "_ticket_data":ticket_data}
        else:
            f, df_, dt_ = calcular_prorrateo(fc_max, ini, fin, dias_mes)
            return {"_dec":"DESCUENTO_PARCIAL", "_factor":f,
                    "_dias_fact":df_, "_dias_total":dt_, "_fc_display":fc_max,
                    "_submenu2":submenu2, "_ticket_data":ticket_data}

    # ── Prioridad 3: NO BLOQUEA + REPOSICION cerrado en el mes → prorrateo ─
    if activos_repos:
        # Tomar el cierre más tardío (mayor días bloqueados)
        fc_max      = max(t[0] for t in activos_repos)
        submenu2    = next(t[1] for t in activos_repos if t[0] == fc_max)
        ticket_data = next(t[2] for t in activos_repos if t[0] == fc_max)
        f, df_, dt_ = calcular_prorrateo(fc_max, ini, fin, dias_mes)
        return {"_dec":"REPOSICION_PARCIAL", "_factor":f,
                "_dias_fact":df_, "_dias_total":dt_, "_fc_display":fc_max,
                "_submenu2":submenu2, "_ticket_data":ticket_data}

    # ── Prioridad 4: NO BLOQUEA + DAÑO (técnico) cerrado en el mes → prorrateo
    # Puede haber estado abierto desde meses anteriores; se prorratea igual
    # que BLOQUEA: desde el inicio efectivo del bloqueo hasta FechaCierre.
    if activos_dano:
        inicio_min  = min(t[0] for t in activos_dano)
        fc_max      = max(t[1] for t in activos_dano)
        submenu2    = next(t[2] for t in activos_dano if t[1] == fc_max)
        ticket_data = next(t[3] for t in activos_dano if t[1] == fc_max)

        inicio_bloq_efectivo = max(inicio_min, ini)
        fin_bloq_efectivo    = min(fc_max, fin)
        dias_bloqueados      = (fin_bloq_efectivo - inicio_bloq_efectivo).days + 1
        dias_fact            = dias_mes - dias_bloqueados
        if dias_fact < 0: dias_fact = 0
        factor = round(dias_fact / dias_mes, 6)
        return {"_dec":"DANO_PARCIAL", "_factor":factor,
                "_dias_fact":int(dias_fact), "_dias_total":dias_mes,
                "_fc_display":fin_bloq_efectivo, "_submenu2":submenu2,
                "_ticket_data":ticket_data}

    # ── Sin tickets relevantes activos → factura completo ──────────────────
    return {"_dec":"SIN_NOVEDAD_SAC", "_factor":1.0,
            "_dias_fact":dias_mes, "_dias_total":dias_mes, "_fc_display":None,
            "_submenu2":"", "_ticket_data":{}}


def clasificar_nui_hurtos_con_sac(nui: str,
                                   df_sac_nui: pd.DataFrame,
                                   df_hurto_nui: pd.DataFrame,
                                   ini: pd.Timestamp,
                                   fin: pd.Timestamp,
                                   dias_mes: int) -> dict:
    """
    Determina la facturación de un NUI que aparece en la base de hurtos,
    cruzando con sus tickets SAC para buscar un ticket de reposición.

    Regla:
      1. Si NO existe ticket con "REPOSICION" en Concatenado (en SAC) → No facturar
      2. Si existe ticket de reposición pero su FechaCierre es ANTERIOR a la
         FechaCreacion del hurto → se ignora (reposición de un evento previo)
      3. Si existe ticket de reposición ABIERTO (Semáforo abierto) → No facturar
      4. Si existe ticket de reposición CERRADO:
           a. FechaCierre antes del mes → Sí facturar (mes completo)
           b. FechaCierre dentro del mes → Prorrateo
           c. FechaCierre después del mes → No facturar (cerró luego)

    "Reposición" se identifica por la palabra "REPOSICI" en el campo Concatenado.
    La reposición solo es válida si su FechaCierre > FechaCreacion del hurto.
    """
    REPOS = REPOSICION_KEYWORD

    # Datos de respaldo tomados de la base de Hurtos (cuando SAC no aporta ticket)
    if len(df_hurto_nui):
        fila_h = df_hurto_nui.iloc[0]
        hurto_ticket_data = {c: str(fila_h.get(c, "")) for c in EXTRA_COLS_TICKET}
        hurto_submenu2    = str(fila_h.get(COL_SUBMENU2, "")).strip()
        hurto_seccional   = str(fila_h.get(COL_SECCIONAL, "")).strip()
    else:
        hurto_ticket_data = {}
        hurto_submenu2    = ""
        hurto_seccional   = ""

    # Calcular la FechaCierre más reciente del hurto como referencia.
    # La reposición solo es válida si su FechaCierre es POSTERIOR a la
    # FechaCierre del hurto más reciente (indica que el equipo fue repuesto
    # después de haber sido reportado como hurtado/cerrado el caso).
    fc_hurto_cierre = pd.NaT
    if len(df_hurto_nui):
        fc_hurto_col = parsear_fechas(df_hurto_nui[COL_FC])             if COL_FC in df_hurto_nui.columns             else pd.Series([], dtype="datetime64[ns]")
        validas = fc_hurto_col.dropna()
        if len(validas):
            fc_hurto_cierre = validas.max()

    # Filtrar tickets del NUI en SAC que tengan "REPOSICION" en Concatenado
    # Excluir tickets con FechaCreacion posterior al fin del mes
    df_sac_valido = df_sac_nui[
        df_sac_nui["_fc_creac"].isna() | (df_sac_nui["_fc_creac"] <= fin)
    ]
    tickets_repos_raw = df_sac_valido[
        df_sac_valido["_concat_upper"].str.contains(REPOS, na=False)
    ]

    # Filtrar reposiciones cuya FechaCierre sea POSTERIOR a la FechaCierre del hurto.
    # Una reposición anterior al cierre del hurto corresponde a un evento distinto.
    if pd.notna(fc_hurto_cierre) and len(tickets_repos_raw):
        tickets_repos = tickets_repos_raw[
            tickets_repos_raw["_fc"].isna() |            # sin fecha: no se descarta aún
            (tickets_repos_raw["_fc"] > fc_hurto_cierre) # posterior al cierre del hurto
        ]
    else:
        tickets_repos = tickets_repos_raw

    # ── Sin ticket de reposición → No facturar ────────────────────────────
    if tickets_repos.empty:
        return {"_factor": 0.0, "_dias_fact": 0, "_tipo": "SIN_REPOSICION", "_fc": None,
                "_submenu2": hurto_submenu2, "_ticket_data": hurto_ticket_data,
                "_seccional": hurto_seccional}

    # ── Evaluar estado de los tickets de reposición ───────────────────────
    # Semáforo normalizado (sin tildes, mayúsculas)
    abiertos  = tickets_repos[tickets_repos["_semaforo_n"].isin(SEMAFOROS_ABIERTOS)]
    cerrados  = tickets_repos[~tickets_repos["_semaforo_n"].isin(SEMAFOROS_ABIERTOS)]

    # SubMenu2 representativo (del primer ticket de reposición encontrado)
    submenu2_repos = str(tickets_repos.iloc[0].get(COL_SUBMENU2, "")).strip()
    ticket_repos   = extraer_datos_ticket(tickets_repos.iloc[0])

    # Si hay reposición abierta → No facturar (aunque también haya cerradas)
    if not abiertos.empty:
        sub2_ab = str(abiertos.iloc[0].get(COL_SUBMENU2, "")).strip()
        ticket_ab = extraer_datos_ticket(abiertos.iloc[0])
        return {"_factor": 0.0, "_dias_fact": 0, "_tipo": "REPOS_ABIERTA", "_fc": None,
                "_submenu2": sub2_ab, "_ticket_data": ticket_ab, "_seccional": hurto_seccional}

    # Solo reposiciones cerradas → evaluar FechaCierre
    if cerrados.empty:
        return {"_factor": 0.0, "_dias_fact": 0, "_tipo": "SIN_REPOSICION", "_fc": None,
                "_submenu2": hurto_submenu2, "_ticket_data": hurto_ticket_data,
                "_seccional": hurto_seccional}

    # Tomar la FechaCierre más reciente entre las reposiciones cerradas
    fc_max = cerrados["_fc"].dropna().max()
    fila_max = cerrados[cerrados["_fc"] == fc_max]
    if len(fila_max):
        sub2_cerr   = str(fila_max.iloc[0].get(COL_SUBMENU2, "")).strip()
        ticket_cerr = extraer_datos_ticket(fila_max.iloc[0])
    else:
        sub2_cerr, ticket_cerr = submenu2_repos, ticket_repos

    if pd.isna(fc_max):
        # Cerrado sin fecha → No facturar por precaución
        return {"_factor": 0.0, "_dias_fact": 0, "_tipo": "REPOS_SIN_FECHA", "_fc": None,
                "_submenu2": submenu2_repos, "_ticket_data": ticket_repos,
                "_seccional": hurto_seccional}

    if fc_max < ini:
        # Reposición cerrada ANTES del mes → factura completo
        return {"_factor": 1.0, "_dias_fact": dias_mes, "_tipo": "REPOS_CERRADA_ANTES", "_fc": fc_max,
                "_submenu2": sub2_cerr, "_ticket_data": ticket_cerr, "_seccional": hurto_seccional}

    if fc_max <= fin:
        # Reposición cerrada DENTRO del mes → prorrateo
        f, df_, _ = calcular_prorrateo(fc_max, ini, fin, dias_mes)
        return {"_factor": f, "_dias_fact": df_, "_tipo": "REPOS_CERRADA_EN_MES", "_fc": fc_max,
                "_submenu2": sub2_cerr, "_ticket_data": ticket_cerr, "_seccional": hurto_seccional}

    # Reposición cerrada DESPUÉS del mes → No facturar (aún estaba abierta en el mes)
    return {"_factor": 0.0, "_dias_fact": 0, "_tipo": "REPOS_CERRADA_DESPUES", "_fc": fc_max,
            "_submenu2": sub2_cerr, "_ticket_data": ticket_cerr, "_seccional": hurto_seccional}


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
    # Asegurar columna SubMenu2 (opcional)
    if COL_SUBMENU2 not in df.columns:
        df[COL_SUBMENU2] = ""
    else:
        df[COL_SUBMENU2] = df[COL_SUBMENU2].fillna("").astype(str)
    # Asegurar columna SubMenu1 (opcional, usada para identificar DAÑO)
    if COL_SUBMENU1 not in df.columns:
        df[COL_SUBMENU1] = ""
    else:
        df[COL_SUBMENU1] = df[COL_SUBMENU1].fillna("").astype(str)
    # Asegurar columnas extra para reporte de tickets (opcionales)
    for extra_col in EXTRA_COLS_TICKET:
        if extra_col not in df.columns:
            df[extra_col] = ""
        else:
            df[extra_col] = df[extra_col].fillna("").astype(str)
    df = df.dropna(subset=[COL_NUI])

    filas = []
    for nui, grupo in df.groupby(COL_NUI):
        info = clasificar_nui_sac(grupo, ini, fin, dias_mes,
                                  col_creacion=COL_FECHA_CREACION)
        info["NUI"] = nui
        # Capturar NombreSeccionales (mismo valor para todo el NUI)
        if COL_SECCIONAL in grupo.columns:
            vals = grupo[COL_SECCIONAL].dropna()
            info["_seccional"] = vals.iloc[0] if len(vals) else ""
        else:
            info["_seccional"] = ""
        filas.append(info)
    return pd.DataFrame(filas)


def consolidar_hurtos(df_h: pd.DataFrame,
                      df_sac: pd.DataFrame,
                      ini: pd.Timestamp,
                      fin: pd.Timestamp,
                      dias_mes: int) -> dict:
    """
    Consolida la base de hurtos cruzando con SAC para verificar
    la existencia y estado del ticket de reposición por NUI.
    """
    df = df_h.copy()
    df[COL_NUI] = normalizar_nui(df[COL_NUI])
    df = df.dropna(subset=[COL_NUI])

    # Asegurar columnas de respaldo desde hurtos
    for extra_col in EXTRA_COLS_TICKET + [COL_SUBMENU2, COL_SECCIONAL]:
        if extra_col not in df.columns:
            df[extra_col] = ""
        else:
            df[extra_col] = df[extra_col].fillna("").astype(str)

    # Preparar SAC con columnas normalizadas necesarias
    sac = df_sac.copy()
    sac["_concat_upper"]  = sac[COL_CONCAT].fillna("").astype(str).str.upper()
    sac["_semaforo_n"]    = normalizar_texto(sac[COL_SEMAFORO])
    sac["_fc"]            = parsear_fechas(sac[COL_FC])
    sac["_fc_creac"]      = parsear_fechas(sac[COL_FECHA_CREACION])
    if COL_SUBMENU2 not in sac.columns:
        sac[COL_SUBMENU2] = ""
    else:
        sac[COL_SUBMENU2] = sac[COL_SUBMENU2].fillna("").astype(str)
    for extra_col in EXTRA_COLS_TICKET:
        if extra_col not in sac.columns:
            sac[extra_col] = ""
        else:
            sac[extra_col] = sac[extra_col].fillna("").astype(str)

    resultado = {}
    for nui in df[COL_NUI].unique():
        df_sac_nui   = sac[sac[COL_NUI] == nui]
        df_hurto_nui = df[df[COL_NUI] == nui]
        resultado[nui] = clasificar_nui_hurtos_con_sac(
            nui, df_sac_nui, df_hurto_nui, ini, fin, dias_mes
        )
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
    if "_submenu2" not in df.columns:
        df["_submenu2"] = ""
    df["_submenu2"] = df["_submenu2"].fillna("")
    if "_seccional" not in df.columns:
        df["_seccional"] = ""
    df["_seccional"] = df["_seccional"].fillna("")
    if "_ticket_data" not in df.columns:
        df["_ticket_data"] = [dict() for _ in range(len(df))]
    df["_ticket_data"] = df["_ticket_data"].apply(lambda x: x if isinstance(x, dict) else {})

    def _extra_cols(ticket_data: dict, incluir: bool) -> dict:
        """Retorna las 5 columnas extra de tickets; vacías si incluir=False."""
        if not incluir:
            return {c: "" for c in EXTRA_COLS_TICKET}
        return {c: str(ticket_data.get(c, "")) for c in EXTRA_COLS_TICKET}

    filas = []
    for _, row in df.iterrows():
        nui  = row[COL_NUI]
        dec  = row["_dec"]
        fac  = float(row["_factor"])
        df_  = int(row["_dias_fact"])
        dt_  = int(row["_dias_total"])
        fcd  = row["_fc_display"]
        fcs  = fcd.strftime("%d/%m/%Y") if pd.notna(fcd) and fcd is not None else "—"
        sub2 = str(row["_submenu2"]).strip()
        seccional = str(row["_seccional"]).strip()
        tdata = row["_ticket_data"]

        # ── Regla 1: SAC BLOQUEA completo ────────────────────────────────────
        if dec == "BLOQUEA":
            filas.append({"NUI":nui,"Estado de Facturación":"No facturar",
                "Motivo":"Ticket abierto - Bloquea facturación",
                "Fuente de decisión":"SAC","SubMenu2":sub2,"NombreSeccionales":seccional,
                **_extra_cols(tdata, True),
                "Factor":0.0,
                "Días Facturables":0,"Días del Mes":dt_,"Fecha Cierre Bloqueo":"—"})

        # ── Regla 1P: SAC BLOQUEA parcial (prorrateo) ────────────────────────
        elif dec == "BLOQUEA_PARCIAL":
            estado = "Sí facturar" if fac > 0 else "No facturar"
            filas.append({"NUI":nui,"Estado de Facturación":estado,
                "Motivo":"Ticket cerrado en el mes - Bloquea facturación (prorrateo)",
                "Fuente de decisión":"SAC","SubMenu2":sub2,"NombreSeccionales":seccional,
                **_extra_cols(tdata, True),
                "Factor":fac,
                "Días Facturables":df_,"Días del Mes":dt_,"Fecha Cierre Bloqueo":fcs})

        # ── Regla 2: SAC DESCUENTO ───────────────────────────────────────────
        elif dec in ("DESCUENTO","DESCUENTO_PARCIAL"):
            motivo = ("Ticket abierto - Descuento comercial" if dec=="DESCUENTO"
                      else "Ticket cerrado en el mes - Descuento comercial (prorrateo)")
            filas.append({"NUI":nui,"Estado de Facturación":"Sí facturar",
                "Motivo":motivo,"Fuente de decisión":"SAC","SubMenu2":"","NombreSeccionales":seccional,
                **_extra_cols(tdata, False),
                "Factor":fac,
                "Días Facturables":df_,"Días del Mes":dt_,"Fecha Cierre Bloqueo":fcs})

        # ── Regla 2R: NO BLOQUEA + REPOSICION cerrado en mes → prorrateo ────
        elif dec == "REPOSICION_PARCIAL":
            estado = "Sí facturar" if fac > 0 else "No facturar"
            filas.append({"NUI":nui,"Estado de Facturación":estado,
                "Motivo":"Ticket cerrado en el mes - No bloquea / Reposición (prorrateo)",
                "Fuente de decisión":"SAC","SubMenu2":sub2,"NombreSeccionales":seccional,
                **_extra_cols(tdata, True),
                "Factor":fac,
                "Días Facturables":df_,"Días del Mes":dt_,"Fecha Cierre Bloqueo":fcs})

        # ── Regla 2D: NO BLOQUEA + DAÑO (técnico) cerrado en mes → prorrateo ─
        elif dec == "DANO_PARCIAL":
            estado = "Sí facturar" if fac > 0 else "No facturar"
            filas.append({"NUI":nui,"Estado de Facturación":estado,
                "Motivo":"Ticket cerrado en el mes - No bloquea / Daño técnico (prorrateo)",
                "Fuente de decisión":"SAC","SubMenu2":sub2,"NombreSeccionales":seccional,
                **_extra_cols(tdata, True),
                "Factor":fac,
                "Días Facturables":df_,"Días del Mes":dt_,"Fecha Cierre Bloqueo":fcs})

        # ── Reglas 3 y 4: sin decisión SAC → revisar hurtos ──────────────────
        else:
            if nui in hurtos_info:
                h = hurtos_info[nui]
                tipo = h["_tipo"]
                hf   = h["_factor"]
                hdf  = h["_dias_fact"]
                hfc  = h["_fc"]
                hfs  = hfc.strftime("%d/%m/%Y") if hfc is not None else "—"
                hsub2= str(h.get("_submenu2","")).strip()
                htdata = h.get("_ticket_data", {}) or {}
                h_seccional = str(h.get("_seccional","")).strip()
                seccional_final = h_seccional if h_seccional else seccional

                # Motivos por tipo de resolución
                _motivos = {
                    "SIN_REPOSICION":        "Hurto sin ticket de reposición",
                    "REPOS_ABIERTA":         "Hurto - Reposición abierta",
                    "REPOS_SIN_FECHA":       "Hurto - Reposición sin fecha de cierre",
                    "REPOS_CERRADA_DESPUES": "Hurto - Reposición cerrada fuera del mes",
                    "COMPLETO":              "Usuario reportado en hurtos",
                    "REPOS_CERRADA_ANTES":   "Hurto con reposición cerrada (factura completo)",
                    "REPOS_CERRADA_EN_MES":  "Hurto con reposición cerrada en el mes (prorrateo)",
                    "PARCIAL":               "Usuario en hurtos - Ticket cerrado en el mes (prorrateo)",
                }
                motivo = _motivos.get(tipo, "Usuario reportado en hurtos")
                est    = "Sí facturar" if hf > 0 else "No facturar"
                no_fac = est=="No facturar"
                es_prorrateo = 0.0 < hf < 1.0
                incluir = no_fac or es_prorrateo

                filas.append({"NUI":nui,"Estado de Facturación":est,
                    "Motivo":motivo,
                    "Fuente de decisión":"Hurtos",
                    "SubMenu2":hsub2 if incluir else "","NombreSeccionales":seccional_final,
                    **_extra_cols(htdata, incluir),
                    "Factor":hf,
                    "Días Facturables":hdf,"Días del Mes":dt_,"Fecha Cierre Bloqueo":hfs})
            else:
                filas.append({"NUI":nui,"Estado de Facturación":"Sí facturar",
                    "Motivo":"Sin novedades","Fuente de decisión":"Sin coincidencias",
                    "SubMenu2":"","NombreSeccionales":seccional,
                    **_extra_cols({}, False),
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
    st.caption("Columnas: **NUI · Semaforo · SubMenu2 · SubMenu3 · FechaCierre · Concatenado**")
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
                hurtos_info = consolidar_hurtos(df_h, df_s, ini_mes, fin_mes, dias_mes)
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
