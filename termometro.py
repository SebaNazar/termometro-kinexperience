"""
Termómetro Kinexperience — Motor de cálculo mensual
Uso: python3 termometro.py
"""

import os
import sys
import json
from datetime import datetime, date, timedelta
import calendar

import unicodedata

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

load_dotenv()

# ── CONSTANTES ─────────────────────────────────────────────────────────────────

REGISTRO_ID      = "1kQgC5koSq-tgsP7W2Bxah7ilLUZrEgp6tY12XNByN-s"
PESTAÑA_REGISTRO = "Respuestas de formulario 1"

COL_ESTADO       = "Estado de la sesión"
COL_FECHA        = "Fecha de la sesión realizada"
COL_KINE         = "Nombre del Kinesiólogo"
COL_PACIENTE     = "Nombre del Paciente"
COL_TIMESTAMP    = "Marca temporal"

ESTADOS_VALIDOS  = {"Realizada", "Suspendida", "Recuperada",
                    "Evaluación de ingreso", "Sesión Grupal"}

KINE_EXCLUIDO    = "Mauricio Arce"   # socio, nunca entra en métricas grupales

# ── CONEXIÓN ───────────────────────────────────────────────────────────────────

def conectar() -> gspread.Client:
    credentials_path = os.getenv("GOOGLE_CREDENTIALS_PATH")
    if not credentials_path or not os.path.exists(credentials_path):
        sys.exit("ERROR: GOOGLE_CREDENTIALS_PATH no encontrado. Revisa tu .env")
    creds = Credentials.from_service_account_file(
        credentials_path,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )
    return gspread.authorize(creds)


# ── CARGA DE DATOS ─────────────────────────────────────────────────────────────

def cargar_registro(gc: gspread.Client) -> pd.DataFrame:
    print("Conectando al Registro de Sesiones...")
    sh = gc.open_by_key(REGISTRO_ID)
    ws = sh.worksheet(PESTAÑA_REGISTRO)
    datos = ws.get_all_records()
    df = pd.DataFrame(datos)
    df.columns = df.columns.str.strip()  # elimina espacios en nombres de columna
    if COL_PACIENTE in df.columns:
        df[COL_PACIENTE] = (
            df[COL_PACIENTE]
            .astype(str)
            .str.strip()
            .str.upper()
            .apply(lambda x: unicodedata.normalize("NFD", x)
                             .encode("ascii", "ignore")
                             .decode("ascii"))
        )
    print(f"  {len(df)} filas cargadas.")
    return df


def cargar_config() -> dict:
    config_path = os.path.join(os.path.dirname(__file__), "config_mes.json")
    if not os.path.exists(config_path):
        sys.exit("ERROR: config_mes.json no encontrado.")
    with open(config_path, encoding="utf-8") as f:
        config = json.load(f)
    return config


# ── FILTRADO Y LIMPIEZA ────────────────────────────────────────────────────────

MESES_ES = {
    "Enero": 1, "Febrero": 2, "Marzo": 3, "Abril": 4,
    "Mayo": 5, "Junio": 6, "Julio": 7, "Agosto": 8,
    "Septiembre": 9, "Octubre": 10, "Noviembre": 11, "Diciembre": 12,
}


def filtrar_mes(df: pd.DataFrame, config: dict) -> pd.DataFrame:
    mes_num = MESES_ES[config["mes"]]
    anio    = config["año"]

    # Parsear fecha de sesión
    df = df.copy()
    df["_fecha"] = pd.to_datetime(df[COL_FECHA], dayfirst=True, errors="coerce")
    df["_ts"]    = pd.to_datetime(df[COL_TIMESTAMP], dayfirst=True, errors="coerce")

    mascara = (df["_fecha"].dt.month == mes_num) & (df["_fecha"].dt.year == anio)
    df_mes  = df[mascara].copy()
    print(f"  {len(df_mes)} sesiones en {config['mes']} {anio}.")
    return df_mes


def detectar_fechas_sospechosas(df: pd.DataFrame) -> pd.DataFrame:
    """
    Marca filas donde la fecha de sesión difiere en mes o año de la Marca temporal.
    NO corrige: solo alerta.
    """
    sospechosas = df[
        (df["_fecha"].notna()) & (df["_ts"].notna()) & (
            (df["_fecha"].dt.month != df["_ts"].dt.month) |
            (df["_fecha"].dt.year  != df["_ts"].dt.year)
        )
    ].copy()
    return sospechosas


# ── CÁLCULO DE CAPACIDAD ───────────────────────────────────────────────────────

def dias_habiles_rango(inicio: date, fin: date) -> int:
    """Cuenta días hábiles (lunes a viernes) entre inicio y fin, ambos inclusive."""
    count = 0
    current = inicio
    while current <= fin:
        if current.weekday() < 5:  # 0=lun … 4=vie
            count += 1
        current += timedelta(days=1)
    return count


def obtener_rango_sesiones_kine(
    kine: str, excepciones: dict, mes_inicio: date, mes_fin: date
) -> tuple:
    """
    Retorna (inicio, fin) del rango válido de sesiones para el kine.
    Para excepciones de porcentaje no hay restricción de fechas.
    """
    exc = excepciones.get(kine)
    if exc:
        if exc["tipo"] == "fecha_hasta":
            return mes_inicio, date.fromisoformat(exc["fecha"])
        elif exc["tipo"] == "fecha_desde":
            return date.fromisoformat(exc["fecha"]), mes_fin
    return mes_inicio, mes_fin


def calcular_capacidad(
    kine: str, dias_habiles: int, excepciones: dict,
    mes_inicio: date, mes_fin: date
) -> float:
    exc = excepciones.get(kine)
    if exc:
        if exc["tipo"] == "porcentaje":
            return round(dias_habiles * 5.0 * exc["valor"], 2)
        elif exc["tipo"] == "fecha_hasta":
            fecha = date.fromisoformat(exc["fecha"])
            dh = dias_habiles_rango(mes_inicio, min(fecha, mes_fin))
            return float(dh * 5)
        elif exc["tipo"] == "fecha_desde":
            fecha = date.fromisoformat(exc["fecha"])
            dh = dias_habiles_rango(max(fecha, mes_inicio), mes_fin)
            return float(dh * 5)
    return dias_habiles * 5.0


# ── MÉTRICAS POR KINE ──────────────────────────────────────────────────────────

def calcular_metricas_kine(df_kine: pd.DataFrame, kine: str,
                            dias_habiles: int, excepciones: dict,
                            mes_inicio: date, mes_fin: date) -> dict:
    realizadas          = (df_kine[COL_ESTADO] == "Realizada").sum()
    suspendidas         = (df_kine[COL_ESTADO] == "Suspendida").sum()
    recuperadas         = (df_kine[COL_ESTADO] == "Recuperada").sum()
    evaluaciones        = (df_kine[COL_ESTADO] == "Evaluación de ingreso").sum()
    grupales_raw        = (df_kine[COL_ESTADO] == "Sesión Grupal").sum()
    grupales_ponderadas = grupales_raw * 1.5

    canceladas_real = suspendidas - recuperadas
    efectivas       = realizadas + recuperadas + evaluaciones + grupales_ponderadas
    programadas     = realizadas + suspendidas + evaluaciones + grupales_ponderadas

    capacidad        = calcular_capacidad(kine, dias_habiles, excepciones, mes_inicio, mes_fin)
    TOP_individual   = round(programadas / capacidad, 4) if capacidad else 0
    TOE_individual   = round(efectivas   / capacidad, 4) if capacidad else 0
    ratio_efectividad = round(efectivas  / programadas, 4) if programadas else 0

    pacientes_unicos = df_kine[COL_PACIENTE].nunique() if COL_PACIENTE in df_kine.columns else 0

    return {
        "kine":               kine,
        "pacientes_unicos":   pacientes_unicos,
        "evaluaciones":       int(evaluaciones),
        "realizadas":         int(realizadas),
        "suspendidas":        int(suspendidas),
        "recuperadas":        int(recuperadas),
        "canceladas_real":    int(canceladas_real),
        "grupales":           int(grupales_raw),
        "efectivas":          float(efectivas),
        "programadas":        float(programadas),
        "capacidad":          float(capacidad),
        "TOP_individual":     TOP_individual,
        "TOE_individual":     TOE_individual,
        "ratio_efectividad":  ratio_efectividad,
    }


# ── MÉTRICAS GRUPALES ──────────────────────────────────────────────────────────

def calcular_grupales(resultados_staff: list[dict], config: dict) -> dict:
    total_efectivas  = sum(r["efectivas"]   for r in resultados_staff)
    total_programadas = sum(r["programadas"] for r in resultados_staff)
    total_capacidad  = sum(r["capacidad"]   for r in resultados_staff)

    # Potencial total incluye refuerzo (se calcula fuera) — aquí solo staff
    TOP_grupal = round(total_programadas / total_capacidad, 4) if total_capacidad else 0
    TOE_grupal = round(total_efectivas   / total_capacidad, 4) if total_capacidad else 0

    meta_TOE = config.get("meta_TOE", 0)
    meta_TOP = config.get("meta_TOP", 0)

    return {
        "potencial_TOP":          float(total_capacidad),
        "total_efectivas_staff":  float(total_efectivas),
        "total_programadas_staff": float(total_programadas),
        "TOP_grupal":             TOP_grupal,
        "TOE_grupal":             TOE_grupal,
        "meta_TOE":               meta_TOE,
        "meta_TOP":               meta_TOP,
        "vs_meta_TOE":            round(TOE_grupal - meta_TOE, 4),
        "vs_meta_TOP":            round(TOP_grupal - meta_TOP, 4),
    }


# ── OUTPUT CSV ─────────────────────────────────────────────────────────────────

def guardar_csv(resultados: list[dict], config: dict) -> str:
    output_dir = os.path.join(os.path.dirname(__file__), "docs")
    os.makedirs(output_dir, exist_ok=True)

    mes_num = str(MESES_ES[config["mes"]]).zfill(2)
    anio    = config["año"]
    nombre  = f"resumen_{mes_num}{anio}.csv"
    ruta    = os.path.join(output_dir, nombre)

    df = pd.DataFrame(resultados)
    df.to_csv(ruta, index=False, encoding="utf-8-sig")
    return ruta


# ── OUTPUT HTML ────────────────────────────────────────────────────────────────

def _pct(valor: float) -> str:
    return f"{valor * 100:.1f}%"


def _color_semaforo(valor: float, meta: float) -> str:
    diff = valor - meta
    if diff >= 0:
        return "#22c55e"   # verde
    elif diff >= -0.05:
        return "#f59e0b"   # amarillo
    else:
        return "#ef4444"   # rojo


def guardar_html(resultados_staff: list[dict], grupales: dict,
                 config: dict, sospechosas_count: int) -> str:
    mes     = config["mes"]
    anio    = config["año"]
    notas   = config.get("notas_mes", "")
    meta_TOE = grupales["meta_TOE"]
    meta_TOP = grupales["meta_TOP"]

    filas_kines = ""
    for r in sorted(resultados_staff, key=lambda x: x["TOE_individual"], reverse=True):
        color_TOE = _color_semaforo(r["TOE_individual"], meta_TOE)
        color_TOP = _color_semaforo(r["TOP_individual"], meta_TOP)
        filas_kines += f"""
        <tr>
          <td>{r['kine']}</td>
          <td>{r['pacientes_unicos']}</td>
          <td>{r['evaluaciones']}</td>
          <td>{r['efectivas']:.0f}</td>
          <td>{r['programadas']:.0f}</td>
          <td>{r['suspendidas']}</td>
          <td>{r['recuperadas']}</td>
          <td>{r['canceladas_real']}</td>
          <td>{r['capacidad']:.0f}</td>
          <td style="color:{color_TOP};font-weight:bold">{_pct(r['TOP_individual'])}</td>
          <td style="color:{color_TOE};font-weight:bold">{_pct(r['TOE_individual'])}</td>
          <td>{_pct(r['ratio_efectividad'])}</td>
        </tr>"""

    color_TOE_g = _color_semaforo(grupales["TOE_grupal"], meta_TOE)
    color_TOP_g = _color_semaforo(grupales["TOP_grupal"], meta_TOP)
    alerta_fechas = (
        f'<div class="alerta">⚠️ {sospechosas_count} sesiones con fecha sospechosa detectadas. Revisar antes de cerrar el mes.</div>'
        if sospechosas_count > 0 else ""
    )

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Termómetro Kinexperience — {mes} {anio}</title>
  <style>
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{ font-family: 'Segoe UI', sans-serif; background: #0f172a; color: #e2e8f0; padding: 2rem; }}
    h1 {{ font-size: 1.8rem; margin-bottom: 0.25rem; }}
    .subtitulo {{ color: #94a3b8; margin-bottom: 2rem; font-size: 0.9rem; }}
    .alerta {{ background: #7c2d12; border-left: 4px solid #f97316; padding: 0.75rem 1rem;
               border-radius: 4px; margin-bottom: 1.5rem; }}
    .kpis {{ display: flex; gap: 1.5rem; margin-bottom: 2rem; flex-wrap: wrap; }}
    .kpi {{ background: #1e293b; border-radius: 8px; padding: 1.25rem 1.75rem; flex: 1; min-width: 160px; }}
    .kpi-label {{ font-size: 0.75rem; color: #94a3b8; text-transform: uppercase; letter-spacing: .05em; }}
    .kpi-value {{ font-size: 2rem; font-weight: 700; margin-top: 0.25rem; }}
    .kpi-meta {{ font-size: 0.75rem; color: #64748b; margin-top: 0.15rem; }}
    table {{ width: 100%; border-collapse: collapse; background: #1e293b; border-radius: 8px; overflow: hidden; }}
    thead tr {{ background: #334155; }}
    th {{ padding: 0.75rem 0.9rem; text-align: left; font-size: 0.75rem;
          text-transform: uppercase; letter-spacing: .05em; color: #94a3b8; }}
    td {{ padding: 0.65rem 0.9rem; font-size: 0.875rem; border-top: 1px solid #0f172a; }}
    tr:hover td {{ background: #263347; }}
    .notas {{ margin-top: 2rem; color: #64748b; font-size: 0.8rem; }}
    .footer {{ margin-top: 3rem; color: #334155; font-size: 0.7rem; text-align: center; }}
    .btn-actualizar {{ display: inline-block; margin-top: 0.5rem; padding: 0.5rem 1.25rem;
      background: #3b82f6; color: #fff; border-radius: 6px;
      font-size: 0.85rem; text-decoration: none; transition: background 0.2s; }}
    .btn-actualizar:hover {{ background: #2563eb; }}
    .header {{ display: flex; align-items: center; justify-content: space-between; margin-bottom: 2rem; }}
    .header-logo {{ height: 55px; width: auto; flex-shrink: 0; object-fit: contain; }}
    .header-text h1 {{ font-size: 1.8rem; }}
  </style>
</head>
<body>
  <div class="header">
    <div class="header-text">
      <h1>Termómetro Kinexperience</h1>
      <div class="subtitulo">{mes} {anio} &nbsp;·&nbsp; {config['dias_habiles']} días hábiles
        {f'&nbsp;·&nbsp; {notas}' if notas else ''}
      </div>
    </div>
    <img class="header-logo"
      src="https://www.kinexperience.cl/_next/image?url=%2F_next%2Fstatic%2Fmedia%2FLOGO-FONDO-OSCURO.81c302e1.png&w=128&q=75"
      alt="Kinexperience">
  </div>
  <a class="btn-actualizar" href="https://github.com/SebaNazar/termometro-kinexperience/actions/workflows/termometro.yml" target="_blank" rel="noopener">Actualizar datos</a>

  {alerta_fechas}

  <div class="kpis">
    <div class="kpi">
      <div class="kpi-label">TOE Grupal</div>
      <div class="kpi-value" style="color:{color_TOE_g}">{_pct(grupales['TOE_grupal'])}</div>
      <div class="kpi-meta">Meta: {_pct(meta_TOE)} &nbsp; ({'+' if grupales['vs_meta_TOE'] >= 0 else ''}{_pct(grupales['vs_meta_TOE'])})</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">TOP Grupal</div>
      <div class="kpi-value" style="color:{color_TOP_g}">{_pct(grupales['TOP_grupal'])}</div>
      <div class="kpi-meta">Meta: {_pct(meta_TOP)} &nbsp; ({'+' if grupales['vs_meta_TOP'] >= 0 else ''}{_pct(grupales['vs_meta_TOP'])})</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Efectivas Staff</div>
      <div class="kpi-value">{grupales['total_efectivas_staff']:.0f}</div>
      <div class="kpi-meta">de {grupales['total_programadas_staff']:.0f} programadas</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Potencial TOP</div>
      <div class="kpi-value">{grupales['potencial_TOP']:.0f}</div>
      <div class="kpi-meta">capacidad total staff</div>
    </div>
  </div>

  <table>
    <thead>
      <tr>
        <th>Kinesiólogo</th>
        <th>Pacientes</th>
        <th>Eval.</th>
        <th>Efectivas</th>
        <th>Programadas</th>
        <th>Suspendidas</th>
        <th>Recuperadas</th>
        <th>Canceladas</th>
        <th>Capacidad</th>
        <th>TOP</th>
        <th>TOE</th>
        <th>% Efectividad</th>
      </tr>
    </thead>
    <tbody>
      {filas_kines}
    </tbody>
  </table>

  <div class="notas">* TOE = Efectivas / Capacidad &nbsp;|&nbsp; TOP = Programadas / Capacidad &nbsp;|&nbsp; % Efectividad = Efectivas / Programadas (uso interno)</div>
  <div class="footer">Generado el {datetime.now().strftime('%d/%m/%Y %H:%M')} por termometro.py</div>

</body>
</html>"""

    output_dir = os.path.join(os.path.dirname(__file__), "docs")
    os.makedirs(output_dir, exist_ok=True)
    mes_num = str(MESES_ES[config["mes"]]).zfill(2)
    nombre  = f"termometro_{mes_num}{anio}.html"
    ruta    = os.path.join(output_dir, nombre)
    with open(ruta, "w", encoding="utf-8") as f:
        f.write(html)
    return ruta


# ── EXPORT JSON PRESENTACIÓN ──────────────────────────────────────────────────

def exportar_json_presentacion(resultados_kines: list[dict], resumen_grupal: dict, config: dict) -> str:
    potencial_total = sum(r["capacidad"]   for r in resultados_kines)
    potencial_top   = sum(r["programadas"] for r in resultados_kines)
    toe_grupal      = round(resumen_grupal["TOE_grupal"] * 100, 2)
    top_grupal      = round(resumen_grupal["TOP_grupal"] * 100, 2)

    # Cargar TOE del mes anterior para calcular tendencia
    mes_num_actual = MESES_ES[config["mes"]]
    anio_actual    = config["año"]
    if mes_num_actual == 1:
        mes_num_ant, anio_ant = 12, anio_actual - 1
    else:
        mes_num_ant, anio_ant = mes_num_actual - 1, anio_actual
    ruta_csv_ant = os.path.join(
        os.path.dirname(__file__), "docs",
        f"resumen_{str(mes_num_ant).zfill(2)}{anio_ant}.csv"
    )
    toe_anterior = {}
    if os.path.exists(ruta_csv_ant):
        df_ant = pd.read_csv(ruta_csv_ant, encoding="utf-8-sig")
        toe_anterior = dict(zip(df_ant["kine"], df_ant["TOE_individual"]))

    datos = {
        "potencial_total": potencial_total,
        "potencial_top":   potencial_top,
        "toe_grupal":      toe_grupal,
        "top_grupal":      top_grupal,
        "trayectoria_mes_actual": {
            "mes": config["mes"],
            "toe": toe_grupal,
        },
        "kines": [
            {
                "nombre":      r["kine"],
                "pacientes":   r["pacientes_unicos"],
                "evaluaciones": r["evaluaciones"],
                "realizadas":  r["realizadas"],
                "efectivas":   r["efectivas"],
                "suspendidas": r["suspendidas"],
                "recuperadas": r["recuperadas"],
                "canceladas":  r["canceladas_real"],
                "top":         round(r["TOP_individual"] * 100, 2),
                "toe":         round(r["TOE_individual"] * 100, 2),
                "tendencia":   (
                    "↑" if r["TOE_individual"] > toe_anterior.get(r["kine"], r["TOE_individual"]) + 0.02
                    else "↓" if r["TOE_individual"] < toe_anterior.get(r["kine"], r["TOE_individual"]) - 0.02
                    else "→"
                ),
            }
            for r in sorted(resultados_kines, key=lambda x: x["TOE_individual"], reverse=True)
        ],
    }

    output_dir = os.path.join(os.path.dirname(__file__), "docs")
    os.makedirs(output_dir, exist_ok=True)
    ruta = os.path.join(output_dir, "datos_presentacion.json")
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=2)
    return ruta


# ── MAIN ───────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  TERMÓMETRO KINEXPERIENCE")
    print("=" * 60)

    config = cargar_config()
    print(f"\nMes configurado: {config['mes']} {config['año']}")
    print(f"Días hábiles:    {config['dias_habiles']}")
    print(f"Staff oficial:   {len(config['kines_staff'])} kines")

    gc = conectar()
    df_raw = cargar_registro(gc)

    df_mes = filtrar_mes(df_raw, config)

    if df_mes.empty:
        print("\nNo hay sesiones para el mes configurado. Verifica config_mes.json.")
        return

    # Validación de fechas
    sospechosas = detectar_fechas_sospechosas(df_mes)
    if not sospechosas.empty:
        print(f"\n{'='*60}")
        print(f"  ADVERTENCIA: {len(sospechosas)} sesiones con fecha sospechosa")
        print(f"  (Fecha de sesión difiere en mes/año de Marca temporal)")
        print(f"{'='*60}")
        pd.set_option("display.max_columns", None)
        pd.set_option("display.width", 120)
        cols_mostrar = [COL_KINE, COL_FECHA, COL_TIMESTAMP, COL_ESTADO]
        cols_presentes = [c for c in cols_mostrar if c in sospechosas.columns]
        print(sospechosas[cols_presentes].to_string(index=False))
        print()
        respuesta = input("¿Continuar con el análisis de todas formas? [s/N]: ").strip().lower()
        if respuesta not in ("s", "si", "sí", "y", "yes"):
            print("Análisis cancelado. Corrige las fechas en el Registro de Sesiones y vuelve a correr.")
            return

    # Calcular métricas por kine (todos los activos del mes)
    kines_activos = df_mes[COL_KINE].dropna().unique()
    kines_staff   = config["kines_staff"]
    excepciones   = config.get("excepciones_capacidad", {})
    dias_habiles  = config["dias_habiles"]

    mes_num   = MESES_ES[config["mes"]]
    anio      = config["año"]
    mes_inicio = date(anio, mes_num, 1)
    mes_fin    = date(anio, mes_num, calendar.monthrange(anio, mes_num)[1])

    print(f"\nKines activos en el mes: {len(kines_activos)}")
    kines_no_staff = [k for k in kines_activos
                      if k not in kines_staff and k != KINE_EXCLUIDO
                      and k not in config.get("kines_refuerzo", [])]
    if kines_no_staff:
        print(f"  (fuera del Termómetro: {', '.join(kines_no_staff)})")

    resultados_staff = []
    resultados_todos = []

    for kine in kines_activos:
        # Filtrar sesiones al rango válido del kine (fecha_hasta / fecha_desde)
        inicio_rango, fin_rango = obtener_rango_sesiones_kine(
            kine, excepciones, mes_inicio, mes_fin
        )
        df_kine = df_mes[
            (df_mes[COL_KINE] == kine) &
            (df_mes["_fecha"] >= pd.Timestamp(inicio_rango)) &
            (df_mes["_fecha"] <= pd.Timestamp(fin_rango))
        ]
        metricas = calcular_metricas_kine(
            df_kine, kine, dias_habiles, excepciones, mes_inicio, mes_fin
        )
        resultados_todos.append(metricas)
        if kine in kines_staff:
            resultados_staff.append(metricas)

    if not resultados_staff:
        print("\nNingún kine del staff tiene sesiones este mes. Revisa config_mes.json.")
        return

    grupales = calcular_grupales(resultados_staff, config)

    # Imprimir resumen en consola
    print(f"\n{'─'*60}")
    print(f"  RESULTADOS — {config['mes'].upper()} {config['año']}")
    print(f"{'─'*60}")
    print(f"  TOE Grupal : {_pct(grupales['TOE_grupal'])}  (meta: {_pct(grupales['meta_TOE'])})")
    print(f"  TOP Grupal : {_pct(grupales['TOP_grupal'])}  (meta: {_pct(grupales['meta_TOP'])})")
    print(f"  Efectivas  : {grupales['total_efectivas_staff']:.0f}")
    print(f"  Capacidad  : {grupales['potencial_TOP']:.0f}")
    print(f"{'─'*60}")
    header = f"  {'Kine':<28} {'TOE':>6}  {'TOP':>6}  {'Ef':>6}  {'Prog':>6}  {'Cap':>6}"
    print(header)
    print(f"  {'─'*28} {'─'*6}  {'─'*6}  {'─'*6}  {'─'*6}  {'─'*6}")
    for r in sorted(resultados_staff, key=lambda x: x["TOE_individual"], reverse=True):
        simbolo_TOE = "✓" if r["TOE_individual"] >= grupales["meta_TOE"] else "✗"
        print(f"  {r['kine']:<28} {_pct(r['TOE_individual']):>6} {simbolo_TOE}"
              f" {_pct(r['TOP_individual']):>6}  {r['efectivas']:>6.0f}  "
              f"{r['programadas']:>6.0f}  {r['capacidad']:>6.0f}")
    print(f"{'─'*60}")

    # Guardar outputs
    ruta_csv  = guardar_csv(resultados_todos, config)
    ruta_html = guardar_html(resultados_staff, grupales, config, len(sospechosas))

    ruta_json = exportar_json_presentacion(resultados_staff, grupales, config)

    print(f"\nOutputs generados:")
    print(f"  CSV : {ruta_csv}")
    print(f"  HTML: {ruta_html}")
    print(f"  JSON: {ruta_json}")
    print()


if __name__ == "__main__":
    main()
