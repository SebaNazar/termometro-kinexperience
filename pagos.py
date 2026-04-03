"""
Pagos Kinexperience — Módulo de novedades para contabilidad
Uso: python3 pagos.py

Outputs:
  - docs/borrador_contabilidad_MMYYYY.txt  (borrador para contabilidad)
  - docs/refuerzo_[nombre]_MMYYYY.txt       (uno por kine de refuerzo con sesiones)
"""

import os
import sys
import json
import calendar as cal
from datetime import date

import pandas as pd

from termometro import (
    conectar,
    cargar_registro,
    cargar_config,
    filtrar_mes,
    calcular_metricas_kine,
    obtener_rango_sesiones_kine,
    dias_habiles_rango,
    MESES_ES,
    COL_KINE,
    COL_ESTADO,
)


# ── TABLAS DE BONOS ────────────────────────────────────────────────────────────
# Cada entrada: (umbral_mínimo_inclusive, monto_bono)
# Se recorre de mayor a menor; se devuelve el bono del primer umbral alcanzado.

TABLA_BONO_KINE_NUEVO = [
    (1.15, 650_000),
    (1.10, 600_000),
    (1.05, 525_000),
    (1.00, 450_000),
    (0.95, 375_000),
    (0.90, 300_000),
    (0.85, 250_000),
    (0.80, 200_000),
    (0.75, 150_000),
    (0.70, 100_000),
    (0.65,  75_000),
    (0.60,  50_000),
]

TABLA_BONO_KINE_ANTIGUO = [
    (1.15, 350_000),
    (1.10, 300_000),
    (1.05, 200_000),
    (1.00, 150_000),
    (0.95, 100_000),
    (0.90,  50_000),
]

TABLA_BONO_MAURICIO = [
    (0.95, 1_200_000),
    (0.90,   800_000),
    (0.85,   450_000),
    (0.80,   300_000),
    (0.75,   150_000),
    (0.00,         0),
]


def _factor_ponderacion(
    kine: str, excepciones: dict, dias_habiles: int,
    mes_inicio: date, mes_fin: date
) -> float:
    """
    Retorna el factor de ponderación para el bono de tabla (0.0 – 1.0).
    Solo aplica a excepciones fecha_hasta / fecha_desde.
    Excepción porcentaje no pondera el bono (ya ajusta la capacidad).
    """
    exc = excepciones.get(kine)
    if exc:
        if exc["tipo"] == "fecha_hasta":
            fecha = date.fromisoformat(exc["fecha"])
            dias_trabajados = dias_habiles_rango(mes_inicio, min(fecha, mes_fin))
            return dias_trabajados / dias_habiles if dias_habiles else 0.0
        elif exc["tipo"] == "fecha_desde":
            fecha = date.fromisoformat(exc["fecha"])
            dias_trabajados = dias_habiles_rango(max(fecha, mes_inicio), mes_fin)
            return dias_trabajados / dias_habiles if dias_habiles else 0.0
    return 1.0


def _buscar_bono(tabla: list, toe: float) -> int:
    for umbral, bono in tabla:
        if toe >= umbral:
            return bono
    return 0


# ── FORMATO DE MONEDA ──────────────────────────────────────────────────────────

def fmt_clp(valor: float) -> str:
    """Formatea un entero como precio chileno: $1.234.567"""
    return f"${int(valor):,}".replace(",", ".")


# ── PERFILES ───────────────────────────────────────────────────────────────────

def determinar_perfil(kine: str, config: dict) -> str:
    perfiles_especiales = config.get("pagos", {}).get("perfiles_especiales", {})
    if kine in perfiles_especiales:
        return perfiles_especiales[kine]["perfil"]
    if kine in config.get("kines_refuerzo", []):
        return "refuerzo"
    return "kine_nuevo"


# ── CÁLCULO DE BONOS STAFF ─────────────────────────────────────────────────────

def calcular_bonos_staff(
    resultados_staff: list,
    config: dict,
    mes_inicio: date,
    mes_fin: date,
) -> dict:
    """
    Calcula bono individual para cada kine del staff.
    Los bonos de tabla se ponderan por (días_trabajados / dias_habiles)
    cuando el kine tiene excepción fecha_hasta o fecha_desde.
    Retorna dict kine → {perfil, TOE_individual, bono, bono_bencina, nota_manual}
    """
    pagos = config.get("pagos", {})
    bono_bencina_por_sesion = pagos.get("bono_bencina_por_sesion", 0)
    notas_manuales = pagos.get("notas_manuales", {})
    excepciones = config.get("excepciones_capacidad", {})
    dias_habiles = config["dias_habiles"]

    resultado = {}
    for r in resultados_staff:
        kine = r["kine"]
        perfil = determinar_perfil(kine, config)
        toe = r["TOE_individual"]
        realizadas = r["realizadas"]

        factor = _factor_ponderacion(kine, excepciones, dias_habiles, mes_inicio, mes_fin)
        bono = 0
        bono_bencina = 0

        if perfil == "kine_nuevo":
            bono = round(_buscar_bono(TABLA_BONO_KINE_NUEVO, toe) * factor)
            bono_bencina = realizadas * bono_bencina_por_sesion

        elif perfil == "kine_antiguo":
            bono = round(_buscar_bono(TABLA_BONO_KINE_ANTIGUO, toe) * factor)

        # mauricio y sebastian se manejan trimestralmente (ver generar_borrador)

        resultado[kine] = {
            "perfil": perfil,
            "TOE_individual": toe,
            "realizadas": realizadas,
            "bono": bono,
            "bono_bencina": bono_bencina,
            "nota_manual": notas_manuales.get(kine, ""),
        }

    return resultado


# ── CÁLCULO DE PAGOS REFUERZO ──────────────────────────────────────────────────

def calcular_pagos_refuerzo(df_mes: pd.DataFrame, config: dict) -> dict:
    """
    Calcula montos para kines de refuerzo con sesiones en el mes.
    Retorna dict kine → {sesiones, monto_bruto, monto_liquido, bono_bencina, total_liquido}
    """
    pagos = config.get("pagos", {})
    valor_bruto = pagos.get("valor_sesion_refuerzo_bruto", 21_053)
    pct_boleta = pagos.get("porcentaje_boleta_honorarios", 0.1525)
    bonos_bencina = pagos.get("bonos_bencina_refuerzo", {})
    kines_refuerzo = config.get("kines_refuerzo", [])

    resultado = {}
    for kine in kines_refuerzo:
        df_kine = df_mes[df_mes[COL_KINE] == kine]
        realizadas   = (df_kine[COL_ESTADO] == "Realizada").sum()
        recuperadas  = (df_kine[COL_ESTADO] == "Recuperada").sum()
        evaluaciones = (df_kine[COL_ESTADO] == "Evaluación de ingreso").sum()
        grupales     = (df_kine[COL_ESTADO] == "Sesión Grupal").sum()
        sesiones_cobro = int(realizadas + recuperadas + evaluaciones + grupales)

        if sesiones_cobro == 0:
            continue

        monto_bruto   = sesiones_cobro * valor_bruto
        monto_liquido = round(monto_bruto * (1 - pct_boleta))
        bono_benc     = bonos_bencina.get(kine, 0)
        total_liquido = monto_liquido + bono_benc

        resultado[kine] = {
            "sesiones":      sesiones_cobro,
            "monto_bruto":   monto_bruto,
            "monto_liquido": monto_liquido,
            "bono_bencina":  bono_benc,
            "total_liquido": total_liquido,
        }

    return resultado


# ── OUTPUT 1: BORRADOR PARA CONTABILIDAD ──────────────────────────────────────

def generar_borrador_contabilidad(
    bonos_staff: dict,
    toe_grupal_mes: float,
    config: dict,
) -> str:
    pagos = config.get("pagos", {})
    mes = config["mes"]
    anio = config["año"]
    trimestre = pagos.get("trimestre_actual", "")
    mes_cierre = pagos.get("mes_cierre_trimestre", "")
    toe_meses_config = pagos.get("toe_grupal_meses_trimestre", [0.0, 0.0, 0.0])
    perfiles_especiales = pagos.get("perfiles_especiales", {})
    notas_manuales = pagos.get("notas_manuales", {})

    es_mes_cierre = (mes == mes_cierre)

    # Construir lista de TOE grupales del trimestre:
    # toe_meses_config tiene 0.0 en la posición del mes actual (placeholder).
    # Tomamos los meses ya cerrados y agregamos el calculado.
    meses_previos = [t for t in toe_meses_config if t > 0.0]
    toe_trimestre = meses_previos + [toe_grupal_mes]
    promedio_trimestral = sum(toe_trimestre) / len(toe_trimestre)

    bono_mauricio = 0
    bono_sebastian = 0
    if es_mes_cierre:
        bono_mauricio = _buscar_bono(TABLA_BONO_MAURICIO, promedio_trimestral)
        bono_sebastian = bono_mauricio // 2

    # Orden del correo: Mauricio → antiguos → nuevos → Sebastián
    MAURICIO  = "Mauricio Arce"
    SEBASTIAN = "Sebastián Nazar"
    antiguos  = [k for k, v in perfiles_especiales.items() if v["perfil"] == "kine_antiguo"]
    todos_staff = list(bonos_staff.keys())

    orden = []
    if MAURICIO in perfiles_especiales:
        orden.append(MAURICIO)
    orden += [k for k in antiguos if k in todos_staff]
    orden += [
        k for k in todos_staff
        if k not in antiguos and k != MAURICIO and k != SEBASTIAN
    ]
    if SEBASTIAN in perfiles_especiales:
        orden.append(SEBASTIAN)

    # Incluir Mauricio/Sebastián aunque no tengan sesiones clínicas
    for especial in (MAURICIO, SEBASTIAN):
        if especial in perfiles_especiales and especial not in orden:
            orden.append(especial)

    lineas = [
        "Estimadas, buenas tardes",
        "",
        f"Junto con saludar, adjunto novedades del mes de {mes} {anio}",
        "",
    ]

    for kine in orden:
        lineas.append(f"{kine}:")

        if kine == MAURICIO:
            nota = notas_manuales.get(MAURICIO, "")
            if nota:
                lineas.append(f"* {nota}")
            acumulado_str = ", ".join(f"{t*100:.1f}%" for t in toe_trimestre)
            lineas.append(f"* TOE grupal acumulada ({trimestre}): {acumulado_str}")
            lineas.append(f"* Promedio TOE grupal trimestral: {promedio_trimestral*100:.1f}%")
            if es_mes_cierre:
                lineas.append(
                    f"* CORRESPONDE PAGO DE BONO TRIMESTRAL: {fmt_clp(bono_mauricio)} (imponible)"
                )
            else:
                lineas.append(f"* (bono trimestral se paga en {mes_cierre})")

        elif kine == SEBASTIAN:
            nota = notas_manuales.get(SEBASTIAN, "")
            if nota:
                lineas.append(f"* {nota}")
            if es_mes_cierre:
                lineas.append(
                    f"* Bono desempeño trimestral imponible: {fmt_clp(bono_sebastian)}"
                )

        else:
            info = bonos_staff.get(kine, {})
            bono = info.get("bono", 0)
            bono_bencina = info.get("bono_bencina", 0)
            nota = info.get("nota_manual", "")
            if bono > 0:
                lineas.append(f"* Bono desempeño imponible: {fmt_clp(bono)}")
            if bono_bencina > 0:
                lineas.append(f"* Bono bencina: {fmt_clp(bono_bencina)}")
            if nota:
                lineas.append(f"* {nota}")

        lineas.append("")

    lineas.append("Muchas gracias")

    return "\n".join(lineas)


# ── OUTPUT 2: MENSAJES PARA KINES DE REFUERZO ─────────────────────────────────

def generar_mensaje_refuerzo(kine: str, datos: dict, config: dict) -> str:
    mes = config["mes"]
    anio = config["año"]
    rut_empresa = config.get("pagos", {}).get("rut_empresa", "RUT PENDIENTE")
    nombre_corto = kine.split()[0]

    mes_num = MESES_ES[mes]
    ultimo_dia = cal.monthrange(anio, mes_num)[1]

    lineas = [
        f"Hola {nombre_corto},",
        "",
        f"Te escribo para informarte las sesiones del mes de {mes}:",
        "",
        f"Sesiones realizadas: {datos['sesiones']}",
        f"Monto bruto: {fmt_clp(datos['monto_bruto'])}",
        f"Monto líquido: {fmt_clp(datos['monto_liquido'])} (considera 15.25% de retención de honorarios)",
    ]

    if datos["bono_bencina"] > 0:
        lineas.append(f"Bono bencina: {fmt_clp(datos['bono_bencina'])}")

    lineas += [
        f"Total líquido: {fmt_clp(datos['total_liquido'])}",
        "",
        f"Por favor emite tu boleta de honorarios por {fmt_clp(datos['monto_bruto'])} "
        f"a nombre de Kinexperience SpA, RUT {rut_empresa}, "
        f"Providencia 1208, oficina 207, comuna de Providencia, Santiago, Chile. "
        f"Glosa: Atenciones kinesiológicas - {mes} {anio}. "
        f"La boleta debe tener fecha {ultimo_dia} de {mes} de {anio}.",
        "",
        "Saludos,",
    ]

    return "\n".join(lineas)


# ── GUARDAR ARCHIVOS ───────────────────────────────────────────────────────────

def guardar_txt(contenido: str, nombre: str) -> str:
    output_dir = os.path.join(os.path.dirname(__file__), "docs")
    os.makedirs(output_dir, exist_ok=True)
    ruta = os.path.join(output_dir, nombre)
    with open(ruta, "w", encoding="utf-8") as f:
        f.write(contenido)
    return ruta


# ── MAIN ───────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  PAGOS KINEXPERIENCE")
    print("=" * 60)

    config = cargar_config()
    if "pagos" not in config:
        sys.exit("ERROR: Falta la sección 'pagos' en config_mes.json.")

    mes = config["mes"]
    anio = config["año"]
    dias_habiles = config["dias_habiles"]
    kines_staff = config["kines_staff"]
    excepciones = config.get("excepciones_capacidad", {})

    print(f"\nMes configurado: {mes} {anio}")
    print(f"Días hábiles:    {dias_habiles}")

    gc = conectar()
    df_raw = cargar_registro(gc)
    df_mes = filtrar_mes(df_raw, config)

    if df_mes.empty:
        print("\nNo hay sesiones para el mes configurado. Verifica config_mes.json.")
        return

    mes_num    = MESES_ES[mes]
    mes_inicio = date(anio, mes_num, 1)
    mes_fin    = date(anio, mes_num, cal.monthrange(anio, mes_num)[1])

    # ── Métricas de staff ──────────────────────────────────────────────────────
    resultados_staff = []
    for kine in kines_staff:
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
        resultados_staff.append(metricas)

    # TOE grupal del mes (solo staff, igual que termometro.py)
    total_efectivas = sum(r["efectivas"] for r in resultados_staff)
    total_capacidad = sum(r["capacidad"] for r in resultados_staff)
    toe_grupal_mes  = round(total_efectivas / total_capacidad, 4) if total_capacidad else 0.0

    # ── Calcular bonos ─────────────────────────────────────────────────────────
    bonos_staff    = calcular_bonos_staff(resultados_staff, config, mes_inicio, mes_fin)
    pagos_refuerzo = calcular_pagos_refuerzo(df_mes, config)

    # ── Imprimir resumen en consola ────────────────────────────────────────────
    print(f"\n{'─'*60}")
    print(f"  BONOS STAFF — {mes.upper()} {anio}")
    print(f"{'─'*60}")
    print(f"  TOE grupal del mes: {toe_grupal_mes*100:.1f}%")
    print()
    for kine in kines_staff:
        info   = bonos_staff.get(kine, {})
        perfil = info.get("perfil", "?")
        toe    = info.get("TOE_individual", 0.0)
        bono   = info.get("bono", 0)
        bb     = info.get("bono_bencina", 0)
        extras = []
        if bono > 0:
            extras.append(f"bono={fmt_clp(bono)}")
        if bb > 0:
            extras.append(f"bencina={fmt_clp(bb)}")
        extras_str = "  " + "  ".join(extras) if extras else "  (sin novedad)"
        print(f"  {kine:<28} [{perfil:<12}] TOE={toe*100:.1f}%{extras_str}")

    if pagos_refuerzo:
        print(f"\n{'─'*60}")
        print(f"  REFUERZO")
        print(f"{'─'*60}")
        for kine, datos in pagos_refuerzo.items():
            print(
                f"  {kine:<28} {datos['sesiones']} ses."
                f"  bruto={fmt_clp(datos['monto_bruto'])}"
                f"  líquido={fmt_clp(datos['total_liquido'])}"
            )

    # ── Guardar outputs ────────────────────────────────────────────────────────
    mes_str = str(mes_num).zfill(2) + str(anio)

    borrador  = generar_borrador_contabilidad(bonos_staff, toe_grupal_mes, config)
    ruta_borrador = guardar_txt(borrador, f"borrador_contabilidad_{mes_str}.txt")

    print(f"\nOutputs generados:")
    print(f"  Borrador contabilidad : {ruta_borrador}")

    for kine, datos in pagos_refuerzo.items():
        msg         = generar_mensaje_refuerzo(kine, datos, config)
        nombre_kine = kine.replace(" ", "_").lower()
        ruta_ref    = guardar_txt(msg, f"refuerzo_{nombre_kine}_{mes_str}.txt")
        print(f"  Mensaje refuerzo      : {ruta_ref}")

    print()


if __name__ == "__main__":
    main()
