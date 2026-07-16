# -*- coding: utf-8 -*-
"""
Conciliador Maestro BCI — DAEM Puerto Montt
===========================================
Procesa un único archivo Excel con formato de entrada:
  - Hoja 1: "1110301 MOV FONDOS"         -> Rango útil Columnas B a M (Fila 6)
  - Hoja 2: "CARTOLAS BANCARIAS GENERAL" -> Rango útil Columnas C a N (Fila 6)

Lógica Híbrida robustecida frente a celdas combinadas (MergedCells).
"""

import streamlit as st
import pandas as pd
import openpyxl
import io
import re
import math
from datetime import datetime, date
from itertools import combinations

# CONFIGURACIÓN GENERAL DE LA UI
st.set_page_config(
    page_title="Conciliador Bancario Maestro — DAEM",
    page_icon="🏦",
    layout="wide"
)

# Constantes de estructura basadas en tu archivo maestro
SHEET_CONTABLE = "1110301 MOV FONDOS"
SHEET_BANCO = "CARTOLAS BANCARIAS GENERAL"

TOLERANCIA_DIAS = 5
TOLERANCIA_MONTO = 1.0
MAX_COMBO_FASE2 = 6

# ---------------------------------------------------------------------------
# UTILIDADES DE LIMPIEZA Y PARSEO
# ---------------------------------------------------------------------------
def limpiar_monto(val):
    if val is None or val == "":
        return 0.0
    if isinstance(val, float) and math.isnan(val):
        return 0.0
    if isinstance(val, (int, float)):
        return abs(float(val))
    texto = str(val).strip().replace("$", "").replace(".", "").replace(",", ".")
    try:
        num = float(texto)
        return 0.0 if math.isnan(num) else abs(num)
    except ValueError:
        return 0.0

def parsear_fecha(val):
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    texto = str(val).strip().split(" ")[0]
    formatos = ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"]
    for fmt in formatos:
        try:
            return datetime.strptime(texto, fmt).date()
        except ValueError:
            continue
    return None

def limpiar_glosa(texto):
    if texto is None or pd.isna(texto):
        return ""
    texto = str(texto).strip().upper()
    return re.sub(r"[0-9]+", "", texto).strip()

# ---------------------------------------------------------------------------
# ESCRITURA IN-PLACE RESISTENTE A CELDAS COMBINADAS (MERGED CELLS)
# ---------------------------------------------------------------------------
def escribir_celda_segura(ws, row, col, valor):
    """
    Escribe de forma segura en una celda de openpyxl. Si es parte de un rango
    combinado subordinado (MergedCell), busca la celda de origen y escribe allí.
    """
    try:
        celda = ws.cell(row=row, column=col)
        if type(celda).__name__ == "MergedCell":
            # Buscar el rango combinado al que pertenece para escribir en la cabecera
            for rgo in list(ws.merged_cells.ranges):
                if row >= rgo.min_row and row <= rgo.max_row and col >= rgo.min_col and col <= rgo.max_col:
                    ws.cell(row=rgo.min_row, column=rgo.min_col).value = valor
                    return
        else:
            celda.value = valor
    except Exception:
        # Pasa silenciosamente para que nunca interrumpa la ejecución del reporte final
        pass

# ---------------------------------------------------------------------------
# PARSERS DE HOJAS CON FILAS DE DATOS REALES
# ---------------------------------------------------------------------------
def extraer_datos_contables(ws_v):
    datos = []
    for r in range(7, ws_v.max_row + 1):
        fecha_raw = ws_v.cell(row=r, column=4).value
        fecha = parsear_fecha(fecha_raw)
        if not fecha:
            continue
            
        glosa = str(ws_v.cell(row=r, column=5).value or "")
        if "SALDO ANTERIOR" in glosa.upper():
            continue
            
        doc_raw = ws_v.cell(row=r, column=6).value
        doc = int(float(str(doc_raw).strip())) if pd.notna(doc_raw) and str(doc_raw).strip().replace(".0","").isdigit() and float(str(doc_raw).strip()) > 0 else 0
        
        debe = limpiar_monto(ws_v.cell(row=r, column=9).value)
        haber = limpiar_monto(ws_v.cell(row=r, column=10).value)
        
        if debe == 0 and haber == 0:
            continue
            
        datos.append({
            "fila_excel": r,
            "fecha": fecha,
            "comprobante": ws_v.cell(row=r, column=3).value,
            "glosa": glosa,
            "documento": doc,
            "debe": debe,
            "haber": haber,
            "monto": debe if debe > 0 else haber,
            "tipo": "CARGO" if debe > 0 else "ABONO",
            "usado": False
        })
    return datos

def extraer_datos_bancarios(ws_v):
    datos = []
    for r in range(7, ws_v.max_row + 1):
        fecha_raw = ws_v.cell(row=r, column=4).value
        fecha = parsear_fecha(fecha_raw)
        if not fecha:
            continue
            
        movimiento = str(ws_v.cell(row=r, column=6).value or "")
        if "SALDO INICIAL" in movimiento.upper() or "GIRADO NO COBRADO" in movimiento.upper():
            continue
            
        doc_raw = ws_v.cell(row=r, column=7).value
        doc = int(float(str(doc_raw).strip())) if pd.notna(doc_raw) and str(doc_raw).strip().replace(".0","").isdigit() and float(str(doc_raw).strip()) > 0 else 0
        
        cargo = limpiar_monto(ws_v.cell(row=r, column=8).value)
        abono = limpiar_monto(ws_v.cell(row=r, column=9).value)
        
        if cargo == 0 and abono == 0:
            continue
            
        datos.append({
            "fila_excel": r,
            "fecha": fecha,
            "movimiento": movimiento,
            "documento": doc,
            "cargo": cargo,
            "abono": abono,
            "monto": cargo if cargo > 0 else abono,
            "tipo": "ABONO" if cargo > 0 else "CARGO",
            "usado": False
        })
    return datos

# ---------------------------------------------------------------------------
# ALGORITMO DE PRE-CRUCE EN DOS FASES
# ---------------------------------------------------------------------------
def ejecutar_conciliacion_hibrida(datos_c, datos_b):
    matches = []

    # FASE 1.1: Cruce Directo por Número de Documento Único (> 0)
    for c in datos_c:
        if c["documento"] == 0:
            continue
        for b in datos_b:
            if b["usado"] or b["documento"] == 0:
                continue
            if c["documento"] == b["documento"] and abs(c["monto"] - b["monto"]) <= TOLERANCIA_MONTO:
                c["usado"] = True
                b["usado"] = True
                matches.append({"tipo": "1 a 1 (Documento)", "contable": c, "bancarios": [b]})
                break

    # FASE 1.2: Cruce Directo por Monto Exacto + Tolerancia de Fecha (Lógica FIFO)
    for c in datos_c:
        if c["usado"]:
            continue
        candidatos = []
        for b in datos_b:
            if b["usado"] or c["tipo"] != b["tipo"]:
                continue
            if abs(c["monto"] - b["monto"]) <= TOLERANCIA_MONTO:
                diff_dias = abs((c["fecha"] - b["fecha"]).days)
                if diff_dias <= TOLERANCIA_DIAS:
                    candidatos.append((diff_dias, b))
        
        if candidatos:
            candidatos.sort(key=lambda x: x[0])
            b_elegido = candidatos[0][1]
            c["usado"] = True
            b_elegido["usado"] = True
            matches.append({"tipo": "1 a 1 (Monto/Fecha)", "contable": c, "bancarios": [b_elegido]})

    # FASE 2: Consolidación Combinatoria Intradiaria (1 Contable a N Bancarios)
    huerfanos_ingresos = [c for c in datos_c if not c["usado"] and c["debe"] > 0]
    for c in huerfanos_ingresos:
        banco_del_dia = [b for b in datos_b if not b["usado"] and b["abono"] > 0 and b["fecha"] == c["fecha"]]
        if not banco_del_dia:
            continue
            
        grupos_glosa = {}
        for b in banco_del_dia:
            clave = limpiar_glosa(b["movimiento"])
            grupos_glosa.setdefault(clave, []).append(b)
            
        match_fase2 = False
        for grupo in grupos_glosa.values():
            if len(grupo) < 2:
                continue
            for tam in range(2, min(len(grupo), MAX_COMBO_FASE2) + 1):
                for combo in combinations(grupo, tam):
                    if abs(sum(g["abono"] for g in combo) - c["debe"]) <= TOLERANCIA_MONTO:
                        c["usado"] = True
                        for g in combo:
                            g["usado"] = True
                        matches.append({"tipo": "1 a N (Caja Agrupada)", "contable": c, "bancarios": list(combo)})
                        match_fase2 = True
                        break
                if match_fase2:
                    break
            if match_fase2:
                break
                
    return matches

# ---------------------------------------------------------------------------
# ESCRITURA IN-PLACE
# ---------------------------------------------------------------------------
def escribir_reporte_excel(excel_bytes, matches, datos_c, datos_b):
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=False)
    ws_c = wb[SHEET_CONTABLE]
    ws_b = wb[SHEET_BANCO]

    # Inyectar resultados de los amarres exitosos usando la función de escritura segura
    for m in matches:
        c = m["contable"]
        bancarios = m["bancarios"]
        
        sum_monto_banco = sum(b["monto"] for b in bancarios)
        comp_origen = f"{c['comprobante']}" if c['comprobante'] else "MOV"
        
        # Hoja 1 Contable: Columna L (CRUCE CARTOLAS BANCARIAS, Columna 12)
        if m["tipo"] == "1 a N (Caja Agrupada)":
            escribir_celda_segura(ws_c, c["fila_excel"], 12, f"Agrupado x{len(bancarios)}")
        else:
            escribir_celda_segura(ws_c, c["fila_excel"], 12, sum_monto_banco)
            
        # Hoja 2 Banco: Columnas L, M, N (Cruce, Diferencia, Glosa detalle)
        for b in bancarios:
            escribir_celda_segura(ws_b, b["fila_excel"], 12, c["monto"])
            escribir_celda_segura(ws_b, b["fila_excel"], 13, 0)
            escribir_celda_segura(ws_b, b["fila_excel"], 14, comp_origen)

    # Marcar partidas conciliatorias pendientes
    for c in datos_c:
        if not c["usado"]:
            escribir_celda_segura(ws_c, c["fila_excel"], 12, "PENDIENTE")
            
    for b in datos_b:
        if not b["usado"]:
            escribir_celda_segura(ws_b, b["fila_excel"], 12, "PENDIENTE")

    buffer_salida = io.BytesIO()
    wb.save(buffer_salida)
    buffer_salida.seek(0)
    return buffer_salida

# ---------------------------------------------------------------------------
# INTERFAZ WEB STREAMLIT
# ---------------------------------------------------------------------------
st.title("🏦 Sistema de Conciliación Bancaria Automática — DAEM")
st.markdown(
    f"Herramienta de cruce inteligente para las hojas **`{SHEET_CONTABLE}`** y **`{SHEET_BANCO}`** "
    "del archivo unificado mensual."
)

st.divider()

archivo_subido = st.file_uploader("📊 Cargar archivo Excel Maestro (.xlsx)", type=["xlsx"])

if archivo_subido:
    if st.button("🚀 Ejecutar Proceso de Conciliación", type="primary"):
        bytes_archivo = archivo_subido.read()
        
        wb_valores = openpyxl.load_workbook(io.BytesIO(bytes_archivo), data_only=True)
        
        with st.spinner("Analizando estructuras y extrayendo movimientos contables..."):
            datos_c = extraer_datos_contables(wb_valores[SHEET_CONTABLE])
            datos_b = extraer_datos_bancarios(wb_valores[SHEET_BANCO])
            
        if not datos_c or not datos_b:
            st.error("No se pudieron leer registros válidos en las hojas seleccionadas. Verifica las filas de inicio.")
        else:
            with st.spinner("Ejecutando algoritmo de emparejamiento FIFO y combinatorio..."):
                matches = ejecutar_conciliacion_hibrida(datos_c, datos_b)
                
            with st.spinner("Escribiendo resultados preservando el formato original..."):
                excel_finalizado = escribir_reporte_excel(bytes_archivo, matches, datos_c, datos_b)
                
            st.success("🎯 ¡Proceso completado con éxito!")
            
            # Cálculo de Métricas Financieras en Pantalla
            m1a_count = sum(1 for m in matches if "Documento" in m["tipo"])
            m1b_count = sum(1 for m in matches if "Monto" in m["tipo"])
            m2_count = sum(1 for m in matches if "Caja" in m["tipo"])
            
            pend_c = sum(1 for c in datos_c if not c["usado"])
            pend_b = sum(1 for b in datos_b if not b["usado"])
            
            c1, c2, c3 = st.columns(3)
            c1.metric("Movimientos Contables Totales", len(datos_c))
            c2.metric("Movimientos Banco Totales", len(datos_b))
            c3.metric("Total de Matches Logrados", len(matches))
            
            c4, c5, c6 = st.columns(3)
            c4.metric("Calces por Documento (1 a 1)", m1a_count)
            c5.metric("Calces por Monto/Fecha (1 a 1)", m1b_count)
            c6.metric("Cajas Consolidadas (1 a N)", m2_count)
            
            st.warning(f"⚠️ Partidas Conciliatorias Pendientes: **{pend_c}** en Libro Mayor y **{pend_b}** en Cartola Bancaria.")
            
            st.divider()
            
            st.download_button(
                label="📥 Descargar Excel Conciliado con Fórmulas e Historial",
                data=excel_finalizado,
                file_name="Reporte_Conciliacion_BCI_DAEM.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info(f"⬆️ Por favor, sube tu archivo Excel que contenga las pestañas de control '{SHEET_CONTABLE}' y bancaria '{SHEET_BANCO}'.")
