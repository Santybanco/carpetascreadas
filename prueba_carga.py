# Ejecuta esto desde la raíz del proyecto (ISCE)

import sys
import time
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parents[0]))

import pandas as pd
from cargadores.cargador_excel import CargadorExcel
from indicadores.generador_indicadores import GeneradorIndicadores
from exportadores.exportador_excel import ExportadorExcel
from procesadores.procesador_temporales import ProcesadorTemporales
from datos.rutas import obtener_ruta_salida

NOMBRE_HOJA_MES = "enero 2026"
RUTA_SP_STR = r"C:\Users\santcord\OneDrive - Grupo Bancolombia\Administrativo_M365 - PRUEBAS ENERO"
RUTA_SP = Path(RUTA_SP_STR) if Path(RUTA_SP_STR).exists() else None

# Flags para pruebas rápidas
EJECUTAR_TEMPORALES = True  # ponlo en False para iterar rápido sin reprocesar

def tic(): return time.perf_counter()
def lap(t0, label): t = time.perf_counter() - t0; print(f"⏱️ {label}: {t:0.2f} s"); return time.perf_counter()

t_total = tic()

# 1) Carga
t = tic()
cargador = CargadorExcel()
data, meta = cargador.cargar_todo()
print("\n📊 Resumen"); print(cargador.resumen(data))
for tipo, dfs in data.items():
    if dfs:
        print(f"\n🔎 Primeras filas de '{tipo}':"); print(dfs[0].head())
t = lap(t, "Carga y resumen de archivos")

# 2) ALCON
t_alcon = tic()
archivos_alcon = meta.get("alcon", [])
exportador_ind = ExportadorExcel(obtener_ruta_salida())
if archivos_alcon:
    archivo_alcon = archivos_alcon[0]
    print(f"\n🧩 Usando ALCON: {archivo_alcon.name}")
    gen = GeneradorIndicadores()
    df_alertas = gen.extraer_atencion_alertas_calidad(archivo_alcon)
    destino_ind = exportador_ind.exportar_atencion_alertas_calidad(
        df_alertas, nombre_hoja_mes=NOMBRE_HOJA_MES, nombre_archivo="Indicadores_operacion_nuevo.xlsx")
    print(f"\n📤 Exportado ALCON en: {destino_ind}"); print(f"   Hoja: {NOMBRE_HOJA_MES} | Encabezados desde A83")
else:
    print("⚠️ No hay archivo ALCON en la carpeta de entrada.")
lap(t_alcon, "ALCON → Indicadores (A83)")

# 3) HISTÓRICO
t_hist = tic()
archivos_historico = meta.get("historico", [])
if archivos_historico:
    archivo_hist = archivos_historico[0]
    print(f"\n🧩 Usando HISTÓRICO: {archivo_hist.name}")
    gen = GeneradorIndicadores()
    try:
        df_hist = gen.extraer_historico_certificacion(archivo_hist)
        destino_ind2 = exportador_ind.exportar_historico_certificacion(
            df_hist, nombre_hoja_mes=NOMBRE_HOJA_MES, nombre_archivo="Indicadores_operacion_nuevo.xlsx",
            fila_inicio_excel=180, col_inicio_excel=1, escribir_promedio=True)
        print(f"\n📤 Histórico exportado a: {destino_ind2}"); print("   Hoja:", NOMBRE_HOJA_MES, "| Celda inicio: A180")
    except Exception as e:
        print(f"❌ Error procesando Histórico: {e}")
else:
    print("⚠️ No hay archivo 'Historico Indicador Certificación Gerentes.xlsx'.")
lap(t_hist, "HISTÓRICO → Indicadores (A180)")

# 4) TEMPORALES
if EJECUTAR_TEMPORALES:
    t_temp = tic()
    archivos_temporales = meta.get("temporales", [])
    if not archivos_temporales:
        print("⚠️ No se encontró el archivo de temporales en datos/entrada.")
        sys.exit(1)
    archivo_temp = archivos_temporales[0]
    print(f"\n🧩 Procesando TEMPORALES desde: {archivo_temp.name}")
    proc = ProcesadorTemporales(ruta_sharepoint=RUTA_SP)
    destino_proc = proc.procesar_y_exportar(archivo_temporales=archivo_temp, crear_db_cr_sabana=True)
    print(f"📤 Libro generado (TEMPORALES): {destino_proc}")
    lap(t_temp, "TEMPORALES: procesar y exportar")
else:
    destino_proc = Path(RUTA_SP_STR) / "Informe Cuentas Temporales Ene_procesado.xlsx"
    print("⚡ TEMPORALES omitido (modo prueba rápida)")

# 5) TD Saldo -> A4
t_td = tic()
try:
    df_td = pd.read_excel(destino_proc, sheet_name="TD Saldo", engine="openpyxl")
    destino_final = exportador_ind.exportar_td_saldo_en_indicadores(
        df_td=df_td, nombre_hoja_mes=NOMBRE_HOJA_MES, nombre_archivo="Indicadores_operacion_nuevo.xlsx", celda_inicio="A4")
    print(f"\n📎 TD Saldo pegado en: {destino_final} -> hoja '{NOMBRE_HOJA_MES}' desde A4")
except Exception as e:
    print(f"❌ Error pegando TD Saldo en Indicadores: {e}")
lap(t_td, "Pegar TD Saldo en Indicadores (A4)")

# 6) TD SABANA (conteo) -> E4
t_cnt = tic()
try:
    df_td_cnt = pd.read_excel(destino_proc, sheet_name="TD SABANA", engine="openpyxl")
    destino_final_cnt = exportador_ind.exportar_td_sabana_en_indicadores(
        df_td_sabana=df_td_cnt, nombre_hoja_mes=NOMBRE_HOJA_MES,
        nombre_archivo="Indicadores_operacion_nuevo.xlsx", celda_inicio="E4")
    print(f"\n📎 TD SABANA (conteo) pegada en: {destino_final_cnt} -> hoja '{NOMBRE_HOJA_MES}' desde E4")
except Exception as e:
    print(f"❌ Error pegando TD SABANA (conteo) en Indicadores: {e}")
lap(t_cnt, "Pegar TD SABANA (conteo) en Indicadores (E4)")

# 7) TD SABANA (DB) -> H4
t_db = tic()
try:
    df_td_db = pd.read_excel(destino_proc, sheet_name="TD SABANA", engine="openpyxl", header=31)  # A32 (31-based)
    cols_ok = [c for c in df_td_db.columns if str(c).strip() in ("Etiquetas de fila", "Total general", "SI", "SÍ")]
    df_td_db = df_td_db[cols_ok].dropna(how="all").copy()
    destino_final_db = exportador_ind.exportar_td_sabana_db_en_indicadores(
        df_td_db=df_td_db, nombre_hoja_mes=NOMBRE_HOJA_MES,
        nombre_archivo="Indicadores_operacion_nuevo.xlsx", celda_inicio="H4", formato_porcentaje="0.0%")
    print(f"\n📎 TD SABANA (DB) pegada en: {destino_final_db} -> hoja '{NOMBRE_HOJA_MES}' desde H4")
except Exception as e:
    print(f"❌ Error pegando TD SABANA (DB) en Indicadores: {e}")
lap(t_db, "Pegar TD SABANA (DB) en Indicadores (H4)")

# 8) TD SABANA (CR) -> K4 (leer encabezados en fila 59)
t_cr = tic()
try:
    df_td_cr = pd.read_excel(destino_proc, sheet_name="TD SABANA", engine="openpyxl", header=58)  # A59
    cols_ok = [c for c in df_td_cr.columns if str(c).strip() in ("Etiquetas de fila", "Total general", "SI", "SÍ", "NO")]
    df_td_cr = df_td_cr[cols_ok].dropna(how="all").copy()
    destino_final_cr = exportador_ind.exportar_td_sabana_cr_en_indicadores(
        df_td_cr=df_td_cr, nombre_hoja_mes=NOMBRE_HOJA_MES,
        nombre_archivo="Indicadores_operacion_nuevo.xlsx", celda_inicio="K4", formato_porcentaje="0.0%")
    print(f"\n📎 TD SABANA (CR) pegada en: {destino_final_cr} -> hoja '{NOMBRE_HOJA_MES}' desde K4")
except Exception as e:
    print(f"❌ Error pegando TD SABANA (CR) en Indicadores: {e}")
lap(t_cr, "Pegar TD SABANA (CR) en Indicadores (K4)")

lap(t_total, "⏱️ Tiempo TOTAL del script")
print("✅ Finalizado.")
