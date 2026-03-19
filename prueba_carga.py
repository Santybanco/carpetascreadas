# Ejecuta esto desde la raíz del proyecto (ISCE)

import sys
import time
from pathlib import Path

# Asegurar que Python conozca la raíz del proyecto
sys.path.append(str(Path(__file__).resolve().parents[0]))

# Imports del proyecto
import pandas as pd
from cargadores.cargador_excel import CargadorExcel
from indicadores.generador_indicadores import GeneradorIndicadores
from exportadores.exportador_excel import ExportadorExcel
from procesadores.procesador_temporales import ProcesadorTemporales
from datos.rutas import obtener_ruta_salida


# ---------------------------------------------
# 0) Parámetros de ejecución
# ---------------------------------------------
NOMBRE_HOJA_MES = "enero 2026"  # Ajusta el mes aquí una sola vez

# Si quieres guardar TEMPORALES directo en OneDrive/SharePoint, pon la ruta local sincronizada:
# Si no existe, el procesador guardará en datos/salida automáticamente.
RUTA_SP_STR = r"C:\Users\santcord\OneDrive - Grupo Bancolombia\Administrativo_M365 - PRUEBAS ENERO"
RUTA_SP = Path(RUTA_SP_STR) if Path(RUTA_SP_STR).exists() else None


# ---------------------------------------------
# Utilidades simples de medición
# ---------------------------------------------
def tic():
    return time.perf_counter()

def lap(t0: float, label: str):
    t = time.perf_counter() - t0
    print(f"⏱️ {label}: {t:0.2f} s")
    return time.perf_counter()  # devuelve nuevo punto de partida


# ---------------------------------------------
# Helper robusto: leer TD SABANA (DB) desde la hoja homónima
#   - Intenta header=31 (fila 32) como nos definimos
#   - Si falla, busca dinámicamente la fila donde está 'Etiquetas de fila' + 'Total general'
# ---------------------------------------------
def leer_td_sabana_db(destino_proc: Path) -> pd.DataFrame:
    hoja = "TD SABANA"
    # 1) Primer intento: header=31 (fila 32)
    try:
        df = pd.read_excel(destino_proc, sheet_name=hoja, engine="openpyxl", header=31)
        cols = [c for c in df.columns]
        if any(str(c).strip() in ("Etiquetas de fila", "Total general", "SI", "SÍ") for c in cols):
            return df
    except Exception:
        pass

    # 2) Fallback: detectar la fila de encabezados dinámicamente
    raw = pd.read_excel(destino_proc, sheet_name=hoja, engine="openpyxl", header=None)
    header_row = None
    for i in range(min(len(raw), 200)):  # escaneo de primeras 200 filas
        fila = [str(x).strip() for x in list(raw.iloc[i].values)]
        if ("Etiquetas de fila" in fila) and ("Total general" in fila):
            if ("SI" in fila) or ("SÍ" in fila) or any((isinstance(s, str) and s.upper() == "SI") for s in fila):
                header_row = i
                break
    if header_row is None:
        # Último intento: usa header=31 y dejamos que el exportador valide columnas
        return pd.read_excel(destino_proc, sheet_name=hoja, engine="openpyxl", header=31)

    df = pd.read_excel(destino_proc, sheet_name=hoja, engine="openpyxl", header=header_row)
    return df


# ---------------------------------------------
# INICIO
# ---------------------------------------------
t_total = tic()

# ---------------------------------------------
# 1) Cargar archivos (resumen)
# ---------------------------------------------
t = tic()
cargador = CargadorExcel()
data, meta = cargador.cargar_todo()

print("\n📊 Resumen")
print(cargador.resumen(data))

# (Opcional) Ver primeras filas por tipo
for tipo, dfs in data.items():
    if not dfs:
        continue
    print(f"\n🔎 Primeras filas de '{tipo}':")
    print(dfs[0].head())

t = lap(t, "Carga y resumen de archivos")


# ---------------------------------------------
# 2) ALCON → Indicadores (A83)
# ---------------------------------------------
t_alcon = tic()
archivos_alcon = meta.get("alcon", [])
exportador_ind = ExportadorExcel(obtener_ruta_salida())  # instancia única

if not archivos_alcon:
    print("⚠️ No hay archivo ALCON en la carpeta de entrada.")
else:
    archivo_alcon = archivos_alcon[0]
    print(f"\n🧩 Usando ALCON: {archivo_alcon.name}")

    gen = GeneradorIndicadores()
    df_alertas = gen.extraer_atencion_alertas_calidad(archivo_alcon)

    print("\n✅ Tabla extraída (primeras filas):")
    print(df_alertas.head())
    print("\n🔚 Últimas filas:")
    print(df_alertas.tail())
    print(f"\n📏 Filas totales (incluye 'Total general'): {len(df_alertas)}")

    destino_ind = exportador_ind.exportar_atencion_alertas_calidad(
        df_alertas,
        nombre_hoja_mes=NOMBRE_HOJA_MES,
        nombre_archivo="Indicadores_operacion_nuevo.xlsx",
    )
    print(f"\n📤 Exportado ALCON en: {destino_ind}")
    print(f"   Hoja: {NOMBRE_HOJA_MES} | Encabezados desde A83")
lap(t_alcon, "ALCON → Indicadores (A83)")


# ---------------------------------------------
# 3) HISTÓRICO → Indicadores (A180)
# ---------------------------------------------
t_hist = tic()
archivos_historico = meta.get("historico", [])

if not archivos_historico:
    print("⚠️ No hay archivo 'Historico Indicador Certificación Gerentes.xlsx' en la carpeta de entrada.")
else:
    archivo_hist = archivos_historico[0]
    print(f"\n🧩 Usando HISTÓRICO: {archivo_hist.name}")

    gen = GeneradorIndicadores()
    try:
        df_hist = gen.extraer_historico_certificacion(archivo_hist)
        print("\n✅ Histórico (primeras filas):")
        print(df_hist.head())

        destino_ind2 = exportador_ind.exportar_historico_certificacion(
            df_hist,
            nombre_hoja_mes=NOMBRE_HOJA_MES,
            nombre_archivo="Indicadores_operacion_nuevo.xlsx",
            fila_inicio_excel=180,  # Encabezados en A180
            col_inicio_excel=1,     # Columna A
            escribir_promedio=True  # Promedio dinámico de INDICADOR (formato %)
        )
        print(f"\n📤 Histórico exportado a: {destino_ind2}")
        print("   Hoja:", NOMBRE_HOJA_MES, "| Celda inicio: A180")
    except Exception as e:
        print(f"❌ Error procesando Histórico: {e}")
lap(t_hist, "HISTÓRICO → Indicadores (A180)")


# ---------------------------------------------
# 4) TEMPORALES → procesar (TEMPORAL, TEMPORAL2, TD Saldo, Sábana, TD SABANA)
# ---------------------------------------------
t_temp = tic()
archivos_temporales = meta.get("temporales", [])
if not archivos_temporales:
    print("⚠️ No se encontró el archivo de temporales en datos/entrada.")
else:
    archivo_temp = archivos_temporales[0]
    print(f"\n🧩 Procesando TEMPORALES desde: {archivo_temp.name}")

    proc = ProcesadorTemporales(ruta_sharepoint=RUTA_SP)  # si RUTA_SP es None, guarda en datos/salida
    destino_proc = proc.procesar_y_exportar(
        archivo_temporales=archivo_temp,
        crear_db_cr_sabana=True   # pon False si quieres “omitir/comentar” DB/CR
    )
    print(f"📤 Libro generado (TEMPORALES): {destino_proc}")
lap(t_temp, "TEMPORALES: procesar y exportar")


# ---------------------------------------------
# 5) TD Saldo → Indicadores (A4 en la hoja del mes)
# ---------------------------------------------
t_td = tic()
try:
    df_td = pd.read_excel(destino_proc, sheet_name="TD Saldo", engine="openpyxl")

    destino_final = exportador_ind.exportar_td_saldo_en_indicadores(
        df_td=df_td,
        nombre_hoja_mes=NOMBRE_HOJA_MES,
        nombre_archivo="Indicadores_operacion_nuevo.xlsx",
        celda_inicio="A4"                       # encabezados en A4
    )
    print(f"\n📎 TD Saldo pegado en: {destino_final} -> hoja '{NOMBRE_HOJA_MES}' desde A4")
except Exception as e:
    print(f"❌ Error pegando TD Saldo en Indicadores: {e}")
lap(t_td, "Pegar TD Saldo en Indicadores (A4)")


# ---------------------------------------------
# 6) TD SABANA (conteo) → Indicadores (E4 en la hoja del mes)
# ---------------------------------------------
t_tdsab_cnt = tic()
try:
    # La TD de conteo está al inicio de la hoja "TD SABANA" (encabezados en la primera fila)
    df_td_cnt = pd.read_excel(destino_proc, sheet_name="TD SABANA", engine="openpyxl")

    # Pegamos: E = Total general, F = SI, G = % (IFERROR)
    destino_final_cnt = exportador_ind.exportar_td_sabana_en_indicadores(
        df_td_sabana=df_td_cnt,
        nombre_hoja_mes=NOMBRE_HOJA_MES,
        nombre_archivo="Indicadores_operacion_nuevo.xlsx",
        celda_inicio="E4"                       # 👈 como tu plantilla
    )
    print(f"\n📎 TD SABANA (conteo) pegada en: {destino_final_cnt} -> hoja '{NOMBRE_HOJA_MES}' desde E4")
except Exception as e:
    print(f"❌ Error pegando TD SABANA (conteo) en Indicadores: {e}")
lap(t_tdsab_cnt, "Pegar TD SABANA (conteo) en Indicadores (E4)")


# ---------------------------------------------
# 7) TD SABANA (DB) → Indicadores (H4 en la hoja del mes)
# ---------------------------------------------
t_db = tic()
try:
    # Leer de forma robusta la TD SABANA (DB) que insertamos desde A32 en la hoja TD SABANA
    df_td_db = leer_td_sabana_db(destino_proc)

    # Quedarnos solo con las columnas relevantes (acepta 'SI' o 'SÍ')
    cols_ok = []
    for c in df_td_db.columns:
        c_clean = str(c).strip()
        if c_clean in ("Etiquetas de fila", "Total general", "SI", "SÍ"):
            cols_ok.append(c)
    df_td_db = df_td_db[cols_ok].copy()
    df_td_db = df_td_db.dropna(how="all")

    destino_final_db = exportador_ind.exportar_td_sabana_db_en_indicadores(
        df_td_db=df_td_db,
        nombre_hoja_mes=NOMBRE_HOJA_MES,
        nombre_archivo="Indicadores_operacion_nuevo.xlsx",
        celda_inicio="H4",               # 👈 como en tu plantilla
        formato_porcentaje="0.0%"        # cambia a "0.00%" si prefieres 2 decimales
    )
    print(f"\n📎 TD SABANA (DB) pegada en: {destino_final_db} -> hoja '{NOMBRE_HOJA_MES}' desde H4")
except Exception as e:
    print(f"❌ Error pegando TD SABANA (DB) en Indicadores: {e}")
lap(t_db, "Pegar TD SABANA (DB) en Indicadores (H4)")


# ---------------------------------------------
# FIN
# ---------------------------------------------
lap(t_total, "⏱️ Tiempo TOTAL del script")
print("✅ Finalizado.")