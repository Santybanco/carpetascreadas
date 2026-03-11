# Ejecuta esto desde la raíz del proyecto (ISCE)

import sys
from pathlib import Path
sys.path.append(str(Path(__file__).resolve().parents[0]))
from cargadores.cargador_excel import CargadorExcel

cargador = CargadorExcel()
data, meta = cargador.cargar_todo()

print("\n📊 Resumen")
print(cargador.resumen(data))

# Si quieres ojear las primeras filas de cada tipo:
for tipo, dfs in data.items():
    if not dfs:
        continue
    print(f"\n🔎 Primeras filas de '{tipo}':")
    print(dfs[0].head())

# --- PRUEBA ALCON: extraer y exportar ---

from indicadores.generador_indicadores import GeneradorIndicadores
from exportadores.exportador_excel import ExportadorExcel
from datos.rutas import obtener_ruta_salida

# 1) Tomamos el primer archivo ALCON detectado por el cargador
archivos_alcon = meta.get("alcon", [])
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

    # 2) Exportamos al archivo final, hoja del mes (ajusta el nombre según tu mes)
    nombre_hoja_mes = "enero 2026"  # <- cámbialo por el mes que estés construyendo
    exportador = ExportadorExcel(obtener_ruta_salida())
    destino = exportador.exportar_atencion_alertas_calidad(
        df_alertas, nombre_hoja_mes=nombre_hoja_mes,
        nombre_archivo="Indicadores_operacion_nuevo.xlsx",
        fila_inicio_excel=81
    )

    print(f"\n📤 Exportado a: {destino}")
    print(f"   Hoja: {nombre_hoja_mes} | Fila inicio: 83 | Columnas: A:E")


    # --- PRUEBA HISTÓRICO: extraer y exportar ---
from indicadores.generador_indicadores import GeneradorIndicadores
from exportadores.exportador_excel import ExportadorExcel
from datos.rutas import obtener_ruta_salida

gen = GeneradorIndicadores()
exportador = ExportadorExcel(obtener_ruta_salida())

# 1) Buscar el archivo histórico
archivos_historico = meta.get("historico", [])
if not archivos_historico:
    print("⚠️ No hay archivo 'Historico Indicador Certificación Gerentes.xlsx' en la carpeta de entrada.")
else:
    archivo_hist = archivos_historico[0]
    print(f"\n🧩 Usando HISTÓRICO: {archivo_hist.name}")

    # 2) Extraer las 4 columnas tal cual
    df_hist = gen.extraer_historico_certificacion(archivo_hist)

    print("\n✅ Histórico (primeras filas):")
    print(df_hist.head())

    # 3) Exportar a la misma hoja del mes en A180
    nombre_hoja_mes = "enero 2026"  # 👈 Ajusta según el mes
    destino2 = exportador.exportar_historico_certificacion(
        df_hist,
        nombre_hoja_mes=nombre_hoja_mes,
        nombre_archivo="Indicadores_operacion_nuevo.xlsx",
        fila_inicio_excel=180,  # Encabezados en A180
        col_inicio_excel=1      # Columna A
    )

    print(f"\n📤 Histórico exportado a: {destino2}")
    print("   Hoja:", nombre_hoja_mes, "| Celda inicio: A180")
