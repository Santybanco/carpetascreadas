# Ejecuta esto desde la raíz del proyecto (ISCE)

import sys
from pathlib import Path
sys.path.append(str(Path(__file__).resolve().parents[0]))
from cargadores.cargador_excel import CargadorExcel
from pathlib import Path
from procesadores.procesador_temporales import ProcesadorTemporales

from pathlib import Path
import pandas as pd
from exportadores.exportador_excel import ExportadorExcel
from datos.rutas import obtener_ruta_salida

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

from procesadores.procesador_temporales import ProcesadorTemporales
from config.configuracion import Configuracion

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

    

# 1) Ruta SharePoint (si ya la tienes exacta):
#    Reemplaza esta línea cuando me confirmes tu ruta real:
ruta_sharepoint = None  # Path(r"C:\Users\...\OneDrive - ...\Administrativo_M365\Documentos\PRUEBAS ENERO\Todos los documentos")

# 2) Tomar el archivo de temporales detectado por el cargador
archivos_temporales = meta.get("temporales", [])
if not archivos_temporales:
    print("⚠️ No se encontró el archivo de temporales en datos/entrada.")
else:
    archivo_temp = archivos_temporales[0]
    print(f"\n🧩 Procesando TEMPORALES desde: {archivo_temp.name}")

    proc = ProcesadorTemporales(ruta_sharepoint=ruta_sharepoint)
    destino = proc.procesar_y_exportar(archivo_temp)

    print(f"📤 Libro generado: {destino}")
    print("   Hojas: TEMPORAL, TEMPORAL2 (filtrada), TD Saldo, Sábana Temporales")

    from pathlib import Path

def resolver_ruta_sharepoint_local() -> Path:
    """
    Busca la carpeta local sincronizada que corresponde a:
      Administrativo125 - Documentos compartidos\General\Indicadores\Indicadores Operación\2026\PRUEBAS ENERO
    Retorna la primera que exista. Si no encuentra, retorna Path() vacío.
    """
    # Posibles raíces de OneDrive/SharePoint en Windows
    bases = [
        Path.home() / "OneDrive - Bancolombia S.A",
        Path.home() / "OneDrive - Bancolombia",
        Path.home() / "Bancolombia",
        Path.home() / "OneDrive - Grupo Bancolombia",
    ]

    # Dos patrones frecuentes de cómo se “monta” la biblioteca en el Explorador
    relativos = [
        Path(r"Administrativo125 - Documentos compartidos\General\Indicadores\Indicadores Operación\2026\PRUEBAS ENERO"),
        Path(r"Administrativo125\Documentos compartidos\General\Indicadores\Indicadores Operación\2026\PRUEBAS ENERO"),
    ]

    for base in bases:
        for rel in relativos:
            candidato = base / rel
            if candidato.exists():
                return candidato
    return Path()



# 👇 Pega aquí la ruta EXACTA que copiaste del Explorador:
RUTA_SP = Path(r"C:\Users\santcord\OneDrive - Grupo Bancolombia\Administrativo_M365 - PRUEBAS ENERO")

archivos_temporales = meta.get("temporales", [])
if not archivos_temporales:
    print("⚠️ No se encontró el archivo de temporales en datos/entrada.")
else:
    archivo_temp = archivos_temporales[0]
    print(f"\n🧩 Procesando TEMPORALES desde: {archivo_temp.name}")

    # Instancia el procesador apuntando a la carpeta de OneDrive/SharePoint
    proc = ProcesadorTemporales(ruta_sharepoint=RUTA_SP)

    destino = proc.procesar_y_exportar(
        archivo_temporales=archivo_temp,
        nombre_archivo_salida="Informe Cuentas Temporales Ene_procesado.xlsx"
    )

    print(f"📤 Libro generado: {destino}")

    # --- DESPUÉS de procesar temporales y generar el archivo procesado en SharePoint ---


# 1) Ruta donde quedó el "Informe Cuentas Temporales Ene_procesado.xlsx" (tu SharePoint)
RUTA_SP = Path(r"C:\Users\santcord\OneDrive - Grupo Bancolombia\Administrativo_M365 - PRUEBAS ENERO")

# 2) Leemos la hoja TD Saldo del archivo procesado
archivo_procesado = RUTA_SP / "Informe Cuentas Temporales Ene_procesado.xlsx"
df_td = pd.read_excel(archivo_procesado, sheet_name="TD Saldo", engine="openpyxl")

# 3) Pegamos en el libro final "Indicadores_operacion_nuevo.xlsx", hoja "enero 2026", empezando en A4
exportador = ExportadorExcel(obtener_ruta_salida())
destino_final = exportador.exportar_td_saldo_en_indicadores(
    df_td=df_td,
    nombre_hoja_mes="enero 2026",     # ajusta si cambias de mes
    nombre_archivo="Indicadores_operacion_nuevo.xlsx",
    celda_inicio="A4"
)

print(f"📎 TD Saldo pegado en: {destino_final} -> hoja 'enero 2026' desde A4")
