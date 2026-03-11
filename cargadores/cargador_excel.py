# -*- coding: utf-8 -*-
# cargadores/cargador_excel.py
from pathlib import Path
from typing import Dict, List, Tuple
import pandas as pd
from datos.rutas import obtener_ruta_entrada

class CargadorExcel:
    """
    Carga archivos Excel de la carpeta de entrada y los organiza por tipo.
    Tipos esperados: 'temporales', 'cxc', 'cxp', 'alcon'
    """
    def __init__(self, ruta_entrada: Path = None):
        self.ruta_entrada = ruta_entrada or obtener_ruta_entrada()
        # Patrones simples para identificar cada tipo por nombre del archivo
        self.patrones = {
            "temporales": ["cuentas temporales", "temporales"],
            "cxc": ["cxc", "por cobrar"],
            "cxp": ["cxp", "por pagar"],
            "alcon": ["indicadores_alcon", "alcon"],
            
            "historico": ["historico indicador certificación gerentes",
                  "historico indicador certificacion gerentes",
                  "certificación gerentes", "certificacion gerentes", "historico"]


        }

    # -------------------- Utilidades internas --------------------

    def _normalizar_nombre(self, nombre: str) -> str:
        return nombre.strip().lower().replace("  ", " ")

    def _clasificar_tipo(self, nombre_archivo: str) -> str:
        """
        Retorna el tipo ('temporales'|'cxc'|'cxp'|'alcon') según patrones, o '' si no matchea.
        """
        n = self._normalizar_nombre(nombre_archivo)
        for tipo, patrones in self.patrones.items():
            for p in patrones:
                if p in n:
                    return tipo
        return ""

    def _normalizar_columnas(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Minúsculas y reemplazo de espacios por guion bajo.
        """
        df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
        return df

    # -------------------- API pública --------------------

    def listar_archivos(self) -> Dict[str, List[Path]]:
        """
        Lista los archivos por tipo. Solo toma .xlsx
        """
        archivos_por_tipo: Dict[str, List[Path]] = {"temporales": [], "cxc": [], "cxp": [], "alcon": [], "historico": []}
        for p in self.ruta_entrada.glob("*.xlsx"):
            tipo = self._clasificar_tipo(p.name)
            if tipo:
                archivos_por_tipo[tipo].append(p)
        return archivos_por_tipo

    def cargar_dataframe(self, archivo: Path) -> pd.DataFrame:
        """
        Carga un archivo Excel como DataFrame (primer hoja).
        Usa dtype=str para evitar errores de tipos; convertiremos luego por columnas.
        """
        df = pd.read_excel(archivo, dtype=str, engine="openpyxl")
        df = self._normalizar_columnas(df)
        # Limpieza básica: quitar filas completamente vacías
        df = df.dropna(how="all")
        return df

    def cargar_todo(self) -> Tuple[Dict[str, List[pd.DataFrame]], Dict[str, List[Path]]]:
        """
        Carga todos los archivos detectados y retorna:
        - data: dict tipo -> lista de DataFrames
        - meta: dict tipo -> lista de Paths (los mismos en el mismo orden)
        """
        archivos = self.listar_archivos()
        data: Dict[str, List[pd.DataFrame]] = {k: [] for k in archivos.keys()}

        print("📂 Carpeta de entrada:", str(self.ruta_entrada))
        for tipo, lista_paths in archivos.items():
            if not lista_paths:
                print(f"⚠️  No se encontraron archivos de tipo '{tipo}'.")
                continue
            for path in lista_paths:
                try:
                    df = self.cargar_dataframe(path)
                    data[tipo].append(df)
                    print(f"✅ Cargado [{tipo}]: {path.name} -> filas: {len(df)} | columnas: {len(df.columns)}")
                except Exception as e:
                    print(f"❌ Error cargando '{path.name}': {e}")

        return data, archivos

    def resumen(self, data: Dict[str, List[pd.DataFrame]]) -> pd.DataFrame:
        """
        Retorna un DataFrame resumen por tipo con total de archivos y filas.
        """
        resumen_rows = []
        for tipo, dfs in data.items():
            total_archivos = len(dfs)
            total_filas = sum(len(df) for df in dfs)
            resumen_rows.append({"tipo": tipo, "archivos": total_archivos, "filas": total_filas})
        return pd.DataFrame(resumen_rows)
