# -*- coding: utf-8 -*-
# indicadores/generador_indicadores.py

from pathlib import Path
from typing import Optional
import pandas as pd

class GeneradorIndicadores:
    """
    Genera/extrae datos base para los indicadores.
    Incluye la extracción de 'Atención de las alertas contables con calidad'
    desde el archivo ALCON (hoja 'Detalle_Bancolombia').
    """

    def __init__(self, ruta_entrada: Optional[Path] = None):
        self.ruta_entrada = Path(ruta_entrada) if ruta_entrada else None

    # ---------- Utilidad interna ----------

    def _buscar_encabezado_tabla(self, df_raw: pd.DataFrame) -> Optional[tuple]:
        """
        Busca la fila y columna inicial donde aparecen los encabezados exactos.
        """
        headers = ["gerencia", "cantidad alertas", "alertas con reproceso",
                   "alertas sin reproceso", "calidad gerencia"]

        n_rows, n_cols = df_raw.shape
        for r in range(n_rows):
            fila = df_raw.iloc[r].astype(str).str.strip().str.lower().tolist()
            for c in range(n_cols - 4):
                ventana = fila[c:c+5]
                if ventana == headers:
                    return (r, c)
        return None

    # ---------- API pública ----------

    def extraer_atencion_alertas_calidad(self, archivo_alcon: Path) -> pd.DataFrame:
        """
        Lee la hoja 'Detalle_Bancolombia', detecta la tabla de 'CALIDAD GESTION USUARIOS (Gerencia)',
        limpia '(en blanco)', conserva 'Total general', y retorna el DataFrame.
        """
        df_raw = pd.read_excel(
            archivo_alcon,
            sheet_name="Detalle_Bancolombia",
            header=None,
            engine="openpyxl"
        )

        df_norm = df_raw.copy()
        # Se usa map para compatibilidad con versiones recientes de pandas
        df_norm = df_norm.map(lambda x: str(x).strip() if pd.notna(x) else "")

        pos = self._buscar_encabezado_tabla(df_norm)
        if not pos:
            raise ValueError(
                "No encontré la cabecera de la sección 'CALIDAD GESTION USUARIOS (Gerencia)' "
                "en la hoja 'Detalle_Bancolombia'."
            )

        fila_ini, col_ini = pos

        df = df_norm.iloc[fila_ini+1:, col_ini:col_ini+5].copy()
        df.columns = ["Gerencia", "Cantidad alertas", "Alertas con reproceso",
                      "Alertas sin reproceso", "Calidad Gerencia"]

        df["Gerencia"] = (
            df["Gerencia"]
            .astype(str)
            .str.replace(r"[\r\n\t]", "", regex=True)
            .str.strip()
        )

        idx_total = df.index[df["Gerencia"].str.lower().eq("total general")]
        if len(idx_total) > 0:
            df = df.loc[:idx_total[0]]

        mask_en_blanco = df["Gerencia"].str.lower().eq("(en blanco)")
        df = df[~mask_en_blanco].copy()

        for col in ["Cantidad alertas", "Alertas con reproceso", "Alertas sin reproceso"]:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

        df["Calidad Gerencia"] = df["Calidad Gerencia"].astype(str).str.strip()

        return df
    
    def extraer_historico_certificacion(self, archivo_historico: Path, hoja: str = None) -> pd.DataFrame:
        """
        Lee el archivo 'Historico Indicador Certificación Gerentes.xlsx' y devuelve
        únicamente las columnas necesarias, totalmente limpias (sin NA).
        """
        # 1) Leer archivo completo como texto
        read_kwargs = {"dtype": str, "engine": "openpyxl"}
        if hoja:
            read_kwargs["sheet_name"] = hoja
        df = pd.read_excel(archivo_historico, **read_kwargs)

        # 2) Normalizar nombres de columnas
        def norm(s):
            if pd.isna(s):
                return ""
            return str(s).strip().replace("\r", "").replace("\n", "")

        cols_map = {norm(c).lower(): c for c in df.columns}

        nombre_gerencia = cols_map.get("gerencia")
        nombre_fecha_cert = cols_map.get("fecha certificación") or cols_map.get("fecha certificacion")
        nombre_fecha_obj  = cols_map.get("fecha objetivo")
        nombre_indicador  = cols_map.get("indicador")

        faltantes = []
        if not nombre_gerencia: faltantes.append("GERENCIA")
        if not nombre_fecha_cert: faltantes.append("FECHA CERTIFICACIÓN")
        if not nombre_fecha_obj:  faltantes.append("FECHA OBJETIVO")
        if not nombre_indicador:  faltantes.append("INDICADOR")
        
        if faltantes:
            raise ValueError(f"Faltan columnas requeridas: {', '.join(faltantes)}")

        # 3) Seleccionar columnas necesarias
        df_sel = df[[nombre_gerencia, nombre_fecha_cert, nombre_fecha_obj, nombre_indicador]].copy()
        df_sel.columns = ["GERENCIA", "FECHA CERTIFICACIÓN", "FECHA OBJETIVO", "INDICADOR"]

        # 4) Limpieza de texto (CORREGIDO)
        for c in df_sel.columns:
            df_sel[c] = (
                df_sel[c]
                .astype(str)
                .str.replace(r"[\r\n\t]", "", regex=True)
                .str.replace("nan", "", regex=False)
                .str.strip()
            )

        # Convertir INDICADOR a número real, si aplica
        try:
            df_sel["INDICADOR"] = pd.to_numeric(df_sel["INDICADOR"], errors="coerce")
        except:
            pass

        # Llenar vacíos
        df_sel = df_sel.fillna("")

        # 5) Quitar filas totalmente vacías y filas sin GERENCIA
        df_sel = df_sel.replace({"": pd.NA}).dropna(how="all")
        df_sel = df_sel[df_sel["GERENCIA"].notna() & (df_sel["GERENCIA"].str.strip() != "")]

        # 6) Asegurar que no queden NA en ninguna celda
        df_sel = df_sel.fillna("")

        # 7) Resetear índice limpio
        df_sel = df_sel.reset_index(drop=True)

        return df_sel
