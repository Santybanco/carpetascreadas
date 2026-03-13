# -*- coding: utf-8 -*-
# procesadores/procesador_temporales.py

from pathlib import Path
from typing import Optional, Tuple
import re
import unicodedata
import pandas as pd

from datos.rutas import obtener_ruta_salida


class ProcesadorTemporales:
    """
    Procesa el archivo 'Informe Cuentas Temporales ...xlsx':
      - Duplica 'TEMPORAL' → 'TEMPORAL2' y filtra SALDO CONTABLE != 0
      - Genera 'TD Saldo' (conteos por gerencia)
      - Copia 'Sábana Temporales'
      - Exporta a SharePoint (si existe ruta) o a datos/salida/
    """

    def __init__(self, ruta_sharepoint: Optional[Path] = None):
        self.ruta_sharepoint = Path(ruta_sharepoint) if ruta_sharepoint else None

    # -------------------- Utilidades internas --------------------

    def _normalizar(self, s: str) -> str:
        """
        Normaliza strings para matching robusto (minúsculas, sin tildes/espacios extras).
        """
        if s is None:
            return ""
        # quitar saltos/tabulaciones y espacios múltiples
        s = re.sub(r"[\r\n\t]+", " ", str(s)).strip().lower()
        # quitar tildes
        s = "".join(
            ch for ch in unicodedata.normalize("NFD", s)
            if unicodedata.category(ch) != "Mn"
        )
        # colapsar espacios
        s = re.sub(r"\s+", " ", s)
        return s

    def _encontrar_columna(self, columnas: list, objetivo: str) -> Optional[str]:
        """
        Encuentra el nombre real de la columna que matchea al 'objetivo' (insensible a tildes/mayúsculas).
        """
        objetivo_n = self._normalizar(objetivo)
        mapa = {self._normalizar(c): c for c in columnas}
        return mapa.get(objetivo_n, None)

    def _to_numeric_safe(self, serie: pd.Series) -> pd.Series:
        """
        Convierte a numérico de forma segura, limpiando símbolos comunes.
        Maneja miles con punto o coma, y decimales con coma.
        """
        if serie is None:
            return pd.Series(dtype=float)

        s = (
            serie.astype(str)
            .str.replace(r"[\r\n\t]", "", regex=True)
            .str.replace(r"[^\d\-.,]", "", regex=True)  # deja solo dígitos, - . ,
            .str.strip()
        )

        # Si hay coma y punto, asumimos que . = miles y , = decimal
        mask_coma = s.str.contains(",", regex=False)
        mask_punto = s.str.contains(r"\.", regex=True)

        # Caso común ES: "1.234,56" -> "1234.56"
        s = s.where(~(mask_coma & mask_punto), s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
        # Caso "1234,56" -> "1234.56"
        s = s.where(~(mask_coma & ~mask_punto), s.str.replace(",", ".", regex=False))

        return pd.to_numeric(s, errors="coerce")

    # -------------------- Carga y transformación --------------------

    def _cargar_hoja(self, archivo: Path, nombre_hoja: str) -> pd.DataFrame:
        return pd.read_excel(archivo, sheet_name=nombre_hoja, dtype=str, engine="openpyxl")

    def _construir_temporal2(self, df_temporal: pd.DataFrame) -> Tuple[pd.DataFrame, str, str, str]:
        """
        - Devuelve df_temporal2 filtrado, y los nombres REALES de 3 columnas:
          gerencia_responsable, SALDO CONTABLE, PARTIDAS FUERA DE POLITICA_y
        """
        cols = list(df_temporal.columns)

        col_gerencia = self._encontrar_columna(cols, "gerencia_responsable")
        col_saldo = self._encontrar_columna(cols, "SALDO CONTABLE")
        col_fuera = self._encontrar_columna(cols, "PARTIDAS FUERA DE POLITICA_y")  # tal cual tu fuente

        faltan = [n for n, v in {
            "gerencia_responsable": col_gerencia,
            "SALDO CONTABLE": col_saldo,
            "PARTIDAS FUERA DE POLITICA_y": col_fuera
        }.items() if v is None]
        if faltan:
            raise ValueError(f"No encontré columnas requeridas en 'TEMPORAL': {', '.join(faltan)}")

        # Filtrar SALDO CONTABLE != 0 (y no NaN)
        saldo_num = self._to_numeric_safe(df_temporal[col_saldo])
        mask_saldo = saldo_num.fillna(0) != 0
        df_temporal2 = df_temporal.loc[mask_saldo].copy()

        # Asegurar limpieza mínima
        df_temporal2[col_gerencia] = (
            df_temporal2[col_gerencia].astype(str).str.replace(r"[\r\n\t]", "", regex=True).str.strip()
        )

        return df_temporal2, col_gerencia, col_saldo, col_fuera

    def _construir_td_saldo(self, df_temporal2: pd.DataFrame, col_gerencia: str, col_saldo: str, col_fuera: str) -> pd.DataFrame:
        """
        Construye el resumen tipo 'tabla dinámica':
          - Etiquetas de fila: gerencia_responsable
          - B: Cuenta de SALDO CONTABLE  (conteo de filas por gerencia)
          - C: Cuenta de VALOR PARTIDAS FUERA DE POLITICA (conteo de filas con valor != 0)
        """
        # Conteo de filas por gerencia (tras filtro de saldo != 0)
        conteo_saldo = (
            df_temporal2
            .groupby(col_gerencia, dropna=False)
            .size()
            .rename("Cuenta de SALDO CONTABLE")
        )

        # Conteo de filas con PARTIDAS FUERA DE POLITICA_y != 0
        fuera_num = self._to_numeric_safe(df_temporal2[col_fuera])
        df_tmp = df_temporal2.copy()
        df_tmp["_flag_fuera"] = fuera_num.fillna(0).ne(0).astype(int)

        conteo_fuera = (
            df_tmp
            .groupby(col_gerencia, dropna=False)["_flag_fuera"]
            .sum()
            .rename("Cuenta de VALOR PARTIDAS FUERA DE POLITICA")
        )

        # Unir y ordenar
        td = pd.concat([conteo_saldo, conteo_fuera], axis=1).reset_index()
        td = td.rename(columns={col_gerencia: "Etiquetas de fila"}).fillna(0)

        # Totales
        total_row = {
            "Etiquetas de fila": "Total general",
            "Cuenta de SALDO CONTABLE": int(td["Cuenta de SALDO CONTABLE"].sum()),
            "Cuenta de VALOR PARTIDAS FUERA DE POLITICA": int(td["Cuenta de VALOR PARTIDAS FUERA DE POLITICA"].sum())
        }
        td = pd.concat([td, pd.DataFrame([total_row])], ignore_index=True)

        return td

    # -------------------- Exportación --------------------

    def procesar_y_exportar(self, archivo_temporales: Path, nombre_archivo_salida: Optional[str] = None,
                            ruta_destino: Optional[Path] = None) -> Path:
        """
        Orquesta el proceso y guarda el nuevo libro.
        - archivo_temporales: Path al Excel fuente (Informe Cuentas Temporales ...xlsx)
        - nombre_archivo_salida: si no se pasa, usa <nombre_origen>_procesado.xlsx
        - ruta_destino: si se pasa, prioriza esa ruta; si no, intenta SharePoint; si no, datos/salida/
        """
        archivo_temporales = Path(archivo_temporales)

        # 1) Cargar hojas necesarias
        df_temporal = self._cargar_hoja(archivo_temporales, "TEMPORAL")
        df_sabana = self._cargar_hoja(archivo_temporales, "Sábana Temporales")

        # 2) Construir TEMPORAL2 filtrada
        df_temporal2, col_gerencia, col_saldo, col_fuera = self._construir_temporal2(df_temporal)

        # 3) Construir TD Saldo
        df_td = self._construir_td_saldo(df_temporal2, col_gerencia, col_saldo, col_fuera)

        # 4) Resolver destino
        if ruta_destino:
            base_dest = Path(ruta_destino)
        elif self.ruta_sharepoint and self.ruta_sharepoint.exists():
            base_dest = self.ruta_sharepoint
        else:
            base_dest = obtener_ruta_salida()
            print(f"⚠️ SharePoint no disponible. Guardando en salida local: {base_dest}")

        base_dest.mkdir(parents=True, exist_ok=True)

        if not nombre_archivo_salida:
            nombre_archivo_salida = archivo_temporales.stem + "_procesado.xlsx"

        destino = base_dest / nombre_archivo_salida

        # 5) Escribir libro final
        with pd.ExcelWriter(destino, engine="openpyxl", mode="w") as writer:
            # Copias tal cual (valores) + hojas nuevas
            df_temporal.to_excel(writer, sheet_name="TEMPORAL", index=False)
            df_temporal2.to_excel(writer, sheet_name="TEMPORAL2", index=False)
            df_td.to_excel(writer, sheet_name="TD Saldo", index=False)
            df_sabana.to_excel(writer, sheet_name="Sábana Temporales", index=False)

        return destino
