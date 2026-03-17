# -*- coding: utf-8 -*-
# procesadores/procesador_temporales.py

from pathlib import Path
from typing import Optional, Tuple, Dict
import re
import unicodedata
import pandas as pd
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from datos.rutas import obtener_ruta_salida
from shutil import copy2


class ProcesadorTemporales:
    """
    Procesa el archivo 'Informe Cuentas Temporales ...xlsx':
      - Duplica 'TEMPORAL' -> 'TEMPORAL2' y filtra SALDO CONTABLE != 0
      - Genera 'TD Saldo' (conteos por gerencia)
      - Manipula 'Sábana Temporales' (DB/CR opcional) y construye 'TD SABANA'
      - Exporta a SharePoint (si existe ruta) o a datos/salida/ (guardado atómico)
    """

    def __init__(self, ruta_sharepoint: Optional[Path] = None):
        self.ruta_sharepoint = Path(ruta_sharepoint) if ruta_sharepoint else None

    # -------------------- UTILIDADES --------------------

    def _normalizar(self, s: str) -> str:
        """Minúsculas, sin tildes ni saltos, espacios colapsados."""
        if s is None:
            return ""
        s = re.sub(r"[\r\n\t]+", " ", str(s)).strip().lower()
        s = "".join(ch for ch in unicodedata.normalize("NFD", s)
                    if unicodedata.category(ch) != "Mn")
        return re.sub(r"\s+", " ", s)

    def _encontrar_columna(self, columnas: list, objetivo: str) -> Optional[str]:
        """Encuentra el nombre real de una columna sin importar tildes o mayúsculas."""
        objetivo_n = self._normalizar(objetivo)
        mapa = {self._normalizar(c): c for c in columnas}
        return mapa.get(objetivo_n)

    def _to_numeric_safe(self, serie: pd.Series) -> pd.Series:
        """
        Convierte a numérico: limpia símbolos, soporta miles con punto y decimales con coma.
        - '1.234.567,89' -> 1234567.89
        - '1234,56'      -> 1234.56
        """
        if serie is None:
            return pd.Series(dtype=float)
        s = serie.astype(str)
        s = s.str.replace(r"[\r\n\t]", "", regex=True)
        s = s.str.replace(r"[^\d\-.,]", "", regex=True).str.strip()
        # Si hay coma y punto -> quitar puntos (miles) y dejar coma como decimal
        mask_coma = s.str.contains(",", regex=False)
        mask_punto = s.str.contains(r"\.", regex=True)
        s = s.where(~(mask_coma & mask_punto),
                    s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
        # Si solo hay coma -> usarla como decimal
        s = s.where(~(mask_coma & ~mask_punto),
                    s.str.replace(",", ".", regex=False))
        return pd.to_numeric(s, errors="coerce")

    def _limpiar_para_excel(self, df: pd.DataFrame) -> pd.DataFrame:
        """Quita caracteres ilegales para XLSX en columnas de texto."""
        df2 = df.copy()
        for c in df2.select_dtypes(include=["object"]).columns:
            df2[c] = (
                df2[c].astype(str)
                .apply(lambda x: ILLEGAL_CHARACTERS_RE.sub("", x))
                .str.replace(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", regex=True)
            )
        return df2

    def _guardar_atomico(self, destino: Path, hojas: Dict[str, pd.DataFrame]) -> Path:
        """
        Escribe a un temporal .__tmp__.xlsx y luego reemplaza por el definitivo.
        Si el destino ya existe, intenta eliminarlo primero (evita duplicados en OneDrive).
        """
        destino.parent.mkdir(parents=True, exist_ok=True)
        tmp = destino.parent / f"{destino.stem}.__tmp__.xlsx"

        with pd.ExcelWriter(tmp, engine="openpyxl", mode="w") as writer:
            for nombre, df in hojas.items():
                self._limpiar_para_excel(df).to_excel(writer, sheet_name=nombre, index=False)

        if destino.exists():
            try:
                destino.unlink()
            except PermissionError:
                # Si está abierto, replace igual sobrescribe
                pass

        tmp.replace(destino)
        return destino

    def _duplicar_a_salida_local(self, origen: Path) -> Path:
        """
        Duplica el archivo procesado hacia la carpeta local de salida (obtener_ruta_salida()).
        Verifica lectura mínima para asegurar que no quedó vacío/corrupto.
        """
        base_local = obtener_ruta_salida()
        base_local.mkdir(parents=True, exist_ok=True)

        destino_local = base_local / origen.name

        try:
            copy2(str(origen), str(destino_local))
        except Exception as e:
            print(f"⚠️ No se pudo copiar a salida local: {e}")
            return destino_local

        # Verificación rápida de lectura
        try:
            _ = pd.read_excel(destino_local, sheet_name=0, nrows=1, engine="openpyxl")
            print(
                f"🗂️ Copia local OK: {destino_local} "
                f"({round(destino_local.stat().st_size / 1024, 1)} KB)"
            )
        except Exception as e:
            print(f"⚠️ Copia local creada pero no legible: {e}")

        return destino_local

    # -------------------- TRANSFORMACIONES --------------------

    def _cargar_hoja(self, archivo: Path, nombre_hoja: str) -> pd.DataFrame:
        return pd.read_excel(archivo, sheet_name=nombre_hoja, dtype=str, engine="openpyxl")

    def _construir_temporal2(self, df: pd.DataFrame) -> Tuple[pd.DataFrame, str, str, str]:
        """Filtra SALDO CONTABLE != 0 y retorna df2 + nombres reales de columnas clave."""
        col_saldo = self._encontrar_columna(df.columns, "SALDO CONTABLE")
        col_ger   = self._encontrar_columna(df.columns, "gerencia_responsable")
        col_fuera = self._encontrar_columna(df.columns, "PARTIDAS FUERA DE POLITICA_y")
        faltan = [n for n, v in {
            "SALDO CONTABLE": col_saldo,
            "gerencia_responsable": col_ger,
            "PARTIDAS FUERA DE POLITICA_y": col_fuera
        }.items() if v is None]
        if faltan:
            raise ValueError(f"No encontré columnas requeridas en 'TEMPORAL': {', '.join(faltan)}")

        saldo = self._to_numeric_safe(df[col_saldo]).fillna(0)
        df2 = df[saldo != 0].copy()
        # Limpieza leve de gerencia
        df2[col_ger] = df2[col_ger].astype(str).str.replace(r"[\r\n\t]", "", regex=True).str.strip()
        return df2, col_ger, col_saldo, col_fuera

    def _construir_td_saldo(self, df2: pd.DataFrame, col_ger: str, col_saldo: str, col_fuera: str) -> pd.DataFrame:
        """'TD Saldo': conteo de cuentas con saldo y conteo de fuera de política por gerencia."""
        td1 = df2.groupby(col_ger, dropna=False).size().reset_index(name="Cuenta de SALDO CONTABLE")
        flag_fuera = self._to_numeric_safe(df2[col_fuera]).fillna(0).ne(0)
        td2 = flag_fuera.groupby(df2[col_ger]).sum().reset_index(name="Cuenta de VALOR PARTIDAS FUERA DE POLITICA")
        out = td1.merge(td2, on=col_ger, how="left").fillna(0)
        out.rename(columns={col_ger: "Etiquetas de fila"}, inplace=True)
        out.loc[len(out)] = [
            "Total general",
            int(out["Cuenta de SALDO CONTABLE"].sum()),
            int(out["Cuenta de VALOR PARTIDAS FUERA DE POLITICA"].sum())
        ]
        return out

    def _agregar_db_cr_sabana(self, df_sabana: pd.DataFrame) -> pd.DataFrame:
        """
        Desde L 'VALOR PARTIDA PESOS' crea:
          - 'VALOR PARTIDA PESOS DB' (solo > 0, positivo)
          - 'VALOR PARTIDA PESOS CR' (solo < 0, conserva signo negativo)
        Inserta ambas columnas a la derecha de L (sin añadir filas de control).
        """
        cols = list(df_sabana.columns)
        col_valor = self._encontrar_columna(cols, "VALOR PARTIDA PESOS")
        if not col_valor:
            raise ValueError("No encontré 'VALOR PARTIDA PESOS' en 'Sábana Temporales'.")

        serie = self._to_numeric_safe(df_sabana[col_valor]).fillna(0.0)
        db = serie.where(serie > 0, 0.0)
        cr = serie.where(serie < 0, 0.0)  # conserva signo negativo

        df_out = df_sabana.copy()
        for c in ["VALOR PARTIDA PESOS DB", "VALOR PARTIDA PESOS CR"]:
            if c in df_out.columns:
                df_out = df_out.drop(columns=[c])

        pos = cols.index(col_valor) + 1
        df_out.insert(pos,     "VALOR PARTIDA PESOS DB", db)
        df_out.insert(pos + 1, "VALOR PARTIDA PESOS CR", cr)

        # Mensaje de control por consola (no en celdas):
        suma_L  = float(serie.sum())
        suma_DB = float(db.sum())
        suma_CR = float(cr.sum())
        delta   = (suma_DB + suma_CR) - suma_L
        if abs(delta) > 0.01:
            print("⚠️ Aviso Sábana: L vs M+N difiere.",
                  f"L={suma_L:,.2f} | M={suma_DB:,.2f} | N={suma_CR:,.2f} | Δ={delta:,.2f}")
        else:
            print("✅ Sumas OK Sábana:",
                  f"L={suma_L:,.2f} == M+N={suma_DB + suma_CR:,.2f} (Δ={delta:,.2f})")
        return df_out

    def _construir_td_sabana(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        'TD SABANA':
          - Filas: gerencia_responsable
          - Columnas: FUERA DE POLITICA (SI/NO)
          - Valores: Cuenta de 'VALOR PARTIDA PESOS'
          - Incluye columna y fila 'Total general'
        """
        col_g = self._encontrar_columna(df.columns, "gerencia_responsable")
        col_p = self._encontrar_columna(df.columns, "FUERA DE POLITICA")
        col_v = self._encontrar_columna(df.columns, "VALOR PARTIDA PESOS")
        faltan = [n for n, v in {
            "gerencia_responsable": col_g, "FUERA DE POLITICA": col_p, "VALOR PARTIDA PESOS": col_v
        }.items() if v is None]
        if faltan:
            raise ValueError(f"No encontré columnas para TD SABANA: {', '.join(faltan)}")

        dfp = df[[col_g, col_p, col_v]].copy()
        dfp[col_g] = dfp[col_g].astype(str).str.replace(r"[\r\n\t]", "", regex=True).str.strip()
        dfp[col_p] = (dfp[col_p].astype(str)
                      .str.replace(r"[\r\n\t]", "", regex=True).str.strip().str.upper()
                      .replace({"SÍ": "SI"}))

        td = pd.pivot_table(dfp, index=col_g, columns=col_p, values=col_v,
                            aggfunc="count", fill_value=0, dropna=False)

        # Orden de columnas esperado
        orden_cols = [c for c in ["NO", "SI"] if c in td.columns]
        td = td[orden_cols] if orden_cols else td

        td["Total general"] = td.sum(axis=1)
        td.loc["Total general"] = td.sum()

        return td.reset_index().rename(columns={col_g: "Etiquetas de fila"})

    def _agregar_filas_control_excel_sabana(self, xlsx_path: Path, hoja: str = "Sábana Temporales") -> None:
        """
        Abre el archivo guardado y añade dos filas de control al final de la hoja 'Sábana Temporales':
          - Fila 'TOTAL SUMA': SUM(L2:Llast), SUBTOTAL(9,M2:Mlast), SUBTOTAL(9,N2:Nlast)
          - Fila 'M+N - L (debe ser 0)': SUBTOTAL(9,M...)+SUBTOTAL(9,N...)-SUM(L...)
        Usa funciones en INGLÉS (SUM, SUBTOTAL) para compatibilidad OOXML.
        Cambia func=9 a func=109 si quieres que ignore filas ocultas por filtros.
        """
        xlsx_path = Path(xlsx_path)
        wb = load_workbook(xlsx_path)
        if hoja not in wb.sheetnames:
            wb.save(xlsx_path); return
        ws = wb[hoja]

        # Mapear encabezados -> índice de columna
        headers = {str(ws.cell(row=1, column=c).value).strip(): c for c in range(1, ws.max_column + 1)}
        cL = headers.get("VALOR PARTIDA PESOS")
        cM = headers.get("VALOR PARTIDA PESOS DB")
        cN = headers.get("VALOR PARTIDA PESOS CR")
        if not all([cL, cM, cN]):
            wb.save(xlsx_path); return

        L = get_column_letter(cL)
        M = get_column_letter(cM)
        N = get_column_letter(cN)

        first_data = 2                      # pandas escribe encabezados en fila 1
        last_data  = ws.max_row             # última fila con datos
        total_row  = last_data + 1
        diff_row   = last_data + 2

        # Etiquetas (en la primera columna de la hoja)
        ws.cell(row=total_row, column=1, value="TOTAL SUMA")
        ws.cell(row=diff_row,  column=1, value="M+N - L (debe ser 0)")

        # Fórmulas (OOXML en inglés)
        func = 9  # usa 109 si quieres ignorar filas ocultas por filtro
        ws.cell(row=total_row, column=cL, value=f"=SUM({L}{first_data}:{L}{last_data})")
        ws.cell(row=total_row, column=cM, value=f"=SUBTOTAL({func},{M}{first_data}:{M}{last_data})")
        ws.cell(row=total_row, column=cN, value=f"=SUBTOTAL({func},{N}{first_data}:{N}{last_data})")

        ws.cell(
            row=diff_row, column=cL,
            value=f"=SUBTOTAL({func},{M}{first_data}:{M}{last_data})+SUBTOTAL({func},{N}{first_data}:{N}{last_data})-SUM({L}{first_data}:{L}{last_data})"
        )

        # Formato de número
        for r in (total_row, diff_row):
            for c in (cL, cM, cN):
                ws.cell(row=r, column=c).number_format = "#,##0.00"

        wb.save(xlsx_path)

    # -------------------- ORQUESTADOR --------------------

    def procesar_y_exportar(self, archivo_temporales: Path, crear_db_cr_sabana: bool = True) -> Path:
        """
        Ejecuta el flujo y guarda el libro procesado (sobrescribe si ya existe).
        Retorna la ruta del archivo generado.
        """
        archivo_temporales = Path(archivo_temporales)

        print("1) Cargando hojas TEMPORAL y Sábana Temporales…")
        df_temporal = self._cargar_hoja(archivo_temporales, "TEMPORAL")
        df_sabana   = self._cargar_hoja(archivo_temporales, "Sábana Temporales")

        print("2) Construyendo TEMPORAL2 (SALDO CONTABLE != 0)…")
        df_temporal2, col_ger, col_saldo, col_fuera = self._construir_temporal2(df_temporal)

        print("3) Construyendo TD Saldo…")
        td_saldo = self._construir_td_saldo(df_temporal2, col_ger, col_saldo, col_fuera)

        if crear_db_cr_sabana:
            print("4) Agregando DB/CR (M y N) en Sábana…")
            df_sabana_proc = self._agregar_db_cr_sabana(df_sabana)
        else:
            print("4) (Omitido) DB/CR en Sábana por configuración.")
            df_sabana_proc = df_sabana

        print("5) Construyendo TD SABANA…")
        td_sabana = self._construir_td_sabana(df_sabana)

        # Resolver destino (SharePoint si existe; si no, salida local)
        base_dest = self.ruta_sharepoint if (self.ruta_sharepoint and self.ruta_sharepoint.exists()) else obtener_ruta_salida()
        if base_dest == obtener_ruta_salida():
            print(f"⚠️ SharePoint no disponible. Guardando en salida local: {base_dest}")

        destino = base_dest / f"{archivo_temporales.stem}_procesado.xlsx"

        # Guardado atómico (SOBREESCRIBE)
        destino = self._guardar_atomico(destino, {
            "TEMPORAL": df_temporal,
            "TEMPORAL2": df_temporal2,
            "TD Saldo": td_saldo,
            "Sábana Temporales": df_sabana_proc,
            "TD SABANA": td_sabana
        })

        # Añadir fórmulas de control al final de 'Sábana Temporales'
        self._agregar_filas_control_excel_sabana(destino, hoja="Sábana Temporales")

        # Copia de seguridad local (además de SharePoint) **ya con fórmulas**
        self._duplicar_a_salida_local(destino)

        # Verificación rápida
        try:
            _ = pd.read_excel(destino, sheet_name="TEMPORAL", nrows=1, engine="openpyxl")
            print(f"✅ Verificación OK: {destino.name} se puede leer.")
        except Exception as e:
            print(f"❌ Alerta: el archivo se guardó pero no se pudo leer con pandas: {e}")

        return destino
