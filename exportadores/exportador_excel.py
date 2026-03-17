# -*- coding: utf-8 -*-
# exportadores/exportador_excel.py

from pathlib import Path
import re
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font


class ExportadorExcel:
    """
    Exporta resultados al libro final:
    - Atención de alertas con calidad (ALCON)
    - Histórico de Certificación de Gerentes
    - TD Saldo (Temporales) a A4 con % calculado
    - TD SABANA (Temporales) a E4 con % calculado
    """

    def __init__(self, ruta_salida: Path):
        self.ruta_salida = Path(ruta_salida)

    # ==============================================================
    #   1) ALCON → hoja del mes, encabezados en A83 (formato %)
    # ==============================================================
    def exportar_atencion_alertas_calidad(
        self,
        df: pd.DataFrame,
        nombre_hoja_mes: str,
        nombre_archivo: str = "Indicadores_operacion_nuevo.xlsx",
        fila_inicio_excel: int = 83  # Encabezados deben ir en A83
    ) -> Path:

        destino = self.ruta_salida / nombre_archivo
        destino.parent.mkdir(parents=True, exist_ok=True)

        startrow = fila_inicio_excel - 1  # 83 -> fila 82 (pandas es 0-based)
        modo = "a" if destino.exists() else "w"

        # 1) Escribir con pandas
        with pd.ExcelWriter(
            destino,
            engine="openpyxl",
            mode=modo,
            if_sheet_exists="replace" if modo == "a" else None
        ) as writer:
            df.to_excel(
                writer,
                sheet_name=nombre_hoja_mes,
                index=False,
                startrow=startrow,
                startcol=0
            )

        # 2) Formato % en 'Calidad Gerencia' (columna 5 = E)
        wb = load_workbook(destino)
        ws = wb[nombre_hoja_mes]

        first_data_row = startrow + 1      # fila de encabezados (Excel)
        data_start_row = first_data_row + 1
        data_end_row = data_start_row + len(df) - 1

        if len(df) > 0:
            col_index = 5  # E
            for r in range(data_start_row, data_end_row + 1):
                cell = ws.cell(row=r, column=col_index)
                try:
                    if isinstance(cell.value, str) and cell.value.strip() != "":
                        cell.value = float(cell.value)
                except Exception:
                    pass
                cell.number_format = "0.00%"

        wb.save(destino)
        return destino

    # ==============================================================
    #   2) HISTÓRICO → A180, % en INDICADOR y PROMEDIO dinámico
    # ==============================================================
    def exportar_historico_certificacion(
        self,
        df: pd.DataFrame,
        nombre_hoja_mes: str,
        nombre_archivo: str = "Indicadores_operacion_nuevo.xlsx",
        fila_inicio_excel: int = 180,  # encabezados en A180
        col_inicio_excel: int = 1,     # columna A
        escribir_promedio: bool = True
    ) -> Path:
        """Escribe el histórico y opcionalmente calcula el promedio de INDICADOR."""
        destino = self.ruta_salida / nombre_archivo
        destino.parent.mkdir(parents=True, exist_ok=True)

        df = df.fillna("")

        wb = load_workbook(destino) if destino.exists() else Workbook()

        ws = wb[nombre_hoja_mes] if nombre_hoja_mes in wb.sheetnames else wb.create_sheet(title=nombre_hoja_mes)

        # Encabezados
        headers = list(df.columns)
        for j, h in enumerate(headers, start=col_inicio_excel):
            ws.cell(row=fila_inicio_excel, column=j, value=h)

        # Datos
        first_data_row = fila_inicio_excel + 1
        for i in range(len(df)):
            for j, h in enumerate(headers, start=col_inicio_excel):
                value = df.iloc[i][h]
                cell = ws.cell(row=first_data_row + i, column=j)

                if h == "INDICADOR" and value not in ("", None):
                    try:
                        num = float(value)
                        cell.value = num
                        cell.number_format = "0.00%"
                    except Exception:
                        cell.value = value
                else:
                    if pd.isna(value):
                        value = ""
                    cell.value = value

        # PROMEDIO dinámico de INDICADOR (en %)
        if escribir_promedio:
            try:
                idx_ind = headers.index("INDICADOR")
            except ValueError:
                idx_ind = None

            if idx_ind is not None and len(df) > 0:
                def col_to_letter(col_num: int) -> str:
                    letters = ""
                    while col_num > 0:
                        col_num, remainder = divmod(col_num - 1, 26)
                        letters = chr(65 + remainder) + letters
                    return letters

                col_ind_excel = col_inicio_excel + idx_ind
                col_ind_letra = col_to_letter(col_ind_excel)
                last_data_row = first_data_row + len(df) - 1
                promedio_row = last_data_row + 1

                if col_ind_excel - 1 >= 1:
                    ws.cell(row=promedio_row, column=col_ind_excel - 1, value="Promedio INDICADOR")

                formula = f"=AVERAGE({col_ind_letra}{first_data_row}:{col_ind_letra}{last_data_row})"
                ws.cell(row=promedio_row, column=col_ind_excel, value=formula)
                ws.cell(row=promedio_row, column=col_ind_excel).number_format = "0.00%"

        wb.save(destino)
        return destino

    # ==============================================================
    #   3) TD Saldo (Temporales) → A4 con % calculado (IFERROR)
    # ==============================================================
    def exportar_td_saldo_en_indicadores(
        self,
        df_td: pd.DataFrame,
        nombre_hoja_mes: str,
        nombre_archivo: str = "Indicadores_operacion_nuevo.xlsx",
        celda_inicio: str = "A4"
    ) -> Path:
        """
        Pega TD Saldo en la hoja 'nombre_hoja_mes' del libro final.
        - Encabezados en A4: Area | TOTAL CUENTAS TEMPORALES CON SALDO | CUENTAS TEMPORALES FUERA DE POLITICA | %
        - Datos desde A5 hacia abajo.
        - Columna % con fórmula OOXML válida: =IFERROR(Cfila/Bfila,0)
        - Da formato: números enteros (B,C) y porcentaje 0.00% (D).
        """
        destino = self.ruta_salida / nombre_archivo
        if not destino.exists():
            raise FileNotFoundError(
                f"No existe el archivo final '{destino}'. Genéralo primero (ALCON/HISTÓRICO) y vuelve a intentar."
            )

        wb = load_workbook(destino)
        ws = wb[nombre_hoja_mes] if nombre_hoja_mes in wb.sheetnames else wb.create_sheet(title=nombre_hoja_mes)

        # Parsear celda inicio (ej. 'A4' -> col=1, row=4)
        m = re.match(r"^([A-Za-z]+)(\d+)$", celda_inicio)
        if not m:
            raise ValueError(f"Celda inicio inválida: {celda_inicio}")
        col_letters, row_str = m.groups()

        def col_letter_to_index(col: str) -> int:
            col = col.upper()
            n = 0
            for ch in col:
                n = n * 26 + (ord(ch) - 64)
            return n

        start_col = col_letter_to_index(col_letters)
        start_row = int(row_str)

        # Renombrar columnas según plantilla
        mapa = {
            "Etiquetas de fila": "Area",
            "Cuenta de SALDO CONTABLE": "TOTAL CUENTAS TEMPORALES CON SALDO",
            "Cuenta de VALOR PARTIDAS FUERA DE POLITICA": "CUENTAS TEMPORALES FUERA DE POLITICA",
        }
        df_es = df_td.rename(columns=mapa).copy()

        # Validación mínima
        requeridas = [
            "Area",
            "TOTAL CUENTAS TEMPORALES CON SALDO",
            "CUENTAS TEMPORALES FUERA DE POLITICA",
        ]
        for col in requeridas:
            if col not in df_es.columns:
                raise ValueError(f"Falta la columna requerida en TD Saldo: '{col}'")

        # Limpiar bloque (A4:D500) para evitar residuos
        for r in range(start_row, start_row + 500):
            for c in range(start_col, start_col + 4):
                ws.cell(row=r, column=c, value=None)

        # Encabezados (A4:D4)
        headers = [
            "Area",
            "TOTAL CUENTAS TEMPORALES CON SALDO",
            "CUENTAS TEMPORALES FUERA DE POLITICA",
            "%"
        ]
        for j, h in enumerate(headers, start=start_col):
            cell = ws.cell(row=start_row, column=j, value=h)
            cell.font = Font(bold=True)

        # Datos y fórmula %
        first_data_row = start_row + 1
        for i in range(len(df_es)):
            area = df_es.iloc[i]["Area"]

            # B y C numéricos (0 si vacío)
            try:
                tot_cuentas = int(str(df_es.iloc[i]["TOTAL CUENTAS TEMPORALES CON SALDO"]).strip() or "0")
            except Exception:
                tot_cuentas = 0
            try:
                tot_fuera = int(str(df_es.iloc[i]["CUENTAS TEMPORALES FUERA DE POLITICA"]).strip() or "0")
            except Exception:
                tot_fuera = 0

            fila_excel = first_data_row + i

            # A: texto
            ws.cell(row=fila_excel, column=start_col + 0, value=area)

            # B y C: números + formato miles
            cB = ws.cell(row=fila_excel, column=start_col + 1, value=tot_cuentas)
            cC = ws.cell(row=fila_excel, column=start_col + 2, value=tot_fuera)
            cB.number_format = "#,##0"
            cC.number_format = "#,##0"

            # D = % -> fórmula OOXML válida (inglés y coma)
            colB = self._col_letra(start_col + 1)  # B
            colC = self._col_letra(start_col + 2)  # C
            formula = f"=IFERROR({colC}{fila_excel}/{colB}{fila_excel},0)"
            cD = ws.cell(row=fila_excel, column=start_col + 3, value=formula)
            cD.number_format = "0.00%"

        wb.save(destino)
        return destino

    # ==============================================================
    #   4) TD SABANA → E4 con % calculado (IFERROR)
    # ==============================================================
    def exportar_td_sabana_en_indicadores(
        self,
        df_td_sabana: pd.DataFrame,
        nombre_hoja_mes: str,
        nombre_archivo: str = "Indicadores_operacion_nuevo.xlsx",
        celda_inicio: str = "E4",
        formato_porcentaje: str = "0.0%"  # usa "0.00%" si prefieres 2 decimales
    ) -> Path:
        """
        Pega la TD SABANA en la hoja del mes a partir de E4:
          E: TOTAL PARTIDAS
          F: PARTIDAS FUERA DE POLITICA (columna 'SI')
          G: % = IFERROR(F/E,0)
        - Respeta el archivo final existente; no toca el resto de la hoja.
        """
        destino = self.ruta_salida / nombre_archivo
        if not destino.exists():
            raise FileNotFoundError(
                f"No existe el archivo final '{destino}'. Genéralo primero (ALCON/HISTÓRICO/TD Saldo) y vuelve a intentar."
            )

        wb = load_workbook(destino)
        ws = wb[nombre_hoja_mes] if nombre_hoja_mes in wb.sheetnames else wb.create_sheet(title=nombre_hoja_mes)

        # Parsear 'E4' -> (col_ini, fila_ini)
        m = re.match(r"^([A-Za-z]+)(\d+)$", celda_inicio)
        if not m:
            raise ValueError(f"Celda inicio inválida: {celda_inicio}")
        col_letters, row_str = m.groups()

        def col_letter_to_index(col: str) -> int:
            col = col.upper()
            n = 0
            for ch in col:
                n = n * 26 + (ord(ch) - 64)
            return n

        start_col = col_letter_to_index(col_letters)  # E
        start_row = int(row_str)                      # 4

        # Preparar datos (Total Partidas, Partidas fuera de política)
        cols = {str(c).strip(): c for c in df_td_sabana.columns}
        col_total = cols.get("Total general")
        col_si    = cols.get("SI")
        col_no    = cols.get("NO")

        if col_total is None and (col_si is None and col_no is None):
            raise ValueError("TD SABANA no tiene columnas suficientes: se espera 'Total general' o 'SI'/'NO'.")

        if col_total is not None:
            total_partidas = pd.to_numeric(df_td_sabana[col_total], errors="coerce").fillna(0).astype(int)
        else:
            si_vals = pd.to_numeric(df_td_sabana[col_si], errors="coerce").fillna(0)
            no_vals = pd.to_numeric(df_td_sabana[col_no], errors="coerce").fillna(0) if col_no is not None else 0
            total_partidas = (si_vals + no_vals).astype(int)

        fuera_politica = pd.to_numeric(df_td_sabana[col_si], errors="coerce").fillna(0).astype(int) if col_si is not None else pd.Series([0]*len(df_td_sabana))

        # Limpiar bloque previo (E4:G500 por defecto)
        max_filas_borrar = max(500, len(df_td_sabana) + 20)
        for r in range(start_row, start_row + max_filas_borrar):
            for c in range(start_col, start_col + 3):  # E, F, G
                ws.cell(row=r, column=c, value=None)

        # Encabezados en E4:F4:G4
        headers = ["TOTAL PARTIDAS", "PARTIDAS FUERA DE POLITICA", "%"]
        for j, h in enumerate(headers, start=start_col):
            cell = ws.cell(row=start_row, column=j, value=h)
            cell.font = Font(bold=True)

        # Escribir datos desde E5 y fórmula en G
        first_data_row = start_row + 1
        for i in range(len(df_td_sabana)):
            fila_excel = first_data_row + i

            # E y F como enteros
            e_cell = ws.cell(row=fila_excel, column=start_col + 0, value=int(total_partidas.iloc[i]))
            f_cell = ws.cell(row=fila_excel, column=start_col + 1, value=int(fuera_politica.iloc[i]))
            e_cell.number_format = "#,##0"
            f_cell.number_format = "#,##0"

            # G = IFERROR(F/E,0) -> OOXML inglés
            colE = _col_letter(start_col + 0)
            colF = _col_letter(start_col + 1)
            g_cell = ws.cell(row=fila_excel, column=start_col + 2, value=f"=IFERROR({colF}{fila_excel}/{colE}{fila_excel},0)")
            g_cell.number_format = formato_porcentaje

        wb.save(destino)
        return destino

    # Helper: número de columna -> letra (A, B, ..., Z, AA, AB, ...)
    def _col_letra(self, col_num: int) -> str:
        letters = ""
        while col_num > 0:
            col_num, rem = divmod(col_num - 1, 26)
            letters = chr(65 + rem) + letters
        return letters


# Helper alterno (fuera de la clase) si lo necesitas en otros módulos
def _col_letter(col_num: int) -> str:
    letters = ""
    while col_num > 0:
        col_num, rem = divmod(col_num - 1, 26)
        letters = chr(65 + rem) + letters
    return letters
