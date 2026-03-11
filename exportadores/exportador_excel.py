# -*- coding: utf-8 -*-
# exportadores/exportador_excel.py

from pathlib import Path
import pandas as pd
from openpyxl import load_workbook, Workbook


class ExportadorExcel:
    """
    Exporta resultados al libro final:
    - Atención de alertas con calidad (ALCON)
    - Histórico de Certificación de Gerentes
    """

    def __init__(self, ruta_salida: Path):
        self.ruta_salida = Path(ruta_salida)

    # ==============================================================
    #   1. EXPORTAR TABLA DE ALCON (A83)
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

        startrow = fila_inicio_excel - 1  # 83 -> fila 82 pandas
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

        # 2) Abrir con openpyxl para aplicar formato porcentaje en 'Calidad Gerencia'
        wb = load_workbook(destino)
        ws = wb[nombre_hoja_mes]

        # Calculamos filas reales de datos escritos
        first_data_row = startrow + 1  # fila Excel donde quedaron los encabezados (A83)
        data_start_row = first_data_row + 1  # primera fila de datos
        data_end_row = data_start_row + len(df) - 1
        
        if len(df) > 0:
            # 'Calidad Gerencia' es la 5ta columna dentro del bloque que pegamos (A..E)
            col_index = 5  # E
            for r in range(data_start_row, data_end_row + 1):
                cell = ws.cell(row=r, column=col_index)
                # Asegurar que es número (por si algo vino raro)
                try:
                    if isinstance(cell.value, str) and cell.value.strip() != "":
                        cell.value = float(cell.value)
                except:
                    pass
                cell.number_format = "0.00%"

        wb.save(destino)
        return destino

    # ==============================================================
    #   2. EXPORTAR HISTÓRICO DE CERTIFICACIÓN (A180)
    # ==============================================================

    def exportar_historico_certificacion(
        self,
        df: pd.DataFrame,
        nombre_hoja_mes: str,
        nombre_archivo: str = "Indicadores_operacion_nuevo.xlsx",
        fila_inicio_excel: int = 180,  # encabezados en A180
        col_inicio_excel: int = 1,     # columna A
        escribir_promedio: bool = True # 👈 activar/desactivar promedio
    ) -> Path:
        """
        Escribe el histórico y opcionalmente calcula el promedio de la columna INDICADOR.
        """
        destino = self.ruta_salida / nombre_archivo
        destino.parent.mkdir(parents=True, exist_ok=True)

        # Asegurar que no existan NA
        df = df.fillna("")

        # Abrir o crear libro
        if destino.exists():
            wb = load_workbook(destino)
        else:
            wb = Workbook()

        # Hoja del mes
        if nombre_hoja_mes in wb.sheetnames:
            ws = wb[nombre_hoja_mes]
        else:
            ws = wb.create_sheet(title=nombre_hoja_mes)

        # --- Escribir encabezados ---
        headers = list(df.columns)
        for j, h in enumerate(headers, start=col_inicio_excel):
            ws.cell(row=fila_inicio_excel, column=j, value=h)

        # --- Escribir datos ---
        first_data_row = fila_inicio_excel + 1
        for i in range(len(df)):
            for j, h in enumerate(headers, start=col_inicio_excel):
                value = df.iloc[i][h]
                
                # Guardar la celda con la posición correcta
                cell = ws.cell(row=first_data_row + i, column=j)

                # Columna INDICADOR -> aplicar formato porcentaje
                if h == "INDICADOR" and value not in ("", None):
                    try:
                        num = float(value)
                        cell.value = num
                        cell.number_format = "0.00%"
                    except:
                        cell.value = value
                else:
                    if pd.isna(value):
                        value = ""
                    cell.value = value

        # --- (Opcional) Escribir el PROMEDIO de la columna INDICADOR ---
        if escribir_promedio:
            try:
                idx_ind = headers.index("INDICADOR")
            except ValueError:
                idx_ind = None

            if idx_ind is not None:
                def col_to_letter(col_num: int) -> str:
                    letters = ""
                    while col_num > 0:
                        col_num, remainder = divmod(col_num - 1, 26)
                        letters = chr(65 + remainder) + letters
                    return letters

                col_ind_excel = col_inicio_excel + idx_ind
                col_ind_letra = col_to_letter(col_ind_excel)

                last_data_row = first_data_row + len(df) - 1
                promedio_row  = last_data_row + 1

                if col_ind_excel - 1 >= 1:
                    ws.cell(row=promedio_row, column=col_ind_excel - 1, value="Promedio INDICADOR")

                formula = f"=AVERAGE({col_ind_letra}{first_data_row}:{col_ind_letra}{last_data_row})"
                ws.cell(row=promedio_row, column=col_ind_excel, value=formula)
                ws.cell(row=promedio_row, column=col_ind_excel).number_format = "0.00%"

        wb.save(destino)
        return destino
