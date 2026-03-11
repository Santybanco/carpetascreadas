# -*- coding: utf-8 -*-
# config/configuracion.py
from pathlib import Path

class Configuracion:
    """
    Configuración central de rutas. Mantener simple.
    - Por defecto trabaja en carpetas locales (datos/entrada, datos/salida).
    - Para producción, puedes apuntar a la ruta de red (UNC) descomentando RUTA_ENTRADA_RED.
    """
    BASE_DIR = Path(__file__).resolve().parents[1]  # carpeta raíz del proyecto (ISCE)

    # Rutas locales (para desarrollo/pruebas)
    RUTA_ENTRADA_LOCAL = BASE_DIR / "datos" / "entrada"
    RUTA_SALIDA_LOCAL  = BASE_DIR / "datos" / "salida"

    # Ejemplo de ruta de red (cuando pasemos a productivo)
    # RUTA_ENTRADA_RED = Path(r"\\servidor\carpeta\...\Indicadores\Entrada")

    # Usa local por ahora
    RUTA_ENTRADA = RUTA_ENTRADA_LOCAL
    RUTA_SALIDA  = RUTA_SALIDA_LOCAL
