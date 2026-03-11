# -*- coding: utf-8 -*-
# datos/rutas.py
from pathlib import Path
from config.configuracion import Configuracion

def obtener_ruta_entrada() -> Path:
    """
    Retorna la ruta de entrada y la crea si no existe.
    """
    ruta = Configuracion.RUTA_ENTRADA
    ruta.mkdir(parents=True, exist_ok=True)
    return ruta

def obtener_ruta_salida() -> Path:
    """
    Retorna la ruta de salida y la crea si no existe.
    """
    ruta = Configuracion.RUTA_SALIDA
    ruta.mkdir(parents=True, exist_ok=True)
    return ruta