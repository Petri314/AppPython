import pandas as pd
import re
import json
import sys
from typing import List, Dict
from collections import defaultdict

def normalizar_texto(texto: str) -> str:
    """Normaliza el texto eliminando espacios extra y convirtiendo a mayúsculas."""
    return " ".join(texto.split()).upper()

def obtener_horarios_turno(turno: str) -> tuple:
    """Devuelve los horarios de inicio, break y fin según el turno."""
    if turno.lower() == "noche":
        return "22:00", "02:45", "03:15", "06:05"
    elif turno.lower() == "tarde":
        return "14:15", "18:40", "19:10", "21:55"
    elif turno.lower() == "mañana":
        return "06:35", "10:30", "11:00", "14:10"
    return None, None, None, None

def formatear_nombre_apilador(nombre_completo: str) -> str:
    """Formatea la primera letra de cada nombre propio en mayúscula."""
    nombres = nombre_completo.split()
    nombres_formateados = [nombre.capitalize() for nombre in nombres]
    return " ".join(nombres_formateados)

def procesar_excel(excel_file_path: str, sheet_name: str = "noche", turno: str = "Noche") -> tuple[Dict[str, List[Dict]], List[str]]:
    """
    Procesa el Excel para emparejar apiladores y devuelve un diccionario de tareas por día
    y la lista de tareas clave.
    """
    try:
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name, header=1)
    except Exception as e:
        print(f"Error al leer el archivo Excel: {e}")
        return {}, []

    df.columns = [col.strip() for col in df.columns]

    tareas_clave_base = [
        "T", "P", "R", "U", "SECO", "TROPICALES-XDOCK",
        "V1-V2 Y1-Y2", "V3-V4 Y3-Y4", "V5-V9", "Y5-Y9", "W1-W4 Z1-Z5",
        "N", "H", "AA-AG", "AH-AJ", "BA-BG", "BH-BJ"
    ]
    tareas_clave_noche = tareas_clave_base + ["EKONO"]
    tareas_clave_otros = tareas_clave_base

    camaras_por_tarea = {
        "T": "Congelado", "P": "Congelado", "R": "Congelado", "U": "Congelado",
        "SECO": "Seco", "TROPICALES-XDOCK": "Tropicales",
        "V1-V2 Y1-Y2": "Panaderia", "V3-V4 Y3-Y4": "Panaderia",
        "V5-V9": "Vegetales", "Y5-Y9": "Vegetales", "W1-W4 Z1-Z5": "Vegetales",
        "N": "Carnes", "H": "Carnes",
        "AA-AG": "Fiambreria", "AH-AJ": "Fiambreria",
        "BA-BG": "Fiambreria", "BH-BJ": "Fiambreria",
        "EKONO": "Ekono"
    }

    resultados_por_dia = defaultdict(list)
    dias_semana_procesar = ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO"]
    hora_inicio, hora_break_inicio, hora_break_fin, hora_fin = obtener_horarios_turno(turno)
    apilador_col = next((col for col in df.columns if 'apilador' in col.lower()), 'Apilador')

    tareas_procesadas_dia = defaultdict(set)

    tareas_clave_seleccionada = tareas_clave_noche if turno.lower() == "noche" else tareas_clave_otros

    for dia_excel in dias_semana_procesar:
        if dia_excel not in df.columns:
            print(f"Columna no encontrada para el día: {dia_excel}")
            continue

        for index, row in df.iterrows():
            apiladores_primarios_str = normalizar_texto(str(row[apilador_col]).strip())
            apiladores_primarios = [a.strip() for a in apiladores_primarios_str.split("/")]
            tareas_primarias_str = normalizar_texto(str(row[dia_excel]).strip())
            tareas_primarias = [t.strip() for t in tareas_primarias_str.split("/")]

            for i, apilador_primario in enumerate(apiladores_primarios):
                if i < len(tareas_primarias):
                    primera_tarea_raw = tareas_primarias[i]
                    tareas_individuales = [t.strip() for t in primera_tarea_raw.split(" / ")]

                    for tarea_primaria in tareas_individuales:
                        for tarea_clave in tareas_clave_seleccionada:
                            if re.search(r"\b" + re.escape(tarea_clave) + r"\b", tarea_primaria):
                                if (tarea_clave, apilador_primario) not in tareas_procesadas_dia[dia_excel]:
                                    apiladores_emparejados = [apilador_primario]
                                    for idx_sec, row_sec in df.iterrows():
                                        if idx_sec != index:
                                            apiladores_secundarios_str = normalizar_texto(str(row_sec[apilador_col]).strip())
                                            apiladores_secundarios = [a.strip() for a in apiladores_secundarios_str.split("/")]
                                            tareas_secundarias_str = normalizar_texto(str(row_sec[dia_excel]).strip())
                                            tareas_secundarias = [t.strip() for t in tareas_secundarias_str.split("/")] if "/" in tareas_secundarias_str else [tareas_secundarias_str]

                                            for j, apilador_secundario in enumerate(apiladores_secundarios):
                                                if apilador_secundario != apilador_primario and (tarea_clave, apilador_secundario) not in tareas_procesadas_dia[dia_excel]:
                                                    for tarea_secundaria in tareas_secundarias:
                                                        if re.search(r"\b" + re.escape(tarea_clave) + r"\b", tarea_secundaria):
                                                            apiladores_emparejados.append(apilador_secundario)
                                                            tareas_procesadas_dia[dia_excel].add((tarea_clave, apilador_secundario))
                                                            break
                                                    else:
                                                        continue
                                                    break

                                    apiladores_formateados = [formatear_nombre_apilador(apilador) for apilador in apiladores_emparejados]
                                    apilador_str = " \\/ ".join(apiladores_formateados)
                                    camara = camaras_por_tarea.get(tarea_clave, "Desconocido")
                                    resultados_por_dia[dia_excel].append({
                                        "Día": dia_excel,
                                        "Turno": turno.capitalize(),
                                        "Camara": camara,
                                        "Apilador": apilador_str,
                                        "Hora Inicio": hora_inicio,
                                        "Hora break inicio": hora_break_inicio,
                                        "Hora break fin": hora_break_fin,
                                        "Hora fin": hora_fin,
                                        "Pasillo": tarea_clave
                                    })
                                    tareas_procesadas_dia[dia_excel].add((tarea_clave, apilador_primario))

    return resultados_por_dia, tareas_clave_seleccionada

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Uso: python extractor.py <noche|tarde|mañana>")
        sys.exit(1)

    turno_seleccionado = sys.argv[1].lower()
    excel_file = ""
    sheet_name = ""
    nombre_archivo_json = f"resultados_{turno_seleccionado}.json"

    if turno_seleccionado == "noche":
        excel_file = "turnonoche.xlsx"
        sheet_name = "noche"
    elif turno_seleccionado == "tarde":
        excel_file = "turnotarde.xlsx"
        sheet_name = "tarde"
    elif turno_seleccionado == "mañana":
        excel_file = "turnomanana.xlsx"
        sheet_name = "mañana"
    else:
        print("Turno inválido. Debe ser 'noche', 'tarde' o 'mañana'")
        sys.exit(1)

    print(f"--- TURNO {turno_seleccionado.upper()} ---")
    resultados_por_dia, tareas_clave = procesar_excel(excel_file, sheet_name=sheet_name, turno=turno_seleccionado)

    resultados_final_lista = []
    for dia, tareas in resultados_por_dia.items():
        resultados_final_lista.extend(tareas)

    # Definir el orden de los días de la semana
    orden_dias = {"LUNES": 1, "MARTES": 2, "MIÉRCOLES": 3, "JUEVES": 4, "VIERNES": 5, "SÁBADO": 6}

    # Ordenar la lista final primero por día y luego por cámara
    resultados_finales_ordenados = sorted(
        resultados_final_lista,
        key=lambda x: (orden_dias.get(x['Día'], 7), x['Camara'])
    )

    # Corrección para la barra invertida
    resultados_json_str = json.dumps(resultados_finales_ordenados, indent=2, ensure_ascii=False)
    resultados_json_str = resultados_json_str.replace("\\\\/", "\\/")

    with open(nombre_archivo_json, 'w', encoding='utf-8') as archivo_json:
        archivo_json.write(resultados_json_str)

    print(f"Se guardo el archivo {nombre_archivo_json}")

    print("\n--- Resumen de Tareas por Día ---")
    for dia in ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO"]:
        tareas_dia = resultados_por_dia.get(dia, [])
        conteo_tareas = len(tareas_dia)
        print(f"{dia} = {conteo_tareas} TAREAS", end="")

        tareas_encontradas = set(tarea['Pasillo'] for tarea in tareas_dia)
        pendientes = [tarea for tarea in tareas_clave if tarea not in tareas_encontradas]

        if pendientes:
            print(f", {len(pendientes)} PENDIENTE ({', '.join(pendientes)})")
        else:
            print()