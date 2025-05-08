import pandas as pd
import re
import json
import sys
from typing import List, Dict

def normalizar_texto(texto: str) -> str:
    """Normaliza el texto eliminando espacios extra y convirtiendo a mayúsculas."""
    return " ".join(texto.split()).upper()

def obtener_horarios_turno(turno: str) -> tuple:
    """Devuelve los horarios de inicio, break y fin según el turno."""
    if turno.lower() == "noche":
        return "22:00", "02:45", "03:15", "06:05"
    elif turno.lower() == "tarde":
        return "14:10", "18:40", "19:10", "21:55"
    elif turno.lower() == "mañana":
        return "06:40", "10:30", "11:00", "14:10"
    return None, None, None, None

def procesar_excel(excel_file_path: str, sheet_name: str = "noche", turno: str = "Noche") -> List[Dict]:
    """
    Procesa el Excel para emparejar apiladores manteniendo el orden de aparición
    y la secuencia de tareas (antes/después del break).
    """
    try:
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name, header=1)
    except Exception as e:
        print(f"Error al leer el archivo Excel: {e}")
        return []

    df.columns = [col.strip() for col in df.columns]

    tareas_clave = [
        "T", "P", "R", "U", "SECO", "TROPICALES-XDOCK",
        "V1-V2 Y1-Y2", "V3-V4 Y3-Y4", "V5-V9", "Y5-Y9", "W1-W4 Z1-Z5",
        "N", "H", "AA-AG", "AH-AJ", "BA-BG", "BH-BJ", "EKONO"
    ]

    camaras_por_tarea = {
        "T": "Congelado", "P": "Congelado", "R": "Congelado", "U": "Congelado",
        "SECO": "Seco", "TROPICALES-XDOCK": "Tropicales",
        "V1-V2 Y1-Y2": "Panaderia", "V3-V4 Y3-Y4": "Vegetales",
        "V5-V9": "Vegetales", "Y5-Y9": "Vegetales", "W1-W4 Z1-Z5": "Vegetales",
        "N": "Carnes", "H": "Carnes",
        "AA-AG": "Fiambreria", "AH-AJ": "Fiambreria",
        "BA-BG": "Fiambreria", "BH-BJ": "Fiambreria",
        "EKONO": "Ekono"
    }

    resultados_finales = []
    dias_semana_procesar = ["LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO"]
    hora_inicio, hora_break_inicio, hora_break_fin, hora_fin = obtener_horarios_turno(turno)
    apilador_col = next((col for col in df.columns if 'apilador' in col.lower()), 'Apilador')

    tareas_procesadas_dia = {}

    for dia in dias_semana_procesar:
        if dia not in df.columns:
            print(f"Columna no encontrada para el día: {dia}")
            continue

        tareas_procesadas_dia[dia] = set()

        for index, row in df.iterrows():
            apiladores_primarios_str = normalizar_texto(str(row[apilador_col]).strip())
            apiladores_primarios = [a.strip() for a in apiladores_primarios_str.split("/")]
            tareas_primarias_str = normalizar_texto(str(row[dia]).strip())
            tareas_primarias = [t.strip() for t in tareas_primarias_str.split("/")]

            for i, apilador_primario in enumerate(apiladores_primarios):
                if i < len(tareas_primarias):
                    primera_tarea_raw = tareas_primarias[i]
                    tareas_individuales = [t.strip() for t in primera_tarea_raw.split(" / ")]

                    for tarea_primaria in tareas_individuales:
                        for tarea_clave in tareas_clave:
                            if re.search(r"\b" + re.escape(tarea_clave) + r"\b", tarea_primaria):
                                if (dia, tarea_clave, apilador_primario) not in tareas_procesadas_dia[dia]:
                                    apiladores_emparejados = [apilador_primario]
                                    for idx_sec, row_sec in df.iterrows():
                                        if idx_sec != index:
                                            apiladores_secundarios_str = normalizar_texto(str(row_sec[apilador_col]).strip())
                                            apiladores_secundarios = [a.strip() for a in apiladores_secundarios_str.split("/")]
                                            tareas_secundarias_str = normalizar_texto(str(row_sec[dia]).strip())
                                            tareas_secundarias = [t.strip() for t in tareas_secundarias_str.split("/")] if "/" in tareas_secundarias_str else [tareas_secundarias_str]

                                            for j, apilador_secundario in enumerate(apiladores_secundarios):
                                                if apilador_secundario != apilador_primario and (dia, tarea_clave, apilador_secundario) not in tareas_procesadas_dia[dia]:
                                                    for tarea_secundaria in tareas_secundarias:
                                                        if re.search(r"\b" + re.escape(tarea_clave) + r"\b", tarea_secundaria):
                                                            apiladores_emparejados.append(apilador_secundario)
                                                            tareas_procesadas_dia[dia].add((dia, tarea_clave, apilador_secundario))
                                                            break
                                                    else:
                                                        continue
                                                    break

                                    apilador_str = " \\/ ".join(apiladores_emparejados)

                                    camara = camaras_por_tarea.get(tarea_clave, "Desconocido")
                                    resultados_finales.append({
                                        "Día": dia,
                                        "Turno": turno.capitalize(),
                                        "Camara": camara,
                                        "Apilador": apilador_str,
                                        "Hora Inicio": hora_inicio,
                                        "Hora break inicio": hora_break_inicio,
                                        "Hora fin break": hora_break_fin,
                                        "Hora fin": hora_fin,
                                        "Pasillo": tarea_clave
                                    })
                                    tareas_procesadas_dia[dia].add((dia, tarea_clave, apilador_primario))

    return resultados_finales

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
    resultados = procesar_excel(excel_file, sheet_name=sheet_name, turno=turno_seleccionado)

    with open(nombre_archivo_json, 'w', encoding='utf-8') as archivo_json:
        json.dump(resultados, archivo_json, indent=2, ensure_ascii=False)

    print(f"Los resultados han sido guardados en el archivo: {nombre_archivo_json}")