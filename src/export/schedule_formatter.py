"""
Formateadores de horarios visuales
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from export.styles import *


def generar_colores_monitores(monitores):
    """
    Genera un diccionario de colores únicos para cada monitor
    Usa solo 3 colores distintivos: azul, verde y naranja
    """
    # Paleta de 3 colores bonitos y distintivos en formato hex (sin el #)
    paleta_colores = [
        "5DADE2",  # Azul brillante
        "58D68D",  # Verde primavera
        "F5B041",  # Naranja mango
    ]
    
    # Asignar colores a monitores únicos
    monitores_unicos = sorted(list(set(monitores)))
    colores_monitores = {}
    for idx, monitor in enumerate(monitores_unicos):
        color = paleta_colores[idx % len(paleta_colores)]
        colores_monitores[monitor] = color
    
    return colores_monitores


def crear_horario_consolidado(writer, df_asig, salas, cfg_esp):
    """Crea horario visual con todas las salas HORIZONTALMENTE"""
    workbook = writer.book
    worksheet = workbook.create_sheet('Horarios Salas', 0)
    
    dias = ['Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes', 'Sabado']
    
    hora_min = int(df_asig[cfg_esp["col_hora_inicio"]].min())
    hora_max = int(df_asig[cfg_esp["col_hora_fin"]].max())
    
    # Generar colores únicos para cada monitor
    todos_monitores = df_asig['MONITOR'].unique().tolist()
    colores_monitores = generar_colores_monitores(todos_monitores)
    
    columna_actual = 1
    
    # Fila 1: Títulos de salas
    for sala in salas:
        worksheet.merge_cells(
            start_row=1, 
            start_column=columna_actual, 
            end_row=1, 
            end_column=columna_actual + 6
        )
        cell_titulo = worksheet.cell(row=1, column=columna_actual)
        cell_titulo.value = sala
        cell_titulo.fill = COLOR_TITULO_SALA
        cell_titulo.font = FUENTE_TITULO
        cell_titulo.alignment = ALINEACION_CENTRO
        cell_titulo.border = BORDE
        
        # Fila 2: Encabezados de días
        cell_hora_header = worksheet.cell(row=2, column=columna_actual)
        cell_hora_header.value = "Hora"
        cell_hora_header.fill = COLOR_HEADER
        cell_hora_header.font = FUENTE_HEADER
        cell_hora_header.alignment = ALINEACION_CENTRO
        cell_hora_header.border = BORDE
        
        for idx_dia, dia in enumerate(dias, start=1):
            cell = worksheet.cell(row=2, column=columna_actual + idx_dia)
            cell.value = dia
            cell.fill = COLOR_HEADER
            cell.font = FUENTE_HEADER
            cell.alignment = ALINEACION_CENTRO
            cell.border = BORDE
        
        columna_actual += 7
    
    # Filas 3+: Horas y datos
    for fila_hora, hora in enumerate(range(hora_min, hora_max), start=3):
        hora_str = f"{hora}:00-{hora+1}:00"
        
        columna_actual = 1
        
        for sala in salas:
            df_sala = df_asig[df_asig[cfg_esp["col_sala"]] == sala]
            
            cell_hora = worksheet.cell(row=fila_hora, column=columna_actual)
            cell_hora.value = hora_str
            cell_hora.fill = COLOR_HEADER
            cell_hora.font = FUENTE_HEADER
            cell_hora.alignment = ALINEACION_CENTRO
            cell_hora.border = BORDE
            
            for idx_dia, dia in enumerate(dias, start=1):
                cell = worksheet.cell(row=fila_hora, column=columna_actual + idx_dia)
                cell.border = BORDE
                cell.alignment = ALINEACION_CENTRO
                
                dia_norm = dia.lower()
                asignaciones_celda = df_sala[
                    (df_sala['DIA_NORM'] == dia_norm) &
                    (df_sala[cfg_esp["col_hora_inicio"]] <= hora) &
                    (df_sala[cfg_esp["col_hora_fin"]] > hora)
                ]
                
                if len(asignaciones_celda) > 0:
                    asig = asignaciones_celda.iloc[0]
                    curso = asig[cfg_esp["col_curso"]]
                    monitor = asig['MONITOR']
                    
                    if monitor == "SIN MONITOR":
                        cell.value = f"{curso}\n❌ SIN MONITOR"
                        cell.fill = COLOR_SIN_MONITOR
                        cell.font = FUENTE_ERROR
                    else:
                        cell.value = f"{curso}\n{monitor}"
                        cell.font = FUENTE_NEGRA
                        
                        # Usar el color único del monitor
                        color_monitor = colores_monitores.get(monitor, "FFFFFF")
                        cell.fill = PatternFill(start_color=color_monitor, 
                                              end_color=color_monitor, 
                                              fill_type="solid")
                else:
                    cell.value = ""
                    cell.fill = COLOR_VACIO
            
            columna_actual += 7
    
    # Ajustar anchos
    for col_num in range(1, columna_actual):
        col_letter = get_column_letter(col_num)
        if (col_num - 1) % 7 == 0:
            worksheet.column_dimensions[col_letter].width = 11
        else:
            worksheet.column_dimensions[col_letter].width = 25
    
    worksheet.row_dimensions[1].height = 25
    worksheet.row_dimensions[2].height = 20
    for row in range(3, fila_hora + 1):
        worksheet.row_dimensions[row].height = 40


def crear_horario_monitores(writer, monitores, df_asig, cfg_esp):
    """Crea horario detallado de cada monitor"""
    workbook = writer.book
    worksheet = workbook.create_sheet('Horarios Monitores', 1)
    
    dias = ['Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes', 'Sabado']
    
    hora_min = int(df_asig[cfg_esp["col_hora_inicio"]].min())
    hora_max = int(df_asig[cfg_esp["col_hora_fin"]].max())
    
    # Generar colores únicos para cada monitor
    todos_monitores = [m['nombre'] for m in monitores]
    colores_monitores = generar_colores_monitores(todos_monitores)
    
    fila_actual = 1
    
    monitores_activos = [m for m in monitores if m["horas"] > 0]
    monitores_activos.sort(key=lambda x: x["nombre"])
    
    for monitor in monitores_activos:
        nombre_monitor = monitor['nombre']
        color_monitor = colores_monitores.get(nombre_monitor, "FFFFFF")
        
        # Título del monitor con su color único
        worksheet.merge_cells(start_row=fila_actual, start_column=1, 
                             end_row=fila_actual, end_column=7)
        cell_titulo = worksheet.cell(row=fila_actual, column=1)
        cell_titulo.value = f"{nombre_monitor} - {monitor['horas']} horas"
        cell_titulo.fill = PatternFill(start_color=color_monitor, 
                                       end_color=color_monitor, 
                                       fill_type="solid")
        cell_titulo.font = FUENTE_MONITOR
        cell_titulo.alignment = ALINEACION_CENTRO
        cell_titulo.border = BORDE
        fila_actual += 1
        
        # Encabezados
        worksheet.cell(row=fila_actual, column=1).value = "Hora"
        for idx, dia in enumerate(dias, start=2):
            cell = worksheet.cell(row=fila_actual, column=idx)
            cell.value = dia
            cell.fill = COLOR_HEADER
            cell.font = FUENTE_HEADER
            cell.alignment = ALINEACION_CENTRO
            cell.border = BORDE
        
        cell_hora_header = worksheet.cell(row=fila_actual, column=1)
        cell_hora_header.fill = COLOR_HEADER
        cell_hora_header.font = FUENTE_HEADER
        cell_hora_header.alignment = ALINEACION_CENTRO
        cell_hora_header.border = BORDE
        fila_actual += 1
        
        asignaciones_monitor = df_asig[df_asig['MONITOR'] == nombre_monitor]
        
        for hora in range(hora_min, hora_max):
            hora_str = f"{hora}-{hora+1}"
            
            cell_hora = worksheet.cell(row=fila_actual, column=1)
            cell_hora.value = hora_str
            cell_hora.fill = COLOR_HEADER
            cell_hora.font = FUENTE_HEADER
            cell_hora.alignment = ALINEACION_CENTRO
            cell_hora.border = BORDE
            
            for idx_dia, dia in enumerate(dias, start=2):
                cell = worksheet.cell(row=fila_actual, column=idx_dia)
                cell.border = BORDE
                cell.alignment = ALINEACION_CENTRO
                
                dia_norm = dia.lower()
                asig_celda = asignaciones_monitor[
                    (asignaciones_monitor['DIA_NORM'] == dia_norm) &
                    (asignaciones_monitor[cfg_esp["col_hora_inicio"]] <= hora) &
                    (asignaciones_monitor[cfg_esp["col_hora_fin"]] > hora)
                ]
                
                if len(asig_celda) > 0:
                    asig = asig_celda.iloc[0]
                    curso = asig[cfg_esp["col_curso"]]
                    sala = asig[cfg_esp["col_sala"]]
                    
                    cell.value = f"{sala}\n{curso}"
                    # Usar el color del monitor también en sus clases
                    cell.fill = PatternFill(start_color=color_monitor, 
                                          end_color=color_monitor, 
                                          fill_type="solid")
                    cell.font = FUENTE_NORMAL
                else:
                    cell.value = ""
                    cell.fill = COLOR_VACIO
            
            fila_actual += 1
        
        fila_actual += 2
    
    worksheet.column_dimensions['A'].width = 10
    for col in range(2, 8):
        worksheet.column_dimensions[get_column_letter(col)].width = 25
    
    for row in range(1, fila_actual):
        worksheet.row_dimensions[row].height = 40