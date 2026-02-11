"""
Exportación a Excel
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import pandas as pd
from export.schedule_formatter import crear_horario_consolidado, crear_horario_monitores
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


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
    
    # Asignar colores a monitores
    colores_monitores = {}
    for idx, monitor in enumerate(monitores):
        # Si hay más monitores que colores, se repiten
        color = paleta_colores[idx % len(paleta_colores)]
        colores_monitores[monitor] = color
    
    return colores_monitores


def crear_horarios_por_sala(writer, df_resultado, cfg_esp):
    """
    Crea UNA SOLA hoja con todas las salas en HORIZONTAL (lado a lado)
    Al final, calcula la sumatoria total de todas las salas
    """
    # Obtener días de la semana
    dias_semana = ['Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes', 'Sabado']
    
    # Crear la hoja
    workbook = writer.book
    worksheet = workbook.create_sheet('Horarios por Sala')
    
    # Obtener todas las salas únicas
    salas = sorted(df_resultado[cfg_esp["col_sala"]].unique())
    
    # Estilos
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    thin_border = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC')
    )
    
    # Recopilar todos los monitores únicos de TODAS las salas
    todos_monitores = set()
    for sala in salas:
        df_sala = df_resultado[df_resultado[cfg_esp["col_sala"]] == sala]
        todos_monitores.update(df_sala['MONITOR'].unique())
    
    todos_monitores = sorted(list(todos_monitores))
    
    # Generar colores únicos para cada monitor
    colores_monitores = generar_colores_monitores(todos_monitores)
    
    # Diccionario para almacenar matrices de cada sala
    matrices_salas = {}
    totales_por_sala = {}
    
    for sala in salas:
        df_sala = df_resultado[df_resultado[cfg_esp["col_sala"]] == sala].copy()
        
        if df_sala.empty:
            continue
        
        # Crear matriz: Monitor x Día con suma de horas
        matriz_horas = {}
        for monitor in todos_monitores:
            matriz_horas[monitor] = {dia: 0 for dia in dias_semana}
        
        # Calcular horas por monitor y día
        for _, row in df_sala.iterrows():
            monitor = row['MONITOR']
            dia = row['DIA_NORM']
            
            # Convertir dia normalizado a formato display
            dia_display = dia.capitalize() if dia else None
            if dia_display == 'Miercoles':
                dia_display = 'Miercoles'
            
            if dia_display not in dias_semana:
                continue
            
            # Calcular horas del bloque
            duracion = row.get('DURACION', 0)
            if duracion > 0:
                matriz_horas[monitor][dia_display] += duracion
        
        matrices_salas[sala] = matriz_horas
        
        # Calcular total por sala
        total_sala = sum([
            sum(matriz_horas[m].values()) 
            for m in todos_monitores
        ])
        totales_por_sala[sala] = total_sala
    
    # === ESCRIBIR EN EXCEL ===
    
    columna_actual = 1
    num_columnas_por_sala = len(dias_semana) + 1  # días + columna de monitor
    
    # FILA 1: Títulos de salas
    for sala in salas:
        worksheet.merge_cells(
            start_row=1,
            start_column=columna_actual,
            end_row=1,
            end_column=columna_actual + num_columnas_por_sala - 1
        )
        cell_titulo = worksheet.cell(1, columna_actual)
        cell_titulo.value = sala
        cell_titulo.font = Font(size=14, bold=True, color="FFFFFF")
        cell_titulo.fill = header_fill
        cell_titulo.alignment = Alignment(horizontal='center', vertical='center')
        
        columna_actual += num_columnas_por_sala
    
    worksheet.row_dimensions[1].height = 25
    
    # FILA 2: Encabezados (Monitor + Días) para cada sala
    columna_actual = 1
    for sala in salas:
        # Columna "Monitor"
        cell = worksheet.cell(2, columna_actual)
        cell.value = "Monitor"
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = header_border
        
        # Columnas de días
        for idx_dia, dia in enumerate(dias_semana, start=1):
            cell = worksheet.cell(2, columna_actual + idx_dia)
            cell.value = dia
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = header_border
        
        columna_actual += num_columnas_por_sala
    
    worksheet.row_dimensions[2].height = 20
    
    # Agregar encabezado "Subtotal por monitor" al final
    columna_subtotal = (len(salas) * num_columnas_por_sala) + 1
    cell_header_subtotal = worksheet.cell(2, columna_subtotal)
    cell_header_subtotal.value = "Subtotal por monitor"
    cell_header_subtotal.fill = header_fill
    cell_header_subtotal.font = header_font
    cell_header_subtotal.alignment = Alignment(horizontal='center', vertical='center')
    cell_header_subtotal.border = header_border
    
    # FILAS 3+: Datos de monitores
    fila_actual = 3
    for monitor in todos_monitores:
        columna_actual = 1
        
        # Obtener el color único para este monitor
        color_monitor = colores_monitores[monitor]
        
        # Variable para acumular el total de horas del monitor
        total_horas_monitor = 0
        
        for sala in salas:
            matriz = matrices_salas.get(sala, {})
            
            # Columna de monitor - con su color distintivo
            cell_monitor = worksheet.cell(fila_actual, columna_actual)
            cell_monitor.value = monitor
            cell_monitor.alignment = Alignment(horizontal='left', vertical='center')
            cell_monitor.font = Font(size=10, bold=True)
            cell_monitor.fill = PatternFill(start_color=color_monitor, end_color=color_monitor, fill_type="solid")
            cell_monitor.border = thin_border
            
            # Columnas de días
            for idx_dia, dia in enumerate(dias_semana, start=1):
                cell = worksheet.cell(fila_actual, columna_actual + idx_dia)
                
                horas = matriz.get(monitor, {}).get(dia, 0)
                cell.value = int(horas) if horas > 0 else ''
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.font = Font(size=10)
                
                # Acumular horas para el total del monitor
                total_horas_monitor += horas
                
                # Usar el color del monitor si tiene horas asignadas
                if horas > 0:
                    cell.fill = PatternFill(start_color=color_monitor, end_color=color_monitor, fill_type="solid")
                else:
                    cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                
                cell.border = thin_border
            
            columna_actual += num_columnas_por_sala
        
        # Agregar columna "Subtotal por monitor" al final de la fila
        cell_subtotal = worksheet.cell(fila_actual, columna_actual)
        cell_subtotal.value = int(total_horas_monitor)
        cell_subtotal.fill = PatternFill(start_color=color_monitor, end_color=color_monitor, fill_type="solid")
        cell_subtotal.font = Font(bold=True, size=11)
        cell_subtotal.alignment = Alignment(horizontal='center', vertical='center')
        cell_subtotal.border = header_border
        
        fila_actual += 1
    
    # FILA: Subtotal por día (para cada sala)
    columna_actual = 1
    for sala in salas:
        matriz = matrices_salas.get(sala, {})
        
        # Etiqueta "Subtotal por día"
        cell_label = worksheet.cell(fila_actual, columna_actual)
        cell_label.value = "Subtotal por día"
        cell_label.fill = header_fill
        cell_label.font = header_font
        cell_label.alignment = Alignment(horizontal='center', vertical='center')
        cell_label.border = header_border
        
        # Subtotales por día
        for idx_dia, dia in enumerate(dias_semana, start=1):
            subtotal_dia = sum([
                matriz.get(m, {}).get(dia, 0)
                for m in todos_monitores
            ])
            
            cell = worksheet.cell(fila_actual, columna_actual + idx_dia)
            cell.value = int(subtotal_dia) if subtotal_dia > 0 else ''
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = header_border
        
        columna_actual += num_columnas_por_sala
    
    fila_actual += 1
    
    # FILA: Total por semana (para cada sala)
    columna_actual = 1
    for sala in salas:
        # Merge de "Total por semana"
        worksheet.merge_cells(
            start_row=fila_actual,
            start_column=columna_actual,
            end_row=fila_actual,
            end_column=columna_actual + len(dias_semana) - 1
        )
        cell_label = worksheet.cell(fila_actual, columna_actual)
        cell_label.value = "Total por semana"
        cell_label.fill = header_fill
        cell_label.font = header_font
        cell_label.alignment = Alignment(horizontal='center', vertical='center')
        cell_label.border = header_border
        
        # Total de la sala
        cell_total = worksheet.cell(fila_actual, columna_actual + len(dias_semana))
        cell_total.value = int(totales_por_sala.get(sala, 0))
        cell_total.fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
        cell_total.font = Font(bold=True, size=12)
        cell_total.alignment = Alignment(horizontal='center', vertical='center')
        cell_total.border = header_border
        
        columna_actual += num_columnas_por_sala
    
    fila_actual += 2  # Espacio
    
    # === FILA FINAL: SUMATORIA TOTAL DE TODAS LAS SALAS ===
    total_general_todas_salas = sum(totales_por_sala.values())
    
    # Merge de toda la fila menos la última columna
    worksheet.merge_cells(
        start_row=fila_actual,
        start_column=1,
        end_row=fila_actual,
        end_column=(len(salas) * num_columnas_por_sala) - 1
    )
    cell_label_total = worksheet.cell(fila_actual, 1)
    cell_label_total.value = "TOTAL GENERAL (TODAS LAS SALAS)"
    cell_label_total.fill = PatternFill(start_color="203864", end_color="203864", fill_type="solid")
    cell_label_total.font = Font(bold=True, color="FFFFFF", size=13)
    cell_label_total.alignment = Alignment(horizontal='center', vertical='center')
    cell_label_total.border = header_border
    
    # Total general
    cell_total_general = worksheet.cell(fila_actual, len(salas) * num_columnas_por_sala)
    cell_total_general.value = int(total_general_todas_salas)
    cell_total_general.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
    cell_total_general.font = Font(bold=True, size=14, color="FFFFFF")
    cell_total_general.alignment = Alignment(horizontal='center', vertical='center')
    cell_total_general.border = header_border
    
    worksheet.row_dimensions[fila_actual].height = 30
    
    # === AJUSTAR ANCHOS DE COLUMNA ===
    for col_num in range(1, (len(salas) * num_columnas_por_sala) + 2):  # +2 para incluir la columna de subtotal
        col_letter = get_column_letter(col_num)
        
        # Cada primera columna de cada sala (columna de monitores)
        if (col_num - 1) % num_columnas_por_sala == 0:
            worksheet.column_dimensions[col_letter].width = 35
        # Columna de subtotal por monitor (última columna)
        elif col_num == (len(salas) * num_columnas_por_sala) + 1:
            worksheet.column_dimensions[col_letter].width = 20
        else:
            worksheet.column_dimensions[col_letter].width = 12


def exportar_resultados(ruta, df_resultado, monitores_asignados, cfg_esp):
    """
    Exporta resultados completos a Excel con horarios visuales
    
    Args:
        ruta: Path de destino
        df_resultado: DataFrame con asignaciones
        monitores_asignados: Lista de monitores con asignaciones
        cfg_esp: Configuración de espacios
    """
    # Preparar resumen de monitores
    df_mon = pd.DataFrame([{
        'Monitor': m['nombre'], 
        'Prioridad': m['prioridad'], 
        'Horas': m['horas'],
        'Min': m['min'], 
        'Max': m['max'], 
        'Horarios': len(m['asignaciones']),
        'Estado': '✅' if m['min'] <= m['horas'] <= m['max'] else '⚠️'
    } for m in monitores_asignados]).sort_values('Horas', ascending=False)
    
    salas = sorted(df_resultado[cfg_esp["col_sala"]].unique())
    
    # Crear archivo Excel con múltiples hojas
    with pd.ExcelWriter(ruta, engine='openpyxl') as writer:
        # HOJA 1: Horarios visuales consolidados
        crear_horario_consolidado(writer, df_resultado, salas, cfg_esp)
        
        # HOJA 2: Horarios individuales por monitor
        crear_horario_monitores(writer, monitores_asignados, df_resultado, cfg_esp)
        
        # HOJA 3: Lista de asignaciones
        df_resultado.to_excel(writer, sheet_name='Lista Asignaciones', index=False)
        
        # HOJA 4: Resumen de monitores
        df_mon.to_excel(writer, sheet_name='Resumen Monitores', index=False)
        
        # *** NUEVAS HOJAS: Horarios por sala ***
        crear_horarios_por_sala(writer, df_resultado, cfg_esp)