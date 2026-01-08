import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from collections import defaultdict
import re

def procesar_archivo_privilegios(ruta_archivo_txt, ruta_archivo_excel):
    """
    Convierte archivo TXT de privilegios en Excel organizado por grupos
    """
    
    # Estructura para almacenar datos por grupo
    grupos = defaultdict(lambda: {
        'codigo': '',
        'nombre': '',
        'usuarios': [],
        'opciones': []
    })
    
    print("Leyendo archivo...")
    
    with open(ruta_archivo_txt, 'r', encoding='utf-8', errors='ignore') as f:
        for linea in f:
            linea = linea.strip()
            if not linea:
                continue
                
            partes = linea.split('|')
            
            # Verificar que sea una línea de GRUPO
            if partes[0] == 'GRUPO' and len(partes) > 3:
                codigo_grupo = partes[1]
                nombre_grupo = partes[2]
                
                # Inicializar grupo si no existe
                if not grupos[codigo_grupo]['codigo']:
                    grupos[codigo_grupo]['codigo'] = codigo_grupo
                    grupos[codigo_grupo]['nombre'] = nombre_grupo
                
                # Buscar índice de "OPCIONES POR GRUPO"
                idx_opciones = None
                for i, parte in enumerate(partes):
                    if parte == 'OPCIONES POR GRUPO':
                        idx_opciones = i
                        break
                
                if idx_opciones is not None:
                    # CASO 1: Línea de usuario con "OPCIONES POR GRUPO" al final
                    # Formato: GRUPO|006|ANALISTA MERCADEO|Cod Usuario|Nombre del Usuario|Activo|SVALEN|SEBASTIAN VALENCIA|Y|OPCIONES POR GRUPO|
                    if idx_opciones >= 6:  # Hay datos de usuario antes de "OPCIONES POR GRUPO"
                        cod_usuario = partes[6] if len(partes) > 6 else ''
                        nombre_usuario = partes[7] if len(partes) > 7 else ''
                        activo = partes[8] if len(partes) > 8 else ''
                        
                        # Validar que no sea una línea de encabezado ni de opciones
                        if (cod_usuario and 
                            cod_usuario.strip() and 
                            cod_usuario != 'Cod Usuario' and
                            cod_usuario != 'Formas'):
                            
                            grupos[codigo_grupo]['usuarios'].append({
                                'codigo': cod_usuario,
                                'nombre': nombre_usuario,
                                'activo': activo
                            })
                            print(f"  Usuario agregado al grupo {codigo_grupo}: {cod_usuario} - {nombre_usuario}")
                    
                    # CASO 2: Línea de opciones
                    # Formato: GRUPO|006|...|OPCIONES POR GRUPO|Formas|Insertar|Modificar|...
                    # Verificar si después de "OPCIONES POR GRUPO" están los encabezados de opciones
                    if len(partes) > idx_opciones + 1:
                        siguiente = partes[idx_opciones + 1]
                        
                        # Si después de "OPCIONES POR GRUPO" viene "Formas", es una línea de opciones
                        if siguiente == 'Formas' and len(partes) > idx_opciones + 7:
                            # Extraer los datos de la opción (nombre de la forma y permisos)
                            datos_opcion = partes[idx_opciones + 8:]  # Después de los encabezados
                            
                            # Filtrar elementos vacíos
                            datos_opcion = [d for d in datos_opcion if d.strip()]
                            
                            if len(datos_opcion) >= 2:  # Al menos nombre de forma y un permiso
                                grupos[codigo_grupo]['opciones'].append(datos_opcion)
                                print(f"  Opción agregada al grupo {codigo_grupo}: {datos_opcion[0]}")
                
                else:
                    # Sin "OPCIONES POR GRUPO" - línea de usuario normal
                    if len(partes) > 5:
                        cod_usuario = partes[3]
                        nombre_usuario = partes[4]
                        activo = partes[5]
                        
                        # Evitar líneas de encabezado
                        if cod_usuario and cod_usuario.strip() and cod_usuario != 'Cod Usuario':
                            grupos[codigo_grupo]['usuarios'].append({
                                'codigo': cod_usuario,
                                'nombre': nombre_usuario,
                                'activo': activo
                            })
                            print(f"  Usuario agregado al grupo {codigo_grupo}: {cod_usuario} - {nombre_usuario}")
    
    print(f"\nSe encontraron {len(grupos)} grupos")
    for codigo, datos in grupos.items():
        print(f"  Grupo {codigo}: {len(datos['usuarios'])} usuarios, {len(datos['opciones'])} opciones")
    
    # Crear archivo Excel
    print("\nCreando archivo Excel...")
    wb = Workbook()
    wb.remove(wb.active)  # Remover hoja por defecto
    
    # Estilos
    titulo_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    titulo_font = Font(bold=True, color="FFFFFF", size=12)
    seccion_fill = PatternFill(start_color="B4C7E7", end_color="B4C7E7", fill_type="solid")
    seccion_font = Font(bold=True, size=10)
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    header_font = Font(bold=True, size=10)
    
    for codigo_grupo, datos in sorted(grupos.items()):
        # Crear nombre de hoja (máximo 31 caracteres en Excel)
        nombre_hoja = f"{codigo_grupo} {datos['nombre']}"[:31]
        
        # Limpiar caracteres no permitidos en nombres de hojas
        nombre_hoja = re.sub(r'[\\/*?:\[\]]', '', nombre_hoja)
        
        ws = wb.create_sheet(title=nombre_hoja)
        
        fila_actual = 1
        
        # FILA 1: Título del grupo con formato
        cell = ws.cell(row=fila_actual, column=1, value=f"{datos['codigo']} {datos['nombre']}")
        cell.font = titulo_font
        cell.fill = titulo_fill
        ws.merge_cells(start_row=fila_actual, start_column=1, end_row=fila_actual, end_column=7)
        fila_actual += 1
        
        # SECCIÓN DE USUARIOS (si existen)
        if datos['usuarios']:
            fila_actual += 1  # Línea en blanco
            
            # Encabezados de usuarios
            ws.cell(row=fila_actual, column=1, value="#").font = header_font
            ws.cell(row=fila_actual, column=1).fill = header_fill
            ws.cell(row=fila_actual, column=2, value="Cod Usuario").font = header_font
            ws.cell(row=fila_actual, column=2).fill = header_fill
            ws.cell(row=fila_actual, column=3, value="Nombre del Usuario").font = header_font
            ws.cell(row=fila_actual, column=3).fill = header_fill
            ws.cell(row=fila_actual, column=4, value="Activo").font = header_font
            ws.cell(row=fila_actual, column=4).fill = header_fill
            fila_actual += 1
            
            # Datos de usuarios con numeración
            for idx, usuario in enumerate(datos['usuarios'], start=1):
                ws.cell(row=fila_actual, column=1, value=idx)
                ws.cell(row=fila_actual, column=2, value=usuario['codigo'])
                ws.cell(row=fila_actual, column=3, value=usuario['nombre'])
                ws.cell(row=fila_actual, column=4, value=usuario['activo'])
                fila_actual += 1
            
            fila_actual += 1  # Línea en blanco
        
        # SECCIÓN DE OPCIONES POR GRUPO
        # Título de la sección
        cell = ws.cell(row=fila_actual, column=1, value="OPCIONES POR GRUPO")
        cell.font = seccion_font
        cell.fill = seccion_fill
        ws.merge_cells(start_row=fila_actual, start_column=1, end_row=fila_actual, end_column=7)
        fila_actual += 1
        
        # Encabezados de opciones
        encabezados = ['Formas', 'Insertar', 'Modificar', 'Eliminar', 'Consultar', 'Ejecuta Procesos', 'Otros Procesos']
        
        for col, encabezado in enumerate(encabezados, start=1):
            cell = ws.cell(row=fila_actual, column=col, value=encabezado)
            cell.font = header_font
            cell.fill = header_fill
        fila_actual += 1
        
        # Datos de opciones
        if datos['opciones']:
            for opcion_data in datos['opciones']:
                for col, valor in enumerate(opcion_data, start=1):
                    if col <= 7 and valor and valor.strip():  # Solo 7 columnas
                        ws.cell(row=fila_actual, column=col, value=valor.strip())
                fila_actual += 1
        
        # Ajustar anchos de columna
        ws.column_dimensions['A'].width = 35
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 18
        ws.column_dimensions['D'].width = 18
        ws.column_dimensions['E'].width = 18
        ws.column_dimensions['F'].width = 18
        ws.column_dimensions['G'].width = 18
    
    # Guardar archivo
    wb.save(ruta_archivo_excel)
    print(f"\n✓ Archivo Excel creado exitosamente: {ruta_archivo_excel}")
    print(f"✓ Total de hojas creadas: {len(grupos)}")

# Ejemplo de uso
if __name__ == "__main__":
    # Cambia estas rutas por las tuyas
    archivo_entrada = "privilegios.txt"  # Tu archivo TXT
    archivo_salida = "privilegios_organizados.xlsx"  # Excel de salida
    
    try:
        procesar_archivo_privilegios(archivo_entrada, archivo_salida)
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo '{archivo_entrada}'")
        print("Asegúrate de que el archivo esté en la misma carpeta que este script")
    except Exception as e:
        print(f"Error al procesar el archivo: {str(e)}")
        import traceback
        traceback.print_exc()