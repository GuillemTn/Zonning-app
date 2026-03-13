"""
Módulo lector_datos.py

Responsabilidad única: Cargar y procesar datos desde Excel.
- Carga Datos.xlsx (hojas: Empleados, Zonas)
- Carga Horario semana.xlsx (hoja: Lunes, con columnas Empleados, Horas de trabajo, Hora de entrada)
- Calcula hora de salida (entrada + horas)
- Cruza empleados del horario con Datos.xlsx
- Expone funciones para obtener empleados filtrados, zonas con prioridades, habilidades
"""

import pandas as pd
import os
from datetime import datetime


def normalizar_nombre(s):
    """Normaliza un string: quita espacios extras, convierte a mayúsculas para comparación."""
    return ''.join(ch for ch in str(s).upper() if ch.isalnum())


def cargar_datos_excel(ruta_datos='Datos.xlsx'):
    """
    Carga las hojas 'Empleados' y 'Zonas' desde Datos.xlsx.
    
    Retorna:
        (empleados_df, zonas_df) o (None, None) si hay error
    """
    if isinstance(ruta_datos, str) and not os.path.exists(ruta_datos):
        print(f"Error: No se encuentra '{ruta_datos}'")
        return None, None
    
    try:
        xls = pd.ExcelFile(ruta_datos)
        hojas = xls.sheet_names
        
        # Cargar Empleados (primera hoja si no existe 'Empleados')
        if 'Empleados' in hojas:
            empleados_df = pd.read_excel(xls, sheet_name='Empleados')
        else:
            empleados_df = pd.read_excel(xls, sheet_name=hojas[0])
        
        # Guardar el índice original (fila del Excel) para usarlo en el coloreado
        empleados_df['original_idx'] = empleados_df.index
        
        # Cargar Zonas
        if 'Zonas' in hojas:
            zonas_df = pd.read_excel(xls, sheet_name='Zonas')
        else:
            zonas_df = None
        
        return empleados_df, zonas_df
    
    except Exception as e:
        print(f"Error al cargar {ruta_datos}: {e}")
        return None, None


def cargar_horario_semanal_completo(ruta_excel='Horario semana.xlsx'):
    """
    Carga TODAS las hojas del archivo Horario semana.xlsx (Lunes a Domingo).
    
    Retorna:
        Dict {día: DataFrame} con datos procesados de cada día
        
    Si una hoja no existe o está vacía, se omite silenciosamente.
    """
    if isinstance(ruta_excel, str) and not os.path.exists(ruta_excel):
        print(f"\n[ERROR CRITICO] No se encuentra '{ruta_excel}'")
        return {}
    
    dias = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
    horarios_por_dia = {}
    
    try:
        xls = pd.ExcelFile(ruta_excel)
        hojas_disponibles = xls.sheet_names
    except Exception as e:
        print(f"[ERROR] No se pudo abrir {ruta_excel}: {e}")
        return {}
    
    for dia in dias:
        if dia not in hojas_disponibles:
            continue
        
        try:
            # Pasamos el objeto xls directamente para evitar problemas de re-lectura de buffers
            df_dia = cargar_horario_csv(xls, dia)
            # Filtrar solo empleados con Horas de trabajo > 0
            empleados_activos = df_dia[df_dia['Horas de trabajo'] > 0]
            if not empleados_activos.empty:
                horarios_por_dia[dia] = df_dia
                print(f"  [OK] {dia}: {len(empleados_activos)} empleados activos")
            else:
                print(f"  [-] {dia}: sin empleados activos (vacio)")
        except Exception as e:
            print(f"  [!] {dia}: Error al procesar ({e})")
    
    if not horarios_por_dia:
        print("\n[ADVERTENCIA] No se encontraron hojas validas con empleados")
    
    return horarios_por_dia


def cargar_horario_csv(ruta_excel='Horario semana.xlsx', nombre_hoja='Lunes'):
    """
    Carga un archivo Excel con horarios (ej: 'Horario semana.xlsx', hoja 'Lunes')
    y extrae columnas relevantes: 'Empleados', 'Horas de trabajo', 'Hora de entrada'.

    Calcula la hora de salida: entrada + horas de trabajo.
    Marca 'OFF' si 'Horas de trabajo' es 0 o está vacío.

    Retorna:
        DataFrame con columnas: ['Empleados', 'Horas de trabajo', 'Hora de entrada', 'Hora_salida', 'Estado']
        
    Lanza:
        FileNotFoundError si el archivo no existe
    """
    if isinstance(ruta_excel, str) and not os.path.exists(ruta_excel):
        archivos = ', '.join(os.listdir('.'))
        raise FileNotFoundError(
            f"\n[ERROR CRÍTICO] No se encuentra '{ruta_excel}'\n"
            f"Ruta esperada: {os.path.abspath(ruta_excel)}\n"
            f"Archivos en directorio actual: {archivos}\n"
            f"Verifica que el nombre del archivo sea exacto."
        )

    try:
        df = pd.read_excel(ruta_excel, sheet_name=nombre_hoja)
    except Exception as e:
        raise ValueError(
            f"\n[ERROR CRÍTICO] Error al leer {ruta_excel}, hoja '{nombre_hoja}':\n{e}"
        )

    # Normalizar nombres de columnas para facilitar búsquedas
    cols = {c.strip(): c for c in df.columns}
    # Buscar columnas esperadas
    col_emp = None
    col_horas = None
    col_entrada = None
    for name in cols:
        low = name.lower()
        if 'emple' in low:
            col_emp = cols[name]
        if 'hora' in low and 'entrada' in low:
            col_entrada = cols[name]
        if 'horas' in low:
            col_horas = cols[name]

    # Fallbacks if no exact match
    if col_emp is None and len(df.columns) > 0:
        col_emp = df.columns[0]
    if col_horas is None and len(df.columns) > 1:
        col_horas = df.columns[1]
    if col_entrada is None and len(df.columns) > 2:
        col_entrada = df.columns[2]

    out_rows = []
    for _, row in df.iterrows():
        nombre = str(row[col_emp]).strip() if pd.notna(row[col_emp]) else ''
        horas_val = row[col_horas] if col_horas in row else None
        entrada_val = row[col_entrada] if col_entrada in row else None

        # Determinar horas como número (int/float)
        horas_num = None
        try:
            if pd.isna(horas_val) or str(horas_val).strip() == '':
                horas_num = 0
            else:
                horas_num = float(horas_val)
        except:
            horas_num = 0

        estado = 'OFF' if horas_num == 0 else 'ON'

        hora_salida_str = ''
        hora_entrada_val = str(entrada_val).strip() if pd.notna(entrada_val) else ''
        if estado == 'ON' and pd.notna(entrada_val) and str(entrada_val).strip() != '':
            # parsear hora de entrada (esperamos múltiples formatos: HH:MM:SS, HH:MM, H)
            try:
                entrada_str = str(entrada_val).strip()
                # Intentar primero HH:MM:SS
                t = pd.to_datetime(entrada_str, format='%H:%M:%S', errors='coerce')
                if pd.isna(t):
                    # Intentar HH:MM
                    t = pd.to_datetime(entrada_str, format='%H:%M', errors='coerce')
                if pd.isna(t):
                    # Intentar solo HH
                    t = pd.to_datetime(entrada_str, format='%H', errors='coerce')
                if not pd.isna(t):
                    hora_entrada_val = t.strftime('%H:%M')
                    salida = t + pd.to_timedelta(horas_num, unit='h')
                    hora_salida_str = salida.strftime('%H:%M')
            except Exception:
                pass

        out_rows.append({
            'Empleados': nombre,
            'Horas de trabajo': horas_num,
            'Hora de entrada': hora_entrada_val,
            'Hora_salida': hora_salida_str,
            'Estado': estado
        })

    result_df = pd.DataFrame(out_rows)
    return result_df


def extraer_empleados_hoy(empleados_df, horario_df):
    """
    Cruza el DataFrame principal `empleados_df` con el `horario_df` (CSV) y
    devuelve solo los empleados presentes en ambos. Además añade columnas
    de 'Horas de trabajo', 'Hora de entrada', 'Hora_salida' y 'Estado'.

    Parámetros:
        empleados_df: DataFrame leído desde Datos.xlsx (hoja Empleados)
        horario_df: DataFrame devuelto por `cargar_horario_csv`

    Retorna:
        DataFrame combinado listo para el motor de asignación.
    """
    if horario_df is None or horario_df.empty:
        print("Sin filtro de horario: usando todos los empleados")
        # Añadir columnas vacías para compatibilidad
        empleados_df['Horas de trabajo'] = 0
        empleados_df['Hora de entrada'] = ''
        empleados_df['Hora_salida'] = ''
        empleados_df['Estado'] = 'OFF'
        # Asegurar columna 'Clase' presente (vacía si no existe)
        if 'Clase' not in empleados_df.columns:
            empleados_df['Clase'] = ''
        return empleados_df

    nombre_col = empleados_df.columns[0]
    # Preparar horario: nombres tal cual
    horario_df['Empleados'] = horario_df['Empleados'].astype(str).str.strip()

    # Merge: left inner join empleados que estén en horario
    merged = pd.merge(empleados_df, horario_df, how='inner', left_on=nombre_col, right_on='Empleados')

    # Asegurar tipos y columnas en orden esperado
    if 'Horas de trabajo' not in merged.columns:
        merged['Horas de trabajo'] = 0
    if 'Hora de entrada' not in merged.columns:
        merged['Hora de entrada'] = ''
    if 'Hora_salida' not in merged.columns:
        merged['Hora_salida'] = ''
    if 'Estado' not in merged.columns:
        merged['Estado'] = 'ON'
    # Asegurar columna 'Clase' presente (intentar copiar desde empleados_df si existe, si no crear vacío)
    if 'Clase' not in merged.columns:
        # Buscar nombre de columna equivalente en empleados_df (insensible a mayúsculas)
        clase_col = None
        for c in empleados_df.columns:
            if str(c).strip().lower() == 'clase':
                clase_col = c
                break
        if clase_col and clase_col in merged.columns:
            merged['Clase'] = merged[clase_col]
        else:
            merged['Clase'] = ''

    print(f"Empleados filtrados para hoy: {len(merged)}")
    return merged


def extraer_zonas_y_prioridades(zonas_df):
    """
    Procesa zonas_df y retorna estructura con prioridades y recomendado.
    
    Retorna:
        Lista de dicts: [{'name': str, 'prioridad': int, 'min': int, 'max': int, 'recomendado': int}, ...]
    """
    if zonas_df is None:
        print("Error: No hay datos de zonas")
        return []
    
    zonas = []

    # Intentar detectar columna que clasifica la zona (tipo: operacional/comercial)
    col_tipo_idx = None
    for i, c in enumerate(zonas_df.columns):
        low = str(c).strip().lower()
        if any(x in low for x in ['tipo', 'clas', 'class', 'categoria', 'cat']):
            col_tipo_idx = i
            break

    def _normalizar_tipo(val):
        if pd.isna(val) or val is None:
            return None
        s = str(val).strip().lower()
        if 'oper' in s:
            return 'operacional'
        if 'com' in s:
            return 'comercial'
    
        return s

    # Detectar columnas por nombre (más robusto que asumir posiciones)
    col_map = {
        'name': None,
        'prioridad': None,
        'min': None,
        'max': None,
        'recomendado': None,
        'tipo': col_tipo_idx
    }
    for i, c in enumerate(zonas_df.columns):
        low = str(c).strip().lower()
        if col_map['name'] is None and any(x in low for x in ['nombre', 'name']):
            col_map['name'] = i
            continue
        if col_map['prioridad'] is None and any(x in low for x in ['prior', 'prio']):
            col_map['prioridad'] = i
            continue
        if col_map['min'] is None and any(x in low for x in ['min', 'mín', 'minimo']):
            col_map['min'] = i
            continue
        if col_map['max'] is None and any(x in low for x in ['max', 'máx', 'maximo']):
            col_map['max'] = i
            continue
        if col_map['recomendado'] is None and any(x in low for x in ['recomend', 'rec', 'recommended']):
            col_map['recomendado'] = i
            continue

    # Fallbacks a posiciones conocidas si no se detectó cabecera
    if col_map['name'] is None:
        col_map['name'] = 0
    if col_map['prioridad'] is None:
        col_map['prioridad'] = 1
    if col_map['min'] is None:
        col_map['min'] = 2
    if col_map['max'] is None:
        col_map['max'] = 3
    if col_map['recomendado'] is None:
        col_map['recomendado'] = 4

    def _to_int_safe(v, default=0):
        try:
            if pd.isna(v):
                return default
            s = str(v).strip()
            if s == '':
                return default
            return int(float(s))
        except Exception:
            return default

    for _, row in zonas_df.iterrows():
        zona_name = str(row[zonas_df.columns[col_map['name']]]).strip()
        prioridad = _to_int_safe(row[zonas_df.columns[col_map['prioridad']]], 0)
        min_p = _to_int_safe(row[zonas_df.columns[col_map['min']]], 0)
        max_p = _to_int_safe(row[zonas_df.columns[col_map['max']]], min_p)
        recomendado = _to_int_safe(row[zonas_df.columns[col_map['recomendado']]], min_p)

        tipo = None
        if col_map['tipo'] is not None:
            tipo = _normalizar_tipo(row[zonas_df.columns[col_map['tipo']]])

        zonas.append({
            'name': zona_name,
            'prioridad': prioridad,
            'min': min_p,
            'max': max_p,
            'recomendado': recomendado,
            'tipo': tipo
        })
    
    print(f"Zonas cargadas: {len(zonas)}")
    return zonas


def mapear_habilidades(empleados_df, zonas):
    """
    Crea un mapeo: zona → columna de habilidad en empleados_df.
    
    Estrategia:
    1. Busca coincidencia exacta (normalizada)
    2. Si es zona COM*, mapea a columna Comercial
    3. Retorna diccionario {zona_name: columna_nombre}
    
    Retorna:
        Dict {zona_name: col_nombre_o_None}
    """
    nombre_col = empleados_df.columns[0]
    skill_cols = list(empleados_df.columns[1:])
    
    def normalizar(s):
        return normalizar_nombre(s)
    
    mapeo = {}
    for zona in zonas:
        zona_name = zona['name']
        zona_norm = normalizar(zona_name)
        
        col_encontrada = None
        
        # Intenta coincidencia exacta normalizada
        for col in skill_cols:
            if normalizar(col) == zona_norm:
                col_encontrada = col
                break
        
        # Si es COM*, busca "Comercial"
        if col_encontrada is None and 'COM' in zona_norm:
            for col in skill_cols:
                if 'COMER' in normalizar(col) or 'COM' in normalizar(col):
                    col_encontrada = col
                    break
        
        mapeo[zona_name] = col_encontrada
    
    return mapeo


def obtener_habilidad(valor):
    """Convierte un valor a int (habilidad), retorna 0 si falla."""
    try:
        return int(valor)
    except:
        try:
            return int(float(valor))
        except:
            return 0


# ============================================================================
# FUNCIÓN PRINCIPAL DEL MÓDULO
# ============================================================================

def cargar_todo(ruta_datos='Datos.xlsx', ruta_horario='Horario semana.xlsx'):
    """
    Carga y prepara todos los datos necesarios para el procesamiento semanal.
    
    Retorna:
        {
            'empleados_por_dia': Dict {día: DataFrame de empleados para ese día},
            'zonas': List[dict] (zonas con prioridades),
            'mapeo_habilidades': Dict {zona_name: col_nombre},
            'empleados_df_original': DataFrame (todos los empleados para referencia),
            'zonas_df': DataFrame (todas las zonas para referencia)
        }
        
    Si no hay días válidos, retorna None.
    """
    print("\n=== INICIALIZANDO SISTEMA DE GESTION SEMANAL ===\n")
    
    # Paso 1: Cargar Excel principal (Datos.xlsx)
    print("Paso 1: Cargando datos maestros...")
    empleados_df, zonas_df = cargar_datos_excel(ruta_datos)
    if empleados_df is None:
        print("[ERROR] No se pudieron cargar los datos mestros.")
        return None
    print(f"  [OK] Empleados disponibles: {len(empleados_df)}")
    
    # Paso 2: Cargar horarios semanal completo (todas las hojas)
    print("\nPaso 2: Cargando horarios semanales...")
    horarios_por_dia = cargar_horario_semanal_completo(ruta_horario)
    if not horarios_por_dia:
        print("[ERROR] No se encontraron hojas validas en el horario.")
        return None
    
    # Paso 3: Cruzar empleados con horarios por día
    print("\nPaso 3: Cruzando empleados con horarios por dia...")
    empleados_por_dia = {}
    for dia, horario_df in horarios_por_dia.items():
        empleados_dia = extraer_empleados_hoy(empleados_df, horario_df)
        if not empleados_dia.empty and (empleados_dia['Horas de trabajo'].astype(float) > 0).sum() > 0:
            empleados_por_dia[dia] = empleados_dia
            print(f"  [OK] {dia}: {len(empleados_dia)} empleados para procesar")
        else:
            print(f"  [-] {dia}: sin empleados validos para procesar")
    
    if not empleados_por_dia:
        print("[ERROR] No hay dias validos para procesar.")
        return None
    
    # Paso 4: Procesar zonas
    print("\nPaso 4: Procesando zonas...")
    zonas = extraer_zonas_y_prioridades(zonas_df)
    
    # Paso 5: Mapear habilidades
    print("Paso 5: Mapeando habilidades...")
    mapeo_hab = mapear_habilidades(empleados_df, zonas)
    
    print("\n=== INICIALIZACION COMPLETADA ===\n")
    
    return {
        'empleados_por_dia': empleados_por_dia,
        'zonas': zonas,
        'mapeo_habilidades': mapeo_hab,
        'empleados_df_original': empleados_df,
        'zonas_df': zonas_df
    }


def cargar_todo_legacy(ruta_datos='Datos.xlsx', ruta_horario='Horario_Semana.xlsx'):
    """
    [LEGADO] Carga datos para un único día (Lunes).
    
    Retorna:
        {
            'empleados_hoy': DataFrame (empleados que trabajan hoy),
            'zonas': List[dict] (zonas con prioridades),
            'mapeo_habilidades': Dict {zona_name: col_nombre},
            'empleados_df_original': DataFrame (todos los empleados para referencia)
        }
    """
    # Paso 1: Cargar Excel principal
    empleados_df, zonas_df = cargar_datos_excel(ruta_datos)
    if empleados_df is None:
        return None
    
    # Paso 2: Cargar horario desde Excel (Horario semana.xlsx, hoja Lunes) y filtrar empleados por cruce
    horario_archivo = 'Horario semana.xlsx'
    horario_hoja = 'Lunes'
    try:
        horario_df = cargar_horario_csv(horario_archivo, horario_hoja)
    except (FileNotFoundError, ValueError) as e:
        print(f"{e}")
        print("\n[DETENCIÓN] El programa requiere el archivo de horario para proceder.")
        return None
    
    empleados_hoy = extraer_empleados_hoy(empleados_df, horario_df)
    
    # Paso 3: Procesar zonas
    zonas = extraer_zonas_y_prioridades(zonas_df)
    
    # Paso 4: Mapear habilidades
    mapeo_hab = mapear_habilidades(empleados_df, zonas)
    
    return {
        'empleados_hoy': empleados_hoy,
        'zonas': zonas,
        'mapeo_habilidades': mapeo_hab,
        'empleados_df_original': empleados_df,
        'zonas_df': zonas_df
    }


if __name__ == '__main__':
    # Prueba rápida del módulo
    print("\n=== VERIFICACIÓN DE LECTURA: Horario semanal.xlsx - SEMANAL ===\n")
    
    datos = cargar_todo()
    if datos:
        empleados_por_dia = datos['empleados_por_dia']
        print(f"\nTotal de días procesados: {len(empleados_por_dia)}")
        
        for dia, empleados_dia in empleados_por_dia.items():
            print(f"\n{dia.upper()}:")
            nombre_col = empleados_dia.columns[0]
            print(f"  Empleados: {len(empleados_dia)}")
            print(f"  Activos: {(empleados_dia['Horas de trabajo'].astype(float) > 0).sum()}")
        
        print(f"\nDatos cargados exitosamente")
        print(f"  - Zonas: {len(datos['zonas'])}")
