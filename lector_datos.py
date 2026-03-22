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
import unicodedata
from datetime import datetime


def normalizar_nombre(s):
    """Normaliza un string: quita espacios extras, convierte a mayúsculas para comparación."""
    return ''.join(ch for ch in str(s).upper() if ch.isalnum())


def cargar_datos_excel(ruta_datos):
    """
    Carga las hojas 'Empleados' y 'Zonas' desde Datos.xlsx.
    
    Retorna:
        (empleados_df, zonas_df) o (None, None) si hay error
    """
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
        raise e


def cargar_horario_semanal_completo(ruta_excel):
    """
    Carga TODAS las hojas del archivo Horario semana.xlsx (Lunes a Domingo).
    
    Retorna:
        Dict {día: DataFrame} con datos procesados de cada día
        
    Si una hoja no existe o está vacía, se omite silenciosamente.
    """
    dias = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
    horarios_por_dia = {}
    
    try:
        xls = pd.ExcelFile(ruta_excel)
        hojas_disponibles = xls.sheet_names
    except Exception as e:
        raise e
    
    # Normalizamos (quitamos tildes y pasamos a minúsculas) para coincidencias flexibles
    def _normalizar_hoja(s):
        return unicodedata.normalize('NFKD', str(s)).encode('ASCII', 'ignore').decode('utf-8').lower().strip()

    mapa_hojas = {_normalizar_hoja(h): h for h in hojas_disponibles}

    for dia in dias:
        dia_norm = _normalizar_hoja(dia)
        if dia_norm not in mapa_hojas:
            continue
        
        hoja_real = mapa_hojas[dia_norm]
        try:
            # Pasamos el objeto xls directamente para evitar problemas de re-lectura de buffers
            df_dia = cargar_horario_csv(xls, hoja_real)
            # Filtrar solo empleados con Horas de trabajo > 0
            empleados_activos = df_dia[df_dia['Horas de trabajo'] > 0]
            if not empleados_activos.empty:
                horarios_por_dia[dia] = df_dia
        except Exception as e:
            raise e
    
    return horarios_por_dia

def _parsear_hora_franja(franja_col):
    """
    Extrae la hora de inicio (HH:MM) a partir del nombre de la columna.
    Ejemplos soportados: '07:00-08:00' -> '07:00', '7-8' -> '07:00', '15' -> '15:00'
    """
    import re
    s = str(franja_col).strip()
    match = re.search(r'^(\d{1,2})([:.]\d{2})?', s)
    if match:
        h = int(match.group(1))
        m = match.group(2).replace('.', ':') if match.group(2) else ':00'
        return f"{h:02d}{m}"
    return ""

def cargar_horario_csv(ruta_excel, nombre_hoja='Lunes'):
    """
    Lee el horario en formato de cuadrante/matriz visual diseñado por mánagers.
    - Identifica la columna de Empleados.
    - Itera sobre las franjas horarias y cuenta las marcas ('X') para sacar totales.
    - Calcula Hora de Entrada y de Salida.

    Retorna:
        DataFrame con columnas: ['Empleados', 'Horas de trabajo', 'Hora de entrada', 'Hora_salida', 'Estado']
    """

    try:
        df = pd.read_excel(ruta_excel, sheet_name=nombre_hoja)
    except Exception as e:
        raise e

    # Limpieza previa: descartar filas o columnas completamente vacías
    df = df.dropna(how='all').dropna(axis=1, how='all')

    # 1. Identificar columna de empleados
    col_emp = None
    for col in df.columns:
        if isinstance(col, str) and any(kw in col.lower() for kw in ['empleado', 'nombre', 'name', 'colaborador']):
            col_emp = col
            break
    
    # Fallback: asumir que es la primera columna disponible si no encuentra keywords
    if col_emp is None and len(df.columns) > 0:
        col_emp = df.columns[0]

    # 2. Identificar columnas de franjas horarias (ignorando la del empleado)
    import re
    cols_franjas = [c for c in df.columns if c != col_emp]
    # Comprobar cuáles de ellas empiezan con número (son horas válidas)
    cols_horas = [c for c in cols_franjas if re.search(r'\d{1,2}', str(c))]
    if not cols_horas:
        cols_horas = cols_franjas # Fallback total

    out_rows = []
    for _, row in df.iterrows():
        nombre = str(row[col_emp]).strip() if pd.notna(row[col_emp]) else ''
        
        # Limpieza: ignorar filas vacías o identificadas como sumas de totales
        if not nombre or nombre.lower() in ['nan', 'none', 'total', 'totales']:
            continue

        # 3. Iterar sobre las celdas y contar presencias ("X")
        franjas_activas = []
        for col in cols_horas:
            val = row[col]
            if pd.isna(val):
                continue
            
            val_str = str(val).strip().upper()
            # Identificar como activo si hay marca (soporta "X", "1", "SI", ignora nulos o falsos)
            if val_str and val_str not in ['0', 'FALSE', 'NO', 'NAN', 'NONE', '']:
                franjas_activas.append(col)

        # Calcular el total de horas (Asumimos 1 franja = 1 hora por indicación. Multiplicar por factor si usaran 30 mins)
        horas_trabajo = float(len(franjas_activas))
        estado = 'ON' if horas_trabajo > 0 else 'OFF'

        hora_entrada = ''
        hora_salida = ''

        if horas_trabajo > 0:
            hora_entrada = _parsear_hora_franja(franjas_activas[0])
            
            if hora_entrada:
                try:
                    h_ent, m_ent = map(int, hora_entrada.split(':'))
                    
                    # Deducir hora_salida sumando las horas de trabajo a la hora de entrada
                    h_sal = h_ent + int(horas_trabajo)
                    m_sal = m_ent + int((horas_trabajo % 1) * 60)
                    
                    h_sal += m_sal // 60
                    m_sal = m_sal % 60
                    
                    # Ajustar en caso de cruzar la medianoche
                    h_sal = h_sal % 24 
                    hora_salida = f"{h_sal:02d}:{m_sal:02d}"
                except Exception:
                    pass

        out_rows.append({
            'Empleados': nombre,
            'Horas de trabajo': horas_trabajo,
            'Hora de entrada': hora_entrada,
            'Hora_salida': hora_salida,
            'Estado': estado
        })

    return pd.DataFrame(out_rows)


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

    # Guardar el orden original del horario para respetarlo en la salida
    horario_df['_orden_visual'] = range(len(horario_df))

    # Merge: left inner join empleados que estén en horario
    merged = pd.merge(empleados_df, horario_df, how='inner', left_on=nombre_col, right_on='Empleados')

    # Restaurar el orden visual del mánager y limpiar columna temporal
    merged = merged.sort_values('_orden_visual').drop(columns=['_orden_visual']).reset_index(drop=True)

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

def cargar_todo(ruta_datos, ruta_horario):
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
