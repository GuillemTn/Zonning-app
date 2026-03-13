"""
Módulo motor_logica.py - VERSIÓN 2.2

Responsabilidad única: Lógica de asignación horaria con jerarquía en 3 fases + restricciones.

Arquitectura de Asignación:
1. Franja 07:00-10:00: CIERRE (empleados ACTIVOS en Almacén, OFF excluidos)
2. Franja 10:00-21:00: OPERACIÓN NORMAL (3 fases: Mínimos → Recomendados → Exceso)
3. Franja 21:00-22:00: CIERRE (1 Caja ACTIVO + resto CAMIÓN, OFF excluidos)

Restricciones Adicionales:
- Empleados OFF: No se asignan en todo el día
- Rotación COM: Cada zona COM puede repetirse máximo 2 veces por empleado (excepto COM-PS)
- COM-PS: Excepción a rotación, permite repetición sin límite
- Preferencia: Maximizar variedad de zonas visitadas (priorizar zonas no visitadas)

Motor de Rotación COM: 60 min obligatorio (cambio de zona cada hora)
Regla de Oro: Nunca asignar si habilidad = 0
"""

import pandas as pd
from copy import deepcopy
import random

# Registro de violaciones de continuidad (opción B: marcar cuando no hay candidato viable)
continuity_violations = []

# Registro de rechazos por Regla de Estancia Mínima (look-ahead)
minimum_stay_rejections = []

# Mapa temporal zona_name -> tipo (rellenado por `asignar` si se le pasa `zonas` con 'tipo')
current_zona_tipo_map = {}



# ============================================================================
# UTILIDADES
# ============================================================================

def obtener_habilidad(empleado_row, col_mapeo):
    """Obtiene la habilidad (0-5) de un empleado para una zona. Retorna 0 si no hay mapeo."""
    if col_mapeo is None:
        return 0
    try:
        val = empleado_row[col_mapeo]
        if pd.isna(val):
            return 0
        return int(val)
    except:
        try:
            return int(float(empleado_row[col_mapeo]))
        except:
            return 0


def puede_asignarse(empleado_row, col_mapeo):
    """
    Regla de Oro: Retorna True si el empleado puede asignarse a esa zona (habilidad > 0).
    """
    hab = obtener_habilidad(empleado_row, col_mapeo)
    return hab > 0


def obtener_zona_com_anterior(empleado_name, cuadrante_dict, hora_actual):
    """
    Retorna la zona COM en la que estuvo el empleado en la hora anterior (hora_actual - 1).
    Retorna None si no estaba en una zona COM o si es la primera hora.
    """
    if hora_actual <= 7:
        return None
    zona_anterior = cuadrante_dict.get((empleado_name, hora_actual - 1))
    if zona_anterior and 'COM' in zona_anterior.upper():
        return zona_anterior
    return None


def cuantas_horas_com_consecutivas(empleado_name, cuadrante_dict, hora_actual, zona_com):
    """
    Cuenta cuántas horas CONSECUTIVAS ha estado el empleado en esa misma zona COM.
    Retorna la cuenta (0 si no estaba en esa zona en la hora anterior).
    """
    count = 0
    for h in range(hora_actual - 1, 6, -1):  # Hacia atrás desde hora_actual - 1
        zona_h = cuadrante_dict.get((empleado_name, h))
        if zona_h == zona_com:
            count += 1
        else:
            break
    return count


def obtener_empleados_en_zona_hora(cuadrante_dict, zona_name, hora):
    """Retorna lista de empleados asignados a una zona en una hora específica."""
    return [e for (e, h), z in cuadrante_dict.items() if h == hora and z == zona_name]


def obtener_clase_empleado(empleado_row):
    """Normaliza y retorna la clase del empleado (e.g. 'comercial', 'operacional', 'hibrido')."""
    try:
        c = empleado_row.get('Clase', '')
        if pd.isna(c):
            return ''
        return str(c).strip().lower()
    except:
        return ''


def cuantas_horas_consecutivas_limit(empleado_name, cuadrante_dict, hora_actual, zona_name, min_hour_inclusive=7, max_hour_inclusive=20):
    """
    Cuenta cuántas horas consecutivas ha estado el empleado en `zona_name`
    mirando hacia atrás desde `hora_actual - 1` hasta `min_hour_inclusive`,
    pero ignora horas fuera del rango [min_hour_inclusive, max_hour_inclusive].
    Por ejemplo, para excluir 21-22 y 07-10 usar min=10, max=20.
    """
    count = 0
    for h in range(hora_actual - 1, min_hour_inclusive - 1, -1):
        if h < min_hour_inclusive:
            break
        if h > max_hour_inclusive:
            continue
        zona_h = cuadrante_dict.get((empleado_name, h))
        if zona_h == zona_name:
            count += 1
        else:
            break
    return count


def zona_es_com(zona_name):
    return 'COM' in str(zona_name).upper()


def zona_es_operacional(zona_name):
    """
    Verifica si una zona es OPERACIONAL (requiere estancia mínima).
    
    Zonas Operacionales (Críticas):
    - Cajas
    - C3/ALM
    - Almacén
    - Printing
    
    Retorna: True si es operacional, False si es de apoyo o COM
    """
    # Priorizar la clasificación explícita si está disponible
    try:
        tipo = current_zona_tipo_map.get(zona_name)
        if tipo:
            return tipo == 'operacional'
    except Exception:
        pass

    zona_lower = str(zona_name).lower()
    zonas_op = ['cajas', 'c3/back', 'almacén', 'camión', 'camion', 'printing']
    return any(op in zona_lower for op in zonas_op)


def minimo_horas_para_zona(z):
    """
    Define la estancia mínima requerida para una zona.
    Retorna: int (horas)
    """
    # Priorizar clasificación explícita
    try:
        tipo = current_zona_tipo_map.get(z)
        if tipo:
            if tipo == 'operacional': return 3
            return 1
    except:
        pass
    zn = str(z).lower()
    # Si es operacional o contiene palabras clave, mínimo 3h
    if any(x in zn for x in ['printing', 'cajas', 'c3/alm', 'almacén', 'camión', 'camion']):
        return 3
    return 1


def puede_completar_estancia_minima(empleado_name, zona_name, hora_actual, empleados_hoy, nombre_col):
    """
    REGLA DE ESTANCIA MÍNIMA - Validación Look-Ahead (60 minutos mínimo).
    
    Lógica de Validación:
    - Para ZONAS OPERACIONALES: El empleado debe estar activo en AMBAS horas (T y T+1)
    - Si el empleado TERMINA su turno en la hora actual (T), NO puede asignarse a zona operacional
    - Si la zona no requiere personal en T+1, NO se debe asignar a nadie nuevo en T
    
    Parámetros:
        empleado_name: Nombre del empleado
        zona_name: Nombre de la zona
        hora_actual: Hora actual (T)
        empleados_hoy: DataFrame de empleados
        nombre_col: Nombre de la columna de empleados
    
    Retorna:
        True si puede asignarse (cumple estancia mínima)
        False si NO cumple (sería su última hora - violación)
    """
    # Si no es zona operacional, permite la asignación (zonas COM y Almacén general)
    if not zona_es_operacional(zona_name):
        return True
    
    # Regla de Bloqueo: Zonas operacionales requieren MÍNIMO 2 horas (T y T+1)
    # Obtener la hora de salida del empleado
    emp_row = empleados_hoy[empleados_hoy[nombre_col] == empleado_name]
    if emp_row.empty:
        return False
    
    emp_row = emp_row.iloc[0]
    hora_salida = parsear_hora_a_entero(emp_row.get('Hora_salida', ''))
    
    if hora_salida is None:
        return False
    
    # Validar Look-ahead: ¿Estará empleado en T+1?
    # Regla: Entrada <= T+1 < Salida
    #
    # Si Salida <= T+1 (ej: Salida=22, T=21, T+1=22), el empleado NO estará en T+1
    # Por lo tanto, NO puede asignarse en T a zona operacional
    
    turno_restante = max(0, hora_salida - hora_actual)
    
    # Si le quedan más de 1 hora de turno (es decir, T+1 está dentro de su rango), permitir
    if turno_restante > 1:
        return True
    
    # Si turno_restante == 1, significa que esta es su ÚLTIMA hora
    # NO permitir asignación a zona operacional
    if turno_restante == 1:
        return False
    
    # Si turno_restante == 0, ya está fuera de turno
    return False


def clase_coincide_con_zona(clase_norm, zona_name):
    """Devuelve True si la clase del empleado tiene afinidad con la zona."""
    if not clase_norm:
        return False
    # Usar clasificación explícita si existe
    try:
        tipo = current_zona_tipo_map.get(zona_name)
        if tipo == 'comercial':
            return 'comercial' in clase_norm or 'com' in clase_norm
        if tipo == 'operacional':
            return 'oper' in clase_norm or 'operacional' in clase_norm
    except Exception:
        pass

    zn = str(zona_name).lower()
    # Comerciales prefieren zonas COM-
    if 'com' in zn:
        return 'comercial' in clase_norm or 'com' in clase_norm
    # Operacionales prefieren Cajas, Almacén y Printing
    if any(x in zn for x in ['cajas', 'almacén', 'camión', 'printing', 'camion']):
        return 'oper' in clase_norm or 'operacional' in clase_norm
    return False


def seleccionar_empleado_para_zona(
    empleados_activos, empleados_hoy, nombre_col, mapeo_habilidades,
    zona_name, cuadrante_dict, hora_actual
):
    """
    Selecciona el mejor empleado para `zona_name` siguiendo las reglas de:
    1) Especialistas (habilidad == 4)
    2) Coincidencia por Clase (afinidad)
    3) Mejor habilidad restante

    Respeta rotación COM, regla de oro, excepciones para `Cajas`
    Y APLICA la Regla de Estancia Mínima (look-ahead para zonas operacionales).
    """
    
    # Variable para almacenar el mejor candidato que NO cumple estancia mínima (fallback)
    best_fallback_candidate = None

    # Función auxiliar: Busca el primer candidato viable que cumple la regla de estancia mínima
    def buscar_candidato_viable(lista_candidatos):
        """
        Itera lista_candidatos y retorna el nombre del primero que:
        1. No está ya asignado en esta hora
        2. Cumple la regla de estancia mínima
        
        Si ninguno cumple, retorna None y registra rechazos por estancia mínima.
        """
        nonlocal best_fallback_candidate
        for c in lista_candidatos:
            emp_name = c['nombre']
            # Si ya está asignado, saltar
            if (emp_name, hora_actual) in cuadrante_dict:
                continue
            # Validar regla de estancia mínima
            if puede_completar_estancia_minima(emp_name, zona_name, hora_actual, empleados_hoy, nombre_col):
                return emp_name
            else:
                # Si es zona operacional y no cumple por turno restante, registrar rechazo
                if zona_es_operacional(zona_name):
                    emp_row = empleados_hoy[empleados_hoy[nombre_col] == emp_name].iloc[0]
                    hora_salida = parsear_hora_a_entero(emp_row.get('Hora_salida', ''))
                    turno_restante = max(0, hora_salida - hora_actual) if hora_salida else 0
                    
                    # Guardar como fallback si es el primero (mejor opción disponible por jerarquía)
                    if best_fallback_candidate is None:
                        best_fallback_candidate = emp_name

                    minimum_stay_rejections.append({
                        'hora': hora_actual,
                        'zona': zona_name,
                        'empleado': emp_name,
                        'turno_restante': turno_restante,
                        'reason': 'unsuitable_stay_duration'
                    })
        return None
    
    col_mapeo = mapeo_habilidades.get(zona_name)
    candidatos = []
    for emp_name in empleados_activos:
        # Si ya asignado en esta hora, saltar
        if (emp_name, hora_actual) in cuadrante_dict:
            continue
        emp_row = empleados_hoy[empleados_hoy[nombre_col] == emp_name].iloc[0]

        # Regla de Oro: habilidad > 0 (Validación estricta solicitada)
        hab = obtener_habilidad(emp_row, col_mapeo) if col_mapeo else 0
        if hab <= 0:
            continue

        # Para zonas COM: verificar restricciones de rotación
        if zona_es_com(zona_name):
            if cuantas_horas_com_consecutivas(emp_name, cuadrante_dict, hora_actual, zona_name) >= 1:
                continue
            if not puede_asignarse_com_rotacion(emp_name, zona_name, cuadrante_dict):
                continue

        clase_norm = obtener_clase_empleado(emp_row)
        # Calcular horas restantes en turno
        hs = parsear_hora_a_entero(emp_row.get('Hora_salida', ''))
        rem = 0
        try:
            if hs is not None:
                rem = max(0, hs - hora_actual)
        except:
            rem = 0

        # Penalización por transición FRONT <-> F2 (Preferencia)
        penalty = 0
        prev = cuadrante_dict.get((emp_name, hora_actual - 1))
        if zona_name == 'F2' and prev == 'FRONT':
            penalty = 1
        elif zona_name == 'FRONT' and prev == 'F2':
            penalty = 1

        candidatos.append({
            'nombre': emp_name,
            'hab': hab,
            'clase': clase_norm,
            'row': emp_row,
            'hora_salida_int': hs,
            'rem_horas': rem,
            'penalty': penalty
        })

    if not candidatos:
        return None

    # Funciones auxiliares para posiciones operativas
    def es_posicion_operativa(z):
        # Preferir clasificación por tipo si está disponible
        try:
            tipo = current_zona_tipo_map.get(z)
            if tipo:
                return tipo == 'operacional'
        except Exception:
            pass
        zn = str(z).lower()
        return any(x in zn for x in ['cajas', 'C3/ALM', 'almacén', 'camión', 'camion', 'printing'])

    min_total = minimo_horas_para_zona(zona_name)

    # 0) Si existe un ocupante en la hora previa, priorizar que continúe (no sacar)
    ocupantes_prev = [c for c in candidatos if cuadrante_dict.get((c['nombre'], hora_actual - 1)) == zona_name]
    if ocupantes_prev:
        # Si hay varios, preferir quien tenga más rem_horas y luego mayor habilidad
        ocupantes_prev.sort(key=lambda x: (-x.get('rem_horas', 0), -x['hab']))
        
        # MODIFICACION: Prioridad absoluta a la continuidad.
        # Si ya estaba en la zona, permitimos que termine su turno ahí (aunque le quede 1h).
        # Eximimos a los ocupantes previos de la regla de estancia mínima (look-ahead).
        return ocupantes_prev[0]['nombre']

    candidatos_no_last = []
    for c in candidatos:
        prev_consec = cuantas_horas_consecutivas_limit(c['nombre'], cuadrante_dict, hora_actual, zona_name, min_hour_inclusive=10, max_hour_inclusive=20)
        # Si estuvo en la zona en la hora previa y aún no alcanza el mínimo, debe continuar
        if cuadrante_dict.get((c['nombre'], hora_actual - 1)) == zona_name and (prev_consec + 1) < min_total:
            candidatos_no_last.append(c)
    if candidatos_no_last:
        # restringir candidatos a los que deben continuar
        candidatos = candidatos_no_last
    else:
        # Si no hay empleados que deban continuar obligatoriamente, preferir candidatos
        # que puedan completar el mínimo (no aplicar si hay especialistas)
        especialistas_all = [c for c in candidatos if c['hab'] == 4]
        if not especialistas_all:
            candidatos_long = [c for c in candidatos if c.get('rem_horas', 0) >= min_total]
            if candidatos_long:
                candidatos = candidatos_long
    # Si llegamos aquí y existe un ocupante en la hora previa que no puede continuar
    # porque no hay candidato viable, registramos una violación (opción B: marcar pero permitir)
    prev_occupants_all = [e for (e, h), z in cuadrante_dict.items() if h == hora_actual - 1 and z == zona_name]
    if prev_occupants_all:
        for prev in prev_occupants_all:
            prev_consec = cuantas_horas_consecutivas_limit(prev, cuadrante_dict, hora_actual, zona_name, min_hour_inclusive=10, max_hour_inclusive=20)
            if prev_consec < min_total:
                # Registrar violación: será analizada en la visualización
                continuity_violations.append({
                    'hora': hora_actual,
                    'zona': zona_name,
                    'empleado_prev': prev,
                    'prev_consec': prev_consec,
                    'min_total': min_total,
                    'reason': 'no_candidato_viable_para_continuar'
                })

    # Para zonas COM: preferir candidatos cuyo COM anterior sea distinto (evitar repetir A-A)
    if zona_es_com(zona_name):
        candidatos_no_last = [c for c in candidatos if obtener_zona_com_anterior(c['nombre'], cuadrante_dict, hora_actual) != zona_name]
        if candidatos_no_last:
            candidatos = candidatos_no_last

    # 1) Especialistas: hab == 4
    especialistas = [c for c in candidatos if c['hab'] == 4]
    if especialistas:
        # Preferir especialista con clase coincidente
        espec_clase = [c for c in especialistas if clase_coincide_con_zona(c['clase'], zona_name)]
        if espec_clase:
            # Si hay múltiples, aplicar criterios de variedad para COM
            if zona_es_com(zona_name) and len(espec_clase) > 1:
                def _key_variedad(c):
                    nombre = c['nombre']
                    reps = contar_repeticiones_zona_com(nombre, zona_name, cuadrante_dict)
                    last_same = 1 if obtener_zona_com_anterior(nombre, cuadrante_dict, hora_actual) == zona_name else 0
                    consec = cuantas_horas_com_consecutivas(nombre, cuadrante_dict, hora_actual, zona_name)
                    distinct = len(set([z for (e,h), z in cuadrante_dict.items() if e == nombre and 'COM' in str(z).upper()]))
                    return (c['penalty'], reps, last_same, consec, distinct, -c['hab'])
                espec_clase.sort(key=_key_variedad)
            # Buscar el primero que cumpla estancia mínima
            viable = buscar_candidato_viable(espec_clase)
            if viable:
                return viable
        # Si no hay especialistas con clase coincidente, elegir el primero especialista
        viable = buscar_candidato_viable(especialistas)
        if viable:
            return viable

    # 2) Clase coincidente
    clase_match = [c for c in candidatos if clase_coincide_con_zona(c['clase'], zona_name)]
    if clase_match:
        # Para zonas COM, priorizar variedad (menos repeticiones en la misma zona, evitar último mismo)
        if zona_es_com(zona_name):
            def _key_com(c):
                nombre = c['nombre']
                reps = contar_repeticiones_zona_com(nombre, zona_name, cuadrante_dict)
                last_same = 1 if obtener_zona_com_anterior(nombre, cuadrante_dict, hora_actual) == zona_name else 0
                consec = cuantas_horas_com_consecutivas(nombre, cuadrante_dict, hora_actual, zona_name)
                distinct = len(set([z for (e,h), z in cuadrante_dict.items() if e == nombre and 'COM' in str(z).upper()]))
                # Para posiciones operativas preferir candidatos que pueden completar el mínimo
                needed_more = max(0, minimo_horas_para_zona(zona_name) - (consec + 1))
                rem_ok = 1 if c.get('rem_horas', 0) >= needed_more else 0
                return (c['penalty'], reps, last_same, consec, distinct, -rem_ok, -c['hab'])
            clase_match.sort(key=_key_com)
        else:
            clase_match.sort(key=lambda x: (x['penalty'], -x['hab']))
        # Buscar el primero que cumpla estancia mínima
        viable = buscar_candidato_viable(clase_match)
        if viable:
            return viable

    # 3) Excepción por talento: permitir Operacional en COM si hab >=3 y no hay Comerciales
    if zona_es_com(zona_name):
        oper_high = [c for c in candidatos if ('oper' in c['clase'] or 'operacional' in c['clase']) and c['hab'] >= 3]
        if oper_high:
            # Aplicar también criterio de variedad
            def _key_oper(c):
                nombre = c['nombre']
                reps = contar_repeticiones_zona_com(nombre, zona_name, cuadrante_dict)
                last_same = 1 if obtener_zona_com_anterior(nombre, cuadrante_dict, hora_actual) == zona_name else 0
                consec = cuantas_horas_com_consecutivas(nombre, cuadrante_dict, hora_actual, zona_name)
                distinct = len(set([z for (e,h), z in cuadrante_dict.items() if e == nombre and 'COM' in str(z).upper()]))
                return (c['penalty'], reps, last_same, consec, distinct, -c['hab'])
            oper_high.sort(key=_key_oper)
            # Buscar el primero que cumpla estancia mínima
            viable = buscar_candidato_viable(oper_high)
            if viable:
                return viable

    # 4) Finalmente, escoger por mayor habilidad disponible (aplicar variedad para COM)
    if zona_es_com(zona_name):
        def _key_final(c):
            nombre = c['nombre']
            reps = contar_repeticiones_zona_com(nombre, zona_name, cuadrante_dict)
            last_same = 1 if obtener_zona_com_anterior(nombre, cuadrante_dict, hora_actual) == zona_name else 0
            consec = cuantas_horas_com_consecutivas(nombre, cuadrante_dict, hora_actual, zona_name)
            distinct = len(set([z for (e,h), z in cuadrante_dict.items() if e == nombre and 'COM' in str(z).upper()]))
            return (c['penalty'], reps, last_same, consec, distinct, -c['hab'])
        candidatos.sort(key=_key_final)
    else:
        candidatos.sort(key=lambda x: (x['penalty'], -x['hab']))
    
    # Buscar el primero que cumpla estancia mínima
    res = buscar_candidato_viable(candidatos)
    if res:
        return res
    
    # Si no se encontró candidato ideal, usar fallback si existe (prioridad: cobertura > continuidad)
    return best_fallback_candidate



def es_empleado_off(empleado_row):
    """
    Retorna True si el empleado está OFF (Estado = 'OFF').
    Si no existe columna Estado, retorna False.
    """
    try:
        estado = empleado_row.get('Estado', 'ON')
        return estado == 'OFF'
    except:
        return False


def parsear_hora_a_entero(hora_str):
    """
    Convierte una hora en formato 'HH:MM' a entero (hora en formato 24h).
    
    Ejemplos:
        '07:00' → 7
        '15:30' → 15
        '23:45' → 23
    
    Retorna:
        int: hora (0-23), o None si no puede parsear
    """
    if pd.isna(hora_str) or str(hora_str).strip() == '':
        return None
    
    try:
        hora_str = str(hora_str).strip()
        # Formato HH:MM
        if ':' in hora_str:
            partes = hora_str.split(':')
            return int(partes[0])
        # Formato solo HH
        else:
            return int(hora_str)
    except:
        return None


def empleado_esta_en_turno(empleado_row, hora_actual):
    """
    Verifica si un empleado DEBE estar en turno a una hora específica.
    
    Regla Matemática: Hora de Entrada <= Hora Actual < Hora de Salida
    
    Parámetros:
        empleado_row: fila de empleado con columnas 'Hora de entrada', 'Hora_salida'
        hora_actual: int (0-23, hora en formato 24h)
    
    Retorna:
        True si está en turno, False si está fuera
    """
    # Obtener horas de entrada y salida
    hora_entrada = parsear_hora_a_entero(empleado_row.get('Hora de entrada', ''))
    hora_salida = parsear_hora_a_entero(empleado_row.get('Hora_salida', ''))
    
    # Si no se puede parsear, asumir que está OFF
    if hora_entrada is None or hora_salida is None:
        return False
    
    # Aplicar regla: Entrada <= Actual < Salida
    if hora_entrada <= hora_actual < hora_salida:
        return True
    
    # CASO ESPECIAL: Si trabaja de 22:00 a 23:00 (1 hora), a las 22:00 está dentro
    # pero a las 23:00 debe estar fuera (OFF)
    return False


def obtener_empleados_activos_ahora(empleados_hoy, hora_actual):
    """
    Genera una sub-lista de empleados ACTIVOS en la hora actual.
    
    Criterios:
    1. No están en Estado 'OFF'
    2. Están dentro de su rango horario: Entrada <= Hora Actual < Salida
    
    Parámetros:
        empleados_hoy: DataFrame con todas las filas
        hora_actual: int (0-23)
    
    Retorna:
        Set de nombres de empleados activos ahora
    """
    nombre_col = empleados_hoy.columns[0]
    empleados_activos = set()
    
    for _, emp_row in empleados_hoy.iterrows():
        emp_name = emp_row[nombre_col]
        
        # Saltar si está OFF
        if es_empleado_off(emp_row):
            continue
        
        # Verificar si está en su rango horario
        if empleado_esta_en_turno(emp_row, hora_actual):
            empleados_activos.add(emp_name)
    
    return empleados_activos


def contar_repeticiones_zona_com(empleado_name, zona_com, cuadrante_dict):
    """
    Cuenta cuántas veces el empleado ha sido asignado a una zona COM específica.
    
    Retorna: int (cantidad de horas en esa zona COM)
    """
    if zona_com.upper() == 'COM-PS':
        # COM-PS puede repetirse sin límite, retorna 0 para permitir siempre
        return 0
    
    count = 0
    for (emp_name, h), zona in cuadrante_dict.items():
        if emp_name == empleado_name and zona == zona_com:
            count += 1
    return count


def puede_asignarse_com_rotacion(empleado_name, zona_com, cuadrante_dict):
    """
    Verifica si el empleado puede ser asignado a una zona COM específica.
    
    Regla: Cada zona COM (excepto COM-PS) NO puede ser asignada más de 2 veces 
    al mismo empleado en todo el día.
    COM-PS es excepción y se puede repetir sin límite.
    
    Retorna: True si puede asignarse, False si ya alcanzó el límite de 2 repeticiones.
    """
    # COM-PS es excepción: siempre permitido
    if zona_com.upper() == 'COM-PS':
        return True
    
    # Contar cuántas veces ya fue asignado a esta zona COM específica
    repeticiones = contar_repeticiones_zona_com(empleado_name, zona_com, cuadrante_dict)
    
    # Si ya fue asignado 2 veces a esta zona, denegar tercera asignación
    if repeticiones >= 2:
        return False
    
    return True


# ============================================================================
# ASIGNACIÓN POR HORA CON VALIDACIÓN HORARIA
# ============================================================================

def asignar_empleados_por_hora(
    empleados_hoy, zonas, mapeo_habilidades, cuadrante_dict, hora_actual
):
    """
    Asigna empleados a zonas para UNA HORA ESPECÍFICA (en rango 7-23).
    
    Lógica:
    1. Obtener empleados ACTIVOS ahora (dentro de su rango horario)
    2. Marcar empleados INACTIVOS como 'OFF' (fuera de rango)
    3. Aplicar lógica de asignación según la hora:
       - 07:00-09:59: TODOS a Almacén
       - 10:00-20:59: 3 Fases (Mínimos, Recomendados, Máximos)
    - 21:00-22:59: 1 Caja + Resto CAMIÓN
    
    Parámetros:
        hora_actual: int (7-23)
    """
    nombre_col = empleados_hoy.columns[0]
    
    # PASO 1: Obtener empleados activos y marcar inactivos como OFF
    empleados_activos = obtener_empleados_activos_ahora(empleados_hoy, hora_actual)
    
    # Marcar empleados INACTIVOS como 'OFF'
    for _, emp_row in empleados_hoy.iterrows():
        emp_name = emp_row[nombre_col]
        
        # Si NO está en empleados_activos, marcar 'OFF'
        if emp_name not in empleados_activos:
            cuadrante_dict[(emp_name, hora_actual)] = 'OFF'
    
    # PASO 2: Asignar empleados activos a zonas según la hora
    if 7 <= hora_actual < 10:
        # FRANJA MAÑANA (07:00-09:59): Todos a Almacén
        for emp_name in empleados_activos:
            cuadrante_dict[(emp_name, hora_actual)] = 'Almacén'
    
    elif 10 <= hora_actual < 21:
        # FRANJA OPERACIÓN NORMAL (10:00-20:59): 3 Fases
        _asignar_3_fases(
            empleados_hoy, empleados_activos, zonas, mapeo_habilidades,
            cuadrante_dict, hora_actual
        )
    
    elif 21 <= hora_actual < 23:
        # FRANJA CIERRE (21:00-22:59): 1 Caja + Resto CAMIÓN
        _asignar_cierre(
            empleados_hoy, empleados_activos, zonas, mapeo_habilidades,
            cuadrante_dict, hora_actual
        )


def _asignar_3_fases(
    empleados_hoy, empleados_activos, zonas, mapeo_habilidades,
    cuadrante_dict, hora_actual
):
    """
    Aplica las 3 fases de asignación para una hora específica.
    Solo usa empleados_activos (filtrados por disponibilidad horaria).
    
    Fase 1: Cubrir MÍNIMOS por PRIORIDAD
    Fase 2: Rellenar hasta RECOMENDADO
    Fase 3: Rellenar hasta MÁXIMO
    """
    nombre_col = empleados_hoy.columns[0]
    
    # Rastrear estado de cada zona
    estado_zonas = {}
    for zona in zonas:
        zona_name = zona['name']
        estado_zonas[zona_name] = {
            'asignados': [],
            'min_faltantes': zona['min'],
            'rec_faltantes': zona['recomendado'] - zona['min'],
            'max_faltantes': zona['max'] - zona['recomendado']
        }
    
    # Helper: Si hay alguien en C3/Back (o C3/ALM), no asignar a Back (o Almacén)
    def zona_bloqueada_por_c3(z_name):
        z_up = z_name.upper()
        # Si es zona Back (y no es C3). NO BLOQUEAR ALMACÉN.
        if 'BACK' in z_up and 'C3' not in z_up:
            # Verificar si alguna zona C3 tiene asignados
            for k, v in estado_zonas.items():
                k_up = k.upper()
                if 'C3' in k_up and ('BACK' in k_up or 'ALM' in k_up) and len(v['asignados']) > 0:
                    return True
        return False

    # =======================================================================
    # FASE 0: INERCIA OPERATIVA ABSOLUTA (Refactorización)
    # =======================================================================
    # REGLA: Si un empleado estuvo en una zona operacional en la hora anterior
    # (T-1), se le ancla a esa zona en la hora actual (T) hasta fin de turno.
    # Esto aplica a partir de las 11:00 para permitir la redistribución a las 10:00.
    
    if hora_actual >= 11:  # La lógica de anclaje empieza a las 11:00
        for emp_name in list(empleados_activos):
            # Obtener la zona de la hora anterior (T-1)
            prev_zona = cuadrante_dict.get((emp_name, hora_actual - 1))
            
            # Si no tiene asignación previa o no es una zona operacional, no se ancla
            if not prev_zona or not zona_es_operacional(prev_zona):
                continue
            
            # EXCEPCIÓN: Si tiene habilidad 1, no forzar permanencia (permitir rotación/salida)
            col_mapeo = mapeo_habilidades.get(prev_zona)
            emp_row = empleados_hoy[empleados_hoy[nombre_col] == emp_name].iloc[0]
            if obtener_habilidad(emp_row, col_mapeo) == 1:
                continue
            
            # EXCEPCIÓN 2: Prioridad Habilidad 4 (Especialista) sobre Inercia Operativa.
            # Si es especialista en CUALQUIER zona, liberar para que la asignación lo pueda mover.
            es_especialista_global = False
            for col_h in set(mapeo_habilidades.values()):
                if obtener_habilidad(emp_row, col_h) == 4:
                    es_especialista_global = True
                    break
            if es_especialista_global:
                continue
            
            # REGLA 1 y 2: Anclaje y Exclusión
            # Si estaba en una zona operacional, se le reasigna y se saca de la piscina.
            # Se elimina la condición de `min_total` o `min_req`.
            cuadrante_dict[(emp_name, hora_actual)] = prev_zona
            
            # Excluir de candidatos disponibles para las siguientes fases
            if emp_name in empleados_activos:
                empleados_activos.remove(emp_name)
            
            # Actualizar contadores de la zona para que las fases 1, 2 y 3 lo reconozcan
            if prev_zona in estado_zonas:
                estado_zonas[prev_zona]['asignados'].append(emp_name)
                if estado_zonas[prev_zona]['min_faltantes'] > 0:
                    estado_zonas[prev_zona]['min_faltantes'] -= 1
                elif estado_zonas[prev_zona]['rec_faltantes'] > 0:
                    estado_zonas[prev_zona]['rec_faltantes'] -= 1
                elif estado_zonas[prev_zona]['max_faltantes'] > 0:
                    estado_zonas[prev_zona]['max_faltantes'] -= 1

    # FASE 1: Cubrir MÍNIMOS por PRIORIDAD (Modificado: Operacionales primero)
    def sort_key_prioridad(z):
        is_oper = zona_es_operacional(z['name'])
        is_almacen = 'ALMAC' in z['name'].upper()
        # Orden: 1. Operacionales (True=1, False=0 -> -1 primero), 2. Prioridad, 3. No Almacén
        return (-int(is_oper), -z['prioridad'], 1 if is_almacen else 0, z['name'])

    # FASE 1: Cubrir MÍNIMOS por PRIORIDAD
    zonas_en_prioridad = sorted(
        [z for z in zonas if z['min'] > 0],
        key=sort_key_prioridad
    )

    for zona in zonas_en_prioridad:
        zona_name = zona['name']
        if zona_bloqueada_por_c3(zona_name):
            continue
            
        col_mapeo = mapeo_habilidades.get(zona_name)
        faltantes = estado_zonas[zona_name]['min_faltantes']
        
        for _ in range(faltantes):
            mejor_emp = seleccionar_empleado_para_zona(
                empleados_activos, empleados_hoy, nombre_col, mapeo_habilidades,
                zona_name, cuadrante_dict, hora_actual
            )

            if mejor_emp:
                cuadrante_dict[(mejor_emp, hora_actual)] = zona_name
                estado_zonas[zona_name]['asignados'].append(mejor_emp)
                estado_zonas[zona_name]['min_faltantes'] -= 1
    
    # FASE 2: Rellenar hasta RECOMENDADO (Optimización)
    # Usar orden de prioridad para asegurar que las zonas más importantes se optimicen primero
    zonas_con_rec = sorted([z for z in zonas if z['recomendado'] > z['min']], key=sort_key_prioridad)
    
    for zona in zonas_con_rec:
        zona_name = zona['name']
        if zona_bloqueada_por_c3(zona_name):
            continue
            
        col_mapeo = mapeo_habilidades.get(zona_name)
        faltantes = estado_zonas[zona_name]['rec_faltantes']
        
        for _ in range(faltantes):
            mejor_emp = seleccionar_empleado_para_zona(
                empleados_activos, empleados_hoy, nombre_col, mapeo_habilidades,
                zona_name, cuadrante_dict, hora_actual
            )

            if mejor_emp:
                cuadrante_dict[(mejor_emp, hora_actual)] = zona_name
                estado_zonas[zona_name]['asignados'].append(mejor_emp)
                estado_zonas[zona_name]['rec_faltantes'] -= 1
    
    # FASE 3: Rellenar hasta MÁXIMO (Exceso de personal)
    # Orden: Zonas COM -> Zonas operativas (Cajas y Printing) -> Almacén (último recurso)
    def sort_key_fase3(z):
        zn = z['name'].upper()
        if 'COM' in zn: return 0
        if 'CAJAS' in zn or 'PRINTING' in zn: return 1
        if 'ALMAC' in zn: return 3
        return 2 # Otras operativas

    zonas_con_max = sorted([z for z in zonas if z['max'] > z['recomendado']], key=lambda z: (sort_key_fase3(z), -z['prioridad']))
    
    for zona in zonas_con_max:
        zona_name = zona['name']
        if zona_bloqueada_por_c3(zona_name):
            continue
            
        col_mapeo = mapeo_habilidades.get(zona_name)
        faltantes = estado_zonas[zona_name]['max_faltantes']
        
        for _ in range(faltantes):
            mejor_emp = seleccionar_empleado_para_zona(
                empleados_activos, empleados_hoy, nombre_col, mapeo_habilidades,
                zona_name, cuadrante_dict, hora_actual
            )

            if mejor_emp:
                cuadrante_dict[(mejor_emp, hora_actual)] = zona_name
                estado_zonas[zona_name]['asignados'].append(mejor_emp)
                estado_zonas[zona_name]['max_faltantes'] -= 1

    # =======================================================================
    # FASE 4: RESCATE POR INTERCAMBIO (SWAP)
    # =======================================================================
    # Detectar huérfanos: Activos que no están en cuadrante_dict para esta hora
    huerfanos = [e for e in empleados_activos if (e, hora_actual) not in cuadrante_dict]
    
    if huerfanos:
        # Identificar zonas con huecos (donde max_faltantes > 0)
        zonas_con_hueco = [z_name for z_name, estado in estado_zonas.items() if estado['max_faltantes'] > 0]
        
        for huerfano in huerfanos:
            asignado = False
            huerfano_row = empleados_hoy[empleados_hoy[nombre_col] == huerfano].iloc[0]
            
            # 1. Intentar asignación directa a huecos (sanity check)
            for z_hueco in list(zonas_con_hueco):
                col_hueco = mapeo_habilidades.get(z_hueco)
                if puede_asignarse(huerfano_row, col_hueco):
                    cuadrante_dict[(huerfano, hora_actual)] = z_hueco
                    estado_zonas[z_hueco]['asignados'].append(huerfano)
                    estado_zonas[z_hueco]['max_faltantes'] -= 1
                    if estado_zonas[z_hueco]['max_faltantes'] <= 0:
                        if z_hueco in zonas_con_hueco: zonas_con_hueco.remove(z_hueco)
                    asignado = True
                    break
            
            if asignado: continue

            # 2. Búsqueda de Swap
            # Obtener zonas donde el huérfano tiene habilidad > 0
            zonas_candidatas_huerfano = []
            for z in zonas:
                z_name = z['name']
                col_z = mapeo_habilidades.get(z_name)
                if puede_asignarse(huerfano_row, col_z):
                    zonas_candidatas_huerfano.append(z_name)
            
            swap_done = False
            for z_target in zonas_candidatas_huerfano:
                if swap_done: break
                if z_target not in estado_zonas: continue
                
                # PROTECCIÓN: No desasignar de zonas operacionales (Cajas, Printing, Almacén, C3/Back)
                if zona_es_operacional(z_target) and z_target.upper() in ['C3/BACK', 'C3/ALM']:
                    continue
                
                ocupantes = estado_zonas[z_target]['asignados']
                # Iterar copia de ocupantes
                for ocupante in list(ocupantes):
                    ocupante_row = empleados_hoy[empleados_hoy[nombre_col] == ocupante].iloc[0]
                    
                    # Buscar si ocupante cabe en algún hueco
                    for z_dest in list(zonas_con_hueco):
                        col_dest = mapeo_habilidades.get(z_dest)
                        if puede_asignarse(ocupante_row, col_dest):
                            # EXECUTE SWAP
                            # Mover ocupante z_target -> z_dest
                            cuadrante_dict[(ocupante, hora_actual)] = z_dest
                            estado_zonas[z_target]['asignados'].remove(ocupante)
                            estado_zonas[z_dest]['asignados'].append(ocupante)
                            estado_zonas[z_dest]['max_faltantes'] -= 1
                            if estado_zonas[z_dest]['max_faltantes'] <= 0:
                                if z_dest in zonas_con_hueco: zonas_con_hueco.remove(z_dest)
                            
                            # Asignar huérfano -> z_target
                            cuadrante_dict[(huerfano, hora_actual)] = z_target
                            estado_zonas[z_target]['asignados'].append(huerfano)
                            
                            swap_done = True
                            asignado = True
                            break
                    if swap_done: break
            
            if asignado: continue
            
            # 3. Fallback: Almacén
            almacen_zone = next((z['name'] for z in zonas if 'ALMAC' in z['name'].upper()), None)
            if almacen_zone:
                cuadrante_dict[(huerfano, hora_actual)] = almacen_zone
                if almacen_zone in estado_zonas:
                    estado_zonas[almacen_zone]['asignados'].append(huerfano)
                    if estado_zonas[almacen_zone]['max_faltantes'] > 0:
                        estado_zonas[almacen_zone]['max_faltantes'] -= 1


def _asignar_cierre(
    empleados_hoy, empleados_activos, zonas, mapeo_habilidades,
    cuadrante_dict, hora_actual
):
    """
    Asigna para franja cierre (21:00-22:59).
    1 Caja + Resto CAMIÓN.
    Prioridad para Cajas: 1º quien estaba en Cajas, 2º en C3/Back o C3/ALM.
    """
    nombre_col = empleados_hoy.columns[0]
    col_cajas = mapeo_habilidades.get('Cajas')
    
    # Prioridad 1: Empleados que estaban en Cajas a las 20:00
    prio_1_cajas = [
        e for (e, h), z in cuadrante_dict.items()
        if h == 20 and z == 'Cajas'
    ]
    
    # Prioridad 2: Empleados que estaban en C3/Back o C3/ALM a las 20:00
    prio_2_c3 = [
        e for (e, h), z in cuadrante_dict.items()
        if h == 20 and z.upper() in ('C3/BACK', 'C3/ALM')
    ]

    cajero_asignado = False

    # Función auxiliar para asignar un cajero de una lista de candidatos
    def intentar_asignar_cajero(lista_candidatos):
        nonlocal cajero_asignado
        if cajero_asignado:
            return

        for emp_name in lista_candidatos:
            # Comprobar si el empleado está activo y aún no ha sido asignado en esta hora
            if emp_name in empleados_activos and (emp_name, hora_actual) not in cuadrante_dict:
                emp_row = empleados_hoy[empleados_hoy[nombre_col] == emp_name].iloc[0]
                # Comprobar si tiene habilidad para Cajas
                if puede_asignarse(emp_row, col_cajas):
                    cuadrante_dict[(emp_name, hora_actual)] = 'Cajas'
                    cajero_asignado = True
                    return  # Salir en cuanto se asigne uno

    # Aplicar prioridades
    intentar_asignar_cajero(prio_1_cajas)
    intentar_asignar_cajero(prio_2_c3)

    # Prioridad 3: Si no se asignó a nadie, buscar en el resto de empleados activos
    if not cajero_asignado:
        # Crear una lista de todos los demás para mantener un orden consistente
        resto_activos = [emp for emp in empleados_activos if emp not in prio_1_cajas and emp not in prio_2_c3]
        intentar_asignar_cajero(resto_activos)

    # Resto de empleados activos a CAMIÓN
    for emp_name in empleados_activos:
        if (emp_name, hora_actual) not in cuadrante_dict:
            cuadrante_dict[(emp_name, hora_actual)] = 'CAMIÓN'




# ============================================================================
# FUNCIÓN PRINCIPAL DEL MOTOR
# ============================================================================

def asignar(empleados_hoy, zonas, mapeo_habilidades):
    """
    Genera asignaciones horarias (07:00-22:59) aplicando:
    
    VALIDACIÓN HORARIA INTEGRADA:
    - Cada hora valida si cada empleado está en su rango (Entrada <= Hora < Salida)
    - Empleados fuera de rango: 'OFF'
    - Empleados en rango: se asignan según la franja
    
    FRANJAS:
    1. 07:00-09:59: ALMACÉN (todos activos en rango)
    2. 10:00-20:59: OPERACIÓN NORMAL (3 fases)
    3. 21:00-22:59: 1 CAJA + CAMIÓN
    
    Retorna:
        Dict {(empleado_name, hora): zona_name}
    """
    cuadrante = {}
    # Construir mapa temporal zona -> tipo (si la estructura de zonas lo trae)
    try:
        global current_zona_tipo_map
        current_zona_tipo_map = {z.get('name'): z.get('tipo') for z in zonas if isinstance(z, dict)}
    except Exception:
        current_zona_tipo_map = {}
    
    # Iterar sobre TODAS las horas (7 a 22, inclusive)
    for hora in range(7, 23):
        asignar_empleados_por_hora(empleados_hoy, zonas, mapeo_habilidades, cuadrante, hora)
    
    # POST-PROCESAMIENTO: Desdoblar Cajas -> Cajas 2
    # Se realiza al final para no afectar la lógica de continuidad (que busca 'Cajas')
    last_cajas2 = None
    for h in range(7, 23):
        cajeros = [e for (e, hr), z in cuadrante.items() if hr == h and z == 'Cajas']
        if len(cajeros) >= 2:
            # Si el anterior Cajas 2 sigue en Cajas, mantenerlo. Si no, azar.
            if last_cajas2 in cajeros:
                elegido = last_cajas2
            else:
                elegido = random.choice(cajeros)
            
            cuadrante[(elegido, h)] = 'Cajas 2'
            last_cajas2 = elegido
        else:
            last_cajas2 = None
    
    return cuadrante

# ============================================================================
# FUNCIONES DE VISUALIZACIÓN DEL CUADRANTE
# ============================================================================

def mostrar_cuadrante_visual(empleados_hoy, cuadrante, zonas):
    """
    Genera una visualización clara del cuadrante generado.
    Muestra:
    1. Tabla de empleados con asignaciones por franja
    2. Estadísticas de rotación COM
    3. Validación de restricciones
    """
    nombre_col = empleados_hoy.columns[0]
    empleados_names = empleados_hoy[nombre_col].unique()
    
    print("\n" + "=" * 100)
    print("CUADRANTE DE ASIGNACIONES GENERADO - MOTOR v2.2")
    print("=" * 100)
    
    # Tabla 1: Resumen por empleado
    print("\n1. ASIGNACIONES POR EMPLEADO (Muestreo de activos):\n")
    print(f"{'Empleado':<40} {'Franja 1':<20} {'Franja 2':<20} {'Franja 3':<10}")
    print("-" * 100)
    
    count = 0
    for emp_name in empleados_names:
        if count >= 8:
            break
        
        # Franja 1 (07-10)
        franja1 = [cuadrante.get((emp_name, h), '-') for h in range(7, 10)]
        franja1_str = f"{franja1[0]} x3h" if all(z == franja1[0] for z in franja1) else "Variado"
        
        # Franja 2 (10-21)
        franja2_zonas = set([cuadrante.get((emp_name, h), '-') for h in range(10, 21) if cuadrante.get((emp_name, h))])
        franja2_str = ", ".join(sorted(franja2_zonas)[:3])  # Primeras 3 zonas
        if len(franja2_zonas) > 3:
            franja2_str += f" +{len(franja2_zonas)-3} más"
        
        # Franja 3 (21-22)
        franja3 = [cuadrante.get((emp_name, h), '-') for h in range(21, 23)]
        franja3_str = f"{franja3[0]}" if franja3[0] else "-"
        
        print(f"{emp_name:<40} {franja1_str:<20} {franja2_str:<20} {franja3_str:<10}")
        count += 1
    
    # Tabla 2: Distribución por zona a la hora punta (hora 15)
    print("\n" + "-" * 100)
    print("\n2. DISTRIBUCION POR ZONA - HORA 15:00 (hora punta):\n")
    
    zonas_dict = {z['name']: [] for z in zonas}
    zonas_dict['OFF'] = []  # Agregar 'OFF' para empleados fuera de turno
    for (emp_name, h), zona in cuadrante.items():
        if h == 15:
            if zona in zonas_dict:
                zonas_dict[zona].append(emp_name)
            else:
                # Si es una zona desconocida, agrégala dinámicamente
                zonas_dict[zona] = [emp_name]
    
    print(f"{'Zona':<20} {'Empleados':<40} {'Total':<5}")
    print("-" * 100)
    for zona_name in sorted(zonas_dict.keys()):
        empleados_zona = zonas_dict[zona_name]
        count_str = str(len(empleados_zona))
        emp_nombres = ", ".join([e[:15] for e in empleados_zona[:2]])
        if len(empleados_zona) > 2:
            emp_nombres += f" +{len(empleados_zona)-2} más"
        print(f"{zona_name:<20} {emp_nombres:<40} {count_str:<5}")
    
    # Tabla 3: Validación de restricciones
    print("\n" + "-" * 100)
    print("\n3. VALIDACION DE RESTRICCIONES:\n")
    
    # Verificar empleados OFF
    empleados_off = set([e[nombre_col] for _, e in empleados_hoy.iterrows() if e.get('Estado')=='OFF'])
    off_asignados = len([emp for (emp, h), z in cuadrante.items() if emp in empleados_off])
    
    print(f"OFF no asignados:              {'CUMPLE' if off_asignados == 0 else 'ERROR':.<30} ({off_asignados} asignaciones OFF)")
    
    # Verificar rotacion COM
    violaciones = 0
    for emp_name in empleados_names:
        zonas_com_dict = {}
        for (e, h), z in cuadrante.items():
            if e == emp_name and 'COM' in z.upper() and z != 'COM-PS':
                zonas_com_dict[z] = zonas_com_dict.get(z, 0) + 1
        
        for zona_com, reps in zonas_com_dict.items():
            if reps > 2:
                violaciones += 1
    
    print(f"Max 2 repeticiones COM:        {'CUMPLE' if violaciones == 0 else 'ERROR':.<30} ({violaciones} violaciones)")
    
    # Verificar COM-PS sin limite
    max_com_ps = max([sum(1 for (e, h), z in cuadrante.items() if e == emp and z == 'COM-PS') 
                       for emp in empleados_names], default=0)
    
    print(f"COM-PS sin restriccion:        {'CUMPLE' if max_com_ps > 0 else 'SIN USE':.<30} (max {max_com_ps} horas)")
    # Mostrar violaciones de continuidad registradas por la opción B
    try:
        print(f"Continuity violations (operativas): {len(continuity_violations)}")
    except Exception:
        print("Continuity violations (operativas): 0")
    
    # Mostrar rechazos por Regla de Estancia Mínima (look-ahead)
    print(f"\nEstancia Mínima (look-ahead):  {'CUMPLE' if len(minimum_stay_rejections) == 0 else 'VALIDAR':.<30} ({len(minimum_stay_rejections)} rechazos)")
    if minimum_stay_rejections and len(minimum_stay_rejections) <= 10:
        print("\n  Rechazos por Estancia Mínima:")
        for rej in minimum_stay_rejections[:10]:
            print(f"    - {rej['empleado']:<20} a {rej['zona']:<15} en hora {rej['hora']:>2}:00 (turno restante: {rej['turno_restante']}h)")
    
    # Tabla 4: Estadísticas finales
    print("\n" + "-" * 100)
    print("\n4. ESTADISTICAS FINALES:\n")
    
    activos = len([e for _, e in empleados_hoy.iterrows() if e.get('Estado')=='ON'])
    off = len([e for _, e in empleados_hoy.iterrows() if e.get('Estado')=='OFF'])
    empleados_en_com = len(set([e for (e, h), z in cuadrante.items() if 'COM' in z.upper()]))
    
    f1 = len([(e,h,z) for (e,h),z in cuadrante.items() if 7 <= h < 10])
    f2 = len([(e,h,z) for (e,h),z in cuadrante.items() if 10 <= h < 21])
    f3 = len([(e,h,z) for (e,h),z in cuadrante.items() if 21 <= h < 23])
    
    print(f"Total asignaciones:            {len(cuadrante)}")
    print(f"Empleados ACTIVOS (ON):        {activos}")
    print(f"Empleados OFF (excluidos):     {off}")
    print(f"Empleados en zonas COM:        {empleados_en_com}")
    print(f"Horas cubiertas:               07:00-22:00 (16 horas)")
    print(f"\nDesglose por franja:")
    print(f"  - Franja 1 (07:00-10:00):    {f1} asignaciones")
    print(f"  - Franja 2 (10:00-21:00):    {f2} asignaciones")
    print(f"  - Franja 3 (21:00-22:00):    {f3} asignaciones")
    
    print("\n" + "=" * 100)
    print("FIN DE VISUALIZACION")
    print("=" * 100 + "\n")


if __name__ == '__main__':
    # Prueba: cargar datos y generar cuadrante
    from lector_datos import cargar_todo
    
    datos = cargar_todo()
    if datos:
        empleados = datos['empleados_hoy']
        zonas = datos['zonas']
        mapeo = datos['mapeo_habilidades']
        
        print("\n" + "=" * 100)
        print("GENERANDO CUADRANTE - MOTOR DE LOGICA v2.2")
        print("=" * 100)
        
        cuadrante = asignar(empleados, zonas, mapeo)
        
        # Mostrar visualización
        mostrar_cuadrante_visual(empleados, cuadrante, zonas)
        
        print(f"Cuadrante disponible: {len(cuadrante)} asignaciones (empleado, hora) -> zona")
        print("Listo para exportar a Zonning_Hoy.xlsx")
