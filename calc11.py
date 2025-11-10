import pandas as pd
import re
import os
from math import ceil
from difflib import SequenceMatcher

# Mandamos a llamar el archivo de Excel desde la ruta general
try:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    BASE_DIR = os.getcwd()
RUTA_EXCEL = os.path.join(BASE_DIR, "DuracionBateriasAG.xlsx")

# Normalización mejorada de las palabras del usuario
def _norm(s: str) -> str:
    s = (s or "").strip().lower()
    rep = {"á":"a","é":"e","í":"i","ó":"o","ú":"u","ü":"u","ñ":"n"}
    for k,v in rep.items():
        s = s.replace(k,v)
    return s

def _try_float(x):
    try:
        return float(str(x).replace(",", ".").strip())
    except:
        return 0

# Normalización avanzada para búsqueda inteligente
def _norm_avanzada(s: str) -> str:
    s = (s or "").strip().lower()
    rep = {"á":"a","é":"e","í":"i","ó":"o","ú":"u","ü":"u","ñ":"n"}
    for k,v in rep.items():
        s = s.replace(k,v)
    
    # Eliminar palabras de conexión comunes
    palabras_conexion = {
        'de', 'del', 'la', 'el', 'y', 'en', 'a', 'para', 'por', 'con', 'sin', 
        'sobre', 'bajo', 'entre', 'hacia', 'desde', 'hasta', 'mediante', 'según',
        'como', 'que', 'cuando', 'donde', 'cual', 'quien', 'cuyo', 'cuyas', 'cuyos',
        'unas', 'unos', 'una', 'un', 'lo', 'los', 'las', 'al', 'se', 'su', 'sus',
        'este', 'esta', 'estos', 'estas', 'ese', 'esa', 'esos', 'esas', 'aquel',
        'aquella', 'aquellos', 'aquellas', 'otro', 'otra', 'otros', 'otras',
        'mismo', 'misma', 'mismos', 'mismas', 'todo', 'toda', 'todos', 'todas',
        'cada', 'cualquier', 'cualesquiera', 'varios', 'varias', 'ambos', 'ambas',
        'etc', 'etcétera', 'entre otros', 'entre otras', 'para que', 'de la', 'de los',
        'de las', 'en la', 'en el', 'a la', 'al', 'del', 'y las', 'y los', 'y la', 'y el'
    }
    
    # Eliminar caracteres especiales y dividir en palabras
    s = re.sub(r'[^\w\s]', ' ', s)
    palabras = re.findall(r'\b[a-z0-9]+\b', s)
    
    # Filtrar palabras de conexión y palabras muy cortas sin significado
    palabras_filtradas = [p for p in palabras if p not in palabras_conexion and len(p) > 2]
    
    return ' '.join(palabras_filtradas)

# Función para calcular similitud entre cadenas
def _calcular_similitud(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()

# Función para buscar coincidencias con umbral de similitud
# ... en calc10.py, mejorar la función _buscar_coincidencias:

def _buscar_coincidencias(texto: str, busqueda: str, umbral=0.7) -> bool:
    if not texto or not busqueda:
        return False
    
    texto_norm = _norm_avanzada(texto)
    busqueda_norm = _norm_avanzada(busqueda)
    
    if not texto_norm or not busqueda_norm:
        return False
    
    # **MEJORA: Dividir ambos textos en términos individuales**
    def dividir_terminos(texto):
        # Usar los mismos separadores que en la API
        separadores = [',', ';', '/', '|', ' y ', ' e ']
        texto_para_dividir = texto
        for sep in separadores:
            texto_para_dividir = texto_para_dividir.replace(sep, ',')
        return [t.strip() for t in texto_para_dividir.split(',') if t.strip()]
    
    terminos_texto = dividir_terminos(texto_norm)
    terminos_busqueda = dividir_terminos(busqueda_norm)
    
    # Buscar si algún término de búsqueda coincide con algún término del texto
    for termino_b in terminos_busqueda:
        for termino_t in terminos_texto:
            # Si el término de búsqueda está contenido en el término del texto
            if termino_b in termino_t:
                return True
            # Calcular similitud entre términos individuales
            similitud = _calcular_similitud(termino_b, termino_t)
            if similitud >= umbral:
                return True
    
    # También verificar coincidencia completa por si acaso
    if busqueda_norm in texto_norm:
        return True
    
    similitud_completa = _calcular_similitud(texto_norm, busqueda_norm)
    return similitud_completa >= umbral
    
# Se carga el catálogo de baterías desde la ruta del excel, se lee de la hoja Baterias
def cargar_catalogo_baterias(ruta_excel, hoja="Baterias"):
    try:
        df = pd.read_excel(ruta_excel, sheet_name=hoja, dtype=str)
    except Exception as e:
        print(f"[ERROR] No se pudo leer el archivo '{ruta_excel}': {e}")
        return pd.DataFrame()

    # Normalizar los nombres de las columnas
    df.columns = (
        df.columns.str.strip()
        .str.lower()
        .str.replace(r"[\s\-/]+","_",regex=True)
        .str.replace(r"[()]","",regex=True)
    )

    # Se convierten los valores de las columnas a numéricos
    for col in ['voltaje_v','corriente_ah','capacidad_bateria_wh']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(r"[^0-9.\-]","",regex=True), errors='coerce')

    # Eliminar filas donde todos los valores numéricos importantes son NaN
    columnas_numericas = [c for c in ['voltaje_v','corriente_ah','capacidad_bateria_wh'] if c in df.columns]
    if columnas_numericas:
        df = df.dropna(subset=columnas_numericas, how='all')

    return df

# Función principal de cálculo - VERSIÓN MEJORADA PARA ARREGLOS
def calcular_baterias(cat: pd.DataFrame, voltaje=0, corriente=0, capacidad=0,
                      tipo_bateria="", aplicacion="", autonomia_horas=0, potencia_carga=0,
                      permitir_arreglos=False, umbral_similitud=0.6):
    datos = cat.copy()
    if datos.empty:
        return pd.DataFrame()

    # Crear columna de capacidad si no existe
    if 'capacidad_bateria_wh' not in datos.columns and 'voltaje_v' in datos.columns and 'corriente_ah' in datos.columns:
        datos['capacidad_bateria_wh'] = datos['voltaje_v'] * datos['corriente_ah']

    # Filtro por tipo con normalización mejorada
    if tipo_bateria and 'tipo' in datos.columns:
        tipo_busqueda = _norm(tipo_bateria)
        datos['tipo_norm'] = datos['tipo'].astype(str).apply(_norm)
        datos = datos[datos['tipo_norm'] == tipo_busqueda]

    # Filtro por aplicación con búsqueda inteligente - VERSIÓN MÁS ROBUSTA
    if aplicacion and aplicacion.strip():
        # Buscar en diferentes columnas que puedan contener información de aplicación
        columnas_aplicacion = ['uso', 'aplicacion', 'aplicaciones']
        columna_encontrada = None
        
        for col in columnas_aplicacion:
            if col in datos.columns:
                columna_encontrada = col
                break
        
        if columna_encontrada:
            def aplicar_filtro_aplicacion(fila):
                uso_valor = fila[columna_encontrada] if pd.notna(fila[columna_encontrada]) else ''
                return _buscar_coincidencias(str(uso_valor), aplicacion, umbral=umbral_similitud)
            
            mask = datos.apply(aplicar_filtro_aplicacion, axis=1)
            datos = datos[mask]

    # Calcular capacidad requerida
    capacidad_requerida = capacidad
    if autonomia_horas > 0 and potencia_carga > 0:
        capacidad_requerida = autonomia_horas * potencia_carga
    elif capacidad == 0 and voltaje > 0 and corriente > 0:
        capacidad_requerida = voltaje * corriente

    # Margen que van a tener los filtros numéricos
    parametros_numericos = sum(1 for x in [voltaje, corriente, capacidad_requerida] if x > 0)
    margen = 0.5 if parametros_numericos <= 1 else 0.3

    # --- LÓGICA MEJORADA PARA ARREGLOS ---
    resultados_finales = []
    
    if permitir_arreglos:
        print(f"🔧 Modo arreglos ACTIVADO - Buscando: {voltaje}V, {corriente}A, {capacidad_requerida}Wh")
        
        # Procesar cada batería individual y generar arreglos posibles
        for _, bat in datos.iterrows():
            v_individual = bat['voltaje_v'] if 'voltaje_v' in bat and pd.notna(bat['voltaje_v']) else 0
            a_individual = bat['corriente_ah'] if 'corriente_ah' in bat and pd.notna(bat['corriente_ah']) else 0
            
            if v_individual <= 0 or a_individual <= 0:
                continue
            
            # Calcular configuraciones posibles - LÓGICA MEJORADA
            n_serie = 1
            n_paralelo = 1
            
            # Si se especificó voltaje, calcular serie necesaria
            if voltaje > 0 and v_individual > 0:
                n_serie = max(1, ceil(voltaje / v_individual))
            
            # Si se especificó corriente, calcular paralelo necesario  
            if corriente > 0 and a_individual > 0:
                n_paralelo = max(1, ceil(corriente / a_individual))
            
            # Calcular valores totales del arreglo
            voltaje_total = v_individual * n_serie
            corriente_total = a_individual * n_paralelo
            capacidad_total = voltaje_total * corriente_total
            
            # Crear copia de la batería con información del arreglo
            bateria_con_arreglo = bat.copy()
            bateria_con_arreglo['n_serie'] = n_serie
            bateria_con_arreglo['n_paralelo'] = n_paralelo
            bateria_con_arreglo['voltaje_total_v'] = voltaje_total
            bateria_con_arreglo['corriente_total_ah'] = corriente_total
            bateria_con_arreglo['capacidad_total_wh'] = capacidad_total
            bateria_con_arreglo['es_arreglo'] = n_serie > 1 or n_paralelo > 1
            
            resultados_finales.append(bateria_con_arreglo)
            
        print(f"🔧 Generados {len(resultados_finales)} arreglos posibles")
        
    else:
        print(f"🔧 Modo arreglos DESACTIVADO")
        # Modo sin arreglos - solo baterías individuales
        for _, bat in datos.iterrows():
            bateria_individual = bat.copy()
            bateria_individual['n_serie'] = 1
            bateria_individual['n_paralelo'] = 1
            bateria_individual['voltaje_total_v'] = bateria_individual.get('voltaje_v', 0)
            bateria_individual['corriente_total_ah'] = bateria_individual.get('corriente_ah', 0)
            bateria_individual['capacidad_total_wh'] = bateria_individual.get('voltaje_v', 0) * bateria_individual.get('corriente_ah', 0)
            bateria_individual['es_arreglo'] = False
            resultados_finales.append(bateria_individual)
    
    if not resultados_finales:
        return pd.DataFrame()
        
    datos = pd.DataFrame(resultados_finales)

    # Filtrar por rangos usando los valores totales del arreglo - LÓGICA MEJORADA
    datos_filtrados = datos.copy()
    
    if voltaje > 0 and 'voltaje_total_v' in datos_filtrados.columns:
        rango_min_voltaje = voltaje * (1 - margen)
        rango_max_voltaje = voltaje * (1 + margen)
        datos_filtrados = datos_filtrados[
            datos_filtrados['voltaje_total_v'].between(rango_min_voltaje, rango_max_voltaje)
        ]
        print(f"🔧 Filtro voltaje: {rango_min_voltaje:.1f}V - {rango_max_voltaje:.1f}V, quedan {len(datos_filtrados)}")
    
    if corriente > 0 and 'corriente_total_ah' in datos_filtrados.columns:
        rango_min_corriente = corriente * (1 - margen)
        rango_max_corriente = corriente * (1 + margen)
        datos_filtrados = datos_filtrados[
            datos_filtrados['corriente_total_ah'].between(rango_min_corriente, rango_max_corriente)
        ]
        print(f"🔧 Filtro corriente: {rango_min_corriente:.1f}A - {rango_max_corriente:.1f}A, quedan {len(datos_filtrados)}")
    
    if capacidad_requerida > 0 and 'capacidad_total_wh' in datos_filtrados.columns:
        rango_min_capacidad = capacidad_requerida * (1 - margen)
        rango_max_capacidad = capacidad_requerida * (1 + margen)
        datos_filtrados = datos_filtrados[
            datos_filtrados['capacidad_total_wh'].between(rango_min_capacidad, rango_max_capacidad)
        ]
        print(f"🔧 Filtro capacidad: {rango_min_capacidad:.1f}Wh - {rango_max_capacidad:.1f}Wh, quedan {len(datos_filtrados)}")

    if datos_filtrados.empty:
        print("🔧 No hay resultados después del filtrado")
        return pd.DataFrame()

    # Ordenamiento por relevancia
    if 'capacidad_total_wh' in datos_filtrados.columns:
        datos_filtrados['diff_capacidad'] = abs(datos_filtrados['capacidad_total_wh'] - capacidad_requerida)
    if 'voltaje_total_v' in datos_filtrados.columns:
        datos_filtrados['diff_voltaje'] = abs(datos_filtrados['voltaje_total_v'] - voltaje)
    
    columnas_orden = []
    if 'diff_capacidad' in datos_filtrados.columns:
        columnas_orden.append('diff_capacidad')
    if 'diff_voltaje' in datos_filtrados.columns:
        columnas_orden.append('diff_voltaje')
    
    if columnas_orden:
        res = datos_filtrados.sort_values(by=columnas_orden)
    else:
        res = datos_filtrados

    # Limpiar columnas auxiliares
    res = res.drop(columns=['diff_capacidad','diff_voltaje', 'uso_norm', 'tipo_norm'], errors='ignore')
    
    # Calcular capacidad individual
    if 'voltaje_v' in res.columns and 'corriente_ah' in res.columns:
        res['capacidad_individual_wh'] = res['voltaje_v'] * res['corriente_ah']

    # Reordenar columnas
    cols_finales = ['tipo','no_de_parte','voltaje_v','corriente_ah','capacidad_individual_wh',
                   'n_serie','n_paralelo','voltaje_total_v','corriente_total_ah','capacidad_total_wh','es_arreglo']
    cols_existentes = [c for c in cols_finales if c in res.columns]
    res = res[cols_existentes + [c for c in res.columns if c not in cols_existentes]]

    print(f"🔧 Resultados finales: {len(res)} baterías/arreglos")
    return res.reset_index(drop=True)

def main_baterias():
    print("=== CALCULADORA DE BATERÍAS ===\n")
    print("Si no sabe algún dato, déjelo en blanco y presione Enter.")
    print("Puede buscar solo con un parámetro (ej: solo 12V, solo 100Ah, solo 500Wh)\n")
    
    # Tipos de batería comunes
    print("Tipos de batería comunes: Ácido Plomo, Litio, Lipo, LiFEPO4, Alcalinas, Niquel, Oxido de Plata\n")
    
    # Entrada de usuario
    tipo_bateria = input("Tipo de batería (deje en blanco para cualquier tipo): ").strip()
    aplicacion = input("Aplicación / Uso (ej. 'UPS', 'solar', 'drones', 'vehículos eléctricos'): ").strip()
    voltaje = _try_float(input("Voltaje requerido (V): "))
    corriente = _try_float(input("Corriente/Capacidad (Ah): "))
    capacidad = _try_float(input("Capacidad de energía (Wh): "))
    
    print("\n--- Opciones de autonomía (opcional) ---")
    print("Si conoce el consumo y tiempo deseado, podemos calcular la capacidad necesaria.")
    autonomia_horas = _try_float(input("Autonomía deseada (horas): "))
    potencia_carga = _try_float(input("Potencia de la carga (W): "))
    
    permitir_arreglos = input("\n¿Desea permitir arreglos en serie/paralelo? (s/n): ").strip().lower() == 's'
    
    # Validación de entrada mínima
    parametros_numericos = sum(1 for x in [voltaje,corriente,capacidad,autonomia_horas,potencia_carga] if x>0)
    if parametros_numericos == 0 and not tipo_bateria and not aplicacion:
        print("\n[ERROR] Debe ingresar al menos un criterio de búsqueda.")
        return

    # Cargar catálogo
    cat = cargar_catalogo_baterias(RUTA_EXCEL)
    if cat.empty:
        print("[ERROR] No se pudieron cargar datos del catálogo. Revise la ruta o el formato del Excel.")
        return

    # Cargar los parámetros del cálculo
    res = calcular_baterias(
        cat,
        voltaje=voltaje,
        corriente=corriente,
        capacidad=capacidad,
        tipo_bateria=tipo_bateria,
        aplicacion=aplicacion,
        autonomia_horas=autonomia_horas,
        potencia_carga=potencia_carga,
        permitir_arreglos=permitir_arreglos
    )

    # Mostrar resultados
    if res.empty:
        print("\nNo se encontraron baterías que coincidan con los criterios especificados.")
        print("Sugerencias:")
        print("- Amplíe los rangos de búsqueda")
        print("- Use menos criterios de filtrado")
        print("- Verifique los tipos de batería y aplicaciones")
        return

    print(f"\n=== BATERÍAS RECOMENDADAS ({len(res)} encontradas) ===")
    
    # Columnas a mostrar en consola
    columnas_mostrar = ['tipo','no_de_parte','voltaje_v','corriente_ah','capacidad_individual_wh','n_serie','n_paralelo','voltaje_total_v','capacidad_total_wh']
    columnas_mostrar = [c for c in columnas_mostrar if c in res.columns]
    print(res[columnas_mostrar].head(20).to_string(index=False))  # máximo 20 resultados

    # Guardar resultados en Excel
    salida_archivo = "recomendaciones_baterias.xlsx"
    try:
        res.to_excel(salida_archivo, index=False, sheet_name="Resultados")
        print(f"\nArchivo de resultados guardado como: {salida_archivo}")
    except Exception as e:
        print(f"[ERROR] No se pudo guardar el archivo Excel: {e}")

if __name__ == "__main__":
    main_baterias()