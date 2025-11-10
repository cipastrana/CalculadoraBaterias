from flask import Flask, render_template, request, jsonify
import pandas as pd
import os
import sys
import logging
import re
from difflib import SequenceMatcher
from math import ceil

# Configurar logging (para mostrar peticiones)
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Configuración de rutas
try:
    current_dir = os.path.dirname(os.path.abspath(__file__))
    if current_dir not in sys.path:
        sys.path.append(current_dir)
    
    # Definir RUTA_EXCEL aquí para evitar errores de importación
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    RUTA_EXCEL = os.path.join(BASE_DIR, "DuracionBateriasAG.xlsx")
    
except Exception as e:
    logger.error(f"Error configurando rutas: {e}")
    BASE_DIR = os.getcwd()
    RUTA_EXCEL = os.path.join(BASE_DIR, "DuracionBateriasAG.xlsx")

# Funciones auxiliares para evitar dependencias de calc11.py
def _try_float(x):
    try:
        return float(str(x).replace(",", ".").strip())
    except:
        return 0

def _norm(s: str) -> str:
    s = (s or "").strip().lower()
    rep = {"á":"a","é":"e","í":"i","ó":"o","ú":"u","ü":"u","ñ":"n"}
    for k,v in rep.items():
        s = s.replace(k,v)
    return s

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

def _calcular_similitud(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()

def _buscar_coincidencias(texto: str, busqueda: str, umbral=0.7) -> bool:
    if not texto or not busqueda:
        return False
    
    texto_norm = _norm_avanzada(texto)
    busqueda_norm = _norm_avanzada(busqueda)
    
    if not texto_norm or not busqueda_norm:
        return False
    
    def dividir_terminos(texto):
        separadores = [',', ';', '/', '|', ' y ', ' e ']
        texto_para_dividir = texto
        for sep in separadores:
            texto_para_dividir = texto_para_dividir.replace(sep, ',')
        return [t.strip() for t in texto_para_dividir.split(',') if t.strip()]
    
    terminos_texto = dividir_terminos(texto_norm)
    terminos_busqueda = dividir_terminos(busqueda_norm)
    
    for termino_b in terminos_busqueda:
        for termino_t in terminos_texto:
            if termino_b in termino_t:
                return True
            similitud = _calcular_similitud(termino_b, termino_t)
            if similitud >= umbral:
                return True
    
    if busqueda_norm in texto_norm:
        return True
    
    similitud_completa = _calcular_similitud(texto_norm, busqueda_norm)
    return similitud_completa >= umbral

# Función para cargar catálogo
def cargar_catalogo_baterias(ruta_excel, hoja="Baterias"):
    try:
        df = pd.read_excel(ruta_excel, sheet_name=hoja, dtype=str)
    except Exception as e:
        logger.error(f"[ERROR] No se pudo leer el archivo '{ruta_excel}': {e}")
        return pd.DataFrame()

    # Normalizar los nombres de las columnas
    df.columns = (
        df.columns.str.strip()
        .str.lower()
        .str.replace(r"[\s\-/]+","_",regex=True)
        .str.replace(r"[()]","",regex=True)
    )

    # Convertir valores numéricos
    for col in ['voltaje_v','corriente_ah','capacidad_bateria_wh']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(r"[^0-9.\-]","",regex=True), errors='coerce')

    # Eliminar filas donde todos los valores numéricos importantes son NaN
    columnas_numericas = [c for c in ['voltaje_v','corriente_ah','capacidad_bateria_wh'] if c in df.columns]
    if columnas_numericas:
        df = df.dropna(subset=columnas_numericas, how='all')

    return df

# Función de cálculo de baterías (versión simplificada y robusta)
def calcular_baterias(cat: pd.DataFrame, voltaje=0, corriente=0, capacidad=0,
                      tipo_bateria="", aplicacion="", autonomia_horas=0, potencia_carga=0,
                      permitir_arreglos=False, umbral_similitud=0.6):
    
    logger.info(f"🔧 Iniciando cálculo: voltaje={voltaje}V, corriente={corriente}A, capacidad={capacidad}Wh, arreglos={permitir_arreglos}")
    
    datos = cat.copy()
    if datos.empty:
        logger.warning("❌ Catálogo vacío")
        return pd.DataFrame()

    # Crear columna de capacidad si no existe
    if 'capacidad_bateria_wh' not in datos.columns and 'voltaje_v' in datos.columns and 'corriente_ah' in datos.columns:
        datos['capacidad_bateria_wh'] = datos['voltaje_v'] * datos['corriente_ah']

    # Filtro por tipo
    if tipo_bateria and 'tipo' in datos.columns:
        tipo_busqueda = _norm(tipo_bateria)
        datos['tipo_norm'] = datos['tipo'].astype(str).apply(_norm)
        datos = datos[datos['tipo_norm'] == tipo_busqueda]
        logger.info(f"🔧 Filtrado por tipo '{tipo_bateria}': {len(datos)} baterías")

    # Filtro por aplicación
    if aplicacion and aplicacion.strip():
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
            logger.info(f"🔧 Filtrado por aplicación '{aplicacion}': {len(datos)} baterías")

    # Calcular capacidad requerida
    capacidad_requerida = capacidad
    if autonomia_horas > 0 and potencia_carga > 0:
        capacidad_requerida = autonomia_horas * potencia_carga
    elif capacidad == 0 and voltaje > 0 and corriente > 0:
        capacidad_requerida = voltaje * corriente

    logger.info(f"🔧 Capacidad requerida: {capacidad_requerida}Wh")

    # Margen de búsqueda
    parametros_numericos = sum(1 for x in [voltaje, corriente, capacidad_requerida] if x > 0)
    margen = 0.5 if parametros_numericos <= 1 else 0.3

    # LÓGICA DE ARREGLOS
    resultados_finales = []
    
    if permitir_arreglos:
        logger.info("🔧 Modo arreglos ACTIVADO")
        
        for _, bat in datos.iterrows():
            v_individual = bat['voltaje_v'] if 'voltaje_v' in bat and pd.notna(bat['voltaje_v']) else 0
            a_individual = bat['corriente_ah'] if 'corriente_ah' in bat and pd.notna(bat['corriente_ah']) else 0
            
            if v_individual <= 0 or a_individual <= 0:
                continue
            
            # Calcular configuraciones
            n_serie = 1
            n_paralelo = 1
            
            if voltaje > 0 and v_individual > 0:
                n_serie = max(1, ceil(voltaje / v_individual))
            
            if corriente > 0 and a_individual > 0:
                n_paralelo = max(1, ceil(corriente / a_individual))
            
            # Calcular valores totales
            voltaje_total = v_individual * n_serie
            corriente_total = a_individual * n_paralelo
            capacidad_total = voltaje_total * corriente_total
            
            # Crear batería con arreglo
            bateria_con_arreglo = bat.copy()
            bateria_con_arreglo['n_serie'] = n_serie
            bateria_con_arreglo['n_paralelo'] = n_paralelo
            bateria_con_arreglo['voltaje_total_v'] = voltaje_total
            bateria_con_arreglo['corriente_total_ah'] = corriente_total
            bateria_con_arreglo['capacidad_total_wh'] = capacidad_total
            bateria_con_arreglo['es_arreglo'] = n_serie > 1 or n_paralelo > 1
            
            resultados_finales.append(bateria_con_arreglo)
            
        logger.info(f"🔧 Generados {len(resultados_finales)} arreglos")
        
    else:
        logger.info("🔧 Modo arreglos DESACTIVADO")
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
        logger.warning("❌ No hay resultados después de procesar arreglos")
        return pd.DataFrame()
        
    datos = pd.DataFrame(resultados_finales)

    # Filtrar por rangos
    datos_filtrados = datos.copy()
    
    if voltaje > 0 and 'voltaje_total_v' in datos_filtrados.columns:
        rango_min_voltaje = voltaje * (1 - margen)
        rango_max_voltaje = voltaje * (1 + margen)
        datos_filtrados = datos_filtrados[
            datos_filtrados['voltaje_total_v'].between(rango_min_voltaje, rango_max_voltaje)
        ]
        logger.info(f"🔧 Filtro voltaje: {len(datos_filtrados)} después de filtrar")
    
    if corriente > 0 and 'corriente_total_ah' in datos_filtrados.columns:
        rango_min_corriente = corriente * (1 - margen)
        rango_max_corriente = corriente * (1 + margen)
        datos_filtrados = datos_filtrados[
            datos_filtrados['corriente_total_ah'].between(rango_min_corriente, rango_max_corriente)
        ]
        logger.info(f"🔧 Filtro corriente: {len(datos_filtrados)} después de filtrar")
    
    if capacidad_requerida > 0 and 'capacidad_total_wh' in datos_filtrados.columns:
        rango_min_capacidad = capacidad_requerida * (1 - margen)
        rango_max_capacidad = capacidad_requerida * (1 + margen)
        datos_filtrados = datos_filtrados[
            datos_filtrados['capacidad_total_wh'].between(rango_min_capacidad, rango_max_capacidad)
        ]
        logger.info(f"🔧 Filtro capacidad: {len(datos_filtrados)} después de filtrar")

    if datos_filtrados.empty:
        logger.warning("❌ No hay resultados después del filtrado")
        return pd.DataFrame()

    # Ordenamiento
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

    logger.info(f"✅ Resultados finales: {len(res)} baterías/arreglos")
    return res.reset_index(drop=True)

app = Flask(__name__)

# Página principal
@app.route('/')
def index():
    return render_template('index3.html')

# Endpoints de la API
@app.route('/buscar', methods=['POST'])
def buscar_baterias():
    try:
        data = request.get_json() or {}
        logger.info(f"📥 Datos recibidos: {data}")

        # Obtener datos del formulario
        tipo = data.get('tipo', '').strip()
        aplicacion = data.get('aplicacion', '').strip()
        voltaje_val = _try_float(data.get('voltaje', 0))
        corriente_val = _try_float(data.get('corriente', 0))
        capacidad_wh_val = _try_float(data.get('capacidad_wh', 0))
        autonomia_horas_val = _try_float(data.get('autonomia_horas', 0))
        potencia_carga_val = _try_float(data.get('potencia_carga', 0))
        permitir_arreglos = bool(data.get('permitir_arreglos', False))

        logger.info(f"🔍 Búsqueda: {tipo}, {aplicacion}, {voltaje_val}V, {corriente_val}A, arreglos={permitir_arreglos}")

        # Cargar catálogo
        cat = cargar_catalogo_baterias(RUTA_EXCEL)
        if cat.empty:
            return jsonify({'success': False, 'error': 'No se pudo cargar el catálogo de baterías'})

        # Calcular baterías
        res = calcular_baterias(
            cat=cat,
            voltaje=voltaje_val,
            corriente=corriente_val,
            capacidad=capacidad_wh_val,
            tipo_bateria=tipo,
            aplicacion=aplicacion,
            autonomia_horas=autonomia_horas_val,
            potencia_carga=potencia_carga_val,
            permitir_arreglos=permitir_arreglos,
            umbral_similitud=0.6
        )

        if res.empty:
            return jsonify({'success': True, 'resultados': [], 'total': 0})

        # Construir respuesta
        resultados = []
        for _, bateria in res.iterrows():
            numero_parte = (
                bateria.get('no_de_parte') or 
                bateria.get('no._de_parte') or 
                bateria.get('numero_parte') or
                'N/A'
            )
            
            aplicaciones = (
                bateria.get('uso') or 
                bateria.get('aplicacion') or 
                bateria.get('aplicaciones') or
                'N/A'
            )

            resultados.append({
                'tipo': bateria.get('tipo', 'N/A'),
                'numero_parte': numero_parte,
                'voltaje': bateria.get('voltaje_v', 0),
                'corriente': bateria.get('corriente_ah', 0),
                'capacidad_wh': bateria.get('capacidad_individual_wh', 0),
                'aplicaciones': aplicaciones,
                'n_serie': int(bateria.get('n_serie', 1)),
                'n_paralelo': int(bateria.get('n_paralelo', 1)),
                'voltaje_total': bateria.get('voltaje_total_v', 0),
                'corriente_total': bateria.get('corriente_total_ah', 0),
                'capacidad_total': bateria.get('capacidad_total_wh', 0),
                'es_arreglo': bool(bateria.get('es_arreglo', False))
            })

        capacidad_calculada = autonomia_horas_val * potencia_carga_val if autonomia_horas_val and potencia_carga_val else None

        return jsonify({
            'success': True,
            'resultados': resultados,
            'total': len(resultados),
            'capacidad_calculada': capacidad_calculada,
            'permitir_arreglos': permitir_arreglos
        })

    except Exception as e:
        logger.error(f"❌ Error en búsqueda: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': f'Error interno del servidor: {str(e)}'
        })

# Endpoints auxiliares (tipos, aplicaciones, voltajes)
@app.route('/tipos-baterias')
def obtener_tipos_baterias():
    try:
        cat = cargar_catalogo_baterias(RUTA_EXCEL)
        if 'tipo' in cat.columns:
            tipos = cat['tipo'].dropna().apply(_norm).unique().tolist()
            tipos = sorted([t for t in tipos if t and t.strip()])
        else:
            tipos = []
        return jsonify({'success': True, 'tipos': tipos})
    except Exception as e:
        logger.error(f"Error obteniendo tipos: {e}")
        return jsonify({'success': False, 'tipos': []})

@app.route('/aplicaciones')
def obtener_aplicaciones():
    try:
        cat = cargar_catalogo_baterias(RUTA_EXCEL)
        aplicaciones_set = set()
        
        columnas_aplicaciones = ['uso', 'aplicacion', 'aplicaciones']
        columna_encontrada = None
        
        for col in columnas_aplicaciones:
            if col in cat.columns:
                columna_encontrada = col
                break
        
        if columna_encontrada:
            for uso in cat[columna_encontrada]:
                if pd.notna(uso):
                    uso_str = str(uso).strip()
                    separadores = [',', ';', '/', '|', ' y ', ' e ']
                    for sep in separadores:
                        uso_str = uso_str.replace(sep, ',')
                    terminos = uso_str.split(',')
                    for termino in terminos:
                        termino_limpio = termino.strip()
                        if termino_limpio:
                            termino_normalizado = _norm_avanzada(termino_limpio)
                            if termino_normalizado and len(termino_normalizado) > 2:
                                aplicaciones_set.add(termino_normalizado)
        
        aplicaciones_filtradas = sorted([a for a in aplicaciones_set if a and len(a) > 2])
        return jsonify({'success': True, 'aplicaciones': aplicaciones_filtradas})
    except Exception as e:
        logger.error(f"Error obteniendo aplicaciones: {e}")
        return jsonify({'success': False, 'aplicaciones': []})

@app.route('/aplicaciones-por-tipo')
def obtener_aplicaciones_por_tipo():
    try:
        tipo_bateria = request.args.get('tipo', '').strip()
        if not tipo_bateria:
            return jsonify({'success': False, 'aplicaciones': []})
        
        cat = cargar_catalogo_baterias(RUTA_EXCEL)
        if cat.empty or 'tipo' not in cat.columns:
            return jsonify({'success': False, 'aplicaciones': []})
        
        tipo_normalizado = _norm(tipo_bateria)
        cat['tipo_norm'] = cat['tipo'].astype(str).apply(_norm)
        cat_filtrado = cat[cat['tipo_norm'] == tipo_normalizado]
        
        aplicaciones_set = set()
        columnas_aplicaciones = ['uso', 'aplicacion', 'aplicaciones']
        columna_encontrada = None
        
        for col in columnas_aplicaciones:
            if col in cat_filtrado.columns:
                columna_encontrada = col
                break
        
        if columna_encontrada:
            for uso in cat_filtrado[columna_encontrada]:
                if pd.notna(uso):
                    uso_str = str(uso).strip()
                    separadores = [',', ';', '/', '|', ' y ', ' e ']
                    for sep in separadores:
                        uso_str = uso_str.replace(sep, ',')
                    terminos = uso_str.split(',')
                    for termino in terminos:
                        termino_limpio = termino.strip()
                        if termino_limpio:
                            termino_normalizado = _norm_avanzada(termino_limpio)
                            if termino_normalizado and len(termino_normalizado) > 2:
                                aplicaciones_set.add(termino_normalizado)
        
        aplicaciones_filtradas = sorted([a for a in aplicaciones_set if a and len(a) > 2])
        return jsonify({'success': True, 'aplicaciones': aplicaciones_filtradas})
    except Exception as e:
        logger.error(f"Error obteniendo aplicaciones por tipo: {e}")
        return jsonify({'success': False, 'aplicaciones': []})

@app.route('/voltajes-por-tipo')
def obtener_voltajes_por_tipo():
    try:
        tipo_bateria = request.args.get('tipo', '').strip()
        if not tipo_bateria:
            return jsonify({'success': False, 'voltajes': []})
        
        cat = cargar_catalogo_baterias(RUTA_EXCEL)
        if cat.empty or 'tipo' not in cat.columns or 'voltaje_v' not in cat.columns:
            return jsonify({'success': False, 'voltajes': []})
        
        tipo_normalizado = _norm(tipo_bateria)
        cat['tipo_norm'] = cat['tipo'].astype(str).apply(_norm)
        cat_filtrado = cat[cat['tipo_norm'] == tipo_normalizado]
        
        voltajes = cat_filtrado['voltaje_v'].dropna().unique()
        voltajes = sorted([v for v in voltajes if v is not None and v > 0])
        
        return jsonify({'success': True, 'voltajes': voltajes})
    except Exception as e:
        logger.error(f"Error obteniendo voltajes por tipo: {e}")
        return jsonify({'success': False, 'voltajes': []})

@app.route('/todos-los-voltajes')
def obtener_todos_los_voltajes():
    try:
        cat = cargar_catalogo_baterias(RUTA_EXCEL)
        if cat.empty or 'voltaje_v' not in cat.columns:
            return jsonify({'success': False, 'voltajes': []})
        
        voltajes = cat['voltaje_v'].dropna().unique()
        voltajes = sorted([v for v in voltajes if v is not None and v > 0])
        
        return jsonify({'success': True, 'voltajes': voltajes})
    except Exception as e:
        logger.error(f"Error obteniendo todos los voltajes: {e}")
        return jsonify({'success': False, 'voltajes': []})

@app.route('/debug')
def debug():
    try:
        cat = cargar_catalogo_baterias(RUTA_EXCEL)
        info = {
            'archivo_existe': os.path.exists(RUTA_EXCEL),
            'catalogo_cargado': not cat.empty,
            'total_baterias': len(cat) if not cat.empty else 0,
            'columnas': cat.columns.tolist() if not cat.empty else [],
            'ruta_excel': RUTA_EXCEL
        }
        return jsonify(info)
    except Exception as e:
        return jsonify({'error': str(e)})

if __name__ == '__main__':
    logger.info("🚀 Iniciando servidor Flask")
    app.run(debug=True, host='0.0.0.0', port=5000)