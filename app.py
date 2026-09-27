# --- 1. CONEXIÓN ROBUSTA Y DINÁMICA A POSTGRESQL ---
def obtener_conexion():
    database_url = os.environ.get('DATABASE_URL')
    try:
        if database_url:
            return psycopg2.connect(database_url, sslmode='require')
        else:
            return psycopg2.connect(
                host="localhost",
                database="SISTEMAGUARDIA1",
                user="postgres",
                password="CHOPA",
                connect_timeout=5
            )
    except psycopg2.Error as e:
        print(f"ERROR DE CONEXIÓN: {e}")
        return None

# --- 2. VERIFICACIÓN DE SESIÓN (SEGURIDAD) ---
@app.before_request
def verificar_sesion():
    # Protege todas las rutas excepto login y archivos estáticos
    rutas_libres = ['login', 'static']
    if 'logged_in' not in session and request.endpoint not in rutas_libres:
        return redirect(url_for('login'))

# --- 3. VISTA PRINCIPAL CON FILTROS Y CONTEO ---
@app.route('/')
def index():
    if not session.get('logged_in'): return redirect(url_for('login'))
    
    # 1. CAPTURAR EL FILTRO (Nueva función para los botones)
    filtro = request.args.get('situacion') 
    
    conn = obtener_conexion()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    
    # 2. CONSULTA SQL ROBUSTA CON O SIN FILTRO
    if filtro and filtro != 'TOTAL':
        filtro_clean = filtro.strip().upper()
        
        # Normalizamos los casos con o sin "S" al final
        if filtro_clean in ['DISPONIBLE', 'DISPONIBLES']:
            cur.execute("SELECT ce, dni, apellido_nombre, aula, grado, situacion, total_guardias FROM aspirantes WHERE UPPER(TRIM(situacion)) LIKE 'DISPONIBL%'")
        elif filtro_clean in ['AUTORIZADO', 'AUTORIZADOS']:
            cur.execute("SELECT ce, dni, apellido_nombre, aula, grado, situacion, total_guardias FROM aspirantes WHERE UPPER(TRIM(situacion)) LIKE 'AUTORIZAD%'")
        else:
            cur.execute("SELECT ce, dni, apellido_nombre, aula, grado, situacion, total_guardias FROM aspirantes WHERE UPPER(TRIM(situacion)) = %s", (filtro_clean,))
    else:
        cur.execute("SELECT ce, dni, apellido_nombre, aula, grado, situacion, total_guardias FROM aspirantes")
    
    aspirantes_raw = cur.fetchall()
    
    # 3. MANTENEMOS TU FUNCIÓN DE LIMPIEZA (Función crítica para evitar el .0)
    def limpiar_dato(valor):
        # 1. Manejo de nulos o vacíos
        if valor is None or str(valor).strip().lower() in ['nan', 'none', '']:
            return ""
        
        # 2. Convertimos a texto
        v_str = str(valor).strip()
        
        # 3. Quitamos el .0 (Ojo con los : al final y los 4 espacios de abajo)
        if v_str.endswith('.0'):
            return v_str[:-2]
            
        # 4. Retornamos el valor final
        return v_str
    aspirantes = []
    for a in aspirantes_raw:
        aspirantes.append({
            'ce': limpiar_dato(a['ce']),
            'dni': limpiar_dato(a['dni']),
            'apellido_nombre': a['apellido_nombre'] if a['apellido_nombre'] else "SIN NOMBRE",
            'aula': limpiar_dato(a['aula']),
            'grado': a['grado'] if a['grado'] else "ASP III AÑO",
            'situacion': a['situacion'],
            'total_guardias': a['total_guardias']
        })

    # 4. MANTENEMOS TU ORDENAMIENTO CRÍTICO (Aula y Apellido)
    aspirantes.sort(key=lambda x: (str(x['aula']), x['apellido_nombre']))

    # 5. LÓGICA DE CONTEO ROBUSTA (Acepta variaciones en la BD)
    cur.execute("SELECT UPPER(TRIM(situacion)) AS sit_clean, COUNT(*) AS cantidad FROM aspirantes GROUP BY UPPER(TRIM(situacion))")
    filas_conteo = cur.fetchall()
    
    # Mapeo en un diccionario auxiliar
    conteo_raw = {f['sit_clean']: f['cantidad'] for f in filas_conteo}
    
    cur.execute("SELECT COUNT(*) AS total FROM aspirantes")
    total_res = cur.fetchone()
    
    # Extraemos el total
    cant_total = total_res['total'] if total_res and 'total' in total_res else 0
    
    conteo = {
        'DISPONIBLES': conteo_raw.get('DISPONIBLE', 0) or conteo_raw.get('DISPONIBLES', 0),
        'SSD': conteo_raw.get('SSD', 0),
        'ART': conteo_raw.get('ART', 0),
        'LAO': conteo_raw.get('LAO', 0),
        'AUTORIZADOS': conteo_raw.get('AUTORIZADO', 0) or conteo_raw.get('AUTORIZADOS', 0),
        'TOTAL': total_res['total'] if total_res else 0
    }
    cur.close()
    conn.close()
    return render_template('index.html', aspirantes=aspirantes, conteo=conteo, busqueda='')

@app.route('/importar_excel', methods=['POST'])
def importar_excel():
    # 1. CANDADO DE SEGURIDAD: Solo el Administrador puede cargar datos masivos
    if session.get('rol') != 'admin':
        from flask import flash
        flash("Acceso denegado: Solo el Administrador puede importar archivos Excel.", "danger")
        return redirect(url_for('index'))
        
    file = request.files.get('file')
    if file:
        try:
            import pandas as pd
            df = pd.read_excel(file)
            
            # Normalizamos los encabezados del Excel a mayúsculas y sin espacios
            df.columns = [str(c).strip().upper() for c in df.columns]
            
            conn = obtener_conexion()
            cur = conn.cursor()
            
            for _, row in df.iterrows():
                # --- BUSCADORES FLEXIBLES DE COLUMNAS ---
                
                # Buscador de CE o CC
                ce_detectado = None
                for col in df.columns:
                    if col in ['CE', 'CC', 'CODIGO', 'ESTADISTICO', 'ID']:
                        ce_detectado = str(row[col]).strip()
                        break
                if not ce_detectado or ce_detectado == 'nan': 
                    continue # Si la fila no tiene un identificador, salta a la siguiente
                
                # Buscador de DNI
                dni_detectado = "0"
                for col in df.columns:
                    if 'DNI' in col or 'DOCUMENTO' in col:
                        dni_detectado = str(row[col]).strip()
                        break

                # Buscador de Nombre/Apellido
                nombre_detectado = "SIN NOMBRE"
                for col in df.columns:
                    if "NOMBRE" in col or "APELLIDO" in col or "ASPIRANTE" in col:
                        if pd.notna(row[col]):
                            nombre_detectado = str(row[col]).strip().upper()
                            break

                # Buscador de Aula
                aula_detectada = "SIN AULA"
                for col in df.columns:
                    if "AULA" in col or "CURSO" in col or "SECCION" in col:
                        aula_detectada = str(row[col]).strip()
                        break

                # Buscador de Grado
                grado_detectado = "ASP III AÑO"
                for col in df.columns:
                    if "GRADO" in col or "AÑO" in col:
                        grado_detectado = str(row[col]).strip().upper()
                        break

                # 2. Inserción o actualización en PostgreSQL con los datos limpios
                cur.execute("""
                    INSERT INTO aspirantes (ce, dni, apellido_nombre, aula, grado, situacion, total_guardias)
                    VALUES (%s, %s, %s, %s, %s, 'DISPONIBLE', 0)
                    ON CONFLICT (ce) DO UPDATE SET 
                        apellido_nombre = EXCLUDED.apellido_nombre, 
                        dni = EXCLUDED.dni, 
                        aula = EXCLUDED.aula,
                        grado = EXCLUDED.grado;
                """, (
                    ce_detectado, 
                    dni_detectado, 
                    nombre_detectado, 
                    aula_detectada, 
                    grado_detectado
                ))
            
            conn.commit()
            cur.close()
            conn.close()
            
            from flask import flash
            flash("Excel importado con éxito. Base de datos actualizada.", "success")

        except Exception as e:
            print(f"Error al importar: {e}")
            from flask import flash
            flash(f"Error técnico al procesar el Excel: {e}", "danger")
            
    return redirect(url_for('index'))

@app.route('/eliminar_excel', methods=['POST'])
def eliminar_excel():
    if session.get('rol') == 'admin':
        conn = obtener_conexion()
        cur = conn.cursor()
        cur.execute("DELETE FROM aspirantes")
        conn.commit()
        cur.close()
        conn.close()
    return redirect(url_for('index'))

# --- 4. SISTEMA DE LOGIN ---

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        user = request.form.get('username').strip()
        pw = request.form.get('password').strip()
        
        conn = obtener_conexion()
        cur = conn.cursor()
        cur.execute("SELECT username, rol FROM usuarios WHERE username=%s AND password=%s", (user, pw))
        usuario_db = cur.fetchone()
        cur.close()
        conn.close()

        if usuario_db:
            session.update({'logged_in': True, 'usuario': usuario_db[0], 'rol': usuario_db[1]})
            return redirect(url_for('index'))
        flash("Credenciales incorrectas", "danger")
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# --- 6. GESTIÓN DE USUARIOS DEL SISTEMA ---
@app.route('/gestion_usuarios')
def gestion_usuarios():
    if session.get('rol') != 'admin':
        flash("Zona restringida", "danger")
        return redirect(url_for('index'))
    
    conn = obtener_conexion()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    cur.execute("SELECT id, username, rol FROM usuarios ORDER BY username ASC")
    usuarios = cur.fetchall()
    cur.close()
    conn.close()
    return render_template('gestion_usuarios.html', usuarios=usuarios)

@app.route('/eliminar_usuario/<int:id>', methods=['POST'])
def eliminar_usuario(id):
    if session.get('rol') != 'admin':
        flash("Acceso no autorizado", "danger")
        return redirect(url_for('index'))
    
    conn = obtener_conexion()
    cur = conn.cursor()
    try:
        cur.execute("DELETE FROM usuarios WHERE id = %s", (id,))
        conn.commit()
        flash("Usuario eliminado correctamente", "success")
    except Exception as e:
        conn.rollback()
        flash(f"Error al eliminar usuario: {e}", "danger")
    finally:
        cur.close()
        conn.close()
    
    return redirect(url_for('gestion_usuarios'))

@app.route('/eliminar_aspirante/<ce>', methods=['POST'])
def eliminar_aspirante(ce):
    if session.get('rol') != 'admin':
        flash("Acceso denegado: Solo el Administrador puede eliminar registro.", "danger")
        return redirect(url_for('index'))
    
    conn = obtener_conexion()
    cur = conn.cursor()
    try:
        cur.execute("DELETE FROM aspirantes WHERE ce = %s", (ce,))
        conn.commit()
        flash("Aspirante eliminado de la base de datos", "warning")
    except Exception as e:
        conn.rollback()
        flash(f"Error al eliminar: {e}", "danger")
    finally:
        cur.close()
        conn.close()
    return redirect(url_for('index'))

@app.route('/buscar_por_fecha', methods=['GET', 'POST'])
def buscar_por_fecha():
    if not session.get('logged_in'): return redirect(url_for('login'))
    
    resultados = []
    fecha_buscada = ""
    fecha_para_sql = ""
    
    if request.method == 'POST':
        fecha_buscada = request.form.get('fecha', '').strip()
        fecha_para_sql = fecha_buscada
        
        # --- CONVERTIDOR INTELIGENTE DE FORMATO ---
        # Caso 1: Si el formulario envía DD/MM/AAAA (con barras)
        if "/" in fecha_buscada:
            partes = fecha_buscada.split("/")
            if len(partes) == 3:
                fecha_para_sql = f"{partes[2]}-{partes[1]}-{partes[0]}"
        
        # Caso 2: Si el input type="date" envía AAAA-MM-DD (con guiones)
        # Nos aseguramos de que mantenga ese formato limpio para PostgreSQL
        elif "-" in fecha_buscada:
            partes = fecha_buscada.split("-")
            if len(partes) == 3 and len(partes[0]) == 4:
                fecha_para_sql = fecha_buscada # Ya está en formato AAAA-MM-DD

        if fecha_para_sql:
            conn = obtener_conexion()
            cur = conn.cursor(cursor_factory=RealDictCursor)
            
            # Ejecutamos la consulta con la fecha estandarizada
            cur.execute("""
                SELECT r.ce, a.apellido_nombre, a.aula, r.tipo_guardia, r.turno, r.cubrio, r.observaciones, r.fecha_guardia
                FROM registro_servicios r
                JOIN aspirantes a ON r.ce = a.ce
                WHERE r.fecha_guardia = %s
                ORDER BY a.apellido_nombre ASC
            """, (fecha_para_sql,))
            
            resultados = cur.fetchall()
            cur.close()
            conn.close()
        
    return render_template('buscar_fecha.html', resultados=resultados, fecha_buscada=fecha_buscada)

# --- 9. CREAR USUARIO (POR SI TAMBIÉN FALTA) ---
@app.route('/crear_usuario', methods=['GET', 'POST'])
def crear_usuario():
    if session.get('rol') != 'admin':
        return redirect(url_for('index'))
    
    if request.method == 'POST':
        u = request.form.get('username')
        p = request.form.get('password')
        r = request.form.get('rol')
        
        conn = obtener_conexion()
        cur = conn.cursor()
        
        try:
            # Intentamos insertar el nuevo usuario
            cur.execute("INSERT INTO usuarios (username, password, rol) VALUES (%s, %s, %s)", (u, p, r))
            conn.commit()
            flash("Usuario creado correctamente", "success")
            return redirect(url_for('gestion_usuarios'))
            
        except Exception as e:
            # Si el usuario ya existe, PostgreSQL enviará un error. 
            # Aquí lo capturamos para que el programa no "explote".
            conn.rollback() # Limpiamos la transacción fallida
            flash(f"Error: El nombre de usuario '{u}' ya existe o hubo un problema.", "danger")
            return redirect(url_for('gestion_usuarios'))
            
        finally:
            cur.close()
            conn.close()
            
    return render_template('crear_usuario.html')
# --- 9. HISTORIAL DE FECHAS ---
@app.route('/historial/<int:ce>')
def ver_historial(ce):
    conn = obtener_conexion()
    # Usamos RealDictCursor para que el HTML reconozca los nombres como registro.fecha
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    
    # 1. Obtenemos los datos básicos del aspirante
    cur.execute("SELECT ce, apellido_nombre, situacion FROM aspirantes WHERE ce = %s", (ce,))
    aspirante = cur.fetchone()
    
    # 2. Obtenemos su historial (ajustá el nombre de tu tabla de servicios)
    cur.execute("""
        SELECT fecha, tipo_servicio, observaciones, usuario_registro 
        FROM historial_servicios 
        WHERE ce_aspirante = %s 
        ORDER BY fecha DESC
    """, (ce,))
    historial = cur.fetchall()
    
    cur.close()
    conn.close()
    
    return render_template('historial.html', aspirante=aspirante, historial=historial)

# --- 10. EDICIÓN Y ELIMINACIÓN INDIVIDUAL ---
# --- RUTA 1: EDITAR DATOS Y SITUACIÓN (UNIFICADO) ---
@app.route('/editar_aspirante/<ce>', methods=['GET', 'POST'])
def editar_aspirante(ce):
    if session.get('rol') not in ['admin', 'operador']:
        flash("Acceso denegado: Su usuario es de Solo Lectura.", "danger")
        return redirect(url_for('index'))
    
    conn = obtener_conexion()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    if request.method == 'POST':
        cur.execute("""
            UPDATE aspirantes 
            SET apellido_nombre=%s, grado=%s, aula=%s, situacion=%s 
            WHERE ce=%s
        """, (request.form['apellido_nombre'], request.form['grado'], 
              request.form['aula'], request.form['situacion'], ce))
        conn.commit()
        conn.close()
        return redirect(url_for('index'))
    
    cur.execute("SELECT * FROM aspirantes WHERE ce = %s", (ce,))
    asp = cur.fetchone()
    conn.close()
    return render_template('editar_aspirante.html', asp=asp)

# --- FUNCIÓN PARA CARGAR ASPIRANTE MANUALMENTE ---
@app.route('/cargar_aspirante', methods=['GET', 'POST'])
def cargar_aspirante():
    if not session.get('logged_in'): 
        return redirect(url_for('login'))

    if request.method == 'POST':
        # 1. Capturamos los datos del formulario (html 'name' tags)
        ce = request.form.get('ce')
        dni = request.form.get('dni')
        apellido_nombre = request.form.get('apellido_nombre').upper() # Lo guardamos en mayúsculas
        aula = request.form.get('aula')
        grado = request.form.get('grado', 'Aspirante de III Año')

        try:
            conn = obtener_conexion()
            cur = conn.cursor()
            
            # 2. Insertamos en la tabla aspirantes según tu esquema SQL
            cur.execute("""
                INSERT INTO aspirantes (ce, dni, apellido_nombre, aula, grado, situacion, total_guardias) 
                VALUES (%s, %s, %s, %s, %s, 'DISPONIBLE', 0)
            """, (ce, dni, apellido_nombre, aula, grado))
            
            # 3. EL PASO CRÍTICO: Confirmar los cambios en la base de datos
            conn.commit() 
            
            flash("Aspirante cargado con éxito", "success")
        except Exception as e:
            if conn: conn.rollback() # Si hay error, cancelamos para no romper la base
            flash(f"Error al cargar: {e}", "danger")
        finally:
            if cur: cur.close()
            if conn: conn.close()
            
        return redirect(url_for('index'))
    
    # Si es GET, mostramos el formulario de carga
    return render_template('cargar.html')

# ==========================================
# VISTA: REGISTRAR SERVICIO INDIVIDUAL
# ==========================================
@app.route('/registrar_servicio/<ce>', methods=['GET', 'POST'])
def registrar_servicio(ce):
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    if session.get('rol') not in ['admin', 'operador']:
        flash("Acceso denegado: Su usuario es de Solo Lectura.", "danger")
        return redirect(url_for('index'))

    # Limpiamos el CE por si viene con decimales o espacios (ej: '12345.0')
    ce_clean = str(ce).strip().split('.')[0]

    conn = obtener_conexion()
    cur = conn.cursor(cursor_factory=RealDictCursor)

    if request.method == 'POST':
        fecha = request.form.get('fecha_guardia')
        tipo = request.form.get('tipo_guardia')
        turno = request.form.get('turno')
        cubrio = request.form.get('cubrio', 'SI')
        obs = request.form.get('observaciones', '')

        try:
            # Si la guardia se da por cubierta de entrada, le suma +1 al total
            if cubrio == 'SI':
                cur.execute("""
                    UPDATE aspirantes 
                    SET total_guardias = COALESCE(total_guardias, 0) + 1 
                    WHERE TRIM(SPLIT_PART(ce::text, '.', 1)) = %s
                """, (ce_clean,))

            cur.execute("""
                INSERT INTO registro_servicios (ce, fecha_guardia, tipo_guardia, turno, cubrio, observaciones) 
                VALUES (%s, %s, %s, %s, %s, %s)
            """, (ce_clean, fecha, tipo, turno, cubrio, obs))

            conn.commit()
            flash('Servicio registrado correctamente.', 'success')
        except Exception as e:
            conn.rollback()
            print(f"ERROR EN BD al registrar servicio: {e}")
            flash(f"Hubo un error al registrar el servicio: {e}", "danger")
        finally:
            cur.close()
            conn.close()

        return redirect(url_for('index'))

    # Si es GET, busca los datos del aspirante para mostrar el formulario
    cur.execute("SELECT * FROM aspirantes WHERE TRIM(SPLIT_PART(ce::text, '.', 1)) = %s", (ce_clean,))
    asp = cur.fetchone()
    cur.close()
    conn.close()

    if not asp:
        flash("Aspirante no encontrado en la base de datos.", "danger")
        return redirect(url_for('index'))

    return render_template('registrar_servicio.html', asp=asp)

# --- FUNCIÓN PARA BUSCAR (Por si el error persiste)
@app.route('/buscar_aspirante', methods=['GET', 'POST'])
def buscar_aspirante():
    if not session.get('logged_in'): 
        return redirect(url_for('login'))

    aspirantes = []
    busqueda = request.args.get('q', '').strip()

    if busqueda:
        conn = obtener_conexion()
        cur = conn.cursor(cursor_factory=RealDictCursor)
        
        query = """
            SELECT * FROM aspirantes 
            WHERE dni = %s OR ce = %s OR apellido_nombre ILIKE %s
            ORDER BY apellido_nombre ASC
        """
        cur.execute(query, (busqueda, busqueda, f'%{busqueda}%'))
        aspirantes_raw = cur.fetchall()

        def limpiar_dato(valor):
            if valor is None or str(valor).strip().lower() in ['nan', 'none', '']:
                return "0"
            v_str = str(valor).strip().upper()
            if v_str.endswith('.0'):
                return v_str[:-2]
            return v_str

        for a in aspirantes_raw:
            cur.execute("""
                SELECT fecha_guardia, tipo_guardia, turno, cubrio, observaciones 
                FROM registro_servicios 
                WHERE ce = %s 
                ORDER BY fecha_guardia DESC
            """, (a['ce'],))
            historial_guardias = cur.fetchall()
            
            historial_limpio = []
            for h in historial_guardias:
                fecha_raw = h['fecha_guardia']
                if fecha_raw:
                    fecha_str = fecha_raw.strftime('%d/%m/%Y') if hasattr(fecha_raw, 'strftime') else str(fecha_raw)
                    fecha_sql = fecha_raw.strftime('%Y-%m-%d') if hasattr(fecha_raw, 'strftime') else str(fecha_raw)
                else:
                    fecha_str = "Sin fecha"
                    fecha_sql = ""
                
                historial_limpio.append({
                    'fecha': fecha_str,
                    'fecha_sql': fecha_sql,
                    'tipo': h['tipo_guardia'],
                    'turno': h['turno'],
                    'cubrio': h['cubrio'],
                    'obs': h['observaciones'] if h['observaciones'] else ""
                })

            aspirantes.append({
                'ce': limpiar_dato(a['ce']),
                'dni': limpiar_dato(a['dni']),
                'apellido_nombre': a['apellido_nombre'],
                'aula': limpiar_dato(a['aula']),
                'grado': a['grado'],
                'situacion': a['situacion'],
                'total_guardias': a['total_guardias'],
                'historial': historial_limpio
            })
        
        cur.close()
        conn.close()

    return render_template('buscar_aspirante.html', aspirantes=aspirantes, busqueda=busqueda)

# ==========================================
# ELIMINAR SERVICIO POR C.E. Y VOLVER A DISPONIBLE
# ==========================================
@app.route('/eliminar_servicio/<ce>', methods=['POST'])
def eliminar_servicio(ce):
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    if session.get('rol') not in ['admin', 'operador']:
        flash("Acceso denegado.", "danger")
        return redirect(url_for('index'))

    # Limpiamos el C.E. (ejemplo: '111901.0' -> '111901')
    ce_clean = str(ce).strip().split('.')[0]
    
    fecha_guardia = request.form.get('fecha_guardia')
    turno = request.form.get('turno')
    cubrio = str(request.form.get('cubrio', '')).upper()

    conn = obtener_conexion()
    cur = conn.cursor(cursor_factory=RealDictCursor)

    try:
        # 1. Eliminar el servicio específico de la tabla registro_servicios
        if fecha_guardia and turno:
            cur.execute("""
                DELETE FROM registro_servicios 
                WHERE TRIM(SPLIT_PART(ce::text, '.', 1)) = %s 
                  AND fecha_guardia = %s 
                  AND turno = %s
            """, (ce_clean, fecha_guardia, turno))
        else:
            # Respaldo en caso de que no lleguen fecha/turno
            cur.execute("""
                DELETE FROM registro_servicios 
                WHERE TRIM(SPLIT_PART(ce::text, '.', 1)) = %s
            """, (ce_clean,))

        # 2. Devolver SIEMPRE la situación a 'DISPONIBLE' en la tabla aspirantes
        cur.execute("""
            UPDATE aspirantes 
            SET situacion = 'DISPONIBLE' 
            WHERE TRIM(SPLIT_PART(ce::text, '.', 1)) = %s
        """, (ce_clean,))

        # 3. Si el servicio figuraba como cubierto ('SI'), restar 1 al contador total
        if cubrio == 'SI':
            cur.execute("""
                UPDATE aspirantes 
                SET total_guardias = GREATEST(0, COALESCE(total_guardias, 0) - 1)
                WHERE TRIM(SPLIT_PART(ce::text, '.', 1)) = %s
            """, (ce_clean,))

        conn.commit()
        flash("Servicio eliminado correctamente. La situación volvió a DISPONIBLE.", "success")

    except Exception as e:
        conn.rollback()
        print(f"Error al eliminar servicio: {e}")
        flash(f"Error al eliminar el servicio: {e}", "danger")
    finally:
        cur.close()
        conn.close()

    return redirect(request.referrer or url_for('index'))                                                                                                                               

# ==========================================
# VISTA DE PLANIFICACIÓN DE SERVICIO ENTRANTE
# ==========================================
@app.route('/planificar_guardia', methods=['GET'])
def planificar_guardia():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    
    conn = obtener_conexion()
    cur = conn.cursor(cursor_factory=RealDictCursor)
    
    try:
        # 1. Obtenemos los aspirantes
        cur.execute("SELECT ce, dni, apellido_nombre, aula, grado, situacion, total_guardias FROM aspirantes")
        aspirantes_raw = cur.fetchall()
        
        # 2. Función de limpieza de datos
        def limpiar_dato(valor):
            if valor is None or str(valor).strip().lower() in ['nan', 'none', '']:
                return ""
            v_str = str(valor).strip()
            if v_str.endswith('.0'):
                return v_str[:-2]
            return v_str

        aspirantes = []
        for a in aspirantes_raw:
            aspirantes.append({
                'ce': limpiar_dato(a['ce']),
                'dni': limpiar_dato(a['dni']),
                'apellido_nombre': a['apellido_nombre'] if a['apellido_nombre'] else "SIN NOMBRE",
                'aula': limpiar_dato(a['aula']),
                'grado': a['grado'] if a['grado'] else "ASP III AÑO",
                'situacion': a['situacion'],
                'total_guardias': a['total_guardias']
            })

        # Ordenamiento por Aula y Apellido
        aspirantes.sort(key=lambda x: (x['total_guardias'], str(x['aula']), x['apellido_nombre']))

        # 3. Lógica de conteo de situaciones
        cur.execute("SELECT UPPER(TRIM(situacion)) AS sit_clean, COUNT(*) AS cantidad FROM aspirantes GROUP BY UPPER(TRIM(situacion))")
        filas_conteo = cur.fetchall()
        conteo_raw = {f['sit_clean']: f['cantidad'] for f in filas_conteo}
        
        cur.execute("SELECT COUNT(*) AS total FROM aspirantes")
        total_res = cur.fetchone()
        
        conteo = {
            'DISPONIBLES': conteo_raw.get('DISPONIBLE', 0) or conteo_raw.get('DISPONIBLES', 0),
            'SSD': conteo_raw.get('SSD', 0),
            'ART': conteo_raw.get('ART', 0),
            'LAO': conteo_raw.get('LAO', 0),
            'AUTORIZADOS': conteo_raw.get('AUTORIZADO', 0) or conteo_raw.get('AUTORIZADOS', 0),
            'TOTAL': total_res['total'] if total_res else 0
        }
    except Exception as e:
        print(f"Error en planificar_guardia: {e}")
        flash(f"Error al cargar la planificación: {e}", "danger")
        aspirantes = []
        conteo = {}
    finally:
        cur.close()
        conn.close()
    
    # Fecha sugerida (mañana)
    fecha_manana = (date.today() + timedelta(days=1)).strftime('%Y-%m-%d')
    
    return render_template(
        'planificar_guardia.html', 
        aspirantes=aspirantes, 
        conteo=conteo, 
        fecha_sugerida=fecha_manana
    )
# ==========================================
# 1. GUARDAR PLANIFICACIÓN DE GUARDIA
# ==========================================
@app.route('/guardar_planificacion_guardia', methods=['POST'])
def guardar_planificacion_guardia():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    
    fecha_guardia = request.form.get('fecha_guardia')
    tipo_guardia = request.form.get('tipo_guardia')
    turno = request.form.get('turno')
    aspirantes_seleccionados = request.form.getlist('ces_seleccionados')
    observaciones = request.form.get('observaciones', '')
    
    if not fecha_guardia or not aspirantes_seleccionados:
        flash('Debe seleccionar una fecha y al menos un aspirante.', 'warning')
        return redirect(url_for('planificar_guardia'))
    
    conn = obtener_conexion()
    cur = conn.cursor()
    try:
        for ce_raw in aspirantes_seleccionados:
            # Limpiamos el CE para asegurar compatibilidad de tipo (ej: '123.0' -> '123')
            ce_clean = str(ce_raw).strip().split('.')[0]
            
            # Inserta el registro de servicio como pendiente ('NO')
            cur.execute("""
                INSERT INTO registro_servicios (ce, fecha_guardia, tipo_guardia, turno, cubrio, observaciones)
                VALUES (%s, %s, %s, %s, 'NO', %s)
            """, (ce_clean, fecha_guardia, tipo_guardia, turno, observaciones))
            
            # Cambia la situación del aspirante usando casteo flexible
            cur.execute("""
                UPDATE aspirantes 
                SET situacion = 'DE SERVICIO' 
                WHERE TRIM(SPLIT_PART(ce::text, '.', 1)) = %s
            """, (ce_clean,))
            
        conn.commit()
        flash(f'Se planificó el servicio correctamente para {len(aspirantes_seleccionados)} aspirante(s).', 'success')
    except Exception as e:
        conn.rollback()
        print(f"Error al planificar servicio: {e}")
        flash('Ocurrió un error al guardar la planificación.', 'danger')
    finally:
        cur.close()
        conn.close()
        
    return redirect(url_for('planificar_guardia'))


# ==========================================
# 2. CONSULTAR SERVICIOS PENDIENTES POR FECHA
# ==========================================
@app.route('/confirmar_servicio', methods=['GET'])
def confirmar_servicio():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    conn = obtener_conexion()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

    fecha_busqueda = request.args.get('fecha', date.today().strftime('%Y-%m-%d')).strip()
    if not fecha_busqueda:
        fecha_busqueda = date.today().strftime('%Y-%m-%d')

    query = """
        SELECT 
            TRIM(SPLIT_PART(r.ce::text, '.', 1)) AS ce, 
            COALESCE(a.grado, 'ASP III AÑO') AS grado, 
            COALESCE(a.apellido_nombre, 'SIN NOMBRE') AS apellido_nombre, 
            COALESCE(a.aula, '0') AS aula, 
            r.tipo_guardia, 
            r.turno, 
            COALESCE(a.total_guardias, 0) AS total_guardias
        FROM registro_servicios r
        LEFT JOIN aspirantes a ON TRIM(SPLIT_PART(a.ce::text, '.', 1)) = TRIM(SPLIT_PART(r.ce::text, '.', 1))
        WHERE r.fecha_guardia::text LIKE %s
          AND (r.cubrio IS NULL OR UPPER(TRIM(r.cubrio::text)) NOT IN ('SI', 'TRUE', '1'))
        ORDER BY a.apellido_nombre ASC
    """
    
    servicios = []
    try:
        cur.execute(query, (f"{fecha_busqueda}%",))
        servicios = cur.fetchall()
    except Exception as e:
        conn.rollback()
        print(f"Error al consultar confirmar_servicio: {e}")
        flash(f"Error al consultar registros: {e}", "danger")
    finally:
        cur.close()
        conn.close()

    return render_template('confirmar_servicio.html', servicios=servicios, fecha_busqueda=fecha_busqueda)


# ==========================================
# 3. PROCESAR CONFIRMACIÓN (+1 Y DISPONIBLE)
# ==========================================
@app.route('/procesar_confirmacion_guardia', methods=['POST'])
def procesar_confirmacion_guardia():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    if session.get('rol') not in ['admin', 'operador']:
        flash("Acceso denegado: Su usuario es de Solo Lectura.", "danger")
        return redirect(url_for('index'))

    ces_confirmados = request.form.getlist('ces_confirmados')
    fecha_busqueda = request.form.get('fecha_busqueda', date.today().strftime('%Y-%m-%d'))

    if not ces_confirmados:
        flash('Debe seleccionar al menos un aspirante para confirmar la guardia.', 'warning')
        return redirect(url_for('confirmar_servicio', fecha=fecha_busqueda))

    try:
        conn = obtener_conexion()
        cur = conn.cursor()

        for ce in ces_confirmados:
            ce_clean = str(ce).strip().split('.')[0]

            # Actualiza el registro en el historial para la fecha dada
            cur.execute("""
                UPDATE registro_servicios 
                SET cubrio = 'SI' 
                WHERE TRIM(SPLIT_PART(ce::text, '.', 1)) = %s 
                  AND fecha_guardia::text LIKE %s
            """, (ce_clean, f"{fecha_busqueda}%"))

            # Suma +1 al total de guardias y lo reestablece a DISPONIBLE
            cur.execute("""
                UPDATE aspirantes 
                SET total_guardias = COALESCE(total_guardias, 0) + 1,
                    situacion = 'DISPONIBLE'
                WHERE TRIM(SPLIT_PART(ce::text, '.', 1)) = %s
            """, (ce_clean,))

        conn.commit()
        flash('Se confirmó el servicio del personal seleccionado (+1 guardia, estado DISPONIBLE).', 'success')
    except Exception as e:
        conn.rollback()
        print(f"Error en procesar_confirmacion_guardia: {e}")
        flash(f'Error al procesar la confirmación: {e}', 'danger')
    finally:
        cur.close()
        conn.close()

    return redirect(url_for('confirmar_servicio', fecha=fecha_busqueda))

# --- CIERRE DEL ARCHIVO ---
if __name__ == '__main__':
    # Aseguramos que corra en el puerto 5000 como en tu captura
    app.run(debug=True, port=5000)
