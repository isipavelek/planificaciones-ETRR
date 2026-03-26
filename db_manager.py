import sqlite3

DB_NAME = "planificaciones.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    # Tabla Materias
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS materias (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT UNIQUE NOT NULL
        )
    ''')
    
    # Tabla Docentes
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS docentes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT UNIQUE NOT NULL,
            email TEXT,
            coordinador TEXT
        )
    ''')
    
    # Migración: Agregar columna coordinador si no existe
    try:
        cursor.execute("ALTER TABLE docentes ADD COLUMN coordinador TEXT")
    except:
        pass
    
    # Relación Materia - Docente (Una materia puede tener varios docentes)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS materia_docente (
            materia_id INTEGER,
            docente_id INTEGER,
            PRIMARY KEY (materia_id, docente_id),
            FOREIGN KEY (materia_id) REFERENCES materias(id) ON DELETE CASCADE,
            FOREIGN KEY (docente_id) REFERENCES docentes(id) ON DELETE CASCADE
        )
    ''')
    
    conn.commit()

    # Tabla Config
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS config (
            clave TEXT PRIMARY KEY,
            valor TEXT
        )
    ''')
    # Valor inicial (vacio) para sharepoint_url
    default_url = "https://etrrar-my.sharepoint.com/:x:/g/personal/ipavelek_etrr_edu_ar/IQAiT6ypndWyTZJmKJ1JQVQrASLW1LQ99lU3kU0wHWTzrq0?e=jDXiWS"
    cursor.execute("INSERT OR IGNORE INTO config (clave, valor) VALUES (?, ?)", ("sharepoint_url", default_url))
    
    # Si esta vacia (por ejecuciones anteriores), actualizarla
    cursor.execute("UPDATE config SET valor = ? WHERE clave = 'sharepoint_url' AND (valor = '' OR valor IS NULL)", (default_url,))
    
    conn.commit()
    conn.close()

def get_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def limpiar_db():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM materia_docente")
    cursor.execute("DELETE FROM materias")
    cursor.execute("DELETE FROM docentes")
    # Resetear contadores de autoincremento
    cursor.execute("UPDATE sqlite_sequence SET seq = 0 WHERE name = 'materias'")
    cursor.execute("UPDATE sqlite_sequence SET seq = 0 WHERE name = 'docentes'")
    conn.commit()
    conn.close()

def agregar_materia(nombre):
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO materias (nombre) VALUES (?)", (nombre.strip(),))
        conn.commit()
        materia_id = cursor.lastrowid
    except sqlite3.IntegrityError:
        cursor.execute("SELECT id FROM materias WHERE nombre = ?", (nombre.strip(),))
        materia_id = cursor.fetchone()[0]
    finally:
        conn.close()
    return materia_id

def agregar_docente(nombre, email="", coordinador=""):
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO docentes (nombre, email, coordinador) VALUES (?, ?, ?)", 
                       (nombre.strip(), email.strip(), coordinador.strip()))
        conn.commit()
        docente_id = cursor.lastrowid
    except sqlite3.IntegrityError:
        # Si ya existe, actualizamos email y coordinador
        updates = []
        params = []
        if email.strip():
            updates.append("email = ?")
            params.append(email.strip())
        if coordinador.strip():
            updates.append("coordinador = ?")
            params.append(coordinador.strip())
        
        if updates:
            params.append(nombre.strip())
            cursor.execute(f"UPDATE docentes SET {', '.join(updates)} WHERE nombre = ?", params)
            conn.commit()
            
        cursor.execute("SELECT id FROM docentes WHERE nombre = ?", (nombre.strip(),))
        docente_id = cursor.fetchone()[0]
    finally:
        conn.close()
    return docente_id

def asignar_docente_a_materia(materia_id, docente_id):
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO materia_docente (materia_id, docente_id) VALUES (?, ?)", (materia_id, docente_id))
        conn.commit()
    except sqlite3.IntegrityError:
        pass # Ya existe la relación
    finally:
        conn.close()

def obtener_todas_materias():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, nombre FROM materias ORDER BY nombre")
    res = cursor.fetchall()
    conn.close()
    return res

def eliminar_materia(materia_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM materias WHERE id = ?", (materia_id,))
    conn.commit()
    conn.close()

def actualizar_materia_por_id(materia_id, nuevo_nombre):
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("UPDATE materias SET nombre = ? WHERE id = ?", (nuevo_nombre.strip(), materia_id))
        conn.commit()
    except sqlite3.IntegrityError:
        pass # Ignorar si el nuevo nombre ya está en uso
    finally:
        conn.close()

def obtener_todos_docentes():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, nombre, email, coordinador FROM docentes ORDER BY nombre")
    res = cursor.fetchall()
    conn.close()
    return res

def eliminar_docente(docente_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM docentes WHERE id = ?", (docente_id,))
    conn.commit()
    conn.close()

def actualizar_docente_por_id(docente_id, nuevo_nombre, nuevo_email, nuevo_coordinador):
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("UPDATE docentes SET nombre = ?, email = ?, coordinador = ? WHERE id = ?", 
                       (nuevo_nombre.strip(), nuevo_email.strip(), nuevo_coordinador.strip(), docente_id))
        conn.commit()
    except sqlite3.IntegrityError:
        pass # Ignorar si el nuevo nombre ya está en uso por otro
    finally:
        conn.close()

def obtener_materias_con_docentes():
    """Devuelve un diccionario donde la clave es el nombre de la materia y el valor es una lista de nombres de docentes."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT m.nombre, d.nombre 
        FROM materias m
        LEFT JOIN materia_docente md ON m.id = md.materia_id
        LEFT JOIN docentes d ON md.docente_id = d.id
    ''')
    filas = cursor.fetchall()
    conn.close()
    
    resultado = {}
    for materia, docente in filas:
        if materia not in resultado:
            resultado[materia] = []
        if docente:
            resultado[materia].append(docente)
            
    return resultado

def eliminar_relaciones_materia(materia_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM materia_docente WHERE materia_id = ?", (materia_id,))
    conn.commit()
    conn.close()

def obtener_materias_con_docentes_detallado():
    """Devuelve la información completa: materia, docente, email, coordinador."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT m.nombre, d.nombre, d.email, d.coordinador
        FROM materias m
        LEFT JOIN materia_docente md ON m.id = md.materia_id
        LEFT JOIN docentes d ON md.docente_id = d.id
    ''')
    filas = cursor.fetchall()
    conn.close()
    
    resultado = {}
    for mat, doc, email, coord in filas:
        if mat not in resultado:
            resultado[mat] = []
        if doc:
            resultado[mat].append({
                "nombre": doc,
                "email": email if email else "",
                "coordinador": coord if coord else ""
            })
    return resultado

def set_config(clave, valor):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("INSERT OR REPLACE INTO config (clave, valor) VALUES (?, ?)", (clave, valor))
    conn.commit()
    conn.close()

def get_config(clave, default=""):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT valor FROM config WHERE clave = ?", (clave,))
    row = cursor.fetchone()
    conn.close()
    return row[0] if row else default

if __name__ == "__main__":
    init_db()
    print("Base de datos inicializada correctamente.")
