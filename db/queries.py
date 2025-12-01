from db.connection import get_connection

def registrar_envio(nombre_archivo, base, curso, fecha_envio):
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO historial (nombre_archivo, base, curso, fecha_envio)
        VALUES (%s, %s, %s, %s);
    """, (nombre_archivo, base, curso, fecha_envio))

    conn.commit()
    cursor.close()
    conn.close()

def obtener_historial():
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM historial ORDER BY fecha_envio DESC;")
    rows = cursor.fetchall()

    cursor.close()
    conn.close()

    return rows
