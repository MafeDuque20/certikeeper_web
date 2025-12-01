from db.connection import get_connection

def create_tables():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS historial (
            id SERIAL PRIMARY KEY,
            nombre_archivo TEXT,
            base TEXT,
            curso TEXT,
            fecha_envio TIMESTAMP
        );
    """)
    conn.commit()
    cursor.close()
    conn.close()
