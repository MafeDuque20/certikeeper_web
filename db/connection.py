# db/connection.py
import os
import psycopg2
from psycopg2.extras import RealDictCursor

def get_connection():
    """
    Devuelve una conexión a la base de datos PostgreSQL.
    Compatible con Neon / Heroku / cualquier PostgreSQL remoto.
    """
    try:
        conn = psycopg2.connect(
            host=os.getenv("DB_HOST"),        # Ej: 'ep-wispy-sky-123456.us-east-1.aws.neon.tech'
            dbname=os.getenv("DB_NAME"),      # Nombre de la base de datos
            user=os.getenv("DB_USER"),        # Usuario
            password=os.getenv("DB_PASSWORD"),# Contraseña
            port=os.getenv("DB_PORT", 5432),  # Puerto (por defecto 5432)
            sslmode='require',                # IMPORTANTE: Neon y otros requieren SSL
            cursor_factory=RealDictCursor     # Para que los SELECT devuelvan diccionarios
        )
        return conn
    except Exception as e:
        print("❌ Error al conectar a la base de datos:", e)
        raise e
