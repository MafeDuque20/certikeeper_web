import psycopg2
import os
from dotenv import load_dotenv

load_dotenv()

def get_connection():
    return psycopg2.connect(
        host=os.getenv("db.zbbsgshqfsdzflqogssn.supabase.co"),       # Endpoint de Supabase
        database=os.getenv("postgres"),   # Nombre de la base de datos
        user=os.getenv("postgres"),       # Usuario de Supabase
        password=os.getenv("mmTMk9Zl7Dsbnqb1"), # Contraseña
        port=5432
    )
