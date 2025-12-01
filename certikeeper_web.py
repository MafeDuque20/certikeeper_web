from supabase import create_client, Client
from datetime import datetime
import os

url: str = os.environ.get("SUPABASE_URL")
key: str = os.environ.get("SUPABASE_KEY")
supabase: Client = create_client(url, key)

# Ejemplo de insertar un registro
data = {
    "nombre_archivo": "Julian_Perez_FACTORES_HUMANOS.pdf",
    "base": "CLO",
    "curso": "Factores Humanos",
    "fecha_envio": datetime.now().isoformat()
}
supabase.table("certificados").insert(data).execute()

# Ejemplo de leer registros
historial = supabase.table("certificados").select("*").execute()
print(historial.data)
