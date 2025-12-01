from db.models import create_tables
from db.queries import registrar_envio, obtener_historial
from datetime import datetime

# 1️⃣ Crear tablas (solo se ejecuta la primera vez)
create_tables()
print("Tablas creadas exitosamente.")

# 2️⃣ Insertar un registro de prueba
registrar_envio(
    nombre_archivo="Julian_Perez_FACTORES_HUMANOS.pdf",
    base="CLO",
    curso="Factores Humanos",
    fecha_envio=datetime.now()
)
print("Registro de prueba insertado.")

# 3️⃣ Leer registros y mostrar
historial = obtener_historial()
print("Historial actual:")
for fila in historial:
    print(fila)
