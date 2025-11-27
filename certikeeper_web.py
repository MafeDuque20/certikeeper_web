# =========================
# EXTRAER SOLO PRIMER NOMBRE + PRIMER APELLIDO (MEJORADO)
# =========================
def extraer_primer_nombre_apellido(nombre_completo):
    if not nombre_completo:
        return None, None

    limpio = nombre_completo.replace("\n", " ").replace("-", " ")
    limpio = " ".join(limpio.split())
    partes = limpio.split()

    primer_nombre = None
    primer_apellido = None

    # Buscar primer nombre válido
    for palabra in partes:
        if palabra not in palabras_invalidas:
            primer_nombre = palabra
            break

    # Buscar primer apellido válido después del primer nombre
    if primer_nombre:
        for palabra in partes[partes.index(primer_nombre)+1:]:
            if palabra not in palabras_invalidas:
                primer_apellido = palabra
                break

    return primer_nombre, primer_apellido
