import re
from datetime import datetime

mapeo_cursos = {
    "Humanización": "HUMANIZACIÓN DE LA ATENCIÓN EN SALUD",
    "Gestión del duelo": "GESTIÓN Y MANEJO DEL DUELO EN LOS SERVICIOS DE SALUD",
    "Violencia sexual": "ATENCIÓN INTEGRAL EN SALUD A VÍCTIMAS DE VIOLENCIA SEXUAL",
    "Agentes químicos": "ATENCIÓN INTEGRAL EN SALUD A VÍCTIMAS DE ATAQUES CON AGENTES QUÍMICOS",
    "Soporte vital básico": "SOPORTE VITAL BÁSICO",
    "Soporte vital avanzado": "SOPORTE VITAL AVANZADO",
    "Cuidado del donante": "DETECCIÓN Y CUIDADO DEL DONANTE DE ÓRGANOS Y TEJIDOS",
    # Agrega más mapeos según sea necesario
}

def obtener_curso_completo(curse):
    return mapeo_cursos.get(curse, curse)

def fpuntoscedula(numerocedula):
    reverseID = numerocedula[::-1]
    partes = [reverseID[i:i+3] for i in range(0, len(reverseID), 3)]
    cedula2 = ".".join(partes)[::-1]
    return cedula2

def formatear_fecha(fecha_str):
    try:
        fecha_obj = datetime.strptime(fecha_str, "%Y-%m-%d %H:%M:%S")
    except ValueError:
        try:
            fecha_obj = datetime.strptime(fecha_str, "%d/%m/%Y")
        except ValueError as e:
            print(f"❌ Error al formatear la fecha {fecha_str}: {e}")
            return fecha_str
            exit
    return fecha_obj.strftime("%d DE %B DE %Y").upper()

def obtener_parametro(curse_completo):
    if curse_completo in ["SOPORTE VITAL BÁSICO", "SOPORTE VITAL AVANZADO"]:
        return "BAJO CONSENSOS ILCOR 2020, RECOMENDACIONES Y ACTUALIZACIONES AHA 2020, Y LO REGLAMENTADO EN LA RESOLUCIÓN 3100 DE 2019 DEL MINISTERIO DE SALUD Y PROTECCIÓN SOCIAL."
    elif curse_completo == "DETECCIÓN Y CUIDADO DEL DONANTE DE ÓRGANOS Y TEJIDOS":
        return "SEGÚN LO REGLAMENTADO EN LA LEY 1805 DE 2016, RESOLUCIÓN 3100 DE 2019 DEL MINISTERIO DE SALUD Y PROTECCIÓN SOCIAL, RESOLUCIÓN 0156 DE 2021 Y RESOLUCIÓN 0317 DE 2022 DEL INSTITUTO NACIONAL DE SALUD."
    elif curse_completo == "GESTIÓN Y MANEJO DEL DUELO EN LOS SERVICIOS DE SALUD":
        return "SEGÚN LO REGLAMENTADO EN LA RESOLUCIÓN 3100 DE 2019 DEL MINISTERIO DE SALUD Y DE LA PROTECCIÓN SOCIAL."
    elif curse_completo == "HUMANIZACIÓN DE LA ATENCIÓN EN SALUD":
        return "BASADO EN EL PLAN NACIONAL DE MEJORAMIENTO DE LA CALIDAD EN SALUD Y EN LA POLÍTICA NACIONAL DE HUMANIZACIÓN EN SALUD DEL MINISTERIO DE SALUD Y PROTECCIÓN SOCIAL, Y LO REGLAMENTADO EN LA LEY 1438 DE 2011"
    elif curse_completo == "ATENCIÓN INTEGRAL EN SALUD A VÍCTIMAS DE VIOLENCIA SEXUAL":
        return "BAJO PARÁMETROS ESTABLECIDOS EN LA LEY 1146 DE 2007 Y LA RESOLUCIÓN 0459 DE 2012, Y LO REGLAMENTADO EN LA RESOLUCIÓN 3100 DE 2019 DEL MINISTERIO DE SALUD Y PROTECCIÓN SOCIAL."
    elif curse_completo == "ATENCIÓN INTEGRAL EN SALUD A VÍCTIMAS DE ATAQUES CON AGENTES QUÍMICOS":
        return "EN CUMPLIMIENTO DE LA LEY 1971 DE 2019, PARÁMETROS ESTABLECIDOS EN LA RESOLUCIÓN 4568 DE 2014, Y LO REGLAMENTADO EN LA RESOLUCIÓN 3100 DE 2019 DEL MINISTERIO DE SALUD Y PROTECCIÓN SOCIAL"
    return "-"