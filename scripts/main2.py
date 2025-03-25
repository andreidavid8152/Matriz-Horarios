import pandas as pd
import random
from datetime import datetime, timedelta

# 1. Leer el Excel completo (sin filtrar por TIPO_SALA_DESC)
archivo_excel = "programacion.xlsx"
hoja = "Data"
df = pd.read_excel(archivo_excel, sheet_name=hoja)

# Limpiar espacios en columnas clave
df["SIGLA"] = df["SIGLA"].astype(str).str.strip()
df["SSBSECT_SCHD_CODE"] = df["SSBSECT_SCHD_CODE"].astype(str).str.strip()
df["HORA_INICIO"] = df["HORA_INICIO"].astype(str).str.strip()
df["HORA_FIN"] = df["HORA_FIN"].astype(str).str.strip()
df["SALA"] = df["SALA"].astype(str).str.strip()
df["MATERIA_TITULO_PA"] = df["MATERIA_TITULO_PA"].astype(str).str.strip()

print("Total registros en el Excel:", len(df))

# 2. Definir las materias (se separa la materia mixta en teórica y práctica)
materias = [
    {
        "sigla": "MEDZ4378",
        "nombre": "D-HIST DEL CONOCIM Y LA PRAC M",
        "horas": 2,
        "tipo": "teo",
    },
    {
        "sigla": "VIDA0001",
        "nombre": "D-APREDIZ ESTRATEGIC Y LIDERAZ / D-APREDIZ ESTRATEGIC Y LIDERA",
        "horas": 3,
        "tipo": "teo",
    },
    {"sigla": "MEDZ4376", "nombre": "D-HISTOLOGIA", "horas": 4, "tipo": "teo"},
    {"sigla": "MEDZ4377", "nombre": "D-PSICOLOGIA MEDICA", "horas": 3, "tipo": "teo"},
    {
        "sigla": "MEDZ4374",
        "nombre": "D-QUIMICA GENERAL PARA MEDICIN",
        "horas": 4,
        "tipo": "teo",
    },
    {
        "sigla": "MEDZ4375",
        "nombre": "D-ANATOM CLINIC E IMAGENOL I - TEO",
        "horas": 6,
        "tipo": "teo",
    },
    {
        "sigla": "MEDZ4375",
        "nombre": "D-ANATOM CLINIC E IMAGENOL I - PRA",
        "horas": 2,
        "tipo": "pra",
    },
    {
        "sigla": "VIDA0002",
        "nombre": "D-COMUNICACION EFECTIVA",
        "horas": 3,
        "tipo": "teo",
    },
]

# Se asigna el número de grupos según la capacidad:
# Teóricas: 30 alumnos -> 10 grupos; Prácticas: 15 alumnos -> 20 grupos.
for m in materias:
    m["grupos"] = 10 if m["tipo"] == "teo" else 20

# 3. Generar la lista de sesiones (cada sesión = 1 hora)
sesiones = []
for m in materias:
    for i in range(m["horas"]):
        sesiones.append({"sigla": m["sigla"], "nombre": m["nombre"], "tipo": m["tipo"]})

print("Total de sesiones generadas:", len(sesiones))  # Debe ser 27


# 4. Generar los slots horarios hasta las 21:50
def generar_slots():
    slots = []
    hora_actual = datetime.strptime("07:00", "%H:%M")
    end_time = datetime.strptime("21:50", "%H:%M")
    while True:
        fin_dt = hora_actual + timedelta(minutes=60)
        if fin_dt > end_time:
            break
        inicio = hora_actual.strftime("%H%M")
        fin = fin_dt.strftime("%H%M")
        slots.append({"inicio": inicio, "fin": fin})
        # Desde las 07:00 hasta las 17:45 se añade receso de 5 minutos; luego sin receso
        if hora_actual < datetime.strptime("17:45", "%H:%M"):
            hora_actual = fin_dt + timedelta(minutes=5)
        else:
            hora_actual = fin_dt
    return slots


slots = generar_slots()
print("Slots generados:")
for s in slots:
    print(s)

# Días de la semana para la cuadrícula
dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]

# 5. Distribuir las sesiones en una matriz (round-robin)
matriz = [[None for _ in range(len(dias))] for _ in range(len(slots))]
indice = 0
for r in range(len(slots)):
    for c in range(len(dias)):
        if indice < len(sesiones):
            matriz[r][c] = sesiones[indice]
            indice += 1


# Función para asignar aula según disponibilidad
def get_aula(session, slot):
    current_inicio = slot["inicio"]
    current_fin = slot["fin"]
    code = "TEO" if session["tipo"] == "teo" else "PRA"

    # Primero, buscar candidatos para esa materia en ese horario y tipo
    candidates = df[
        (df["SIGLA"] == session["sigla"])
        & (df["SSBSECT_SCHD_CODE"] == code)
        & (df["HORA_INICIO"] == current_inicio)
        & (df["HORA_FIN"] == current_fin)
    ]
    available_aulas = candidates["SALA"].unique().tolist()
    if available_aulas:
        return random.choice(available_aulas)
    else:
        # Fallback: buscar en todo el Excel candidatos para ese horario y tipo, sin filtrar por materia
        fallback_candidates = df[
            (df["SSBSECT_SCHD_CODE"] == code)
            & (df["HORA_INICIO"] == current_inicio)
            & (df["HORA_FIN"] == current_fin)
        ]
        fallback_aulas = fallback_candidates["SALA"].unique().tolist()
        if fallback_aulas:
            return random.choice(fallback_aulas)
        else:
            return "No Aula"


# 6. Generar los horarios para cada grupo (20 hojas)
horarios = {}
num_horarios = 20

for grupo in range(1, num_horarios + 1):
    # Crear copia de la matriz para este grupo
    horario_matrix = [[None for _ in range(len(dias))] for _ in range(len(slots))]
    # Diccionario para almacenar aula asignada en una misma jornada (para sesiones consecutivas)
    aula_assigned = {}
    for r in range(len(slots)):
        for c in range(len(dias)):
            session = matriz[r][c]
            if session is None:
                horario_matrix[r][c] = ""
            else:
                label = "TEO" if session["tipo"] == "teo" else "PRA"
                # Si la sesión anterior del mismo día es la misma materia, reusar el aula asignada previamente
                if r > 0 and matriz[r - 1][c] is not None:
                    prev_session = matriz[r - 1][c]
                    prev_aula = aula_assigned.get((c, r - 1))
                    if (
                        prev_session["sigla"] == session["sigla"]
                        and prev_session["tipo"] == session["tipo"]
                    ):
                        aula = prev_aula
                    else:
                        aula = None
                else:
                    aula = None

                if aula is None:
                    aula = get_aula(session, slots[r])
                # Guardar el aula asignada para este slot en el día (para reusar en sesiones consecutivas)
                aula_assigned[(c, r)] = aula

                # Obtener el nombre de la materia usando la columna MATERIA_TITULO_PA del Excel
                # Se filtra por la sala asignada (si se encontró) y el horario
                if aula != "No Aula":
                    candidate_row = df[
                        (df["SIGLA"] == session["sigla"])
                        & (
                            df["SSBSECT_SCHD_CODE"]
                            == ("TEO" if session["tipo"] == "teo" else "PRA")
                        )
                        & (df["HORA_INICIO"] == slots[r]["inicio"])
                        & (df["HORA_FIN"] == slots[r]["fin"])
                        & (df["SALA"] == aula)
                    ]
                    if not candidate_row.empty:
                        materia_titulo_pa = candidate_row.iloc[0]["MATERIA_TITULO_PA"]
                    else:
                        materia_titulo_pa = session["nombre"]
                else:
                    materia_titulo_pa = session["nombre"]

                # Formato de la celda: "MATERIA_TITULO_PA (TEO/PRA) - AULA"
                horario_matrix[r][c] = f"{materia_titulo_pa} ({label}) - {aula}"
    horarios[f"Horario_{grupo}"] = horario_matrix

# 7. Escribir el resultado en un Excel, una hoja por grupo
with pd.ExcelWriter("horarios.xlsx", engine="openpyxl") as writer:
    for nombre_hoja, matrix in horarios.items():
        data = []
        for idx, slot in enumerate(slots):
            fila = {"Hora Inicio": slot["inicio"], "Hora Fin": slot["fin"]}
            total = 0
            for i, dia in enumerate(dias):
                valor = matrix[idx][i]
                fila[dia] = valor
                if valor != "":
                    total += 1
            fila["Total general"] = total
            data.append(fila)
        df_horario = pd.DataFrame(data)
        totales = {
            dia: df_horario[dia].apply(lambda x: 1 if x != "" else 0).sum()
            for dia in dias
        }
        totales["Hora Inicio"] = "Total general"
        totales["Hora Fin"] = ""
        totales["Total general"] = df_horario["Total general"].sum()
        df_totales = pd.DataFrame([totales])
        df_horario = pd.concat([df_horario, df_totales], ignore_index=True)
        df_horario.to_excel(writer, sheet_name=nombre_hoja, index=False)

    # Forzar que al menos una hoja esté visible
    if writer.book.worksheets:
        writer.book.active = writer.book.worksheets[0]

print("Archivo 'horarios.xlsx' generado con éxito.")
