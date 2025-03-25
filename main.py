import pandas as pd
import random
from datetime import datetime, timedelta

# =============================================================================
# 1. Cargar y filtrar datos
# =============================================================================
archivo = "Programacion.xlsx"
df = pd.read_excel(archivo, sheet_name="Data")
df.columns = df.columns.str.strip()

# Definir las materias de interés
materias_buscadas = [
    "D-HIST DEL CONOCIM Y LA PRAC M",
    "D-APREDIZ ESTRATEGIC Y LIDERAZ",
    "D-HISTOLOGIA",
    "D-PSICOLOGIA MEDICA",
    "D-QUIMICA GENERAL PARA MEDICIN",
    "D-ANATOM CLINIC E IMAGENOL I",
    "D-COMUNICACION EFECTIVA",
]

# Filtrar solo las filas de las materias deseadas, tipo de sala y campus UP
tipos_sala_validos = ["AULA", "MORFOFUNCION O LAB. DESTREZAS"]
filtro_materias = df["MATERIA_TITULO_PA"].isin(materias_buscadas)
filtro_salas = df["TIPO_SALA_DESC"].isin(tipos_sala_validos)
filtro_campus = df["CODIGO_CAMPUS"] == "UP"
df_filtrado = df[filtro_materias & filtro_salas & filtro_campus]

# =============================================================================
# 2. Generar las llaves de aulas candidatas para cada materia y modalidad
# =============================================================================
subject_classrooms = {}
for materia in materias_buscadas:
    if materia == "D-ANATOM CLINIC E IMAGENOL I":
        aulas_teo = (
            df_filtrado[
                (df_filtrado["MATERIA_TITULO_PA"] == materia)
                & (df_filtrado["SSBSECT_SCHD_CODE"] == "TEO")
            ]["SALA"]
            .dropna()
            .unique()
            .tolist()
        )
        aulas_pra = (
            df_filtrado[
                (df_filtrado["MATERIA_TITULO_PA"] == materia)
                & (df_filtrado["SSBSECT_SCHD_CODE"] == "PRA")
            ]["SALA"]
            .dropna()
            .unique()
            .tolist()
        )
        subject_classrooms[materia + " TEO"] = aulas_teo
        subject_classrooms[materia + " PRA"] = aulas_pra
    else:
        aulas = (
            df_filtrado[df_filtrado["MATERIA_TITULO_PA"] == materia]["SALA"]
            .dropna()
            .unique()
            .tolist()
        )
        subject_classrooms[materia] = aulas

# =============================================================================
# 3. Definir las horas requeridas por materia (diferenciando Anatom TEO y PRA)
# =============================================================================
subject_hours = {
    "D-HIST DEL CONOCIM Y LA PRAC M": 2,
    "D-APREDIZ ESTRATEGIC Y LIDERAZ": 3,
    "D-HISTOLOGIA": 4,
    "D-PSICOLOGIA MEDICA": 3,
    "D-QUIMICA GENERAL PARA MEDICIN": 4,
    "D-ANATOM CLINIC E IMAGENOL I TEO": 6,
    "D-ANATOM CLINIC E IMAGENOL I PRA": 2,
    "D-COMUNICACION EFECTIVA": 3,
}


# =============================================================================
# 4. Generar la lista de franjas horarias con el formato solicitado
# =============================================================================
def generate_time_slots():
    slots = []
    # Fase 1: Desde 07:00 hasta un cierto umbral (17:50) con bloques de 60 min + 5 min de break
    start_phase1 = datetime.strptime("07:00", "%H:%M")
    threshold = datetime.strptime("17:50", "%H:%M")
    slot_duration1 = timedelta(minutes=60)
    break_duration1 = timedelta(minutes=5)

    current = start_phase1
    while current + slot_duration1 <= threshold:
        slot_end = current + slot_duration1
        slots.append((current.strftime("%H:%M"), slot_end.strftime("%H:%M")))
        current = slot_end + break_duration1

    # Fase 2: Desde threshold en adelante, bloques de 59 min + 1 min de break
    slot_duration2 = timedelta(minutes=59)
    break_duration2 = timedelta(minutes=1)
    current = threshold
    final_end = datetime.strptime("21:49", "%H:%M")

    while current + slot_duration2 <= final_end:
        slot_end = current + slot_duration2
        slots.append((current.strftime("%H:%M"), slot_end.strftime("%H:%M")))
        current = slot_end + break_duration2
    return slots


time_slots = generate_time_slots()

# =============================================================================
# 5. Días de la semana (columnas). Agregamos Sábado
# =============================================================================
days = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"]


# =============================================================================
# 6. Función para verificar la disponibilidad de un aula en el DataFrame original
# =============================================================================
def is_aula_available(aula, day, slot_start, slot_end, df_existing):
    """
    Revisa, según el df de la programación original, si esa aula está
    disponible en ese día y franja (no se empalme con algo existente).
    """
    # Mapeo de días, incluyendo sábado como día 6
    day_mapping = {
        "Lunes": 1,
        "Martes": 2,
        "Miércoles": 3,
        "Jueves": 4,
        "Viernes": 5,
        "Sábado": 6,
    }
    day_id = day_mapping[day]
    base_date = "2025-01-01"

    start_dt = datetime.strptime(base_date + " " + slot_start, "%Y-%m-%d %H:%M")
    end_dt = datetime.strptime(base_date + " " + slot_end, "%Y-%m-%d %H:%M")

    df_day = df_existing[df_existing["DAY_ID"] == day_id]
    for _, row in df_day.iterrows():
        if row["SALA"] != aula:
            continue
        row_start_str = f"{int(row['HORA_INICIO']):04d}"
        row_end_str = f"{int(row['HORA_FIN']):04d}"
        row_start_dt = datetime.strptime(
            base_date + " " + row_start_str[:2] + ":" + row_start_str[2:],
            "%Y-%m-%d %H:%M",
        )
        row_end_dt = datetime.strptime(
            base_date + " " + row_end_str[:2] + ":" + row_end_str[2:], "%Y-%m-%d %H:%M"
        )
        # Checamos si se traslapan
        if not (end_dt <= row_start_dt or start_dt >= row_end_dt):
            return False
    return True


# =============================================================================
# 7. (Opcional) Set de asignaciones usadas si quieres evitar que repitan combos
#    (Si no quieres evitar repetición, comenta estas líneas)
# =============================================================================
used_assignments = set()


# =============================================================================
# 8. Función que genera un horario usando un algoritmo "greedy" de asignación
#    Si falla alguna materia, regresa None (sin warnings).
# =============================================================================
def generate_schedule():
    # Crear la estructura vacía del horario
    index_slots = [f"{s}-{e}" for s, e in time_slots]
    schedule_df = pd.DataFrame(
        index=index_slots, columns=["Hora Inicio", "Hora Fin"] + days
    )
    for i, (s, e) in enumerate(time_slots):
        schedule_df.loc[schedule_df.index[i], "Hora Inicio"] = s
        schedule_df.loc[schedule_df.index[i], "Hora Fin"] = e

    for day in days:
        schedule_df[day] = ""

    # Llevar el conteo de horas asignadas por materia en cada día
    subject_schedule_count = {
        subject: {day: 0 for day in days} for subject in subject_hours.keys()
    }

    # Algoritmo greedy para asignar bloques (máximo 2 horas consecutivas por día)
    for subject, hours_needed in subject_hours.items():
        candidate_aulas = subject_classrooms.get(subject, [])

        # Mientras esta materia necesite horas
        while hours_needed > 0:
            assigned_in_iteration = False

            # Mezclamos el orden de los días para no asignar siempre en el mismo orden
            day_order = days[:]
            random.shuffle(day_order)

            for day in day_order:
                # Si ese sujeto ya tiene 2 horas en este día, pasamos
                if subject_schedule_count[subject][day] >= 2:
                    continue

                # Tamaño posible del bloque a asignar (máx 2, o lo que falte de la materia)
                block_size = min(2 - subject_schedule_count[subject][day], hours_needed)

                # 1) Intentar encontrar un bloque consecutivo de tamaño block_size
                for i in range(len(time_slots) - block_size + 1):
                    # Revisamos que todas esas celdas estén libres en la schedule_df
                    slots_free = True
                    for j in range(block_size):
                        row_label = f"{time_slots[i+j][0]}-{time_slots[i+j][1]}"
                        if schedule_df.at[row_label, day] != "":
                            slots_free = False
                            break
                    if not slots_free:
                        continue

                    # Para cada aula candidata, revisamos disponibilidad
                    for aula in candidate_aulas:
                        available_for_all = True
                        for j in range(block_size):
                            slot_start, slot_end = time_slots[i + j]

                            # 1) que el aula esté libre en el DF original
                            # 2) que la (subject,day,slot,aula) no haya sido usada antes (si usas used_assignments)
                            combo = (subject, day, slot_start, slot_end, aula)
                            if (
                                not is_aula_available(
                                    aula, day, slot_start, slot_end, df
                                )
                            ) or (combo in used_assignments):
                                available_for_all = False
                                break

                        if available_for_all:
                            # Asignamos ese bloque
                            for j in range(block_size):
                                slot_start, slot_end = time_slots[i + j]
                                row_label = f"{slot_start}-{slot_end}"
                                schedule_df.at[row_label, day] = f"{subject} - {aula}"
                                used_assignments.add(
                                    (subject, day, slot_start, slot_end, aula)
                                )

                            subject_schedule_count[subject][day] += block_size
                            hours_needed -= block_size
                            assigned_in_iteration = True
                            break  # Salimos del bucle de aulas
                    if assigned_in_iteration:
                        break  # Salimos del bucle de slots

                # 2) Si no se encontró un bloque consecutivo y no asignamos nada aún:
                if (not assigned_in_iteration) and (
                    subject_schedule_count[subject][day] < 2
                ):
                    # Intentar asignaciones individuales
                    for i in range(len(time_slots)):
                        row_label = f"{time_slots[i][0]}-{time_slots[i][1]}"
                        if schedule_df.at[row_label, day] == "":
                            slot_start, slot_end = time_slots[i]
                            for aula in candidate_aulas:
                                combo = (subject, day, slot_start, slot_end, aula)
                                if is_aula_available(
                                    aula, day, slot_start, slot_end, df
                                ) and (combo not in used_assignments):
                                    schedule_df.at[row_label, day] = (
                                        f"{subject} - {aula}"
                                    )
                                    used_assignments.add(combo)

                                    subject_schedule_count[subject][day] += 1
                                    hours_needed -= 1
                                    assigned_in_iteration = True
                                    break
                        if assigned_in_iteration or hours_needed <= 0:
                            break

                if assigned_in_iteration or hours_needed <= 0:
                    break  # Salimos del bucle de días

            if not assigned_in_iteration:
                # No hubo forma de asignar ni un bloque ni una hora suelta en ningún día
                # => retornamos None en lugar de imprimir warning
                return None

    # Si asignamos todo sin problemas, regresamos el dataframe
    return schedule_df


# =============================================================================
# 9. Generar 20 horarios. Si algún horario sale None, paramos (ya no hay sentido).
# =============================================================================
unique_schedules = []
num_schedules_required = 20

for i in range(num_schedules_required):
    sched = generate_schedule()
    if sched is None:
        print(
            "No se pudo asignar una de las materias; se aborta la generación de más horarios."
        )
        break
    unique_schedules.append(sched)
    print(f"Horario {len(unique_schedules)} generado.")

if len(unique_schedules) < num_schedules_required:
    print("No se alcanzaron los 20 horarios por falta de disponibilidad de aulas.")
else:
    print("¡Se generaron 20 horarios exitosamente!")

# =============================================================================
# 10. Exportar todos los horarios a un archivo Excel, cada uno en una hoja distinta
# =============================================================================
if len(unique_schedules) > 0:
    output_file = "Horarios_Unicos.xlsx"
    with pd.ExcelWriter(output_file) as writer:
        for idx, sched in enumerate(unique_schedules, start=1):
            sheet_name = f"Horario_{idx}"
            sched.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"\nHorarios generados y guardados en '{output_file}'")
