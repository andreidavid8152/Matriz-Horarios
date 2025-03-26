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
#    Puedes ajustar estos valores a tus necesidades.
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
# 4. Generar la lista de franjas horarias (slots) con tus reglas
# =============================================================================
def generate_time_slots():
    slots = []
    # Fase 1: Desde 07:00 hasta ~17:50 con bloques de 60 min + 5 min de break
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
# 5. Días de la semana (columnas). Incluye Sábado
# =============================================================================
days = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"]


# =============================================================================
# 6. Función para verificar la disponibilidad de un aula en la programación real
# =============================================================================
def is_aula_available(aula, day, slot_start, slot_end, df_existing):
    """
    Verifica, en el df de la programación original, si un aula está disponible
    en un día y franja dada (que no se empalme con otra materia real).
    """
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
# 7. Set global de asignaciones (para evitar colisiones entre distintos horarios)
#    Guardamos combos sin la materia, así ningún horario futuro reutiliza
#    (día, slot, aula) ya usado.
# =============================================================================
used_assignments = set()  # (day, slot_start, slot_end, aula)


# =============================================================================
# 8. Función para checar si existe un tramo de 5 horas consecutivas ocupadas
#    en el día (independientemente de la materia).
# =============================================================================
def exceeds_4_consecutive_any_subject(schedule_df, day):
    """
    Retorna True si en el día dado (columna `day` del DF) se detectan 5 o más
    franjas consecutivas ocupadas, False en caso contrario.
    """
    assigned_series = schedule_df[day] != ""  # True si la celda está ocupada
    consecutive_count = 0

    for assigned in assigned_series:
        if assigned:
            consecutive_count += 1
            if consecutive_count >= 5:
                return True
        else:
            consecutive_count = 0

    return False


# =============================================================================
# 9. Generar horario (greedy), con tope de 4 horas consecutivas por materia
#    y además sin permitir 5 horas consecutivas de clase en un día (cualquier materia)
# =============================================================================
def generate_schedule():
    # Crear estructura vacía
    index_slots = [f"{s}-{e}" for s, e in time_slots]
    schedule_df = pd.DataFrame(
        index=index_slots, columns=["Hora Inicio", "Hora Fin"] + days
    )
    for i, (s, e) in enumerate(time_slots):
        schedule_df.loc[schedule_df.index[i], "Hora Inicio"] = s
        schedule_df.loc[schedule_df.index[i], "Hora Fin"] = e

    for d in days:
        schedule_df[d] = ""

    # Conteo de horas asignadas por materia en cada día (por si quieres mantener tope 4 consecutivas por materia)
    subject_schedule_count = {
        subject: {day: 0 for day in days} for subject in subject_hours.keys()
    }
    max_consecutive_for_subject = 4  # ajusta si quieres otro tope por materia

    for subject, hours_needed in subject_hours.items():
        candidate_aulas = subject_classrooms.get(subject, [])

        while hours_needed > 0:
            assigned_in_iteration = False

            # Orden aleatorio de días para no asignar siempre en el mismo orden
            day_order = days[:]
            random.shuffle(day_order)

            for day in day_order:
                # Si ya tiene 4 horas en este día, no dejamos meter más (para este subject)
                if subject_schedule_count[subject][day] >= max_consecutive_for_subject:
                    continue

                # Cuántas horas podría meter hoy sin pasar el tope de 4 para este subject
                max_today_subject = (
                    max_consecutive_for_subject - subject_schedule_count[subject][day]
                )
                if max_today_subject <= 0:
                    continue

                # No metas más horas de las que realmente faltan
                block_size = min(max_today_subject, hours_needed)

                # 1) Intentar un bloque consecutivo de 'block_size' (si es > 1)
                found_block = False
                if block_size > 1:
                    for i in range(len(time_slots) - block_size + 1):
                        # Revisar que todas esas celdas estén libres
                        all_free = True
                        for j in range(block_size):
                            row_label = f"{time_slots[i+j][0]}-{time_slots[i+j][1]}"
                            if schedule_df.at[row_label, day] != "":
                                all_free = False
                                break
                        if not all_free:
                            continue

                        # Ahora revisamos cada aula candidata
                        aula_block_ok = None
                        for aula in candidate_aulas:
                            available_for_all = True
                            for j in range(block_size):
                                slot_start, slot_end = time_slots[i + j]
                                combo = (day, slot_start, slot_end, aula)
                                if (combo in used_assignments) or (
                                    not is_aula_available(
                                        aula, day, slot_start, slot_end, df
                                    )
                                ):
                                    available_for_all = False
                                    break
                            if available_for_all:
                                aula_block_ok = aula
                                break

                        if aula_block_ok:
                            # Asignamos tentativamente
                            assigned_slots = []
                            for j in range(block_size):
                                slot_start, slot_end = time_slots[i + j]
                                row_label = f"{slot_start}-{slot_end}"
                                schedule_df.at[row_label, day] = (
                                    f"{subject} - {aula_block_ok}"
                                )
                                used_assignments.add(
                                    (day, slot_start, slot_end, aula_block_ok)
                                )
                                assigned_slots.append((row_label, day, aula_block_ok))

                            # Verificamos si con esto se crearon 5 horas consecutivas
                            if exceeds_4_consecutive_any_subject(schedule_df, day):
                                # Revertir asignación
                                for r_lbl, d_col, a_asig in assigned_slots:
                                    schedule_df.at[r_lbl, d_col] = ""
                                    # Quitamos del used_assignments
                                    slot_start_str, slot_end_str = r_lbl.split("-")
                                    used_assignments.discard(
                                        (d_col, slot_start_str, slot_end_str, a_asig)
                                    )
                                # No asignamos este bloque
                                continue
                            else:
                                # Asignación válida
                                subject_schedule_count[subject][day] += block_size
                                hours_needed -= block_size
                                found_block = True
                                assigned_in_iteration = True
                                break  # rompemos bucle de franjas

                        if found_block:
                            break

                # 2) Si no se encontró un bloque mayor y block_size >= 1 => probamos 1 hora
                if not found_block and block_size >= 1:
                    for i in range(len(time_slots)):
                        row_label = f"{time_slots[i][0]}-{time_slots[i][1]}"
                        if schedule_df.at[row_label, day] == "":
                            slot_start, slot_end = time_slots[i]
                            # Probar cada aula
                            aula_ok = None
                            for aula in candidate_aulas:
                                combo = (day, slot_start, slot_end, aula)
                                if (
                                    combo not in used_assignments
                                ) and is_aula_available(
                                    aula, day, slot_start, slot_end, df
                                ):
                                    # Asignar tentativo
                                    schedule_df.at[row_label, day] = (
                                        f"{subject} - {aula}"
                                    )
                                    used_assignments.add(combo)
                                    aula_ok = aula
                                    break
                            if aula_ok:
                                # Revisamos si hay 5 seguidas
                                if exceeds_4_consecutive_any_subject(schedule_df, day):
                                    # revertir
                                    schedule_df.at[row_label, day] = ""
                                    used_assignments.discard(
                                        (day, slot_start, slot_end, aula_ok)
                                    )
                                    aula_ok = None
                                    # y seguimos buscando otro slot
                                else:
                                    # Perfecto
                                    subject_schedule_count[subject][day] += 1
                                    hours_needed -= 1
                                    assigned_in_iteration = True
                                    # No asignamos más horas sueltas en este día (rompemos)
                                    break
                    # fin del for de slots
                # fin if not found_block

                if assigned_in_iteration or hours_needed <= 0:
                    break  # terminamos con este día o materia

            if not assigned_in_iteration:
                # No hubo forma de asignar nada en esta iteración => fallamos
                return None

    return schedule_df


# =============================================================================
# 10. Generar 20 horarios. Si alguno sale None, paramos.
# =============================================================================
unique_schedules = []
num_schedules_required = 5

for i in range(num_schedules_required):
    sched = generate_schedule()
    if sched is None:
        print(
            "No se pudo asignar alguna materia; se aborta la generación de más horarios."
        )
        break
    unique_schedules.append(sched)
    print(f"Horario {len(unique_schedules)} generado.")

if len(unique_schedules) < num_schedules_required:
    print(
        "No se alcanzaron los 20 horarios por falta de disponibilidad o restricciones."
    )
else:
    print("¡Se generaron 20 horarios exitosamente!")

# =============================================================================
# 11. Exportar todos los horarios a un archivo Excel
# =============================================================================
if len(unique_schedules) > 0:
    output_file = "Horarios_Unicos.xlsx"
    with pd.ExcelWriter(output_file) as writer:
        for idx, sched in enumerate(unique_schedules, start=1):
            sheet_name = f"Horario_{idx}"
            sched.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"\nHorarios generados y guardados en '{output_file}'")
