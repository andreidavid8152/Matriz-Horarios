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
    "D-BIOLOGIA CELULAR Y MOLECULAR",
    "D-PROCEDIMIENTOS BASICOS I",
    "D-ANATOM CLINIC E IMAGENOL II",
    "D-EMBRIOLOGIA",
    "D-NEUROANATOMIA",
    "D-PENSAMIENTO CRITICO APLICADO",
    "D-INTERAC EFECT EN SISTEMAS SO",
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
    if materia in ["D-ANATOM CLINIC E IMAGENOL II", "D-EMBRIOLOGIA", "D-NEUROANATOMIA"]:
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
# 3. Definir las horas requeridas por materia (diferenciando TEO y PRA)
# =============================================================================
subject_hours = {
    "D-BIOLOGIA CELULAR Y MOLECULAR": 6,
    "D-PROCEDIMIENTOS BASICOS I": 2,
    "D-ANATOM CLINIC E IMAGENOL II TEO": 3,
    "D-ANATOM CLINIC E IMAGENOL II PRA": 2,
    "D-EMBRIOLOGIA TEO": 2,
    "D-EMBRIOLOGIA PRA": 1,
    "D-NEUROANATOMIA TEO": 2,
    "D-NEUROANATOMIA PRA": 1,
    "D-PENSAMIENTO CRITICO APLICADO": 3,
    "D-INTERAC EFECT EN SISTEMAS SO": 3,
}


# =============================================================================
# 4. Generar la lista de franjas horarias con el formato solicitado
# =============================================================================
def generate_time_slots():
    slots = []
    # Fase 1: Desde 07:00 hasta 17:50 con bloques de 60 min + 5 min break
    start_phase1 = datetime.strptime("07:00", "%H:%M")
    threshold = datetime.strptime("17:50", "%H:%M")
    slot_duration1 = timedelta(minutes=60)
    break_duration1 = timedelta(minutes=5)

    current = start_phase1
    while current + slot_duration1 <= threshold:
        slot_end = current + slot_duration1
        slots.append((current.strftime("%H:%M"), slot_end.strftime("%H:%M")))
        current = slot_end + break_duration1

    # Fase 2: desde threshold hasta ~21:49 con 59 min + 1 min break
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
        if not (end_dt <= row_start_dt or start_dt >= row_end_dt):
            return False
    return True


# =============================================================================
# 7. Set global de asignaciones (sin la materia) para que no se repita (día, franja, aula)
# =============================================================================
used_assignments = set()  # (day, slot_start, slot_end, aula)


# =============================================================================
# 8. Función para verificar si existen 5 franjas consecutivas ocupadas en un día
# =============================================================================
def exceeds_4_consecutive_any_class(schedule_df, day):
    assigned_series = schedule_df[day] != ""
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
# 9. Función que genera un horario cumpliendo:
#    - Máximo 2 horas por materia en un mismo día.
#    - Si se asignan dos horas en el mismo día, deben ser consecutivas y en el mismo aula.
#    - Solo se permite 1 hora individual (no consecutiva) por materia en un mismo día.
# =============================================================================
def generate_schedule():
    # Crear el DataFrame de salida
    index_slots = [f"{s}-{e}" for s, e in time_slots]
    schedule_df = pd.DataFrame(
        index=index_slots, columns=["Hora Inicio", "Hora Fin"] + days
    )
    for i, (s, e) in enumerate(time_slots):
        schedule_df.loc[schedule_df.index[i], "Hora Inicio"] = s
        schedule_df.loc[schedule_df.index[i], "Hora Fin"] = e
    for d in days:
        schedule_df[d] = ""

    # Diccionario para almacenar, por día y por materia, las franjas asignadas y el aula (para garantizar la consecutividad)
    # La estructura es: subject_day_info[day][subject] = {"slots": [índices de time_slots], "aula": aula asignada o None}
    subject_day_info = {
        day: {subject: {"slots": [], "aula": None} for subject in subject_hours.keys()}
        for day in days
    }

    for subject, hours_needed in subject_hours.items():
        candidate_aulas = subject_classrooms.get(subject, [])
        while hours_needed > 0:
            assigned_in_iteration = False

            day_order = days[:]
            random.shuffle(day_order)

            for day in day_order:
                info = subject_day_info[day][subject]
                # Si ya hay 2 franjas asignadas en el día, no se puede asignar más
                if len(info["slots"]) == 2:
                    continue

                # Caso A: Aún no se asigna ninguna franja en el día
                if len(info["slots"]) == 0:
                    # Primero, intentar asignar un bloque de 2 horas consecutivas (si se requieren al menos 2 horas)
                    if hours_needed >= 2:
                        for i in range(len(time_slots) - 1):
                            row_label1 = f"{time_slots[i][0]}-{time_slots[i][1]}"
                            row_label2 = f"{time_slots[i+1][0]}-{time_slots[i+1][1]}"
                            if (
                                schedule_df.at[row_label1, day] != ""
                                or schedule_df.at[row_label2, day] != ""
                            ):
                                continue
                            aula_ok = None
                            for aula in candidate_aulas:
                                combo1 = (day, time_slots[i][0], time_slots[i][1], aula)
                                combo2 = (
                                    day,
                                    time_slots[i + 1][0],
                                    time_slots[i + 1][1],
                                    aula,
                                )
                                if (
                                    combo1 in used_assignments
                                    or combo2 in used_assignments
                                ):
                                    continue
                                if not is_aula_available(
                                    aula, day, time_slots[i][0], time_slots[i][1], df
                                ):
                                    continue
                                if not is_aula_available(
                                    aula,
                                    day,
                                    time_slots[i + 1][0],
                                    time_slots[i + 1][1],
                                    df,
                                ):
                                    continue
                                aula_ok = aula
                                break
                            if aula_ok is not None:
                                # Asignar bloque de 2 horas
                                schedule_df.at[row_label1, day] = (
                                    f"{subject} - {aula_ok}"
                                )
                                schedule_df.at[row_label2, day] = (
                                    f"{subject} - {aula_ok}"
                                )
                                used_assignments.add(
                                    (day, time_slots[i][0], time_slots[i][1], aula_ok)
                                )
                                used_assignments.add(
                                    (
                                        day,
                                        time_slots[i + 1][0],
                                        time_slots[i + 1][1],
                                        aula_ok,
                                    )
                                )
                                if exceeds_4_consecutive_any_class(schedule_df, day):
                                    # Revertir asignación
                                    schedule_df.at[row_label1, day] = ""
                                    schedule_df.at[row_label2, day] = ""
                                    used_assignments.discard(
                                        (
                                            day,
                                            time_slots[i][0],
                                            time_slots[i][1],
                                            aula_ok,
                                        )
                                    )
                                    used_assignments.discard(
                                        (
                                            day,
                                            time_slots[i + 1][0],
                                            time_slots[i + 1][1],
                                            aula_ok,
                                        )
                                    )
                                    continue
                                else:
                                    info["slots"] = [i, i + 1]
                                    info["aula"] = aula_ok
                                    hours_needed -= 2
                                    assigned_in_iteration = True
                                    break
                        if assigned_in_iteration:
                            break

                    # Si no se pudo asignar bloque de 2, intentar asignar 1 hora individual (única por día)
                    for i in range(len(time_slots)):
                        row_label = f"{time_slots[i][0]}-{time_slots[i][1]}"
                        if schedule_df.at[row_label, day] != "":
                            continue
                        for aula in candidate_aulas:
                            combo = (day, time_slots[i][0], time_slots[i][1], aula)
                            if combo in used_assignments:
                                continue
                            if not is_aula_available(
                                aula, day, time_slots[i][0], time_slots[i][1], df
                            ):
                                continue
                            schedule_df.at[row_label, day] = f"{subject} - {aula}"
                            used_assignments.add(combo)
                            if exceeds_4_consecutive_any_class(schedule_df, day):
                                schedule_df.at[row_label, day] = ""
                                used_assignments.discard(combo)
                                continue
                            else:
                                info["slots"] = [i]
                                info["aula"] = aula
                                hours_needed -= 1
                                assigned_in_iteration = True
                                break
                        if assigned_in_iteration:
                            break
                    if assigned_in_iteration:
                        break

                # Caso B: Ya hay 1 franja asignada en el día (hora individual) y se busca extender a bloque consecutivo
                elif len(info["slots"]) == 1:
                    existing_index = info["slots"][0]
                    possible_indices = []
                    if existing_index > 0:
                        possible_indices.append(existing_index - 1)
                    if existing_index < len(time_slots) - 1:
                        possible_indices.append(existing_index + 1)
                    for candidate_index in possible_indices:
                        row_label_candidate = f"{time_slots[candidate_index][0]}-{time_slots[candidate_index][1]}"
                        if schedule_df.at[row_label_candidate, day] != "":
                            continue
                        # El aula debe ser la misma que la ya asignada
                        aula_assigned = info["aula"]
                        combo = (
                            day,
                            time_slots[candidate_index][0],
                            time_slots[candidate_index][1],
                            aula_assigned,
                        )
                        if combo in used_assignments:
                            continue
                        if not is_aula_available(
                            aula_assigned,
                            day,
                            time_slots[candidate_index][0],
                            time_slots[candidate_index][1],
                            df,
                        ):
                            continue
                        schedule_df.at[row_label_candidate, day] = (
                            f"{subject} - {aula_assigned}"
                        )
                        used_assignments.add(combo)
                        if exceeds_4_consecutive_any_class(schedule_df, day):
                            schedule_df.at[row_label_candidate, day] = ""
                            used_assignments.discard(combo)
                            continue
                        else:
                            info["slots"].append(candidate_index)
                            info["slots"].sort()
                            hours_needed -= 1
                            assigned_in_iteration = True
                            break
                    if assigned_in_iteration:
                        break

            if not assigned_in_iteration:
                return None

    return schedule_df


# =============================================================================
# 10. Generar 20 horarios (en el ejemplo se generan 5)
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
    print("No se alcanzaron los 20 horarios.")
else:
    print("¡Se generaron 20 horarios exitosamente!")

# =============================================================================
# 11. Exportar
# =============================================================================
if len(unique_schedules) > 0:
    output_file = "Horarios_Unicos_Semestre2.xlsx"
    with pd.ExcelWriter(output_file) as writer:
        for idx, sched in enumerate(unique_schedules, start=1):
            sched.to_excel(writer, sheet_name=f"Horario_{idx}", index=False)

    print(f"\nHorarios generados y guardados en '{output_file}'")
