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
# 9. Función que genera un horario con:
#    - máximo 2 horas consecutivas por materia en el mismo día
#    - no permitir 5 horas seguidas de clase (cualquier materia)
#    - no repetir combos (día, slot, aula) entre distintos horarios
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

    # Diccionario con conteo de horas (máx 2 consecutivas por materia)
    # Estructura: subject_schedule_count[day][subject] = 0 (inicial)
    subject_schedule_count = {
        day: {subject: 0 for subject in subject_hours.keys()} for day in days
    }

    for subject, hours_needed in subject_hours.items():
        candidate_aulas = subject_classrooms.get(subject, [])

        while hours_needed > 0:
            assigned_in_iteration = False

            day_order = days[:]
            random.shuffle(day_order)

            for day in day_order:
                if subject_schedule_count[day][subject] >= 2:
                    # Ya llegó a 2 horas consecutivas para esa materia en ese día
                    continue

                block_size = min(2 - subject_schedule_count[day][subject], hours_needed)

                # 1) Intentar encontrar un bloque consecutivo de block_size
                for i in range(len(time_slots) - block_size + 1):
                    # Revisar celdas libres en schedule_df
                    all_free = True
                    for j in range(block_size):
                        row_label = f"{time_slots[i+j][0]}-{time_slots[i+j][1]}"
                        if schedule_df.at[row_label, day] != "":
                            all_free = False
                            break
                    if not all_free:
                        continue

                    # Revisar cada aula candidata
                    aula_ok = None
                    assigned_slots = []
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
                            aula_ok = aula
                            break

                    if aula_ok:
                        # Asignar tentativamente
                        for j in range(block_size):
                            slot_start, slot_end = time_slots[i + j]
                            row_label = f"{slot_start}-{slot_end}"
                            schedule_df.at[row_label, day] = f"{subject} - {aula_ok}"
                            used_assignments.add((day, slot_start, slot_end, aula_ok))
                            assigned_slots.append((row_label, day, aula_ok))

                        # Revisar si con esto hay 5 horas seguidas
                        if exceeds_4_consecutive_any_class(schedule_df, day):
                            # Revertir
                            for r_lbl, d_col, a_asig in assigned_slots:
                                schedule_df.at[r_lbl, d_col] = ""
                                slot_start_str, slot_end_str = r_lbl.split("-")
                                used_assignments.discard(
                                    (d_col, slot_start_str, slot_end_str, a_asig)
                                )
                            # y no asignamos este bloque
                            continue
                        else:
                            # Válido
                            subject_schedule_count[day][subject] += block_size
                            hours_needed -= block_size
                            assigned_in_iteration = True
                            break  # de lazo de aulas

                    if assigned_in_iteration:
                        break  # de lazo de slots

                # 2) Si no se encontró bloque consecutivo (o block_size=1) y no asignamos nada
                if (not assigned_in_iteration) and (
                    subject_schedule_count[day][subject] < 2
                ):
                    for i in range(len(time_slots)):
                        row_label = f"{time_slots[i][0]}-{time_slots[i][1]}"
                        if schedule_df.at[row_label, day] == "":
                            slot_start, slot_end = time_slots[i]
                            assigned_this_slot = False
                            for aula in candidate_aulas:
                                combo = (day, slot_start, slot_end, aula)
                                if (
                                    combo not in used_assignments
                                ) and is_aula_available(
                                    aula, day, slot_start, slot_end, df
                                ):
                                    # Asignar tentativamente
                                    schedule_df.at[row_label, day] = (
                                        f"{subject} - {aula}"
                                    )
                                    used_assignments.add(combo)

                                    # Revisar 5 horas seguidas
                                    if exceeds_4_consecutive_any_class(
                                        schedule_df, day
                                    ):
                                        # revertir
                                        schedule_df.at[row_label, day] = ""
                                        used_assignments.discard(combo)
                                    else:
                                        subject_schedule_count[day][subject] += 1
                                        hours_needed -= 1
                                        assigned_in_iteration = True
                                        assigned_this_slot = True

                                    break  # salimos de lazo de aulas
                            if assigned_this_slot or hours_needed <= 0:
                                break  # salimos de lazo de slots

                if assigned_in_iteration or hours_needed <= 0:
                    break  # siguiente materia

            if not assigned_in_iteration:
                return None

    return schedule_df


# =============================================================================
# 10. Generar 20 horarios
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
