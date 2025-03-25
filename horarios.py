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
    # Fase 1: Desde 08:05 hasta alcanzar el umbral para cambiar de patrón.
    start_phase1 = datetime.strptime("07:00", "%H:%M")
    threshold = datetime.strptime(
        "17:50", "%H:%M"
    )  # Se define para que la última franja con break termine a las 17:45 y se sume 5 minutos.
    slot_duration1 = timedelta(minutes=60)
    break_duration1 = timedelta(minutes=5)

    current = start_phase1
    # Generar franjas de fase 1 mientras el siguiente inicio sea menor que el threshold
    while current + slot_duration1 <= threshold:
        slot_end = current + slot_duration1
        # Formatear en HHMM
        slots.append((current.strftime("%H:%M"), slot_end.strftime("%H:%M")))
        current = slot_end + break_duration1
    # Fase 2: Desde threshold en adelante, franjas consecutivas sin break de 5, sino con 1 minuto de separación.
    slot_duration2 = timedelta(minutes=59)
    break_duration2 = timedelta(minutes=1)
    # Reiniciamos current al threshold (o al valor obtenido en fase1, que debería ser igual a threshold)
    current = threshold
    final_end = datetime.strptime(
        "21:49", "%H:%M"
    )  # Se fija el final según el ejemplo.
    while current + slot_duration2 <= final_end:
        slot_end = current + slot_duration2
        slots.append((current.strftime("%H:%M"), slot_end.strftime("%H:%M")))
        current = slot_end + break_duration2
    return slots


time_slots = generate_time_slots()
# Ahora time_slots contendrá las 13 franjas según el formato:
# [("0700","0905"), ("0910","1010"), ... , ("2050","2149")]

# =============================================================================
# 5. Definir la estructura base del horario (DataFrame)
# =============================================================================
days = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]


# =============================================================================
# 6. Función para verificar la disponibilidad de un aula en el DataFrame original
# =============================================================================
def is_aula_available(aula, day, slot_start, slot_end, df_existing):
    day_mapping = {"Lunes": 1, "Martes": 2, "Miércoles": 3, "Jueves": 4, "Viernes": 5}
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
# 7. Función que genera un horario usando el algoritmo greedy de asignación
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
        while hours_needed > 0:
            assigned_in_iteration = False
            day_order = days[:]
            random.shuffle(day_order)
            for day in day_order:
                if subject_schedule_count[subject][day] >= 2:
                    continue
                block_size = min(2 - subject_schedule_count[subject][day], hours_needed)
                # Buscar bloque consecutivo de tamaño block_size
                for i in range(len(time_slots) - block_size + 1):
                    slots_free = True
                    for j in range(block_size):
                        row_label = f"{time_slots[i+j][0]}-{time_slots[i+j][1]}"
                        if schedule_df.at[row_label, day] != "":
                            slots_free = False
                            break
                    if not slots_free:
                        continue
                    for aula in candidate_aulas:
                        available_for_all = True
                        for j in range(block_size):
                            slot_start, slot_end = time_slots[i + j]
                            if not is_aula_available(
                                aula, day, slot_start, slot_end, df
                            ):
                                available_for_all = False
                                break
                        if available_for_all:
                            for j in range(block_size):
                                row_label = f"{time_slots[i+j][0]}-{time_slots[i+j][1]}"
                                schedule_df.at[row_label, day] = f"{subject} - {aula}"
                            subject_schedule_count[subject][day] += block_size
                            hours_needed -= block_size
                            assigned_in_iteration = True
                            break
                    if assigned_in_iteration:
                        break
                # Si no se encontró bloque consecutivo, asignar individualmente
                if (
                    not assigned_in_iteration
                    and subject_schedule_count[subject][day] == 0
                ):
                    for i in range(len(time_slots)):
                        row_label = f"{time_slots[i][0]}-{time_slots[i][1]}"
                        if schedule_df.at[row_label, day] == "":
                            for aula in candidate_aulas:
                                if is_aula_available(
                                    aula, day, time_slots[i][0], time_slots[i][1], df
                                ):
                                    schedule_df.at[row_label, day] = (
                                        f"{subject} - {aula}"
                                    )
                                    subject_schedule_count[subject][day] += 1
                                    hours_needed -= 1
                                    assigned_in_iteration = True
                                    break
                        if assigned_in_iteration:
                            break
                if assigned_in_iteration:
                    break
            if not assigned_in_iteration:
                print(
                    f"Warning: No se encontró franja disponible para {subject} (horas restantes: {hours_needed})"
                )
                break
    return schedule_df


# =============================================================================
# 8. Generar 20 horarios únicos
# =============================================================================
unique_schedules = []
unique_hashes = set()
num_schedules_required = 20
attempts = 0
max_attempts = 1000  # Para evitar bucles infinitos

while len(unique_schedules) < num_schedules_required and attempts < max_attempts:
    attempts += 1
    sched = generate_schedule()
    rep = sched.to_csv(index=False)
    if rep not in unique_hashes:
        unique_schedules.append(sched)
        unique_hashes.add(rep)
        print(f"Horario {len(unique_schedules)} generado.")

if len(unique_schedules) < num_schedules_required:
    print("No se pudieron generar 20 horarios únicos en el número máximo de intentos.")
else:
    print("¡Se generaron 20 horarios únicos exitosamente!")

# =============================================================================
# 9. Exportar todos los horarios a un archivo Excel, cada uno en una hoja distinta
# =============================================================================
output_file = "Horarios_Unicos.xlsx"
with pd.ExcelWriter(output_file) as writer:
    for idx, sched in enumerate(unique_schedules, start=1):
        sheet_name = f"Horario_{idx}"
        sched.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"\nHorarios generados y guardados en '{output_file}'")
