import pandas as pd
import re


def extraer_aula(texto):
    """
    Esta función busca en un string algo con el patrón 'UPO/123' o 'UPE/123', etc.
    Ajusta el regex a tus necesidades si tu formato de aula es distinto.
    """
    if not isinstance(texto, str):
        return None
    # Ejemplo de regex: busca algo como UPO/XXX o UPE/XXX
    patron = re.compile(r"(UPO/\d+|UPE/\d+)")
    match = patron.search(texto)
    if match:
        return match.group(0)
    return None


# Carga el archivo Excel
excel_file = "pruebas/horarios.xlsx"
xls = pd.ExcelFile(excel_file)

# Lista donde guardaremos todas las asignaciones de todas las hojas
df_completo = []

for sheet_name in xls.sheet_names:
    # Lee cada hoja
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    # Suponiendo que las columnas son exactamente:
    # [Hora Inicio, Hora Fin, Lunes, Martes, Miércoles, Jueves, Viernes, Sábado]
    # Ajusta los nombres si difieren.

    # Conviertimos a formato largo.
    # La idea es tener:
    # Hora Inicio | Hora Fin | Dia     | Materia/Aula
    # 07:00       | 08:00    | Lunes   | D-EMBRIOLOGIA TEO - UPO/522
    # 07:00       | 08:00    | Martes  | ...
    # etc.
    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"]
    df_largo = df.melt(
        id_vars=["Hora Inicio", "Hora Fin"],
        value_vars=dias,
        var_name="Dia",
        value_name="Asignatura",
    )

    # Extraer aula de la columna 'Asignatura'
    df_largo["Aula"] = df_largo["Asignatura"].apply(extraer_aula)

    # Añadir columna para saber de qué hoja viene (opcional, para más información)
    df_largo["Hoja"] = sheet_name

    # Agregar a la lista
    df_completo.append(df_largo)

# Unimos todo en un único DataFrame
df_completo = pd.concat(df_completo, ignore_index=True)

# Eliminamos filas donde la asignatura (o aula) esté vacía
df_completo.dropna(subset=["Asignatura", "Aula"], how="any", inplace=True)

# Convertir Hora Inicio y Fin a formato de tiempo (si no lo está ya).
# A veces, si Excel los reconoce como horas, puede que pandas ya los haya convertido.
# Si no, aquí un ejemplo forzado de conversión (ajusta según tu formato de hora real):
df_completo["Hora Inicio"] = pd.to_datetime(
    df_completo["Hora Inicio"], format="%H:%M"
).dt.time
df_completo["Hora Fin"] = pd.to_datetime(
    df_completo["Hora Fin"], format="%H:%M"
).dt.time


# Para facilitar la comparación de intervalos, podemos convertirlos a minutos desde la medianoche:
def tiempo_a_minutos(t):
    return t.hour * 60 + t.minute


df_completo["MinInicio"] = df_completo["Hora Inicio"].apply(tiempo_a_minutos)
df_completo["MinFin"] = df_completo["Hora Fin"].apply(tiempo_a_minutos)

# Ahora sí: detección de choques.
# Estrategia:
# 1) Agrupamos por (Dia, Aula).
# 2) En cada grupo, miramos sus intervalos [MinInicio, MinFin].
# 3) Determinamos si hay traslapes.

conflictos = []  # Aquí guardaremos registros de los choques


def intervalos_se_superponen(inicio1, fin1, inicio2, fin2):
    """
    Devuelve True si los intervalos [inicio1, fin1] y [inicio2, fin2] se superponen.
    """
    return not (fin1 <= inicio2 or fin2 <= inicio1)


for (dia, aula), grupo in df_completo.groupby(["Dia", "Aula"]):
    # Ordenamos por hora de inicio para ir comparando más fácilmente
    grupo_ordenado = grupo.sort_values(by=["MinInicio"])

    # Convertimos en lista para comparar pares
    filas = grupo_ordenado.to_dict("records")

    # Comparamos cada asignación contra las siguientes en el mismo grupo
    n = len(filas)
    for i in range(n):
        for j in range(i + 1, n):
            f1 = filas[i]
            f2 = filas[j]
            if intervalos_se_superponen(
                f1["MinInicio"], f1["MinFin"], f2["MinInicio"], f2["MinFin"]
            ):
                # Hay traslape
                conflictos.append(
                    {
                        "Dia": dia,
                        "Aula": aula,
                        "Hoja1": f1["Hoja"],
                        "Asignatura1": f1["Asignatura"],
                        "HoraInicio1": f1["Hora Inicio"],
                        "HoraFin1": f1["Hora Fin"],
                        "Hoja2": f2["Hoja"],
                        "Asignatura2": f2["Asignatura"],
                        "HoraInicio2": f2["Hora Inicio"],
                        "HoraFin2": f2["Hora Fin"],
                    }
                )

# Revisamos si hay conflictos
if conflictos:
    print("¡Se encontraron conflictos de aula!")
    df_conflictos = pd.DataFrame(conflictos)
    print(df_conflictos)
else:
    print("No se encontraron conflictos de aula.")
