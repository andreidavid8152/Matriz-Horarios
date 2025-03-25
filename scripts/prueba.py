import pandas as pd
import io

# Cargar los datos desde el texto proporcionado (simulando la carga desde un archivo CSV)
data = """
PERIODO	NRC	SIGLA	MATERIA_TITULO_PA	CODIGO_CAMPUS	EDIFICIO	SALA	HORA_INICIO	HORA_FIN	DAY_ID	TIPO_SALA	TIPO_SALA_DESC	FACULTAD
202510	3821	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/424	0805	0905	1	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3821	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/424	0910	1010	1	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3823	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/416	1750	1849	1	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3823	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/416	1850	1949	1	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3830	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/417	1330	1430	1	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3830	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/417	1435	1535	1	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3848	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/416	1435	1535	1	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3848	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/416	1540	1640	1	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3857	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/416	0700	0800	1	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3857	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/416	0805	0905	1	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3821	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/424	1435	1535	2	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3821	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/424	1540	1640	2	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3827	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/416	1015	1115	2	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3827	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/416	1120	1220	2	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3838	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/416	1645	1745	2	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3838	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/416	1750	1849	2	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3850	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/416	1225	1325	2	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3850	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/416	1330	1430	2	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3851	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/416	1435	1535	2	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3851	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/416	1540	1640	2	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3853	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/424	1120	1220	2	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3820	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/424	0910	1010	3	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3820	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/424	1015	1115	3	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3826	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/424	1120	1220	3	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3826	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/424	1225	1325	3	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3829	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/424	1330	1430	3	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3829	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/424	1435	1535	3	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3835	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPO	UPO/413	1645	1745	3	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3835	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPO	UPO/413	1750	1849	3	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3836	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/424	0700	0800	3	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3836	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/424	0805	0905	3	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3841	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/416	1015	1115	3	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3841	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/416	1120	1220	3	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3844	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/424	1645	1745	3	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3844	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/424	1750	1849	3	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3817	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/417	0805	0905	4	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3817	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/417	0910	1010	4	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3832	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPO	UPO/410	1540	1640	4	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3832	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPO	UPO/410	1645	1745	4	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3833	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPO	UPO/309	1330	1430	4	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3833	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPO	UPO/309	1435	1535	4	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3839	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/424	1645	1745	4	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3839	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/424	1750	1849	4	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3845	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/416	1435	1535	4	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3845	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/416	1540	1640	4	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3856	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/416	1015	1115	4	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3856	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/416	1120	1220	4	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3818	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/417	0910	1010	5	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3818	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/417	1015	1115	5	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3824	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/417	1750	1849	5	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3824	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/417	1850	1949	5	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3835	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/417	0700	0800	5	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3835	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/417	0805	0905	5	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3842	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/416	0700	0800	5	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3842	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/416	0805	0905	5	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3842	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPO	UPO/413	1015	1115	5	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3842	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPO	UPO/413	1120	1220	5	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3847	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/416	1435	1535	5	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3847	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/416	1540	1640	5	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3853	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/417	1435	1535	5	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3854	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/424	1225	1325	5	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
202510	3854	MEDZ4375	D-ANATOM CLINIC E IMAGENOL I	UP	UPE	UPE/424	1330	1430	5	MOLD	MORFOFUNCION O LAB. DESTREZAS	FAC. MEDICINA
"""  # Esto es una muestra, se puede expandir con el resto de datos si es necesario

# Convertir el texto a un DataFrame
df = pd.read_csv(io.StringIO(data), sep="\t")

# Convertir columnas de hora a formato datetime para facilitar los cálculos
df["HORA_INICIO"] = pd.to_datetime(df["HORA_INICIO"], format="%H%M")
df["HORA_FIN"] = pd.to_datetime(df["HORA_FIN"], format="%H%M")

# Calcular duración de cada clase en horas
df["DURACION_HORAS"] = (df["HORA_FIN"] - df["HORA_INICIO"]).dt.total_seconds() / 3600

""" # Agrupar por NRC y día para sumar las horas de clase por día para cada paralelo
horas_por_dia = df.groupby(["NRC", "DAY_ID"])["DURACION_HORAS"].sum().reset_index()

# Ordenar por DAY_ID ascendentemente
horas_por_dia_ordenado = horas_por_dia.sort_values(by="DAY_ID", ascending=True).reset_index(drop=True)

print(horas_por_dia_ordenado) """

# Calcular el total de horas semanales por NRC (sumando todas las DURACION_HORAS por NRC)
horas_por_semana = df.groupby("NRC")["DURACION_HORAS"].sum().reset_index()

# Opcional: ordenar por más horas a menos, si querés ver cuáles tienen más carga horaria
horas_por_semana = horas_por_semana.sort_values(
    by="DURACION_HORAS", ascending=False
).reset_index(drop=True)

print(horas_por_semana)
