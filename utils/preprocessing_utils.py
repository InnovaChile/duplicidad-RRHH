import os
import pandas as pd
import re
from dateutil.relativedelta import relativedelta    

import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def validar_rut(rut: str) -> bool:
    rut = str(rut)
    # Limpiar el RUT, eliminando puntos y guion
    rut = rut.replace(".", "").replace("-", "").replace(" ", "")
    
    # Asegurarse de que el RUT tiene al menos 2 caracteres
    if len(rut) < 2:
        return False

    # Separar el número del dígito verificador
    numero = rut[:-1]
    digito_verificador = rut[-1].upper()

    # Validar que la parte numérica sea realmente un número
    if not numero.isdigit():
        return False

    # Convertir la parte numérica a una lista de enteros
    numero = list(map(int, numero))

    # Secuencia de multiplicadores para el cálculo
    secuencia = [2, 3, 4, 5, 6, 7]

    # Calcular la suma ponderada
    suma = 0
    for i in range(len(numero)):
        suma += numero[-(i + 1)] * secuencia[i % len(secuencia)]

    # Calcular el dígito verificador según el módulo 11
    resto = suma % 11
    digito_calculado = 11 - resto

    # Ajustar el dígito verificador calculado
    if digito_calculado == 11:
        digito_calculado = "0"
    elif digito_calculado == 10:
        digito_calculado = "K"
    else:
        digito_calculado = str(digito_calculado)

    # Comparar el dígito verificador calculado con el proporcionado
    return digito_calculado == digito_verificador

def replicate_rows(row):
    # Se realiza el comentado de este código por el errot generado en meses = int(row['meses'].iloc[0])
    # row = pd.DataFrame(row).T
    # rows = row.loc[row.index.repeat(row['meses'])].reset_index(drop=True)

    # fecha_inicio = row["fecha_de_inicio"].iloc[0]
    # rows["periodo"] = [fecha_inicio + relativedelta(months=i) for i in range(len(rows))]

    # return rows
    row = pd.DataFrame(row).T
    try:
        meses = int(row['meses'].iloc[0])
        if pd.isna(meses) or meses < 1:
            print(f"Advertencia: valor inválido en 'meses'. Usando 1. Fila: {row.to_dict()}")
            meses = 1
    except Exception:
        print(f"Error al leer 'meses' en fila: {row.to_dict()}, usando 1.")
        meses = 1

    rows = row.loc[row.index.repeat(meses)].reset_index(drop=True)
    fecha_inicio = row["fecha_de_inicio"].iloc[0]
    rows["periodo"] = [fecha_inicio + relativedelta(months=i) for i in range(len(rows))]
    return rows

# Crear un DataFrame expandido con periodo
def generar_periodos(group, fecha_inicio): 

    fechas = list()

    for index, row in group.iterrows():
        fecha_proyecto = fecha_inicio

        if row["tipo"] == "Valida":
            fecha_proyecto = fecha_inicio + relativedelta(months=row["meses_crea"])

        fechas.append(fecha_proyecto)

    # Reconstruimos group añadiendo la fecha de inicio

    group["fecha_de_inicio"] = fechas

    expanded = pd.DataFrame()

    if len(group) > 1:
        for index, row in group.iterrows():
            row = replicate_rows(row)
            expanded = pd.concat([expanded, row], ignore_index=True)
    else:
        # Repetir filas según la columna 'meses'
        expanded = group.loc[group.index.repeat(group['meses'])].reset_index(drop=True)

        # Crear la columna 'periodo' con fechas
        expanded['periodo'] = [
            fecha_inicio + relativedelta(months=i) for i in range(len(expanded))
        ]
        
    # Eliminar la columna 'meses'
    expanded = expanded.drop(columns=['meses'])

    # if len(group) == 1:
    #     display(expanded)

    return expanded

def calcular_valor(df: pd.DataFrame) -> int:
    rows = df[df.eq("Mes de Inicio").any(axis=1)].index
    df = df.iloc[rows[0]:rows[1]-1].reset_index(drop=True)
    # Usar la primera fila como los nombres de las columnas
    df.columns = df.iloc[0]
    # Eliminar la primera fila (que ahora se ha usado como encabezado)
    df = df.drop(0).reset_index(drop=True)
    # Eliminar las columnas que tengan todos los valores NaN
    df = df[["Mes de Inicio", "Mes de Término"]].dropna().reset_index()
    # Obtenemos el máximo de la columna L y el mínimo de la columna K
    try:
        max_L = df['Mes de Término'].max()
        min_K = df['Mes de Inicio'].min()
    except:
        df = df[:-1]
        max_L = df['Mes de Término'].max()
        min_K = df['Mes de Inicio'].min()
    # Aplicamos las condiciones de la fórmula
    if max_L == 0:
        return 0
    elif max_L == min_K or max_L - min_K == 1:
        return 1
    else:
        return max_L - min_K + 1

def df_crea_valida(df: pd.DataFrame, id: str, meses_crea: int) -> pd.DataFrame:

    start_row = df[df.eq('Nombre y Apellido').any(axis=1)].index
    columns_to_keep = ['Nombre y Apellido', 'Rut', 'Dedicación proyecto:\nhoras al mes [A](*)', 
            'N° Meses [B]', 'Costo unitario ($)/HH', 
            'Aporte Innova Chile\n(Subsidio) ($)']

    df_crea = df.iloc[start_row[0]:start_row[1]-1]
    df_crea.reset_index(drop=True, inplace=True)
    df_crea.columns = df_crea.iloc[0]
    df_crea = df_crea.drop(0).reset_index(drop=True) 
    df_crea = df_crea.dropna(subset=['Nombre y Apellido', "Rut"])
    df_crea = df_crea[columns_to_keep]
    df_crea["tipo"] = "Crea" 

    df_valida = df.iloc[start_row[1]:]
    df_valida.reset_index(drop=True, inplace=True)
    df_valida.columns = df_valida.iloc[0]
    df_valida = df_valida.drop(0).reset_index(drop=True) 
    df_valida = df_valida.dropna(subset=['Nombre y Apellido', "Rut"])
    df_valida = df_valida[columns_to_keep]
    df_valida["tipo"] = "Valida"

    df_final = pd.concat([df_crea, df_valida])
    df_final.columns = ("nombre", 'rut', 'horas_mes', 'meses', 'costo_hora', 'aporte_innova', 'tipo')
    df_final['id_postulacion'] = id
    df_final["rut"] = df_final["rut"].str.replace(".", "")
    df_final["rut"] = df_final["rut"].str.replace("-", "")

    # Agregamos la columna de meses_crea, si su valor en tipo es Valida, si no se deja en 0

    df_final["meses_crea"] = df_final.apply(lambda x: meses_crea if x["tipo"] == "Valida" else 0, axis=1)

    return df_final

def generador_postulantes(file_path: str) -> pd.DataFrame:
    """El proceso es distinto si corresponde a Crea, Valida u otro"""

    final_df = pd.DataFrame()

    postulantes = os.listdir(file_path)

    for postulante in postulantes:
        print(f"Extrayendo postulantes de {postulante}")

        path = os.path.join(file_path, postulante)
        try:
            df = pd.read_excel(path, sheet_name="RRHH")

            # Añadir el nombre del codigo_postulacion como columna
            df['codigo_postulacion'] = postulante.replace('.xlsx', '') # Añadir el nombre del codigo_postulacion como columna
            
            tipo = postulante.split(".")[0].split("-")[0]
        
            if "IATS" in tipo or "CH" in tipo:
                pass
            elif re.search(r"CV",tipo):
                id = postulante.split(".")[0]
                meses_crea = calcular_valor(pd.read_excel(path, sheet_name="PLAN DE TRABAJO"))
                df = df_crea_valida(df, id, meses_crea)
                df['codigo_postulacion'] = postulante  # Asegura que 'codigo_postulacion' persista
                final_df = pd.concat([final_df, df], ignore_index=True)
                continue
            
            elif re.search(r"PATI",tipo):
                try:
                    id = postulante.split(".")[0]
                    meses_crea = calcular_valor(pd.read_excel(path, sheet_name="PLAN DE TRABAJO"))
                    df = df_crea_valida(df, id, meses_crea)
                    df['codigo_postulacion'] = postulante  # Asegura que 'codigo_postulacion' persista
                    final_df = pd.concat([final_df, df], ignore_index=True)
                    continue
                except:
                    pass

            # Procesar DataFrame general
            start_row = df[df.eq('Nombre y Apellido').any(axis=1)].index[0]
            df = df.loc[start_row:].reset_index(drop=True)
            df.columns = df.iloc[0]
            df = df.drop(0).reset_index(drop=True)
            df = df.dropna(axis=1, how='all')
            df = df.dropna(subset=['Nombre y Apellido', "Rut"])

            columns_to_keep = ['Nombre y Apellido', 'Rut', 'Dedicación proyecto:\nhoras al mes [A](*)', 
                            'N° Meses [B]', 'Costo unitario ($)/HH', 
                            'Aporte Innova Chile\n(Subsidio) ($)']
            df = df[columns_to_keep]

            df.columns = ["nombre", 'rut', 'horas_mes', 'meses', 'costo_hora', "aporte_innova"]

            if re.search(r"CH", tipo):
                df = df[~df['nombre'].str.contains("Nombre y Apellido", na=False)].reset_index(drop=True)

            df['codigo_postulacion'] = postulante  # Vuelve a agregar 'codigo_postulacion'
            df['id_postulacion'] = postulante.split(".")[0]
            df.loc[:, "rut"] = df["rut"].str.replace(".", "")
            df.loc[:, "rut"] = df["rut"].str.replace("-", "")
            df["tipo"] = "Otro"
            df["meses_crea"] = 0

            final_df = pd.concat([final_df, df], ignore_index=True)
        except Exception as e:
            print(f"Ocurrió un error al procesar {postulante}: {e}")

    print("Postulantes extraidos")
    
    return final_df

def filtrar_db(df: pd.DataFrame) -> pd.DataFrame:
    # Estandarizamos los nombres de las columnas
    df.columns = (
        df.columns.str.strip()
        .str.lower()
        .str.replace(" ", "_")
        .str.replace("(", "")
        .str.replace(")", "")
    )
    # Eliminamos los tildes de las columnas
    df.columns = (
        df.columns.str.normalize("NFKD")
        .str.encode("ascii", errors="ignore")
        .str.decode("utf-8")
    )
    
    # Nos quedamos solo con las columnas que nos sirven
    columns_to_keep = [
        "codigo_de_proyecto",
        "rut_recurso_presupuesto",
        "periodo_presupuesto",
        "horas",
        "nombre_fuente_de_financiamiento"
    ]
    df = df[columns_to_keep].copy()  # Hacemos una copia explícita
    # Pasamos la columna a tipo datetime
    df.loc[:, "periodo_presupuesto"] = pd.to_datetime(
        df["periodo_presupuesto"], errors="coerce"
    )
    # Filtramos para obtener solo valores mayor a 2019 en la Fecha Postulacion Inicial
    df = df[df["periodo_presupuesto"] > pd.Timestamp(2019, 1, 1)]
    df = df.dropna()

    # Le quitamos los puntos y guiones al rut
    df["rut_recurso_presupuesto"] = df["rut_recurso_presupuesto"].str.replace(
        ".", ""
    )
    df["rut_recurso_presupuesto"] = df["rut_recurso_presupuesto"].str.replace(
        "-", ""
    )
    # Renombramos las columnas
    df.columns = [
        "codigo_proyecto",
        "rut",
        "periodo",
        "horas",
        "fuente_de_financiamiento"]
    
    # Transformamos los ruts a string
    df["rut"] = df["rut"].astype(str)

    return df