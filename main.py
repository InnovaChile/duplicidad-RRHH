import os
import time
import pandas as pd
from datetime import datetime
import utils.preprocessing_utils as utils

import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.filterwarnings("ignore", category=FutureWarning)

def main():
    print("Iniciando...")

    # Iniciamos tiempo

    start_time = time.time()

    # Si no esta creada la carpeta Output, la creamos

    if not os.path.exists("Output"):
        print("Creando carpeta Output...")
        os.makedirs("Output")

    # Leer y filtrar DataBase
    try:
        database = utils.filtrar_db(pd.read_excel("Data\\DataBase.xlsx", sheet_name="Presupuesto"))
    except:
        database = utils.filtrar_db(pd.read_excel("Data\\DataBase.xlsx"))

    # Generar DataFrame con postulantes
    df = utils.generador_postulantes("Data\\Postulantes")
    #df = utils.generador_postulantes("Data\\Postulantes - funciona\\CVC")

        # Verificar que 'archivo' esté presente en el DataFrame
    if 'codigo_postulacion' not in df.columns:
        print("Error: La columna 'codigo_postulacion' no está presente en el DataFrame generado. Revise la función generador_postulantes.")
        exit(1)  # Finaliza el programa si 'codigo_postulacion' no está presente
    
    df['codigo_postulacion'] = df['codigo_postulacion'].str.replace('.xlsx', '', regex=False)

    # Separamos a los postulantes en dos grupos, los que tienes rut valido y los que no

    valido = df["rut"].apply(utils.validar_rut)
    invalido = ~valido

    # Postulantes con rut valido

    df_valido = df[valido]

    # Postulantes con rut invalido

    df_invalido = df[invalido]

    # Buscamos NaN en la columna "horas_mes"

    nan_horas = df_valido["horas_mes"].isna()

    # Los borramos de df_valido y lo dejamos en df_invalido

    df_invalido = pd.concat([df_invalido, df_valido[nan_horas]])

    # Borramos los NaN de df_valido

    df_valido = df_valido[~nan_horas]

    # Retornamos ambos DataFrames

    df_valido.to_csv("Output\\Postulantes_validos.csv", index=False)
    print("Se ha guardado el archivo Postulantes_validos.csv en la carpeta Output.")
    df_invalido.to_csv("Output\\Postulantes_invalidos.csv", index=False)
    print("Se ha guardado el archivo Postulantes_invalidos.csv en la carpeta Output.")

    # Adaptamos la estructura de los postulantes válidos.

    # Input de fecha inicial
    # Detenemos el tiempo para que el usuario pueda ingresar la fecha inicial
    aux = True


    # DESCOMENTAR DESPUES ESTO
    while aux:
        try:
            fecha_inicial = input("Ingrese la fecha inicial (YYYY-MM-DD): ")
            fecha_inicio = datetime.strptime(fecha_inicial, "%Y-%m-%d")
            aux = False
        except ValueError:
            print("Fecha inválida. Por favor, ingrese una fecha válida.")

    # Agrupamos por rut

    groups = df_valido.groupby("rut")

    # Aplicar la lógica a cada grupo
    expanded_groups = groups.apply(lambda x: utils.generar_periodos(x, fecha_inicio)).reset_index(drop=True)

    # Nos quedamos solo con id_postulacion, rut, periodo, horas_mes

    expanded_groups = expanded_groups[['id_postulacion', 'rut', 'periodo', 'horas_mes']]

    # Cambiamos el nombre de id_posulacion a codigo_proyecto

    expanded_groups = expanded_groups.rename(columns={"id_postulacion": "codigo_proyecto", "horas_mes": "horas"})

    # Añadimos la fuente de financiemiento, correspondiente a "Postulacion"

    expanded_groups["fuente_de_financiamiento"] = "Postulacion"

    # Filtrar la base de datos para que solo contenga los proyectos de los postulantes válidos en la base de datos

    ruts = df_valido['rut'].unique()
    database = database[database["rut"].isin(ruts)]

    merged_df = pd.concat([expanded_groups, database], ignore_index=True)

    # Agrupar por 'rut' y 'periodo' y crear listas de los códigos de proyecto y horas
    df_final = merged_df.groupby(['rut', 'periodo']).agg({
        'codigo_proyecto': lambda x: list(x),
        'fuente_de_financiamiento': lambda x: list(x),  # Convertir los códigos de proyecto en lista
        'horas': lambda x: list(x)  # Convertir las horas en lista
    }).reset_index()

    # Calcular la suma total de horas por fila
    df_final['total_horas'] = df_final['horas'].apply(sum)

    # Filtrar los ruts con más de 180 horas
    ruts_mas_180 = df_final[df_final['total_horas'] > 180]['rut'].unique()

    # Crear DataFrames separados
    df_mas_180 = df_final[df_final['rut'].isin(ruts_mas_180)]
    df_menos_180 = df_final[~df_final['rut'].isin(ruts_mas_180)]

    # Obtenemos todos los valores de df_mas_180 que son mayores a 180

    df_topes = df_mas_180[df_mas_180['total_horas'] > 180]

    # Usamos los nombres y ruts de df
    try:
        df_topes = pd.merge(df[["nombre", "rut", "codigo_postulacion"]], df_topes, on='rut', how='inner')
    except KeyError as e:
        print(f"Error al hacer el merge: {e}")
        print("Revisa si 'codigo_postulacion' o alguna de las columnas necesarias no están presentes.")
        exit(1)  # Finaliza el programa si hay un error

    # Reordenar las columnas para que 'codigo_postulacion' sea la primera

    columnas = ['codigo_postulacion'] + [col for col in df_topes.columns if col != 'codigo_postulacion']
    df_topes = df_topes[columnas]

    # Obtenemos la fecha actual

    fecha_actual = datetime.now().strftime("%Y-%m-%d")

    # Guardamos los DataFrames en archivos CSV y Excel con la fecha actual

    df_topes.to_csv(f"Output\\Topes_{fecha_actual}.csv", index=False)
    print(f"Se ha guardado el archivo Topes_{fecha_actual}.csv en la carpeta Output.")
    df_topes.to_excel(f"Output\\Topes_{fecha_actual}.xlsx", index=False)
    print(f"Se ha guardado el archivo Topes_{fecha_actual}.xlsx en la carpeta Output.")

    # Acaba el programa y muestra el tiempo de ejecución

    #print(f"El tiempo de ejecución fue de {time.time() - start_time} segundos.")
    input("Presione Enter para terminar...")

if __name__ == "__main__":
    main()