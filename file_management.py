# Modulo encargado de obtener la tabla principal de cada archivo

import pandas as pd

# Leer archivo excel
def read_excel(sede):
    return pd.read_excel(sede['file_address'], sheet_name=sede['sheet_name'], header=None)

# Obtener la tabla principal
def filter_noise(df, column_index):
    # Tomar la fila 1 como nombres de columns
    df.columns = df.iloc[column_index]
    df.columns = df.columns.str.strip()

    # Eliminar filas innecesarias
    df = df.iloc[column_index+1:].reset_index(drop=True)

    return df

# Eliminar columnas innecesarias
def get_relevant_columns(df, sede):
    return df[sede['relevant_columns']]

# CAMANA gestor
def camana_management(df_cam, CAMANA):
    df = filter_noise(df_cam, 0)
    df = get_relevant_columns(df, CAMANA)

    return df

# PEDREGAL gestor
def pedregal_management(df_ped, PEDREGAL):
    df = filter_noise(df_ped, 0)
    df = get_relevant_columns(df, PEDREGAL)
    
    return df

# CHALA gestor
def chala_management(df_cha, CHALA):
    df = filter_noise(df_cha, 0)
    df = get_relevant_columns(df, CHALA)
    
    return df

# ATICO gestor
def atico_management(df_ati, ATICO):
    df = filter_noise(df_ati, 2)
    df = get_relevant_columns(df, ATICO)
    df = df.dropna(axis=1, how='all') # Eliminar columnas vacias (0 non-null)
    df = df[df.notna().all(axis=1)]
    
    return df

def main(
        CAMANA,
        PEDREGAL,
        CHALA,
        ATICO
):
    # Leer excel
    df_cam = read_excel(CAMANA)
    df_ped = read_excel(PEDREGAL)
    df_cha = read_excel(CHALA)
    df_ati = read_excel(ATICO)

    # Obtener datos de la tabla principal
    df_cam = camana_management(df_cam, CAMANA)
    df_ped = pedregal_management(df_ped, PEDREGAL)
    df_cha = chala_management(df_cha, CHALA)
    df_ati = atico_management(df_ati, ATICO)

    return df_cam, df_ped, df_cha, df_ati