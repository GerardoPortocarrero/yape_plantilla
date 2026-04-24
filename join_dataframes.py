# Modulo encargado de unir los archivo y limpiarlos para su ingreso al excel
import pandas as pd
import log_management as log

def clean_df(df_total, INVALID_CDTRA):
    # Filtramos las filas donde la columna 1 no tenga valores inválidos
    # Usamos loc para asegurar que la operación sea explícita
    return df_total[~df_total[1].astype(str).str.strip().isin(INVALID_CDTRA)]

def main(
        PROJECT_ADDRESS,
        LOCATIONS,      # Lista de diccionarios/objetos con la clave "name"
        df_loc,         # Lista de DataFrames en el mismo orden que LOCATIONS
        INVALID_CDTRA
):
    # Validamos que ambas listas tengan la misma longitud para evitar errores de índice
    if len(LOCATIONS) != len(df_loc):
        log.write_log(PROJECT_ADDRESS, '[!] Error: La cantidad de locaciones no coincide con los dataframes')
        return pd.DataFrame()

    df_total_list = [] # Usar una lista para colectar dfs es más eficiente que concat en bucle

    # Iteramos en paralelo usando zip
    for location, df in zip(LOCATIONS, df_loc):
        if df is None or df.empty:
            continue
            
        df_temp = df.copy()
        # Insertar nombre de la sede como primera columna
        df_temp.insert(0, "SEDE", location["name"])  
        
        # Estandarizar nombres de columnas a índices numéricos para evitar conflictos al concatenar
        df_temp.columns = range(df_temp.shape[1])
        
        df_total_list.append(df_temp)

    # Concatenamos todo al final (es más rápido que hacerlo dentro del for)
    if df_total_list:
        df_total = pd.concat(df_total_list, ignore_index=True)
        log.write_log(PROJECT_ADDRESS, '[*] Dataframes concatenados')
        
        # Limpieza de datos
        df_total = clean_df(df_total, INVALID_CDTRA)
        log.write_log(PROJECT_ADDRESS, '[*] Dataframe total procesado')
        
        return df_total
    else:
        log.write_log(PROJECT_ADDRESS, '[!] Advertencia: No se procesaron datos (Dataframes vacíos)')
        return pd.DataFrame()