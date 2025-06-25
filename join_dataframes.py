# Modulo encargado de unir los archivo y limpiarlos para su ingreso al excel
import pandas as pd
import log_management as log

def clean_df(df_total, INVALID_CDTRA):
    # Filtramos las filas donde la columna 1 no tenga valores inválidos
    return df_total[~df_total[1].astype(str).str.strip().isin(INVALID_CDTRA)]

def main(
        PROJECT_ADDRESS,
        CAMANA, 
        PEDREGAL, 
        CHALA, 
        ATICO,
        df_cam, 
        df_ped, 
        df_cha, 
        df_ati,
        INVALID_CDTRA
):
    sede_dataframes = [
        (CAMANA["name"], df_cam),
        (PEDREGAL["name"], df_ped),
        (CHALA["name"], df_cha),
        (ATICO["name"], df_ati),
    ]

    df_total = pd.DataFrame()

    # Juntar todas las sedes
    for name, df in sede_dataframes:
        if df is None or df.empty:
            continue
        df = df.copy()
        df.insert(0, "SEDE", name)  # Insertar como primera columna
        df.columns = range(df.shape[1])  # Ignorar nombres originales de columnas
        df_total = pd.concat([df_total, df], ignore_index=True)
    log.write_log(PROJECT_ADDRESS, '[*] Dataframes concatenados')

    df_total = clean_df(df_total, INVALID_CDTRA)

    log.write_log(PROJECT_ADDRESS, '[*] Dataframe total procesado')

    return df_total