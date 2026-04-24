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

# --- GESTORES ---

def enlace_management(df_raw, LOCATION):
    df = filter_noise(df_raw, 0)
    df = get_relevant_columns(df, LOCATION)
    
    # Supongamos que df tiene 2 columnas en relevant_columns: [Transp_Compacto, Cliente_Compacto]
    # Ejemplo de contenido: "6014-TRANSPORTISTA" y "12345-CLIENTE"
    
    # 1. Separar Transportista (Columna 0 del DF filtrado)
    # n=1 asegura que solo divida en el primer guion
    split_transp = df.iloc[:, 0].astype(str).str.split('-', n=1, expand=True)
    
    # 2. Separar Cliente (Columna 1 del DF filtrado)
    split_cliente = df.iloc[:, 1].astype(str).str.split('-', n=1, expand=True)
    
    # 3. Construir el DataFrame con la estructura final requerida:
    # Col 0: Código Transp | Col 1: Código Cliente | Col 2: Código Cliente (rep) | Col 3: Nombre Cliente
    df_final = pd.DataFrame()
    df_final[0] = split_transp[0]  # Código Transportista
    df_final[1] = split_transp[1] # Código Cliente
    df_final[2] = split_cliente[0] # Código Cliente (para la columna 8 del excel)
    df_final[3] = split_cliente[1] # Nombre Cliente
    
    return df_final

def basis_management(df_raw, LOCATION):
    # Se asume que este archivo ya viene con las columnas separadas
    df = filter_noise(df_raw, 2)
    df = get_relevant_columns(df, LOCATION)
    
    # Limpieza estándar
    df = df.dropna(axis=1, how='all') 
    df = df[df.notna().all(axis=1)]
    
    # IMPORTANTE: Asegurar que las columnas se llamen 0, 1, 2, 3 para que coincidan con enlace_management
    df.columns = range(df.shape[1])
    
    return df

def main(LOCATIONS):
    df_loc_data = []

    for loc in LOCATIONS:
        df = read_excel(loc)
        if df is None or df.empty:
            continue

        if loc.get("enlace"): 
            processed_data = enlace_management(df, loc)
        else:
            processed_data = basis_management(df, loc)

        df_loc_data.append(processed_data)

    return df_loc_data