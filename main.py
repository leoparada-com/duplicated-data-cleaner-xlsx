import pandas as pd
import os

# ==========================================================================================================
# LIMPIEZA DE LOS DATOS  HOJA: [nombre_hoja]
# ==========================================================================================================

# 1: leer archivo excel y cargar la hoja [nombre_hoja]  a un dataframe
# 2: verificar registros repetidos en hoja  [nombre_hoja]
# 3: Guardar archivo filtrado sin repetidos


'''
El script se encarga de la limpieza de datos de una hoja en un archivo de Excel. Primero, carga la hoja del archivo de Excel en un dataframe. A continuación, verifica si hay registros repetidos en la hoja, agrega un ID correlativo a cada registro para poder identificar y eliminar duplicados, y crea un nuevo dataframe con los registros duplicados. Después, filtra los registros únicos, crea un archivo Excel con los datos y lo guarda en el mismo directorio que el archivo principal. También se le agregaron funciones para mejorar la visualizacion de la data, como por ejemplo: imprimir información del dataframe, limpiar la pantalla de la consola y resaltar el texto en la salida.
'''


def filtrar_dataframe(df, columna, valor):
    df_nuevo = df[df[columna] == valor]
    return df_nuevo

def hay_registros_repetidos(df, columna):
    duplicados = df.duplicated(subset=columna)
    return duplicados.any()

def agregar_id_correlativo(df, columna):
    """
    Agrega una columna con ID correlativo a un dataframe, asignando el mismo ID a registros con datos repetidos en una columna.

    :param df: dataframe al que se agregará la columna con ID correlativo
    :param columna: nombre de la columna a analizar en busca de datos repetidos
    :return: el mismo dataframe de entrada con una nueva columna de ID correlativo
    """
    # Creamos una copia del dataframe de entrada para evitar modificar el original
    df_con_id = df.copy()

    # Agregamos una nueva columna al dataframe con un ID correlativo
    df_con_id['ID'] = df_con_id.groupby(columna).ngroup()

    # Devolvemos el dataframe con la nueva columna de ID correlativo
    return df_con_id

def ordenar_por_columna(df, columna):
    df_ordenado = df.sort_values(by=columna)
    return df_ordenado


def crear_dataframe_repetidos(df, columna):
    df_repetidos = df[df.duplicated(subset=columna, keep=False)]
    return df_repetidos

def filtrar_por_columna_id_unico(df, columna):
    valores_previos = set()
    df_nuevo = pd.DataFrame(columns=df.columns)
    
    for index, row in df.iterrows():
        valor = row[columna]
        if valor not in valores_previos:
            df_nuevo = df_nuevo.append(row, ignore_index=True)
            valores_previos.add(valor)
    
    return df_nuevo

def create_excel_file(df, file_name=None, file_path=None):
    """
    Esta función recibe un dataframe y un archivo de destino y crea un archivo Excel con los datos del dataframe.
    Si no se proporciona una ruta de archivo, se guarda en el mismo directorio que el archivo main.py.
    """
    if file_path is None:
        # Obtener el directorio actual del archivo main.py
        dir_path = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(dir_path, '')
    
    if file_name is None:
        file_name = 'data.xlsx'
        
    file_path = os.path.join(file_path, file_name+'.xlsx')
    
    writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()


def cargar_hoja(nombre_archivo, nombre_hoja=None):
    if nombre_hoja is None:
        # Obtener el nombre de la primera hoja en el archivo
        nombre_hoja = list(pd.read_excel(nombre_archivo, sheet_name=None).keys())[0]
    # Cargar el archivo con el nombre de la hoja especificada o la primera hoja del archivo
    df_report_aux = pd.read_excel(nombre_archivo, sheet_name=nombre_hoja)
    return df_report_aux


def print_destacado(str_texto):
    print('==========================================================================================================')
    print(str_texto)
    print('==========================================================================================================')
    return

def limpia_pantalla():
    # Limpiar la pantalla de la consola
    os.system('cls' if os.name == 'nt' else 'clear')
    return

def imprimir_info_dataframe(df_data, df_nombre):
    print('--------------------------------------------------------')
    print(f"Nombre del DataFrame: {df_nombre}")
    print(f"Número de registros: {len(df_data.index)}")
    print('--------------------------------------------------------')


# ==========================================================================================================
# LIMPIEZA DE LOS DATOS  HOJA: [nombre_hoja]
# ==========================================================================================================
limpia_pantalla()
print_destacado('LIMPIEZA DE DATOS')
print()
# ----------------------------------------------------------------------------------------------------
# 1: leer la hoja  [nombre_hoja] del archivo excel [nombre_archivo] y cargarlo a un dataframe
nombre_archivo = 'nombre_archivo.xlsx'
nombre_hoja    = 'nombre_hoja'
df_report_aux  = cargar_hoja(nombre_archivo, nombre_hoja)


imprimir_info_dataframe(df_report_aux,nombre_hoja)
print(df_report_aux.head())

# 2: verificar registros repetidos en hoja [nombre_hoja]
df_report_id                = agregar_id_correlativo(df_report_aux, 'Correo electrónico')
df_report_id_sort           = ordenar_por_columna(df_report_id,'ID')
df_report_id_sort_repetidos = crear_dataframe_repetidos(df_report_id_sort,'ID')


print_destacado('REPETIDOS')
imprimir_info_dataframe(df_report_id_sort_repetidos,'df_report_id_sort_repetidos')
print(df_report_id_sort_repetidos)
print()
print()

print_destacado('DATA FRAME SIN VALORES REPETIDO')
df_report_id_sort_filtrado = filtrar_por_columna_id_unico(df_report_id_sort,'ID')
imprimir_info_dataframe(df_report_id_sort_filtrado ,'df_report_id_sort_filtrado')
print(df_report_id_sort_filtrado)


# 4: Guardar archivo filtrado sin repetidos
create_excel_file(df_report_id_sort_filtrado,nombre_hoja)

print()
print()
