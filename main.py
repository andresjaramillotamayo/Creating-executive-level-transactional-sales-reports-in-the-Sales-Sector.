# Librerias
import os
import datetime
import pyodbc
import sqlite3
import duckdb
from components import GoogleDrive
import pandas as pd
import xlsxwriter

# Rutas para almacenar data
codigo_vendedor = input("Ingrese el código de vendedor: ")
RUTA_PPAL = os.path.dirname(__file__)
RUTA_ESPEC = os.path.join(RUTA_PPAL,'data/')
exc_date = datetime.datetime.today().strftime('%Y-%m-%d-%H-%M')
RUTA_RS = os.path.join(RUTA_PPAL, 'data', 'InformeVentas_GV{}_{}.xlsx'.format(codigo_vendedor,exc_date))


# Función principal
if __name__ == "__main__":
# Descarga de archivos de Google Drive mediante API
    # Transaccional_Ventas = GoogleDrive.descarga_archivo('Transaccional_Ventas.accdb', RUTA_ESPEC)
    # Maestra_Clientes = GoogleDrive.descarga_archivo('Maestra_Clientes.db', RUTA_ESPEC)

# Conversión .accdb a Dataframe
    # Conección a BD a través de pyodbc
    conn_accdb = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=data\Transaccional_Ventas.accdb;'
    conn_pyodbc = pyodbc.connect(conn_accdb)


    # Leer los datos en un DataFrame
    consulta_tv = 'SELECT * FROM Transaccional_Ventas'
    df_tv = pd.read_sql(consulta_tv, conn_pyodbc, index_col=None)

    conn_pyodbc.close()


    # Organizar Dataframe de Transaccional_Ventas
    df_tv = pd.DataFrame([x.split("|") for x in df_tv.values.flatten()])
    df_tv.columns = df_tv.iloc[0]
    df_tv = df_tv.iloc[1:].reset_index(drop=True)
    df_tv['Dtr_ClienteSap_Cod'] = pd.to_numeric(df_tv['Dtr_ClienteSap_Cod'], errors='coerce')
    df_tv['Dtr_ClienteSap_Cod'] = df_tv['Dtr_ClienteSap_Cod'].fillna(0).astype(int)
    

# Conversión .db a Dataframe
    # Conexión a la base de datos SQLite
    conn_db = sqlite3.connect('data\Maestra_Clientes.db')

    # Consulta SQL para extraer los datos
    consulta_dc = 'SELECT * FROM Detalle_Clientes'

    # Ejecutar la consulta y leer los datos en un DataFrame
    df_dc = pd.read_sql_query(consulta_dc, conn_db)
    df_dc['Cliente_Cod'] = pd.to_numeric(df_dc['Cliente_Cod'], errors='coerce')
    df_dc['Cliente_Cod'] = df_dc['Cliente_Cod'].fillna(0).astype(int)
    conn_db.close()

# Uso de libreria Duckdb para implementar consultas almacenables en objetos de Python
    
    # Conexión a la base de datos Duckdb
    conn_duckdb = duckdb.connect(database=':memory:')

    # Adecuación y cruce de bases de datos.
    q_key_tv = '''
    SELECT TV.*, DC.*, 
        CASE WHEN TV.Extraccion_Cod <> '' THEN SUBSTRING(TV.Extraccion_Cod, 1, 2) ELSE NULL END AS Dtr_Pais_Cod
    FROM df_tv TV
    INNER JOIN df_dc DC
        ON TV.Dtr_ClienteSap_Cod = DC.Cliente_Cod
    WHERE LOWER(CASE WHEN TV.Extraccion_Cod <> '' THEN SUBSTRING(TV.Extraccion_Cod, 1, 2) ELSE NULL END) = 
        LOWER(CASE WHEN DC.ClienteCountr <> '' THEN SUBSTRING(DC.ClienteCountr, 1, 2) ELSE NULL END)
    '''
    df_rs = duckdb.query(q_key_tv).to_df()
    conn_duckdb.close()

# Creación de archivo de Excel
    df_filtrado = df_rs.loc[df_rs['GrupoVendedoresCode'] == codigo_vendedor]
    try:
        writer = pd.ExcelWriter(RUTA_RS,engine='xlsxwriter')
        df_result = df_filtrado.to_excel(writer,index=False)
        writer.close()
        print('Informe de Ventas creado')
    except Exception as e:
        print("Error",e)
        print('Informe de Ventas no creado')


