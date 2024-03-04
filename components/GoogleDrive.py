from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

directorio_credenciales = 'data\credentials_module.json'

# Iniciar sesión
def login():
    gauth = GoogleAuth()
    gauth.LoadCredentialsFile(directorio_credenciales)
    
    if gauth.access_token_expired:
        gauth.Refresh()
        gauth.SaveCredentialsFile(directorio_credenciales)
    else:
        gauth.Authorize()

    return GoogleDrive(gauth)

def descarga_archivo(nombre_archivo, ruta_descarga):
    credenciales = login()
    archivos = credenciales.ListFile({'q': f"title = '{nombre_archivo}' and trashed = false"}).GetList()

    if archivos:
        archivo = archivos[0]
        archivo.GetContentFile(ruta_descarga + archivo['title'])
        print(f"Archivo '{archivo['title']}' descargado exitosamente.")
    else:
        print(f"No se encontró el archivo '{nombre_archivo}' en Google Drive.")

Maestra_Clientes = GoogleDrive.descarga_archivo('Maestra_Clientes.db', RUTA_ESPEC)
RUTA_PPAL = os.path.dirname(__file__)
RUTA_ESPEC = os.path.join(RUTA_PPAL,'data/')