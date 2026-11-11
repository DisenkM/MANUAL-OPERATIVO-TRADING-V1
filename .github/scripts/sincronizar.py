import os
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build

# 1. Autenticación con las credenciales de GitHub Secrets
creds_json = json.loads(os.environ['GOOGLE_CREDENTIALS'])
SCOPES = ["https://" + "www.googleapis.com" + "/auth/documents"]
creds = service_account.Credentials.from_service_account_info(creds_json, scopes=SCOPES)
service = build('docs', 'v1', credentials=creds)

DOC_ID = os.environ['DOC_ID'].strip()


# 2. Recopilar el código del repositorio
# Extensiones que quieres incluir (puedes agregar o quitar según tu proyecto)
EXTENSIONES_VALIDAS = ('.py', '.js', '.html', '.css', '.gs', '.json', '.md', '.txt', '.ts', '.sql')
CARPETAS_IGNORADAS = {'.git', '.github', 'node_modules', '__pycache__', 'venv'}

contenido_total = []

for root, dirs, files in os.walk('.'):
    # Filtrar carpetas ignoradas
    dirs[:] = [d for d in dirs if d not in CARPETAS_IGNORADAS]
    
    for file in sorted(files):
        if file.endswith(EXTENSIONES_VALIDAS) and file != 'sincronizar.py':
            ruta_completa = os.path.join(root, file)
            try:
                with open(ruta_completa, 'r', encoding='utf-8') as f:
                    codigo = f.read()
                    encabezado = f"\n{'='*50}\nARCHIVO: {ruta_completa}\n{'='*50}\n\n"
                    contenido_total.append(encabezado + codigo + "\n")
            except Exception as e:
                print(f"No se pudo leer {ruta_completa}: {e}")

texto_final = "".join(contenido_total)
if not texto_final.strip():
    texto_final = "El repositorio no contiene archivos de código válidos."

# 3. Obtener la longitud actual del Google Doc para borrarlo
doc = service.documents().get(documentId=DOC_ID).execute()
content = doc.get('body').get('content')
end_index = content[-1].get('endIndex') - 1

requests = []

# Si el documento no está vacío, borramos todo el texto anterior
if end_index > 1:
    requests.append({
        'deleteContentRange': {
            'range': {
                'startIndex': 1,
                'endIndex': end_index,
            }
        }
    })

# Insertamos el código actualizado
requests.append({
    'insertText': {
        'location': {'index': 1},
        'text': texto_final
    }
})

# 4. Ejecutar la actualización en Google Docs
service.documents().batchUpdate(documentId=DOC_ID, body={'requests': requests}).execute()
print("¡Sincronización con Google Docs completada con éxito!")
