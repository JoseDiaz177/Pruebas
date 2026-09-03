from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
import pandas as pd
from google import genai
from dotenv import load_dotenv
import io
import os 
import pymupdf
import json 
import re


load_dotenv() # Carga las variables del archivo .env


client = genai.Client(api_key=(""))

app = FastAPI(title="API de Cotización Automatizada - Astillero")

app.add_middleware(
    CORSMiddleware, 
    allow_origins=["*"], # Permite que tu HTML local se conecte
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# Modelo de datos para la validación del Frontend (Vista 2)
class DatosBarcoConfirmados(BaseModel):
    cliente: str
    proyecto: str
    eslora: float
    manga: float
    puntal: float


@app.post("/api/extraer-datos-pdf")
async def extraer_datos_pdf(archivo_rfq: UploadFile = File(...)):
    if not archivo_rfq.filename.endswith('.pdf'):
        raise HTTPException(status_code=400, detail="El archivo debe ser un PDF")
    
    # 1. Lectura del PDF
    try:
        contenido_pdf = await archivo_rfq.read()
        doc = pymupdf.open(stream=contenido_pdf, filetype="pdf")
        texto_ocr = ""
        for pagina in doc:
            texto_ocr += pagina.get_text()
            
        if not texto_ocr.strip():
            print("ADVERTENCIA: El PDF parece ser una imagen sin texto digital.")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error leyendo el PDF: {str(e)}")
      

    # Se agregan los nuevos campos obligatorios al formato JSON esperado
    prompt = f"""
    Lee el siguiente texto extraído de un PDF. 
    Busca los valores correspondientes al renglón de "Nombre del Proyecto", "Tipo / Bandera / Clase", las medidas de "Eslora / Manga / Puntal" y "Localización Actual".
    Devuelve la respuesta ESTRICTAMENTE en este formato JSON, usando solo números para las medidas:
    {{
        "cliente": "",
        "proyecto": "",
        "tipo": "",
        "bandera": "",
        "clase": "",
        "eslora": 0.00,
        "manga": 0.00,
        "puntal": 0.00,
        "localizacion": ""
    }}
    Texto OCR: {texto_ocr}
    """

    try:
        response = client.models.generate_content(
            model='gemini-3.6-flash',
            contents=prompt
        )

        texto_crudo = response.text
        print("--- RESPUESTA CRUDA DE GEMINI ---")
        print(texto_crudo)

        match = re.search(r'\{.*\}', texto_crudo, re.DOTALL)
        if not match:
            raise ValueError("El modelo no devolvió una estructura JSON reconocible.")
            
        texto_json = match.group(0)
        datos_json = json.loads(texto_json)

        eslora_bruta = str(datos_json.get("eslora", "0"))
        # Extrae solo números y el punto decimal
        eslora_limpia = re.sub(r'[^\d.]', '', eslora_bruta)
        eslora_num = float(eslora_limpia) if eslora_limpia else 0.0

        if eslora_num >= 115.0:
            datos_json["dique"] = "Dique 5"
        else:
            datos_json["dique"] = "Dique 2"

        datos_json["eslora"] = eslora_num

        return datos_json

    except Exception as e:
        print(f"ERROR EN EL BACKEND: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error procesando los datos: {str(e)}")

        

@app.post("/api/generar-cotizacion")
async def generar_cotizacion(
    archivo_tender: UploadFile = File(...),
    # Recibe los datos validados desde el formulario de la Vista 2
    cliente: str = Form(...),
    eslora: float = Form(...),
    manga: float = Form(...)
):
    """
    Cruza el archivo Tender subido con el Tarifario base (Libro1.xlsx)
    utilizando Pandas y retorna los datos tabulares.
    """
    try:
        # 1. Leer el Tender subido a memoria
        contenido_tender = await archivo_tender.read()
        df_tender = pd.read_excel(io.BytesIO(contenido_tender))
        
        # 2. Leer el Tarifario Base local (asegúrate de que el archivo esté en la misma ruta)
        df_tarifario = pd.read_excel("Libro1.xlsx", sheet_name="Hoja1", header=2)
        
        # 3. Limpieza y estandarización de columnas para el Join
        # Asumiendo que df_tender tiene una columna 'Section' y df_tarifario tiene 'PART\nT N G'
        df_tender = df_tender.rename(columns={'Section': 'ID_Partida', 'Description': 'Descripcion_Cliente', 'Unit price': 'Precio_Cotizado'})
        df_tarifario = df_tarifario.rename(columns={'PART\nT N G': 'ID_Partida', 'Precio Unitario      Unit Price': 'Precio_Tarifa'})
        
        # Convertir a string para evitar errores de cruce por tipos de datos
        df_tender['ID_Partida'] = df_tender['ID_Partida'].astype(str).str.strip()
        df_tarifario['ID_Partida'] = df_tarifario['ID_Partida'].astype(str).str.strip()
        
        # 4. Cruce de Información (Left Join para mantener todo lo solicitado por el cliente)
        df_cruzado = pd.merge(df_tender, df_tarifario[['ID_Partida', 'D   e   s   c   r   i   p   t   i   o   n', 'Precio_Tarifa']], 
                              on='ID_Partida', how='left')
        
        # 5. Cálculos (Ejemplo básico)
        # Si la tarifa requiere un cálculo volumétrico, usar la eslora/manga recibida
        df_cruzado['Cantidad'] = 1 # Valor por defecto si no viene en el Tender
        df_cruzado['Precio_Final'] = df_cruzado['Precio_Tarifa'].fillna(0)
        df_cruzado['Costo_Total'] = df_cruzado['Cantidad'] * df_cruzado['Precio_Final']
        
        gran_total = df_cruzado['Costo_Total'].sum()
        
        # 6. Preparar respuesta (limitado a las columnas útiles para el frontend)
        resultado = df_cruzado[['ID_Partida', 'Descripcion_Cliente', 'Precio_Final', 'Costo_Total']].fillna('N/A').to_dict(orient='records')
        
        return {
            "cliente": cliente,
            "gran_total": gran_total,
            "partidas": resultado
        }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error procesando los archivos: {str(e)}")

@app.post("/api/exportar-excel")
async def exportar_excel(datos: list[dict]):
    
    df_final = pd.DataFrame(datos)
    
    # Crear un buffer en memoria para el archivo Excel
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_final.to_excel(writer, index=False, sheet_name='Cotizacion')
    
    buffer.seek(0)
    
    return StreamingResponse(
        buffer, 
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=cotizacion_final.xlsx"}
    )

@app.get("/")
def estado_servidor():
    return {"mensaje": "Backend de cotizaciones activo y funcionando."}
