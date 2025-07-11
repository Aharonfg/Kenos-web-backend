from fastapi import FastAPI, File, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import pandas as pd
import google.generativeai as genai
import time
import random
import io
import matplotlib.pyplot as plt
import seaborn as sns
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from collections import Counter
import os
import json
import tempfile

# Configura tu clave de API
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
if not GEMINI_API_KEY:
    raise ValueError("No se encontró la variable de entorno GEMINI_API_KEY")

genai.configure(api_key=GEMINI_API_KEY)

modelo = genai.GenerativeModel("gemini-1.5-pro")

app = FastAPI()
RESULTADOS_DIR = "resultados"
os.makedirs(RESULTADOS_DIR, exist_ok=True)
HISTORIAL_PATH = os.path.join(RESULTADOS_DIR, "historial_emociones.json")

app.add_middleware(
    CORSMiddleware, allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

EMOCIONES_VALIDAS = [
    "satisfacción", "frustración", "compromiso", "desmotivación",
    "estrés", "esperanza", "inseguridad", "aprecio", 
    "indiferencia", "agotamiento"
]

def filtrar_emocion_valida(texto):
    texto_limpio = texto.lower().strip()
    if texto_limpio in EMOCIONES_VALIDAS:
        return texto_limpio
    return "Error"

def obtener_emocion(texto, reintentos=3):
    prompt = (
        "Eres una persona de recursos humanos de una consultoría tecnológica."
        f"Frase: \"{texto}\". Responde SOLO con una palabra exacta de la lista: "
        + ", ".join(EMOCIONES_VALIDAS) + "."
    )
    for intento in range(reintentos):
        try:
            respuesta = modelo.generate_content(prompt)
            emocion = filtrar_emocion_valida(respuesta.text)
            time.sleep(random.uniform(1.5, 2.5))
            return emocion
        except Exception:
            time.sleep(3 + intento * 2)
    return "Error"

def construir_prompt(lista):
    texto = (
        "Eres una persona de RRHH. Para cada frase, responde con UNA palabra "
        "exacta entre: " + ", ".join(EMOCIONES_VALIDAS) + ".\n\n"
    )
    for i, f in enumerate(lista, 1):
        texto += f"{i}. \"{f}\"\n"
    texto += "\nResponde así:\n1. emoción\n2. emoción..."
    return texto

def obtener_emociones_lote(frases, reintentos=3):
    prompt = construir_prompt(frases)
    for intento in range(reintentos):
        try:
            respuesta = modelo.generate_content(prompt)
            lineas = respuesta.text.strip().split("\n")
            resultado = []
            for linea in lineas:
                partes = linea.split(". ", 1)
                emocion = filtrar_emocion_valida(partes[1] if len(partes) == 2 else "")
                resultado.append(emocion)
            time.sleep(random.uniform(1.5, 2.5))
            return resultado
        except Exception:
            time.sleep(3 + intento * 2)
    return ["Error"] * len(frases)

@app.post("/analizar")
async def analizar_excel(file: UploadFile = File(...)):
    try:
        contenido = await file.read()
        encuesta = pd.read_excel(io.BytesIO(contenido))
        columnas = [c for c in encuesta.columns if encuesta[c].dropna().astype(str).str.strip().any()]
        df_texto = encuesta[columnas].fillna("Sin respuesta").astype(str)

        respuestas = []
        tmp = []
        total = df_texto.size
        cont = 0
        for fila in df_texto.values:
            for celda in fila:
                cont += 1
                tmp.append(celda.strip())
                if len(tmp) == 10 or cont == total:
                    respuestas.extend(obtener_emociones_lote(tmp))
                    tmp = []

        datos = [respuestas[i:i+len(columnas)] for i in range(0, len(respuestas), len(columnas))]
        df_res = pd.DataFrame(datos, columns=columnas)
        ruta_excel = os.path.join(RESULTADOS_DIR, "emociones_resultado.xlsx")
        df_res.to_excel(ruta_excel, index=False, engine="openpyxl")

        wb = load_workbook(ruta_excel)
        ws = wb.active

        img_paths = []
        ancho = 10*96
        alto = 8*96
        paso = 40
        inicio = len(df_res) + 3

        for i, col in enumerate(columnas):
            plt.figure(figsize=(10, 8))
            sns.countplot(x=df_res[col], order=EMOCIONES_VALIDAS)
            plt.title(f"Distribución: {col}")
            plt.xticks(rotation=45)
            plt.tight_layout()

            tmpf = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
            plt.savefig(tmpf.name, dpi=96)
            plt.close()

            img = XLImage(tmpf.name)
            img.width = ancho
            img.height = alto
            ws.add_image(img, f"A{inicio + i*paso}")
            img_paths.append(tmpf.name)

        wb.save(ruta_excel)

        for p in img_paths:
            try:
                os.unlink(p)
            except Exception as e:
                print(f"No se pudo borrar {p}: {e}")

        emos = [e for e in df_res.values.flatten() if e in EMOCIONES_VALIDAS]
        if emos:
            mas = Counter(emos).most_common(1)[0][0]
            with open(os.path.join(RESULTADOS_DIR, "emocion_global.txt"), "w", encoding="utf-8") as f:
                f.write(mas)

        return FileResponse(ruta_excel, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            filename="emociones_resultado.xlsx")

    except Exception as ex:
        return {"error": str(ex)}

@app.get("/emocion")
def obtener_emocion_global():
    try:
        txt = os.path.join(RESULTADOS_DIR, "emocion_global.txt")
        xls = os.path.join(RESULTADOS_DIR, "emociones_resultado.xlsx")
        if not os.path.exists(txt) or not os.path.exists(xls):
            return {"error": "archivos necesarios no encontrados"}

        emocion = open(txt, encoding="utf-8").read().strip().lower()
        mapa = {
            "satisfacción": "😊", "frustración": "😠", "compromiso": "💪",
            "desmotivación": "😞", "estrés": "😣", "esperanza": "🌟",
            "inseguridad": "😟", "aprecio": "🤝", "indiferencia": "😐",
            "agotamiento": "😩"
        }
        punt = {
            "satisfacción": 1, "compromiso": 1, "aprecio": 1,
            "esperanza": 0.8, "indiferencia": 0,
            "inseguridad": -0.3, "estrés": -1,
            "desmotivación": -0.8, "agotamiento": -1,
            "frustración": -0.5
        }

        df = pd.read_excel(xls)
        emos = [e for e in df.values.flatten() if e in punt]
        if emos:
            media = sum(punt[e] for e in emos)/len(emos)
            pct = round(((media + 1)/2)*100, 2)
        else:
            pct = 0

        if pct <= 20: estado = "😠"
        elif pct <= 40: estado = "😕"
        elif pct <= 60: estado = "😐"
        elif pct <= 80: estado = "🙂"
        else: estado = "😄"

        actual = {
            "fecha": time.strftime("%Y-%m-%d %H:%M:%S"),
            "emocion": emocion if emocion in punt else "Error",
            "emoji": mapa.get(emocion, ""),
            "porcentaje_satisfaccion": pct,
            "estado_general": estado
        }

        hist = []
        if os.path.exists(HISTORIAL_PATH):
            hist = json.load(open(HISTORIAL_PATH, "r", encoding="utf-8"))

        hist.append(actual)
        hist = hist[-2:]
        json.dump(hist, open(HISTORIAL_PATH, "w", encoding="utf-8"), ensure_ascii=False, indent=2)

        return hist

    except Exception:
        return [{
            "fecha": time.strftime("%Y-%m-%d"),
            "emocion": "Error",
            "emoji": "",
            "porcentaje_satisfaccion": "No disponible",
            "estado_general": "❌"
        }]
