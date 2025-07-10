from fastapi import FastAPI, File, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import pandas as pd
import google.generativeai as genai
import time
import random
import traceback
import io
import matplotlib.pyplot as plt
import seaborn as sns
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from collections import Counter
import os
import json
import tempfile

# Configura la clave de la API de Gemini (pon tu clave real aquí)
GEMINI_API_KEY = "TU_API_KEY_AQUI"
genai.configure(api_key=GEMINI_API_KEY)
modelo = genai.GenerativeModel("gemini-1.5-pro")

app = FastAPI()
RESULTADOS_DIR = "resultados"
os.makedirs(RESULTADOS_DIR, exist_ok=True)
HISTORIAL_PATH = os.path.join(RESULTADOS_DIR, "historial_emociones.json")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

EMOCIONES_VALIDAS = [
    "satisfacción", "frustración", "compromiso", "desmotivación",
    "estrés", "esperanza", "inseguridad", "aprecio", "indiferencia", "agotamiento"
]

def filtrar_emocion_valida(texto):
    texto = texto.lower().strip()
    for emocion in EMOCIONES_VALIDAS:
        if emocion in texto:
            return emocion
    return "Error"

def obtener_emocion(texto, reintentos=3):
    prompt = (
        "Eres una persona de recursos humanos de una consultoría tecnológica llamada Kenos Technology. "
        "Debes indicar qué emoción describe mejor esta frase entre las siguientes: satisfacción, frustración, compromiso, desmotivación, "
        "estrés, esperanza, inseguridad, aprecio, indiferencia o agotamiento. "
        f"Frase: \"{texto}\". "
        "Responde SOLO con una palabra exacta de esa lista."
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

def construir_prompt(lista_de_frases):
    prompt = (
        "Eres una persona de recursos humanos de una consultoría tecnológica llamada Kenos Technology. "
        "Para cada frase, responde con solo una palabra entre: satisfacción, frustración, compromiso, desmotivación, "
        "estrés, esperanza, inseguridad, aprecio, indiferencia o agotamiento.\n\n"
    )
    for idx, frase in enumerate(lista_de_frases, 1):
        prompt += f"{idx}. \"{frase}\"\n"
    prompt += "\nResponde así:\n1. emoción\n2. emoción\n..."
    return prompt

def obtener_emociones_lote(frases, reintentos=3):
    prompt = construir_prompt(frases)
    for intento in range(reintentos):
        try:
            respuesta = modelo.generate_content(prompt)
            lineas = respuesta.text.strip().split("\n")
            emociones = []
            for linea in lineas:
                partes = linea.split(". ", 1)
                if len(partes) == 2:
                    emocion = filtrar_emocion_valida(partes[1])
                    emociones.append(emocion)
                else:
                    emociones.append("Error")
            time.sleep(random.uniform(1.5, 2.5))
            return emociones
        except Exception:
            time.sleep(3 + intento * 2)
    return ["Error"] * len(frases)

@app.post("/analizar")
async def analizar_excel(file: UploadFile = File(...)):
    try:
        contenido = await file.read()
        encuesta = pd.read_excel(io.BytesIO(contenido))
        columnas_texto = [col for col in encuesta.columns if encuesta[col].dropna().astype(str).str.strip().any()]
        texto_df = encuesta[columnas_texto].fillna("Sin respuesta").astype(str)

        bloque, respuestas_api = [], []
        total, contador = texto_df.size, 0

        for fila in texto_df.values:
            for respuesta in fila:
                contador += 1
                bloque.append(respuesta.strip())
                if len(bloque) == 10 or contador == total:
                    respuestas_api.extend(obtener_emociones_lote(bloque))
                    bloque = []

        resultados_emociones = [
            respuestas_api[i:i + len(columnas_texto)] for i in range(0, len(respuestas_api), len(columnas_texto))
        ]
        resultados_df = pd.DataFrame(resultados_emociones, columns=columnas_texto)

        excel_base_path = os.path.join(RESULTADOS_DIR, "emociones_resultado.xlsx")
        resultados_df.to_excel(excel_base_path, engine="openpyxl", index=False)

        # Insertar gráficos en Excel
        wb = load_workbook(excel_base_path)
        ws = wb.active

        img_width_px = int(10 * 96)  # 10 pulgadas * 96 dpi
        img_height_px = int(8 * 96)  # 8 pulgadas * 96 dpi
        separacion_filas = 40
        fila_inicial = len(resultados_df) + 3

        for i, col in enumerate(columnas_texto):
            plt.figure(figsize=(10, 8))
            ax = sns.countplot(x=resultados_df[col], order=EMOCIONES_VALIDAS)
            plt.title(f"Distribución de emociones para columna: {col}")
            plt.xlabel("Emoción")
            plt.ylabel("Frecuencia")
            plt.xticks(rotation=45)
            plt.tight_layout()

            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmpfile:
                plt.savefig(tmpfile.name, dpi=96)
                plt.close()

                img = XLImage(tmpfile.name)
                img.width = img_width_px
                img.height = img_height_px

                posicion_fila = fila_inicial + i * separacion_filas
                posicion_celda = f"A{posicion_fila}"
                ws.add_image(img, posicion_celda)

            os.unlink(tmpfile.name)

        wb.save(excel_base_path)

        # Guardar emoción global más común
        todas_emociones = resultados_df.values.flatten()
        emociones_filtradas = [e for e in todas_emociones if e in EMOCIONES_VALIDAS]
        if emociones_filtradas:
            emocion_mas_comun = Counter(emociones_filtradas).most_common(1)[0][0]
            with open(os.path.join(RESULTADOS_DIR, "emocion_global.txt"), "w", encoding="utf-8") as f:
                f.write(emocion_mas_comun)

        return FileResponse(
            path=excel_base_path,
            filename="emociones_resultado.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return {"error": str(e)}

@app.get("/emocion")
def obtener_emocion_global():
    try:
        emocion_txt_path = os.path.join(RESULTADOS_DIR, "emocion_global.txt")
        excel_path = os.path.join(RESULTADOS_DIR, "emociones_resultado.xlsx")
        if not os.path.exists(emocion_txt_path) or not os.path.exists(excel_path):
            return {"error": "archivos necesarios no encontrados"}

        with open(emocion_txt_path, "r", encoding="utf-8") as f:
            emocion = f.read().strip().lower()

        mapa_emoji = {
            "satisfacción": "😊", "frustración": "😠", "compromiso": "💪", "desmotivación": "😞",
            "estrés": "😣", "esperanza": "🌟", "inseguridad": "😟", "aprecio": "🤝",
            "indiferencia": "😐", "agotamiento": "😩"
        }

        puntuacion_emociones = {
            "satisfacción": 1, "compromiso": 1, "aprecio": 1, "esperanza": 0.8,
            "indiferencia": 0, "inseguridad": -0.3, "estrés": -1,
            "desmotivación": -0.8, "agotamiento": -1, "frustración": -0.5
        }

        df = pd.read_excel(excel_path)
        emociones = df.values.flatten()
        emociones_filtradas = [e.strip().lower() for e in emociones if e in puntuacion_emociones]

        if emociones_filtradas:
            valores = [puntuacion_emociones[e] for e in emociones_filtradas]
            media = sum(valores) / len(valores)
            porcentaje_satisfaccion = round(((media + 1) / 2) * 100, 2)
        else:
            porcentaje_satisfaccion = 0

        if porcentaje_satisfaccion <= 20:
            emoji_estado = "😠"
        elif porcentaje_satisfaccion <= 40:
            emoji_estado = "😕"
        elif porcentaje_satisfaccion <= 60:
            emoji_estado = "😐"
        elif porcentaje_satisfaccion <= 80:
            emoji_estado = "🙂"
        else:
            emoji_estado = "😄"

        respuesta_actual = {
            "fecha": time.strftime("%Y-%m-%d %H:%M:%S"),
            "emocion": emocion if emocion in puntuacion_emociones else "Error",
            "emoji": mapa_emoji.get(emocion, ""),
            "porcentaje_satisfaccion": porcentaje_satisfaccion,
            "estado_general": emoji_estado
        }

        if os.path.exists(HISTORIAL_PATH):
            with open(HISTORIAL_PATH, "r", encoding="utf-8") as f:
                historial_emociones = json.load(f)
        else:
            historial_emociones = []

        historial_emociones.append(respuesta_actual)
        historial_emociones = historial_emociones[-2:]

        with open(HISTORIAL_PATH, "w", encoding="utf-8") as f:
            json.dump(historial_emociones, f, ensure_ascii=False, indent=2)

        return historial_emociones

    except Exception:
        return [{
            "fecha": time.strftime("%Y-%m-%d"),
            "emocion": "Error",
            "emoji": "",
            "porcentaje_satisfaccion": "No disponible",
            "estado_general": "❌"
        }]
