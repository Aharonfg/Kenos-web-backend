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
from PIL import Image
import tempfile
import os
from collections import Counter

# Configura la clave de la API de Gemini  

genai.configure(api_key=["AIzaSyCi0vrZPLA8B2DTlrR86P93CVN8A7j-04o"])
modelo = genai.GenerativeModel("gemini-1.5-pro")

# Carpeta fija para guardar resultados compartidos
RESULTADOS_DIR = "resultados"
os.makedirs(RESULTADOS_DIR, exist_ok=True)

# Crear la app FastAPI
app = FastAPI()

# Configurar CORS para permitir frontend externo
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], 
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Columnas originales que vamos a analizar
columnas_objetivo = [
    "Autoevaluación del rendimiento - ¿Qué consideras que necesitarías para avanzar en tu desarrollo profesional?",
    "Valoración de recursos KENOS - Indica las observaciones que consideres sobre el seguimiento de tu persona referente",
    "Valoración de recursos KENOS - Observaciones:",
    "Autoevaluación - ¿Cómo crees que te valoran los/las compañeras de equipo?",
    "Autoevaluación - ¿Cómo crees que te valoran las personas que coordinan tu equipo?",
    "Valoración General - Observaciones y sugerencias"
]

# Nombres simplificados para el Excel final
columnas_finales = [
    "Autoevaluación del rendimiento",
    "Seguimiento de persona referente",
    "Valoración de recursos KENOS",
    "Valoración por compañeras",
    "Valoración por coordinadores",
    "Observaciones y sugerencias"
]

# Función que llama a Gemini para obtener la emoción de una frase
def obtener_emocion(texto, reintentos=3):
    prompt = (
        "Eres una persona de recursos humanos de una consultoría tecnológica llamada Kenos Technology y debes determinar "
        "cuál de las siguientes emociones se relaciona más con esta frase: satisfacción, frustración, compromiso, desmotivación, "
        "estrés, esperanza, inseguridad, aprecio, indiferencia o agotamiento. "
        f"La frase es: \"{texto}\". "
        "Devuélveme solo una palabra: la emoción que más se relacione con la frase dada, sin ninguna palabra o carácter adicional."
    )
    for intento in range(reintentos):
        try:
            respuesta = modelo.generate_content(prompt)
            emocion = respuesta.text.strip().split()[0]
            time.sleep(random.uniform(1.5, 2.5))  # Evita rate limit
            return emocion
        except Exception:
            print(f"Error al procesar con Gemini. Intento {intento+1}")
            print(traceback.format_exc())
            time.sleep(3 + intento * 2)
    return "Error"

@app.post("/analizar")
async def analizar_excel(file: UploadFile = File(...)):
    try:
        # Validar que el archivo sea Excel
        if not file.filename.endswith((".xlsx", ".xls")):
            return {"error": "Por favor sube un archivo Excel válido (.xlsx o .xls)."}

        contenido = await file.read()
        encuesta = pd.read_excel(io.BytesIO(contenido), index_col=0)

        respuestas_encuesta = encuesta[columnas_objetivo].values.tolist()

        # Definir funciones internas para el análisis
        def construir_prompt(lista_de_frases):
            prompt = (
                "Eres una persona de recursos humanos de una consultoría tecnológica llamada Kenos Technology. "
                "A continuación tienes varias frases que debes analizar. Para cada frase, indica solo una emoción relacionada: "
                "satisfacción, frustración, compromiso, desmotivación, estrés, esperanza, inseguridad, aprecio, indiferencia o agotamiento.\n\n"
            )
            for idx, frase in enumerate(lista_de_frases, 1):
                prompt += f"{idx}. \"{frase}\"\n"
            prompt += "\nResponde en formato:\n1. emoción\n2. emoción\n..."
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
                            emociones.append(partes[1].strip())
                        else:
                            emociones.append("Error")
                    time.sleep(random.uniform(1.5, 2.5))  # evitar rate limit
                    return emociones
                except Exception:
                    print(f"Error en intento {intento+1} de obtener emociones lote")
                    print(traceback.format_exc())
                    time.sleep(3 + intento * 2)
            return ["Error"] * len(frases)

        bloque = []
        respuestas_api = []
        contador = 0
        total = len(respuestas_encuesta) * len(columnas_finales)

        for fila in respuestas_encuesta:
            for respuesta in fila:
                contador += 1
                if pd.isna(respuesta) or (isinstance(respuesta, str) and respuesta.strip() == ""):
                    bloque.append("Sin respuesta")
                else:
                    bloque.append(str(respuesta).strip())

                # Cuando bloque llegue a 10 o sea la última respuesta, enviamos el bloque
                if len(bloque) == 10 or contador == total:
                    emociones_lote = obtener_emociones_lote(bloque)
                    respuestas_api.extend(emociones_lote)
                    bloque = []
                    print(f"Procesadas {contador}/{total} respuestas")

        if len(bloque) > 0:
            emociones_lote = obtener_emociones_lote(bloque)
            respuestas_api.extend(emociones_lote)

        # Reconstruir DataFrame
        respuestas_emociones = [
            respuestas_api[i:i+len(columnas_finales)] for i in range(0, len(respuestas_api), len(columnas_finales))
        ]
        resultados_df = pd.DataFrame(respuestas_emociones, columns=columnas_finales)

        # Guardar Excel base en carpeta fija
        excel_base_path = os.path.join(RESULTADOS_DIR, "emociones_resultado.xlsx")
        resultados_df.to_excel(excel_base_path, engine="openpyxl", index=False)

        # Abrir para agregar gráficos
        wb = load_workbook(excel_base_path)
        ws = wb.active
        fila_grafico = len(resultados_df) + 3

        for col in columnas_finales:
            datos = resultados_df[col][~resultados_df[col].isin(["Sin respuesta", "Error"])]
            if datos.empty:
                continue

            conteo = datos.value_counts()
            fig, ax = plt.subplots(figsize=(10, 8))
            sns.barplot(x=conteo.index, y=conteo.values, ax=ax)
            ax.set_title(f"Emociones: {col}")
            ax.set_ylabel("Frecuencia")
            ax.set_xlabel("Emoción")
            plt.xticks(rotation=45)
            plt.tight_layout()

            img_path = os.path.join(RESULTADOS_DIR, f"{col}.png")
            fig.savefig(img_path)
            plt.close()

            pil_img = Image.open(img_path).convert("RGB")
            pil_img.save(img_path)
            img_excel = XLImage(img_path)
            ws.add_image(img_excel, f"A{fila_grafico}")
            fila_grafico += 20

        # Guardar Excel con gráficos en la misma carpeta
        wb.save(excel_base_path)

        # Calcular emoción global
        todas_emociones = resultados_df.values.flatten()
        emociones_filtradas = [e for e in todas_emociones if e not in ['Sin respuesta', 'Error']]
        if emociones_filtradas:
            emocion_mas_comun = Counter(emociones_filtradas).most_common(1)[0][0]
            emocion_txt_path = os.path.join(RESULTADOS_DIR, "emocion_global.txt")
            
            with open(emocion_txt_path, "w", encoding="utf-8") as f:
                f.write(emocion_mas_comun)

        return FileResponse(
            path=excel_base_path,
            filename="emociones_resultado.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        print("Error general en /analizar:")
        print(traceback.format_exc())
        return {"error": "Ha ocurrido un error al procesar el archivo." + e}

@app.get("/emocion")
def obtener_emocion_global():
    try:
        emocion_txt_path = os.path.join(RESULTADOS_DIR, "emocion_global.txt")
        excel_path = os.path.join(RESULTADOS_DIR, "emociones_resultado.xlsx")

        if not os.path.exists(emocion_txt_path) or not os.path.exists(excel_path):
            return {"error": "archivos necesarios no encontrados"}

        with open(emocion_txt_path, "r", encoding="utf-8") as f:
            emocion = f.read().strip().lower()

        # Mapa para el emoji de la emoción dominante
        mapa_emoji = {
            "satisfacción": "😊",
            "frustración": "😠",
            "compromiso": "💪",
            "desmotivación": "😞",
            "estrés": "😣",
            "esperanza": "🌟",
            "inseguridad": "😟",
            "aprecio": "🤝",
            "indiferencia": "😐",
            "agotamiento": "😩"
        }

        # Puntuación para calcular el % de satisfacción
        puntuacion_emociones = {
            "satisfacción": 1,
            "compromiso": 1,
            "aprecio": 1,
            "esperanza": 0.8,
            "indiferencia": 0,
            "inseguridad": -0.3,
            "estrés": -1,
            "desmotivación": -0.8,
            "agotamiento": -1,
            "frustración": -0.5
        }

        df = pd.read_excel(excel_path)
        emociones = df.values.flatten()
       

        if emociones_filtradas:
            valores = [puntuacion_emociones[e] for e in emociones_filtradas]
            media = sum(valores) / len(valores)
            porcentaje_satisfaccion = round(((media + 1) / 2) * 100, 2)  # de [-1,1] a [0,100]
        else:
            porcentaje_satisfaccion = 0

        # Emoji de estado general según el % de satisfacción
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

        return {
            "emoji": mapa_emoji.get(emocion, "❓"),
            "emocion": emocion,
            "porcentaje_satisfaccion": porcentaje_satisfaccion,
            "estado_general": emoji_estado
        }

    except Exception:
        print("Error al obtener la emoción global:")
        print(traceback.format_exc())
        return {
            "emoji": "❌",
            "emocion": "Error",
            "porcentaje_satisfaccion": "No disponible",
            "estado_general": "❌"
        }
