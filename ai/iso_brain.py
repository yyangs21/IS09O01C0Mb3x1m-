import os
import pickle
import tiktoken
from rank_bm25 import BM25Okapi
from openai import OpenAI

# Carga del vectorstore generado por iso_vectorstore.py
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STORE_PATH = os.path.join(BASE_DIR, "..", "data", "iso_store.pkl")

client = OpenAI()

def cargar_vectorstore():
    if not os.path.exists(STORE_PATH):
        return None
    
    with open(STORE_PATH, "rb") as f:
        return pickle.load(f)

vectorstore = cargar_vectorstore()

def buscar_contexto(query, k=3):
    """Retorna los fragmentos más relevantes usando BM25"""
    if not vectorstore:
        return []
    
    bm25 = BM25Okapi(vectorstore["tokens"])
    scores = bm25.get_scores(query.split())
    top_idx = sorted(range(len(scores)), key=lambda i: scores[i], reverse=True)[:k]
    
    return [vectorstore["chunks"][i] for i in top_idx]

def responder_con_iso(pregunta):
    """Construye respuesta integrando los mejores fragmentos del PDF"""
    if not vectorstore:
        return "Aún no se ha generado la base de conocimiento ISO. Ejecuta iso_vectorstore.py"

    contexto = buscar_contexto(pregunta)

    prompt = f"""
Eres un experto en Sistemas de Gestión ISO 9001.
Usa los siguientes fragmentos del documento oficial para responder:

{contexto}

Pregunta del usuario:
{pregunta}

Da una respuesta clara, profesional y basada en el documento.
"""

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.2,
        max_tokens=600
    )

    return response.choices[0].message["content"]
