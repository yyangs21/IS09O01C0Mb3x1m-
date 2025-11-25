import pickle
import numpy as np
from openai import OpenAI
import os

VECTOR_PATH = os.path.join("ai", "iso_store.pkl")

# Cargar vectorstore desde el repositorio
with open(VECTOR_PATH, "rb") as f:
    VECTOR_DATA = pickle.load(f)

def cosine_similarity(a, b):
    a = np.array(a)
    b = np.array(b)
    return np.dot(a, b) / (np.linalg.norm(a) * np.linalg.norm(b))

def obtener_contexto(query, api_key):
    client = OpenAI(api_key=api_key)

    emb = client.embeddings.create(
        input=query,
        model="text-embedding-3-large"
    ).data[0].embedding

    resultados = []
    for item in VECTOR_DATA:
        sim = cosine_similarity(emb, item["embedding"])
        resultados.append((sim, item["texto"]))

    resultados = sorted(resultados, key=lambda x: x[0], reverse=True)

    top = [text for _, text in resultados[:3]]

    return "\n\n".join(top)

def responder_con_iso(query, api_key=None, client_override=None):
    client = client_override or OpenAI(api_key=api_key)

    # Obtener contexto
    emb = client.embeddings.create(
        input=query,
        model="text-embedding-3-large"
    ).data[0].embedding

    resultados = []
    for item in VECTOR_DATA:
        sim = cosine_similarity(emb, item["embedding"])
        resultados.append((sim, item["texto"]))

    resultados = sorted(resultados, key=lambda x: x[0], reverse=True)
    top = [text for _, text in resultados[:3]]
    contexto = "\n\n".join(top)

    # Generar respuesta
    prompt = f"""
Responde usando únicamente el siguiente contexto ISO:

CONTEXTO:
{contexto}

PREGUNTA:
{query}

RESPUESTA:
"""
    response = client.chat.completions.create(
        model="gpt-4.1-mini",
        messages=[{"role": "user", "content": prompt}]
    )

    return response.choices[0].message.content
