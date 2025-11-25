import os
import pickle
import re
import tiktoken
from rank_bm25 import BM25Okapi

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PDF_PATH = os.path.join(BASE_DIR, "..", "data", "DOCUMENTOISOCONTEXTUAL.pdf")
STORE_PATH = os.path.join(BASE_DIR, "..", "data", "iso_store.pkl")

# --------------------------
# Conversión PDF → Texto
# --------------------------
import PyPDF2

def pdf_to_text(pdf_path):
    reader = PyPDF2.PdfReader(open(pdf_path, "rb"))
    text = ""
    for page in reader.pages:
        text += page.extract_text() + "\n"
    return text

def chunk_text(text, max_chars=1200):
    chunks = []
    while len(text) > max_chars:
        part = text[:max_chars]
        last_dot = part.rfind(".")
        if last_dot == -1:
            last_dot = max_chars
        chunks.append(text[:last_dot])
        text = text[last_dot:]
    if text:
        chunks.append(text)
    return chunks

def clean(t):
    return re.sub(r'\s+', ' ', t).strip()

# --------------------------
# Generar vectorstore
# --------------------------

def generar_vectorstore():
    print("Leyendo documento PDF...")
    
    text = pdf_to_text(PDF_PATH)
    text = clean(text)

    print("Creando chunks...")
    chunks = chunk_text(text)

    print("Tokenizando para BM25...")
    tokenized_chunks = [c.lower().split() for c in chunks]

    store = {
        "chunks": chunks,
        "tokens": tokenized_chunks,
    }

    with open(STORE_PATH, "wb") as f:
        pickle.dump(store, f)

    print("Vectorstore creado correctamente en:", STORE_PATH)

if __name__ == "__main__":
    generar_vectorstore()
