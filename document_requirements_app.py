import streamlit as st
import pandas as pd
from pathlib import Path
import re

from pdfminer.high_level import extract_text as extract_pdf_text
from docx import Document

import spacy
from spacy.util import is_package

# Ensure Spanish model is downloaded
model_name = 'es_core_news_sm'
if not is_package(model_name):
    from spacy.cli import download
    download(model_name)

nlp = spacy.load(model_name)

KEYWORDS = [
    'documentaci',  # matches documentacion, documentación
    'requisit',
    'pliego',
    'anexo',
]


def extract_text(file_path: Path) -> str:
    """Extract text from a PDF or DOCX file."""
    if file_path.suffix.lower() == '.pdf':
        return extract_pdf_text(str(file_path))
    elif file_path.suffix.lower() in {'.docx', '.doc'}:
        doc = Document(str(file_path))
        return '\n'.join(p.text for p in doc.paragraphs)
    else:
        raise ValueError('Unsupported file type: %s' % file_path.suffix)


def find_requirements(text: str) -> pd.DataFrame:
    """Return sentences that mention documentation requirements."""
    doc = nlp(text)
    sentences = []
    for sent in doc.sents:
        s = sent.text.strip()
        if any(re.search(k, s, re.IGNORECASE) for k in KEYWORDS):
            sentences.append(s)
    return pd.DataFrame({'Requisito': sentences})


st.title('Extracción de requisitos documentales')

uploaded_file = st.file_uploader('Sube un archivo PDF o DOCX')

if uploaded_file is not None:
    data = uploaded_file.read()
    path = Path(uploaded_file.name)
    temp = Path('uploaded_' + path.name)
    temp.write_bytes(data)
    try:
        text = extract_text(temp)
        df = find_requirements(text)
        if df.empty:
            st.info('No se encontraron referencias a documentación.')
        else:
            st.subheader('Requisitos encontrados')
            st.dataframe(df)
    finally:
        temp.unlink()
