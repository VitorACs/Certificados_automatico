#importar as bibliotecas
from docx import Document
from datetime import datetime
import os
import win32com.client

def PDF():
    wdFormatPDF = 17

    inputFile = os.path.abspath(f"Contrato Atualizado - {nome}.docx")
    outputFile = os.path.abspath(f"Contrato Atualizado - {nome}.pdf")
    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Open(inputFile)
    doc.SaveAs(outputFile, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()


#abrir o documento
certificado = Document(r"C:\Users\costa\OneDrive\Área de Trabalho\Python_curso_Hastag\Certificados\Modelo_certificado.docx")

nome = 'Vitor Augusto'
CPF  = 156356988
nome_curso = 'Curso de python'

dicionario_valores = {
    "XXXX": nome,
    "YYYY": str(CPF),
    "ZZZZ": nome_curso,
    "DD": str(datetime.now().day),
    "MM": str(datetime.now().month),
    "AA": str(datetime.now().year),
}

# tem alguma função replace? -> é executada por parágrafo
# percorrer as linhas
    # se a linha tiver o texto xxxx substitui pelo nome
for paragrafo in certificado.paragraphs:
    # para cada placeholder do dicionario
    for codigo in dicionario_valores:
        if codigo in paragrafo.text:
            paragrafo.text = paragrafo.text.replace(codigo, dicionario_valores[codigo])


certificado.save(f"Contrato Atualizado - {nome}.docx")
PDF()