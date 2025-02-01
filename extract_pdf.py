import pdfplumber
import pandas as pd
import sys 

def extract_text(pdf_path, search_term=None):
    data = []

    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text:
                if search_term and search_term in text: 
                    data.append([f"Side {i+1}", text])
                elif not search_term:
                    data.append([f"Side {i+1}", text])
    df = pd.DataFrame(data, columns=["Side", "Innhold"])
    df.to_excel("Ekstrahert_PDF.xlsx", index=False)

if __name__ == "__main__":
    pdf_file = sys.argv[1]
    search_term = sys.argv[2] if len(sys.argv) > 2 else None
    extract_text(pdf_file, search_term)

    **VBA MAKRO for å kjøre Python Scriptet fra Excel; 

Sub RunPythonScript()
    Dim pdfPath As String
    Dim searchTerm As String
    Dim pythonExe As String
    Dim scriptPath As String
    Dim cmd As String 

' Definer filstien til Python og scriptet 

    pythonExe = """C:\Path\To\Python.exe""" 'Sett inn din Python script
    scriptPath = """C:\Path\To\extract_pdf.py""" ' Sett inn script-path
    pdfPath = """C:\Sti\Til\Din\Fil.pdf"""
    searchTerm = InputBox("Skriv inn søkeord eller la feltet stå tomt for å hente all tekst:")

' Bygg kommandoen
    If searchTerm = "" Then
        cmd = pythonExe & " " & scriptPath & " " & pdfPath
    Else 
        cmd = pythonExe & " " & scriptPath & " " & pdfPath & " """ & searchTerm & """"
    End If

' Kjør Python-scriptet 
    Shell cmd, vbNormalFocus

    MsgBox "Prosessen er ferdig! Sjekk 'Ekstrahert_PDF.xlsx'.", vbInformation
End Sub 