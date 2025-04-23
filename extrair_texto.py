import pdfplumber
import os
from openpyxl import load_workbook

def ler_pdf(pdf_path):   

    pdf = pdfplumber.open(pdf_path)
    paginas = pdf.pages
    pagina = paginas[0]
    texto = pagina.extract_text()
    print(paginas)

    return texto

def extrair_resultado(texto):
    info = {
            "Contrato": None,  #OK
            "COD": [], #OK
            "Desc": [], #OK
            "QTD": [], #OK
            "QTD_Total": None, #OK
            "Cliente": None,  #OK
            "Porto": None,  #OK
            "Destino": None, #OK
            "USD": [], #OK
            "Val": [], #OK
            }
    
    info["Contrato"] = texto.split("INSTRUÇÃO FÁBRICA - ")[1].split("Cliente")[0].strip()
    info["Cliente"] = texto.split("Cliente")[1].split("ADDRESS")[0].strip()
    info["Porto"] = texto.split("Destino")[1].split("Enbarque")[0].strip()
    info["Destino"] = texto.split("Mercado")[1].split("Temperatura")[0].strip()
    info["QTD_Total"] = float(texto.split("Total ")[1].split("US$")[0].strip().replace(",","."))*1000

    produtos = texto.split("Expire Date")[1].split("External Label")[0].strip()
 
    linhas = (produtos.split("\n"))
    for linha in linhas:
        if "CIF" in linha:
            break
        else:
            info["COD"].append(int(linha.split()[0].strip()))
            info["Desc"].append(linha.split(" CARTONS ")[0].split(" ")[1].strip())
            info["QTD"].append(float((linha.split(" CARTONS ")[1].split(" ")[0].strip()).replace(",","."))*1000)
            info["USD"].append(float(linha.split(" CARTONS ")[1].split(" ")[1].strip().replace(".","").replace(",","."))/1000)
            n = len(linha.split(" "))
            info["Val"].append(linha.split(" ")[n-2]+ " " + linha.split(" ")[n-1])

    return info

def registrar_resultados(resultados, caminho_excel):

    wb = load_workbook(caminho_excel)
    ws = wb["RESULTADOS"]

    linha = 1
    while ws.cell(row=linha, column=1).value is not None:
        linha += 1
    
    tamanhos = [len(resultados["COD"]), 
                len(resultados["Desc"]),
                len(resultados["QTD"]), 
                len(resultados["USD"]), 
                len(resultados["Val"])
    ]

    if len(set(tamanhos)) >1:    
        print("Erro")
        return
    else:
        t = tamanhos[0]

    listas = [[resultados["Contrato"]]*t, 
              resultados["COD"], 
              resultados["Desc"], 
              resultados["QTD"], 
              [resultados["QTD_Total"]]*t,
              [resultados["Cliente"]]*t,
              [resultados["Porto"]]*t,
              [resultados["Destino"]]*t,
              resultados["USD"], 
              resultados["Val"]
    ]
    
    for x in range(t):
        for n, lista in enumerate(listas):
            ws.cell(row=linha+x, column=n+1, value=lista[x])
            print(n)

    wb.save(caminho_excel)
    #print(listas)
        



pdf_path = "contrato.pdf"
texto = ler_pdf(pdf_path)
registrar_resultados(extrair_resultado(texto), "resultados.xlsx")