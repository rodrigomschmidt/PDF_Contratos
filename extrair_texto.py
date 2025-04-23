import pdfplumber
import os

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
            "COD": [],
            "Desc": [],
            "QTD": [],
            "QTD_Total": None, #OK
            "Cliente": None,  #OK
            "Porto": None,  #OK
            "Destino": None, #OK
            "Data_prevista": None,
            "USD": [],
            "Val": [],
            }
    
    info["Contrato"] = texto.split("INSTRUÇÃO FÁBRICA - ")[1].split("Cliente")[0].strip()
    info["Cliente"] = texto.split("Cliente")[1].split("ADDRESS")[0].strip()
    info["Porto"] = texto.split("Destino")[1].split("Enbarque")[0].strip()
    info["Destino"] = texto.split("Mercado")[1].split("Temperatura")[0].strip()
    info["QTD_Total"] = float(texto.split("Total ")[1].split("US$")[0].strip().replace(",","."))*1000
    print(info["Contrato"])
    print(info["Cliente"])
    print(info["Porto"])
    print(info["Destino"])
    print(info["QTD_Total"])

    produtos = texto.split("Expire Date")[1].split("External Label")[0].strip()
 
    print(produtos)
    linhas = (produtos.split("\n"))
    for linha in linhas:
        if "CIF" in linha:
            break
        else:
            info["COD"].append(linha.split()[0].strip())
            info["Desc"].append(linha.split(" CARTONS ")[0].split(" ")[1].strip())
            info["QTD"].append(linha.split(" CARTONS ")[1].split(" ")[0].strip())
            info["USD"].append(linha.split(" CARTONS ")[1].split(" ")[1].strip())
            n = len(linha.split(" "))

            info["Val"].append(linha.split(" ")[n-2]+ " " + linha.split(" ")[n-1])


    print(info["COD"])
    print(info["Desc"])
    print(info["QTD"])
    print(info["USD"])
    print(info["Val"])

pdf_path = "contrato.pdf"
texto = ler_pdf(pdf_path)
extrair_resultado(texto)