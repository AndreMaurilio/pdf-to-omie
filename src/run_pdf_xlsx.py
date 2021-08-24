# Script que extrai dados do pdf e preenche a planilha do padrao Omie de lançamento de notas.


import os

import openpyxl
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.layout import LAParams
from pdfminer.converter import TextConverter
from io import StringIO
from pdfminer.pdfpage import PDFPage

WB = ''
PDF_PATH = ''
CNPJ = ''
NPC = ''
CODEMICRO = {
"90126653012":  "PRD00096",
"90126373012":  "PRD00043",
"90126463012":  "PRD00012",
"90126473012":  "PRD00013",
"90126483012":  "PRD00014",
"90126493012":  "PRD00015",
"90126863012":  "PRD00023",
"90126793012":  "PRD00016",
"90126803012":  "PRD00017",
"90126813012":  "PRD00018",
"90126823012":  "PRD00019",
"90205644628":  "PRD00022",
"90205454628":  "PRD00031",
"90205464628":  "PRD00032",
"90205434628":  "PRD00029",
"90205394628":  "PRD00033",
"90205404628":  "PRD00034",
"90205414628":  "PRD00035",
"90205424628":  "PRD00036",
"90126413012":  "PRD00084",
"90126383012":  "PRD00085",
"90126403012":  "PRD00087",
"90126393012":  "PRD00086",
"90126433012":  "PRD00081",
"90126443012":  "PRD00082",
"90126453012":  "PRD00080",
"90126423012":  "PRD00083",
"90126343012":  "PRD00037",
"90205364628":  "PRD00038",
"90205374628":  "PRD00039",
"90205384628":  "PRD00040",
"90126533012":  "PRD00005",
"90126543012":  "PRD00006",
"90126553012":  "PRD00007",
"90126563012":  "PRD00008",
"90105793012":  "PRD00024",
"90126873012":  "PRD00021",
"90126573012":  "PRD00002",
"90461871210":  "PRD00208",
"90473273012":  "PRD00196",
"90103613012":	"PRD00045",
"90205444628":	"PRD00030",
"90205454628":	"PRD00031",


}


def write_pdf_txt(path, pdf):
    with open(path, mode="w") as output_file:
        for page in pdf.pages:
            text = page.extractText()
            output_file.write(text)


def pdf_to_xlsx(wb, pdf,address):
    sh1 = wb.active
    row = sh1.max_row
    r = 22
    global CNPJ
    #   print(wb['Omie_Pedido_Venda'].cell(3, 17).value)
    sh1.cell(row=7, column=4, value=CNPJ)
    sh1.cell(row=7, column=6, value="Suprimentos")
    sh1.cell(row=7, column=7, value="Para 21 dias ")

    CNPJ = ''

    for i in range(r, row + 1):
        for p in range(len(pdf)):
            for k in range(len(pdf[p])):
                pedido = pdf[p][k]
                sh1.cell(row=i, column=3, value=pedido["Material"])
                sh1.cell(row=i, column=4, value=pedido["Produto"])
                sh1.cell(row=i, column=5, value="MAXION (À FATURAR)")
                sh1.cell(row=i, column=6, value=pedido["QTD"])
                sh1.cell(row=i, column=7, value=pedido["Preço"])
                sh1.cell(row=i, column=10, value=pedido["Pedido de Compras"])
                sh1.cell(row=i, column=11, value=pedido["item"])


                i += 1
            p += 1
        else:
            break
    wb.save(os.getcwd()+"\\xlsx\\"+address[:-4]+".xlsx")


def represent_int(s):
    n = s
    try:
        int(n)
        return True
    except ValueError:
        return False


def get_pdf_miner_file(path, wb):
    resource_manager = PDFResourceManager(caching=True)
    out_text = StringIO()
    laParams = LAParams()
    text_converter = TextConverter(resource_manager, out_text, laparams=laParams)
    fp = open(path, 'rb')
    ar = list()
    interpreter = PDFPageInterpreter(resource_manager, text_converter)
    for page in PDFPage.get_pages(fp, pagenos=set(), maxpages=0, password='', caching=True, check_extractable=True):
        interpreter.process_page(page)
    text = out_text.getvalue()
    text = text.split("\n")
    ar.append(collect_pdf_data(text))
    address = path.split("/")[-1]
    pdf_to_xlsx(wb, ar,address)


def collect_pdf_data(lista):
    pedidos = list()
    precos = list()
    n = 0
    global CNPJ
    global NPC
    for i in lista:
        if len(i) > 9:
            if i[:9] == 'C.N.P.J.:' and CNPJ == '':
                CNPJ = i[10:]
        if len(i) == 10 and represent_int(i) and NPC == '':
            NPC = i
        if len(i) == 17:
            i = i.split(" ")
            if len(i) == 2:
                if represent_int(i[0]) and represent_int(i[1]):
                    pedidos.append(create_item(i[0], i[1], NPC, '', '', False))

        if ("PEÇ" in i or "UN" in i) and (8 <= len(i) <= 10):
            i = i.split(" ")
            precos.append(get_price_and_qtd(i[0], lista[n + 2]))
        n += 1
    pedidos = correct_price_and_qtd(pedidos, precos)
    NPC = ''
    return pedidos


def correct_price_and_qtd(lista, incomplete):
    for i in lista:
        if not i["complete"] and len(incomplete) > 0:
            i["QTD"] = incomplete[0]["QTD"]
            i["Preço"] = incomplete[0]["Preço"]
            incomplete.pop(0)
    return lista


def create_item(item, material, pedidocompra, qtd, preco, complete):
    produto = ""
    if material in CODEMICRO:
        produto = CODEMICRO[material]

    return {"item": item, "Material": material,"Produto":produto, "Pedido de Compras": pedidocompra, "QTD": qtd,
            "Frete": "", "Preço": preco, "complete": complete};


def get_price_and_qtd(qtd, preco):
    return {"QTD": qtd, "Preço": preco};


def pdf_to_omie_xlsx(path,pathomie,label):
    global WB, PDF_PATH
    PDF_PATH = path
    WB = openpyxl.load_workbook(pathomie)
    get_pdf_miner_file(PDF_PATH, WB)
    return "Fim!!"



