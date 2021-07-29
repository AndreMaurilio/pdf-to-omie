# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import openpyxl
import requests

import PyPDF2
from PyPDF2 import PdfFileReader
from selenium import webdriver
from openpyxl import Workbook

wb = openpyxl.load_workbook(
    "/pdf-to-omie/Omie_EXTRAIDO.xlsx")
PDF_PATH = '/pdf-to-omie/4500022260.pdf'  # MAXION.PDF, 4500022260.pdf
TXT_PATH = '/home/andre/Downloads/pdf_extraido.txt'
PATH = '/home/andre/Downloads/ChromeDriver91-19/chromedriver'
chromeOptions = webdriver.ChromeOptions()
# chromeOptions.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
chromeOptions.add_argument('--no-sandbox')
# chromeOptions.add_argument("--disable-setuid-sandbox")
chromeOptions.add_argument('--disable-dev-shm-using')
chromeOptions.add_argument('--disable-extensions')
# chromeOptions.add_argument('start-maximized')
chromeOptions.add_argument('disable-infobars')
chromeOptions.add_argument('--headless')
chromeOptions.add_argument('--disable-gpu')
# chromeOptions.headless = True
driver = webdriver.Chrome(executable_path=PATH, chrome_options=chromeOptions)


# opts = webdriver.FirefoxOptions()
# opts.add_argument("--headless")
# opts.add_argument('--no-sandbox')
# # chromeOptions.add_argument("--disable-setuid-sandbox")
# opts.add_argument('--disable-dev-shm-using')
# opts.add_argument('--disable-extensions')
# driver = webdriver.Firefox(executable_path='/home/andre/Downloads/FireFoxDriver28/geckodriver', firefox_options=opts)


def read_pdf(path):
    # Le o arquivo
    pdf = PdfFileReader(str(path))
    # Printa o numero de paginas
    print(pdf.getNumPages())
    # printa todos os textos
    # for page in pdf.pages:
    #     print(page.extractText(), end='\n')
    return pdf


def write_pdf_txt(path, pdf):
    with open(path, mode="w") as output_file:
        for page in pdf.pages:
            text = page.extractText()
            output_file.write(text)


def pdf_to_xlsx(wb, pdf):
    sh1 = wb['Omie_Pedido_Venda']
    sh1 = wb.active
    pedido = {"item": "", "Material": "", "Pedido de Compras": "",
              "QTD": "", "Frete": "", "Preço": ""}
    row = sh1.max_row
    p = 0
    r = 22
    print(wb['Omie_Pedido_Venda'].cell(3, 17).value)
    for i in range(r, row + 1):
        for p in range(len(pdf)):
            for k in range(len(pdf[p])):
                pedido = pdf[p][k]
                sh1.cell(row=i, column=11, value=pedido["item"])
                sh1.cell(row=i, column=3, value=pedido["Material"])
                sh1.cell(row=i, column=10, value=pedido["Pedido de Compras"])
                sh1.cell(row=i, column=6, value=pedido["QTD"])
                # sh1.cell(row=i, column=20, value=pedido["Frete"])
                sh1.cell(row=i, column=7, value=pedido["Preço"])
                i += 1
            p += 1
        else:
            break
    wb.save("/home/andre/Documentos/workspace/Python-Projects/pdf-to-omie/Omie_EXTRAIDO.xlsx")


def print_pdf_txt(path, pdf, wb):
    with open(path, mode="w"):
        ar = list()
        l = 0
        for page in pdf.pages:
            text = page.extractText()
            pedido = {"item": "", "Material": "", "Pedido de Compras": "",
                      "QTD": "", "Frete": "", "Preço": ""}

            text = text.split(
                "ItemDescrição MaterialPedido deCompras /Programa deremessaQUANT.UMPreçoUnitárioFreteUtilização do MaterialMoedaPreço TotalSaldoData de Entrega")
            if (isinstance(text, list) and len(text) > 1):
                ar.append(populate_pedidos(text[1]))

            #                n = 0
            #                while n > len(pedidos):
            #                    print(f'Resultado', pedidos[n], end='\n')
            #                    n += 1
            else:
                print("DESCARTADO \n")
    pdf_to_xlsx(wb, ar)


def split_item_desc(st):
    n = 16
    while (n + 16) < len(st):
        if represent_int(st[n:n + 16]):
            st = list(st)
            st[n - 1] = '&'
            st = ''.join(st)
        n += 1;
    v = st.split("&")
    return v


def populate_pedidos(lista):
    # pedido = {"item": "", "Material": "", "Pedido de Compras": "",
    #           "QTD": "", "Frete": "", "Preço": ""}
    # lista = lista.split("        ")

    lista = split_item_desc(lista)
    pedidos = list()
    if isinstance(lista, list):
        for s in lista:
            aux = s[0:16]
            if (represent_int(aux) == True):
                item = s[0:5]
                material = s[5:16]
                st = ""
                n2 = 17
                stop = False
                for c in s[16:]:
                    n = 0
                    st = c
                    for e in s[n2:]:


                        if n == 9:
                            if represent_int(st):
                                pedidocompra = st
                                stop = True
                                break
                            else:
                                n = 0
                                st = ""
                                break
                        st = st + e
                        n += 1

                    if stop:
                        break
                    n2 = n2 + 1

                if "PEÇ" in s[n2:n2 + 30]:
                    v = s[n2:].split("PEÇ")
                else:
                    v = s[n2:].split("UN")
                qtd = v[0][9:]
                if len(v) > 1:
                    preco = v[1][:6]
                else:
                    preco = "nd"
                pedidos.append({"item": item, "Material": material, "Pedido de Compras": pedidocompra,
                                "QTD": qtd, "Frete": "", "Preço": preco})

    return pedidos


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


def write_browser():
    driver.set_page_load_timeout(30)
    driver.get("https://www.ig.com.br")
    driver.maximize_window()
    print(driver.title)
    print(driver.page_source)
    # driver.maximize_window()


def represent_int(s):
    n = s
    try:
        int(n)
        return True
    except ValueError:
        return False


if __name__ == '__main__':
    pdf = read_pdf(PDF_PATH)
    # write_pdf_txt(TXT_PATH,pdf)
    # write_browser()
    print_pdf_txt(TXT_PATH, pdf, wb)
    print("FIM!!")
