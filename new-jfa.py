import argparse
from unidecode import unidecode
from selenium.webdriver.support.ui import Select
import threading
import subprocess
import os
import time
from tqdm import tqdm
import shutil
import json
from tqdm import tqdm
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from selenium.common.exceptions import *
import re
import sys
import numpy as np
import cv2
import requests

service = Service()
options = webdriver.ChromeOptions()
titulo_arquivo = ""
options.add_argument("--headless=new")
items = []
options.add_argument("--disable-gpu")
options.add_argument("--disable-extensions")
prefs = {"profile.managed_default_content_settings.images": 2}
options.add_experimental_option("prefs", prefs)



driver = webdriver.Chrome(service=service, options=options)
driver.get("https://www.google.com.br/?hl=pt-BR")
time.sleep(3)
try:
    driver.get("https://corp.shoppingdeprecos.com.br/login")
    counter = 0
    while True:
        test = driver.find_elements(By.XPATH, '//*[@id="email"]')
        if test:
            break
        else:
            counter += 1
            if counter > 20:
                break;
            time.sleep(0.5)
    driver.find_element(By.XPATH, '//*[@id="email"]').send_keys("loja@jfaeletronicos.com")
    driver.find_element(By.XPATH, '//*[@id="password"]').send_keys("922982PC")
    driver.find_element(By.XPATH, '//*[@id="btnLogin"]').click()
except TimeoutException as e:
    print(f"Timeout ao tentar carregar a página ou encontrar um elemento: {e}")
except NoSuchElementException as e:
    print(f"Elemento não encontrado na página: {e}")
except WebDriverException as e:
    print(f"Erro no WebDriver: {e}")

time.sleep(3)
driver.get("https://corp.shoppingdeprecos.com.br/vendedores/vendasMarca")

cookies_list = []

cookies = driver.get_cookies()
for cookie in cookies:
    objeto = cookie['name']
    value = cookie['value']
    cookies_list.append(f"{objeto}={value};")

cookie = "".join(cookies_list)


titulo_arquivo = ""
# options.add_argument("--headless=new")

if os.path.exists(r"produtos.xlsx"):
    os.remove(r"produtos.xlsx")
if os.path.exists(r"modelos_jfa.xlsx"):
    os.remove(r"modelos_jfa.xlsx")

headers = {
    "Cookie": cookie
}


produtos = {
    "FONTE 40A": (402.79, 433.00),
    "FONTE 60A": (443.07, 473.28),
    "FONTE 60A LITE": (364.95, 390.43),
    "FONTE 70A": (493.42, 523.63),
    "FONTE 70A LITE": (408.73, 434.42),
    "FONTE 120A": (634.40, 674.68),
    "FONTE 120A LITE": (536.26, 573.36),
    "FONTE 200A": (805.59, 845.87),
    "FONTE 200A LITE": (681.83, 716.71),
    "FONTE 90 BOB": (422.93, 443.07),
    "FONTE 120 BOB": (499.46, 539.74),
    "FONTE 200 BOB": (624.33, 694.82),
    "FONTE 200 MONO": (736.61, 774.88),
}   

def ajustar_preco(preco, percentual):
    return preco * (1 + percentual / 100)

# Calcular novos ranges
novos_ranges = {}
for produto, (min_preco, max_preco) in produtos.items():
    min_ajustado = ajustar_preco(min_preco, -10)
    max_ajustado = ajustar_preco(max_preco, 10)
    novos_ranges[produto] = (round(min_ajustado, 2), round(max_ajustado, 2))

def SelecionarFonte(item):
    nome = item["Produto"].strip().lower()
    price = float(item["Preço Unitário"].replace(".", "").replace(",", "."))
    tipo = unidecode(item["Tipo de Anúncio"].strip().lower())
    total = float(item["Total"].replace(".", "").replace(",", "."))
    if "inversor" in nome or "amplificador" in nome or "processador" in nome or "capa" in nome:
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "OUTROS"})
        return
    
    if "k600" in nome and "fonte" not in nome and "k1200" not in nome:
    
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "CONTROLE K600"})
        return
        
    if "k1200" in nome and "fonte" not in nome and "k600" not in nome:
    
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "CONTROLE K1200"})
        return
        
    if ("controle wr" in nome or "wr" in nome or "redline" in nome) and "fonte" not in nome:
    
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "CONTROLE REDLINE"})
        return
        
    if ("acqua" in nome or "aqua" in nome or "agua" in nome) and "fonte" not in nome:
    
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "CONTROLE ACQUA"})
        return
    
    if "40" in nome or "40a" in nome or "40 amperes" in nome or "40amperes" in nome or "36a" in nome or "36" in nome or "36 amperes" in nome or "36amperes" in nome:
        if novos_ranges["FONTE 40A"][0] < price < novos_ranges["FONTE 40A"][1]:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 40A"})
            return
    if "60" in nome or "60a" in nome or "60 amperes" in nome or "60amperes" in nome or "60 a" in nome or "-60" in nome:
        if novos_ranges["FONTE 60A"][0] < price < novos_ranges["FONTE 60A"][1]:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 60A"})
            return
        if novos_ranges["FONTE 60A LITE"][0] < price < novos_ranges["FONTE 60A LITE"][1]:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 60A LITE"})
            return
    if "70" in nome or "70a" in nome or "70 amperes" in nome or "70amperes" in nome or "70 a" in nome:
        if novos_ranges["FONTE 70A"][0] < price < novos_ranges["FONTE 70A"][1]:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 70A"})
            return
        if novos_ranges["FONTE 70A LITE"][0] < price < novos_ranges["FONTE 70A LITE"][1]:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 70A LITE"})
            return
    if "90" in nome or "90a" in nome or "90 amperes" in nome or "90amperes" in nome or "90 a" in nome:
        if novos_ranges["FONTE 90 BOB"][0] < price < novos_ranges["FONTE 90 BOB"][1]:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE BOB 90A"})
            return
    if "120" in nome or "120a" in nome or "120 amperes" in nome or "120amperes" in nome or "120 a" in nome:
        if novos_ranges["FONTE 120A"][0] < price < novos_ranges["FONTE 120A"][1]:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 120A"})
            return
        if novos_ranges["FONTE 120 BOB"][0] < price < novos_ranges["FONTE 120 BOB"][1]:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 120A BOB"})
            return
        if novos_ranges["FONTE 120A LITE"][0] < price < novos_ranges["FONTE 120A LITE"][1]:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 120A LITE"})
            return
    if "bob" not in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome and "lit" not in nome:
        if "150" in nome or "150a" in nome or "150 amperes" in nome or "150amperes" in nome or "150 a" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 150A"})
            return  
    if "200" in nome or "200a" in nome or "200 amperes" in nome or "200amperes" in nome or "200 a" in nome:
        if novos_ranges["FONTE 200A LITE"][0] < price < novos_ranges["FONTE 200A LITE"][1]:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 200A LITE"})
            return
        if novos_ranges["FONTE 200 BOB"][0] < price < novos_ranges["FONTE 200 BOB"][1]:   
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 200A BOB"})
            return
        if novos_ranges["FONTE 200A"][0] < price < novos_ranges["FONTE 200A"][1]: 
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 200A"})
            return
        if novos_ranges["FONTE 200 MONO"][0] < price < novos_ranges["FONTE 200 MONO"][1]: 
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 200A MONO"})
            return
    if "bob" not in nome and  ("lite" in nome or "light" in nome or "lit" in nome) and "controle" not in nome and ("mono" in nome or "220v" in nome or "monovolt" in nome):
        if "200" in nome or "200a" in nome or "200 amperes" in nome or "200amperes" in nome or "200 a" in nome:
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE LITE 200A MONO"})
            return

    if "fonte" in nome and "nobreak" not in nome:
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "POSSIVEL FONTE"})
        return
    items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "OUTROS"})
                

parser = argparse.ArgumentParser(description='Processar datas de início e fim.')
parser.add_argument('--dia_inicial', type=str, required=True, help='Data inicial no formato AAAA-MM-DD')
parser.add_argument('--dia_final', type=str, required=True, help='Data final no formato AAAA-MM-DD')

args = parser.parse_args()

dia_inicial = args.dia_inicial
dia_final = args.dia_final


urls = ["JFA", "JFA%20ELETRONICOS"]             
for i in urls:
    response = requests.get(f"https://corp.shoppingdeprecos.com.br/vendedores/exportar_vendas_marca?id={i}&ini={dia_inicial}&fim={dia_final}", headers=headers)

    if response.status_code == 200:  

        with open("produtos.xlsx", 'wb') as file:

            file.write(response.content)

    time.sleep(5)



    db = pd.read_excel("produtos.xlsx", engine='openpyxl')
                    
    for index, item in db.iterrows():
        SelecionarFonte(item)

        
df = pd.DataFrame(items)

# Exportar o DataFrame para um arquivo Excel
df.to_excel("modelos_jfa.xlsx", index=False)