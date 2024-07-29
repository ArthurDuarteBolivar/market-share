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

def SelecionarFonte(item):
    nome = unidecode(item["Produto"].strip().lower())
    price = float(item["Preço Unitário"].replace(".", "").replace(",", "."))
    tipo = unidecode(item["Tipo de Anúncio"].strip().lower())
    total = float(item["Total"].replace(".", "").replace(",", "."))
    if "inversor" in nome or "amplificador" in nome or "processador" in nome or "capa" in nome or "nobreak" in nome or "retificadora" in nome or "multimidia" in nome or "gerenciador" in nome or "suspensao" in nome or "stetsom" in nome or "central" in nome:
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "OUTROS"})
        return
    
    if ("k600" in nome or "k6" in nome) and "fonte" not in nome and "k1200" not in nome:
    
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "CONTROLE K600"})
        return
        
    if ("k1200" in nome or "k12" in nome) and "fonte" not in nome and "k600" not in nome:
    
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "CONTROLE K1200"})
        return
        
    if ("controle wr" in nome or "wr" in nome or "redline" in nome or "red line" in nome) and "fonte" not in nome:
    
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "CONTROLE REDLINE"})
        return
        
    if ("acqua" in nome or "aqua" in nome or "agua" in nome) and "fonte" not in nome:
    
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "CONTROLE ACQUA"})
        return
    
    
    if "controle" not in nome and "lite" not in nome and "light" not in nome:
        if "40" in nome or "40a" in nome or "40 amperes" in nome or "40amperes" in nome or "36a" in nome or "36" in nome or "36 amperes" in nome or "36amperes" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 40A"})
            return
            
    if "bob" not in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome:
        if "60" in nome or "60a" in nome or "60 amperes" in nome or "60amperes" in nome or "60 a" in nome or "-60" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 60A"})
            return
            
    if "bob" not in nome and ("lite" in nome or "light" in nome or "lit" in nome) and "controle" not in nome:
        if "60" in nome or "60a" in nome or "60 amperes" in nome or "60amperes" in nome or "60 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE LITE 60A"})
            return
            
    if "bob" not in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome:
        if "70" in nome or "70a" in nome or "70 amperes" in nome or "70amperes" in nome or "70 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 70A"})
            return

    if "bob" not in nome and  ("lite" in nome or "light" in nome or "lit" in nome) and "controle" not in nome:
        if "70" in nome or "70a" in nome or "70 amperes" in nome or "70amperes" in nome or "70 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE LITE 70A"})
            return
            
    if "bob" not in nome and  ("lite" in nome or "light" in nome) and "controle" not in nome:
        if "40" in nome or "40a" in nome or "40 amperes" in nome or "40amperes" in nome or "40 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE LITE 40A"})
            return
            
    if "lite" not in nome and "light" not in nome  and "controle" not in nome:
        if "90" in nome or "90a" in nome or "90 amperes" in nome or "90amperes" in nome or "90 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE BOB 90A"})
            return
            
    if "bob" not in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome and "lit" not in nome:
        if "120" in nome or "120a" in nome or "120 amperes" in nome or "120amperes" in nome or "120 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 120A"})
            return

    if "bob" not in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome and "lit" not in nome:
        if "150" in nome or "150a" in nome or "150 amperes" in nome or "150amperes" in nome or "150 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 150A"})
            return
             
    if "bob" not in nome and  ("lite" in nome or "light" in nome or "lit" in nome) and "controle" not in nome:
        if "120" in nome or "120a" in nome or "120 amperes" in nome or "120amperes" in nome or "120 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE LITE 120A"})
            return
                
    if "bob" in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome and "lit" not in nome:
        if "120" in nome or "120a" in nome or "120 amperes" in nome or "120amperes" in nome or "120 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE BOB 120A"})
            return
                
    if "bob" not in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome and 'mono' not in nome and 'mono' not in nome and 'monovolt' not in nome and '220v' not in nome and "lit" not in nome:
        if "200" in nome or "200a" in nome or "200 amperes" in nome or "200amperes" in nome or "200 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 200A"})
            return
                
    if "bob" not in nome and  ("lite" in nome or "light" in nome or "lit" in nome) and "controle" not in nome and ("mono" in nome or "220v" in nome or "monovolt" in nome):
        if "200" in nome or "200a" in nome or "200 amperes" in nome or "200amperes" in nome or "200 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE LITE 200A MONO"})
            return
        
    if "bob" not in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome and ("mono" in nome or "220v" in nome or "monovolt" in nome):
        if "200" in nome or "200a" in nome or "200 amperes" in nome or "200amperes" in nome or "200 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 200A MONO"})
            return
                
    if "bob" not in nome and  ("lite" in nome or "light" in nome or "lit" in nome) and "controle" not in nome:
        if "200" in nome or "200a" in nome or "200 amperes" in nome or "200amperes" in nome or "200 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE LITE 200A"})
            return
                
    if "bob" in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome and 'mono' not in nome and 'mono' not in nome and 'monovolt' not in nome and '220v' not in nome:
        if "200" in nome or "200a" in nome or "200 amperes" in nome or "200amperes" in nome or "200 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE BOB 200A"})
            return
        
        
    if "bob" not in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome and 'mono' not in nome and 'mono' not in nome and 'monovolt' not in nome and '220v' not in nome and "lit" not in nome:
        if "20" in nome or "20a" in nome or "20 amperes" in nome or "20amperes" in nome or "20 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 200A"})
            return
                
    if "bob" not in nome and  ("lite" in nome or "light" in nome or "lit" in nome) and "controle" not in nome and ("mono" in nome or "220v" in nome or "monovolt" in nome):
        if "20" in nome or "20a" in nome or "20 amperes" in nome or "20amperes" in nome or "20 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE LITE 200A MONO"})
            return
        
    if "bob" not in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome and ("mono" in nome or "220v" in nome or "monovolt" in nome):
        if "20" in nome or "20a" in nome or "20 amperes" in nome or "20amperes" in nome or "20 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 200A MONO"})
            return
                
    if "bob" not in nome and  ("lite" in nome or "light" in nome or "lit" in nome) and "controle" not in nome:
        if "20" in nome or "20a" in nome or "20 amperes" in nome or "20amperes" in nome or "20 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE LITE 200A"})
            return
                
    if "bob" in nome and "lite" not in nome and "light" not in nome  and "controle" not in nome and 'mono' not in nome and 'mono' not in nome and 'monovolt' not in nome and '220v' not in nome:
        if "20" in nome or "20a" in nome or "20 amperes" in nome or "20amperes" in nome or "20 a" in nome:
        
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE BOB 200A"})
            return
    
    
    if "fonte" in nome:
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "POSSIVEL FONTE"})
        return
    
    if "controle" in nome:
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "POSSIVEL CONTROLE"})
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