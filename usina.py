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
if os.path.exists(r"produtos2.xlsx"):
    os.remove(r"produtos2.xlsx")
if os.path.exists(r"modelos_usina.xlsx"):
    os.remove(r"modelos_usina.xlsx")


headers = {
    "Cookie": cookie
}
def SelecionarFonte(item):
    nome = item["Produto"].strip().lower()
    price = float(item["Preço Unitário"].replace(".", "").replace(",", "."))
    tipo = unidecode(item["Tipo de Anúncio"].strip().lower())
    total = float(item["Total"].replace(".", "").replace(",", "."))
    if "inversor" in nome or "kit" in nome or "controle" in nome or "conversor" in nome or "24v"  in nome or "truck"  in nome or "48v" in nome or "48 v" in nome or "fita led" in nome or "máquina" in nome or "fumaça" in nome or "vela" in nome or "refletor" in nome or "moving" in nome:
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "OUTROS"})
        return
        
    if "bob" not in nome:          
        if "30" in nome or " 30a" in nome or " 30 a" in nome or " 30 amperes" in nome or " 30amperes" in nome or "30 amp" in nome or "30amp" in nome:
            
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 30A"})
            return
        
    if "bob" not in nome and "48v" not in nome and "500" not in nome and "150" not in nome:          
        if "50" in nome or " 50a" in nome or " 50 a" in nome or " 50 amperes" in nome or " 50amperes" in nome or "50 amp" in nome or "50amp" in nome:
            
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 50A"})
            return
       
    if "bob" not in nome:          
        if "70" in nome or " 70a" in nome or " 70 a" in nome or " 70 amperes" in nome or " 70amperes" in nome or "70 amp" in nome or "70amp" in nome: 
            
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 70A"})
            return
        
    if "bob" not in nome:          
        if "90" in nome or " 90a" in nome or " 90 a" in nome or " 90 amperes" in nome or " 90amperes" in nome or "90 amp" in nome or "90amp" in nome:
            
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 90A"})
            return
       
    if "bob" not in nome:          
        if "100" in nome or " 100a" in nome or " 100 a" in nome or " 100 amperes" in nome or " 100amperes" in nome or "100 amp" in nome or "100amp" in nome:
            
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 100A"})
            return
        
    if "bob" not in nome:          
        if "120" in nome or " 120a" in nome or " 120 a" in nome or " 120 amperes" in nome or " 120amperes" in nome or "120 amp" in nome or "120amp" in nome:
            
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 120A"})
            return
        
    if "bob" not in nome:          
        if "160" in nome or " 160a" in nome or " 160 a" in nome or " 160 amperes" in nome or " 160amperes" in nome or "160 amp" in nome or "160amp" in nome:
            
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 160A"})
            return
        
    if "bob" not in nome and "mono" not in nome and "monovolt" not in nome:           
        if "200" in nome or " 200a" in nome or " 200 a" in nome or " 200 amperes" in nome or " 200amperes" in nome or "200 amp" in nome or "200amp" in nome:
            
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 200A"})
            return
        
    if "bob" not in nome and "pfc" not in nome and "pro" not in nome and "edition" not in nome and "220v" not in nome:          
        if "220" in nome or " 220a" in nome or " 220 a" in nome or " 220 amperes" in nome or " 220amperes" in nome or "220 amp" in nome or "220amp" in nome:
            
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 220A"})
            return
        
    if "bob" not in nome and "pfc" not in nome and "pro" not in nome and "edition" not in nome:          
        if "240" in nome or " 240a" in nome or " 240 a" in nome or " 240 amperes" in nome or " 240amperes" in nome or "240 amp" in nome or "240amp" in nome:
            
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 240A"})
            return
        
    if "bob" not in nome and "pfc" not in nome and "pro" not in nome and "edition" not in nome:          
        if "260" in nome or " 260a" in nome or " 260 a" in nome or " 260 amperes" in nome or " 260amperes" in nome or "260 amp" in nome or "260amp" in nome:
            
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 260A"})
            return
        
    if "bob" not in nome and "pfc" not in nome and "pro" not in nome and "edition" not in nome:          
        if "300" in nome or " 300a" in nome or " 300 a" in nome or " 300 amperes" in nome or " 300amperes" in nome or "300 amp" in nome or "300amp" in nome:
            
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 300A"})
            return
        
    if "bob" not in nome and "pfc" not in nome and "pro" not in nome and "edition" not in nome:          
        if "320" in nome or " 320a" in nome or " 320 a" in nome or " 320 amperes" in nome or " 320amperes" in nome or "320 amp" in nome or "320amp" in nome:
            
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 320A"})
            return
        
    if "bob" not in nome and "pfc" not in nome and "pro" not in nome and "edition" not in nome:          
        if "400" in nome or " 400a" in nome or " 400 a" in nome or " 400 amperes" in nome or " 400amperes" in nome or "400 amp" in nome or "400amp" in nome:
            
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 400A"})
            return
        
    if ("mono" in nome or "monovolt" in nome) and "bob" not in nome:          
        if "200" in nome or " 200a" in nome or " 200 a" in nome or " 200 amperes" in nome or " 200amperes" in nome or "200 amp" in nome or "200amp" in nome:
            
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE 200A MONO"})
            return
        
    if "bob" in nome:          
        if "60" in nome or " 60a" in nome or " 60 a" in nome or " 60 amperes" in nome or " 60amperes" in nome or "60 amp" in nome or "60amp" in nome:
            
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE BOB 60A"})
            return
        
    if "bob" in nome:          
        if "60" in nome or " 120a" in nome or " 120 a" in nome or " 120 amperes" in nome or " 120amperes" in nome or "120 amp" in nome or "120amp" in nome:
            
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE BOB 120A"})
            return
        
    if "bob" in nome:          
        if "200" in nome or " 200a" in nome or " 200 a" in nome or " 200 amperes" in nome or " 200amperes" in nome or "200 amp" in nome or "30amp" in nome:
            
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE BOB 200A"})
            return
        
    if "bob" not in nome:          
        if "120" in nome or " 120a" in nome or " 120 a" in nome or " 120 amperes" in nome or " 120amperes" in nome or "120" in nome or "120 amp" in nome or "120amp" in nome:
            
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE PFC 120A"})
            return
        
    if "bob" not in nome:          
        if "240" in nome or " 240a" in nome or " 240 a" in nome or " 240 amperes" in nome or " 240amperes" in nome or "240" in nome or "240 amp" in nome or "240amp" in nome: 
            
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE PFC 240A"})
            return
        
    if "bob" not in nome:          
        if "320" in nome or " 320a" in nome or " 320 a" in nome or " 320 amperes" in nome or " 320amperes" in nome or "320" in nome or "320 amp" in nome or "320amp" in nome:
            
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE PFC 320A"})
            return
        
    if "bob" not in nome:          
        if "500" in nome or "500a" in nome or "500 or 500a" in nome or "500 amperes" in nome or "500amperes" in nome or "500" in nome or "500 amp" in nome or "500amp" in nome:
            
            items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "FONTE PFC 500A"})
            return
    

    if "fonte" in nome:
        items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "POSSIVEL FONTE"})    
        return
    
    items.append({"Vendedor": item["Vendedor"], "Produto": nome,"Marca": item["Marca"],"Frete Grátis": item["Frete Grátis"], "Qtde": item["Qtde"], "Preço Unitário": price, "Total": total, "Produto2": "OUTROS"})

parser = argparse.ArgumentParser(description='Processar datas de início e fim.')
parser.add_argument('--dia_inicial', type=str, required=True, help='Data inicial no formato AAAA-MM-DD')
parser.add_argument('--dia_final', type=str, required=True, help='Data final no formato AAAA-MM-DD')

args = parser.parse_args()

dia_inicial = args.dia_inicial
dia_final = args.dia_final
urls = ["USINA%20LIMITED", "USINA", "USINA%20SPARK", "SPARK"]             
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
df.to_excel("modelos_usina.xlsx", index=False)