import re
import time
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    ElementNotInteractableException,
    ElementClickInterceptedException,
    NoSuchElementException,
)

# =========================
# DICIONÁRIO CORRETORES
# (mantido como você enviou)
# =========================
corretores_gerentes = {
    # Versão original
    "almirat": "Dircesantos",
    "apolo": "Storm",
    "ANASTACIA": "Serafina",
    "Anval-RH": "Emy",
    "Berlim": "Silas",
    "Breya-IND": "Emy",
    "Cikai": "camacho",
    "celia": "camacho",
    "city": "storm",
    "Coliseu": "Dircesantos",
    "Jaguar": "Mariano",
    "Lucas": "Mariano",
    "Shirlei": "Sandrafrancino",
    "WAGNER": "Dircesantos",
    "Zuppo": "Deli",

    # Variações anteriores
    "ANVAL": "Emy",
    "APOLO": "Storm",
    "BERLIM": "Silas",
    "BREYA": "Emy",
    "C.IKAI": "camacho",
    "CELIA": "camacho",
    "CITTY": "storm",
    "CITY": "storm",
    "COLISEU": "Dircesantos",
    "JAGUAR": "Mariano",
    "LUCAS": "Mariano",
    "SHIRLEI": "Sandrafrancino",
    "WAGNER": "Dircesantos",
    "ZUPPO": "Deli",

    # Novos corretores da lista
    "Zoe": "Toro",
    "ZOE": "Toro",
    "Vida": "Tuguchi",
    "VIDA": "Tuguchi",
    "Veridiana": "Silas",
    "VERIDIANA": "Silas",
    "Venus": "Franciscojunior",
    "VENUS": "Franciscojunior",
    "Valentina": "Sandrafrancino",
    "VALENTINA": "Sandrafrancino",
    "Tuna": "Silas",
    "TUNA": "Silas",
    "Adam": "Ruka",
    "ADAM": "Ruka",
    "Aguia": "Henrika",
    "AGUIA": "Henrika",
    "Alemanha": "Serafina",
    "ALEMANHA": "Serafina",
    "Aloisio": "Jaar",
    "ALOISIO": "Jaar",
    "Anjos": "Serafina",
    "ANJOS": "Serafina",
    "Ava": "Jaar",
    "AVA": "Jaar",
    "Bill": "Myrna",
    "BILL": "Myrna",
    "Bulgarelli": "Myrna",
    "BULGARELLI": "Myrna",
    "Cartola": "Ruka",
    "CARTOLA": "Ruka",
    "Catania": "Toro",
    "CATANIA": "Toro",
    "Cristiano": "Franciscojunior",
    "CRISTIANO": "Franciscojunior",
    "Daiane": "Serafina",
    "DAIANE": "Serafina",
    "Dario": "Myrna",
    "DARIO": "Myrna",
    "Djalma": "Silas",
    "DJALMA": "Silas",
    "Felipa": "Serafina",
    "FELIPA": "Serafina",
    "Gabbana": "Henrika",
    "GABBANA": "Henrika",
    "Gisele": "Henrika",
    "GISELE": "Henrika",
    "Gloria": "Gloria",
    "GLORIA": "Gloria",
    "Gomes": "Jaar",
    "GOMES": "Jaar",
    "Graziel": "Myrna",
    "GRAZIEL": "Myrna",
    "Grego": "Deli",
    "GREGO": "Deli",
    "Jaci": "Logan",
    "JACI": "Logan",
    "Japa": "Mariano",
    "JAPA": "Mariano",
    "Joana": "Joana",
    "JOANA": "Joana",
    "Juquei": "Serafina",
    "JUQUEI": "Serafina",
    "Kaline": "Ruka",
    "KALINE": "Ruka",
    "Leila": "Jaar",
    "LEILA": "Jaar",
    "Lelo": "Toro",
    "LELO": "Toro",
    "Leticia": "Serafina",
    "LETICIA": "Serafina",
    "Libanesa": "Ruka",
    "LIBANESA": "Ruka",
    "lyra": "Jaar",
    "LYRA": "Jaar",
    "Maurelio": "Ruka",
    "MAURELIO": "Ruka",
    "Monaco": "Deli",
    "MONACO": "Deli",
    "Natasha": "Serafina",
    "NATASHA": "Serafina",
    "Pitt": "Jaar",
    "PITT": "Jaar",
    "Portugal": "Henrika",
    "PORTUGAL": "Henrika",
    "Quartier": "Cesarricardo",
    "QUARTIER": "Cesarricardo",
    "Ramos": "Mariano",
    "RAMOS": "Mariano",
    "Ravena": "Silas",
    "RAVENA": "Silas",
    "Ribono": "Emy",
    "RIBONO": "Emy",
    "Ruby": "Ruka",
    "RUBY": "Ruka",
    "Sabag": "Myrna",
    "SABAG": "Myrna",
    "Serer": "Myrna",
    "SERER": "Myrna",
    "Seth": "Myrna",
    "SETH": "Myrna",
    "Sol": "Ruka",
    "SOL": "Ruka",
    "Tamashiro": "Emy",
    "TAMASHIRO": "Emy",
    "Tatiana": "Silas",
    "TATIANA": "Silas",

    # Outros corretores que apareceram nos logs anteriores
    "CERQUEIRA": "Storm",
    "CHARMER": "Dircesantos",
    "DANTAS": "Mariano",
    "DARC": "Sandrafrancino",
    "ELLEN": "Emy",
    "ELOHIM": "Silas",
    "LELLO": "Serafina",
    "MANJON": "Dircesantos",
    "NECHI": "Storm",
    "PASQUALINA": "Mariano",
    "RAFAELA": "Emy",
    "SENNA": "camacho",
    "SOLIS": "Mariano",
    "SORIANO": "Deli",
    "TIFFANY": "Storm",
    "Stark":"Logan",
    "STARK":"Logan",
    "Debora":"Caioamon",
    "DEBORA":"Caioamon",
    "Ricca":"Tuguchi",
    "RICCA":"Tuguchi",
    "DUCARMO":"Tuguchi",
    "FORTALEZA":"Myrna",
    "GARCIA":"Flaviabrugnara",
    "GISELE":"Henrika",
    "MANJON":"Deli",
    "MOURA":"Storm",
    "NATASHA":"Serafina",
    "NECHI":"Roseane",
    "ALANIS":"storm",
    "Alanis":"STORM",
    "Maroka":"Toro",
    "Maeva":"Tuguchi",
    "irani":"Mariano",
    "Edbol":"Henrika",
    "Avelino":"Dircesantos",
    "Silmara":"Serafina",
    "Bigode":"Reginaldocarpegiane",
    "Katllyn":"Ruka",
    "MION": "SERAFINA",
    "PIETRA": "LOGAN",
    "NEUSA": "TUGUCHI",
    "CHARLOTTE": "MARIANO",
    "FATIMA": "RUKA",
    "MION": "Serafina",
    "PIETRA": "Logan",
    "NEUSA": "Tuguchi",
    "CHARLOTTE": "Mariano",
    "FATIMA": "Ruka",
}

# =========================
# LEITURA DA PLANILHA
# =========================
arquivo_excel = './SP360_2204a08092025_aby.xlsx'
excel_file     = pd.ExcelFile(arquivo_excel)
nome_planilha  = excel_file.sheet_names[0]
df             = pd.read_excel(arquivo_excel, sheet_name=nome_planilha)

df = df.rename(columns={
    'CORRETOR ORIGEM': 'CORRETOR DE ORIGEM',
    'TELEFONE'       : 'FONE2',
    'NOME COMPLETO'  : 'NOME'
})

# =========================
# CHROME + HELPERS
# =========================
driver = webdriver.Chrome()
driver.maximize_window()
wait = WebDriverWait(driver, 20)

def wait_visible(locator, timeout=20):
    return WebDriverWait(driver, timeout).until(
        EC.visibility_of_element_located(locator)
    )

def wait_clickable(locator, timeout=20):
    return WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable(locator)
    )

def scroll_into_view(el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)

def safe_click(locator):
    el = wait_visible(locator)
    scroll_into_view(el)
    try:
        wait_clickable(locator)
        el.click()
    except (ElementClickInterceptedException, ElementNotInteractableException, TimeoutException):
        driver.execute_script("arguments[0].click();", el)

# =========================
# LOGIN
# =========================
driver.get("https://abyara.sigavi360.com.br/Acesso/Login?ReturnUrl=%2F")
time.sleep(2)

wait_visible((By.XPATH, "/html/body/div[2]/section/div[1]/div/div/div/form/div[1]/div[1]/div/input"))\
    .send_keys("pegomessouza@gmail.com")
wait_visible((By.XPATH, "/html/body/div[2]/section/div[1]/div/div/div/form/div[1]/div[2]/div/input"))\
    .send_keys("123456")
safe_click((By.XPATH, "/html/body/div[2]/section/div[1]/div/div/div/form/div[1]/div[3]/div/button"))
time.sleep(2)

# =========================
# LOOP DE CADASTRO
# =========================
driver.get('https://abyara.sigavi360.com.br/CRM/Fac')
time.sleep(2)

# mapa para tolerar variações de caixa no dicionário
mapa_corretores = {str(k).upper(): v for k, v in corretores_gerentes.items()}

for index, row in df.iterrows():
    nome = str(row.get('NOME') or '').strip()

    # sanitiza telefone (só dígitos)
    telefone_raw = str(row.get('FONE2') or '')
    telefone = re.sub(r'\D', '', telefone_raw)

    corretor_original_raw = str(row.get('CORRETOR DE ORIGEM') or '')
    corretor_original_norm = re.sub(r'\s+', ' ', corretor_original_raw).strip().upper()

    if corretor_original_norm not in mapa_corretores:
        print(f"Corretor '{corretor_original_raw}' não encontrado. Pulando {nome}.")
        continue

    gerente  = mapa_corretores[corretor_original_norm]
    corretor = corretor_original_raw.strip()  # mantém como veio na planilha para digitar no dropdown

    print(f"Processando: {nome} - {telefone} | Corretor: {corretor} | Gerente: {gerente}")

    # Página de busca (apenas preenche telefone; não valida duplicidade aqui)
    driver.get('https://abyara.sigavi360.com.br/CRM/Fac')
    time.sleep(3)

    telefone_busca_locator = (By.XPATH, '/html/body/section/section/div/div/div[2]/div/div[1]/form/div[2]/div/div/div/div[10]/div[4]/input')
    telefone_elem_busca = wait_visible(telefone_busca_locator)
    scroll_into_view(telefone_elem_busca)

    # limpa e digita
    ActionChains(driver)\
        .click(on_element=telefone_elem_busca)\
        .key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL)\
        .send_keys(Keys.DELETE)\
        .send_keys(telefone)\
        .perform()

    # navegação leve (como no seu script)
    ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
    time.sleep(0.5)
    ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
    time.sleep(0.5)

    # Vai direto ao cadastro
    driver.get('https://abyara.sigavi360.com.br/CRM/Fac/Cadastro')
    time.sleep(2)

    # === BLOCO CADASTRO ===
    wait_visible((By.ID, 'Nome')).send_keys(nome)
    time.sleep(1)

    # Abre o bloco de telefones
    safe_click((By.XPATH, '/html/body/div[2]/form/div[2]/div/div/div[1]/div[2]/div[1]/div/div/a'))
    time.sleep(1)

    # Seleciona "Celular" no tipo (setas + enter)
    celular_combo_locator = (By.XPATH, '/html/body/div[2]/form/div[2]/div/div/div[1]/div[2]/div[1]/div/table/tbody/tr/td[1]/span[1]/span/span[1]')
    safe_click(celular_combo_locator)
    ActionChains(driver).send_keys(Keys.ARROW_DOWN, Keys.ARROW_DOWN, Keys.ENTER).perform()

    # Preenche número
    telefone_grid_input_locator = (By.XPATH, '/html/body/div[2]/form/div[2]/div/div/div[1]/div[2]/div[1]/div/table/tbody/tr/td[3]/input')
    tel_input = wait_visible(telefone_grid_input_locator)
    ActionChains(driver)\
        .click(on_element=tel_input)\
        .key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL)\
        .send_keys(Keys.DELETE).send_keys(telefone)\
        .perform()
    time.sleep(1)

    # Adiciona telefone (ícone de +/confirmar)
    safe_click((By.XPATH, '/html/body/div[2]/form/div[2]/div/div/div[1]/div[2]/div[1]/div/table/tbody/tr/td[4]/a[1]/span'))
    time.sleep(1)

    # Canal (SMS)
    sms_combo_locator = (By.XPATH, '/html/body/div[2]/form/div[3]/div/div/div[1]/div[1]/div[1]/div[1]/div[1]/span[2]/span/span[1]')
    safe_click(sms_combo_locator)
    ActionChains(driver).send_keys(Keys.ARROW_DOWN, Keys.ENTER).perform()
    time.sleep(1)

    # Mídia
    midia_combo_locator = (By.XPATH, '/html/body/div[2]/form/div[3]/div/div/div[1]/div[1]/div[1]/div[1]/div[2]/span[2]/span/span[1]')
    safe_click(midia_combo_locator)
    # várias setas para selecionar (mantido)
    ActionChains(driver)\
        .send_keys(Keys.ARROW_DOWN, Keys.ARROW_DOWN, Keys.ARROW_DOWN, Keys.ARROW_DOWN,
                   Keys.ARROW_DOWN, Keys.ARROW_DOWN, Keys.ARROW_DOWN, Keys.ARROW_DOWN,
                   Keys.ARROW_DOWN, Keys.ARROW_DOWN, Keys.ARROW_DOWN, Keys.ARROW_DOWN,
                   Keys.ARROW_DOWN, Keys.ARROW_DOWN, Keys.ARROW_DOWN, Keys.ARROW_DOWN,
                   Keys.ARROW_DOWN, Keys.ARROW_DOWN, Keys.ARROW_DOWN, Keys.ENTER)\
        .perform()
    time.sleep(1)

    # Equipe (gerente)
    equipe_combo_locator = (By.XPATH, '/html/body/div[2]/form/div[3]/div/div/div[1]/div[1]/div[1]/div[2]/div[1]/span[1]/span/span[1]')
    safe_click(equipe_combo_locator)
    ActionChains(driver).send_keys(gerente, Keys.ENTER).perform()
    time.sleep(0.8)

    # Corretor
    corretor_combo_locator = (By.XPATH, '/html/body/div[2]/form/div[3]/div/div/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/span[1]/span/span[1]')
    safe_click(corretor_combo_locator)
    ActionChains(driver).send_keys(corretor, Keys.ENTER).perform()
    time.sleep(0.8)

    # Abre modal "Imóvel/Origem"
    safe_click((By.XPATH, "/html/body/div[2]/form/div[3]/div/div/div[1]/div[2]/div[2]/a/span"))
    modal_container = wait_visible((By.XPATH, "/html/body/div[2]/div[2]/div/div"))

    # Seleciona a opção dentro do modal
    safe_click((By.XPATH, "/html/body/div[2]/div[2]/div/div/div[2]/div[1]/div/div/label[2]"))

    # Preenche o código/descrição
    your_code_input_locator = (By.XPATH, "/html/body/div[2]/div[2]/div/div/div[2]/div[3]/div[1]/input")
    your_code = wait_visible(your_code_input_locator)
    your_code.clear()
    your_code.send_keys("SP")
    time.sleep(0.5)

    # Confirma modal
    safe_click((By.XPATH, "/html/body/div[2]/div[2]/div/div/div[2]/div[3]/div[2]/button"))
    time.sleep(1)

    # Botão que às vezes fica atrás de overlay
    safe_click((By.CSS_SELECTOR, "#dvImovelOrigemComando a"))
    time.sleep(1)

    # 1) Confirmação inicial do formulário
    safe_click((By.XPATH, "/html/body/div[2]/form/div[1]/div/div[1]/button[2]"))
    time.sleep(2)

    # 2) Tenta fechar popup de duplicidade (se existir)
    try:
        safe_click((By.XPATH, '//*[@id="popVerificaDuplicidade"]/div/div/div[3]/button'))
        time.sleep(0.5)
    except Exception:
        pass

    # 3) Salvar
    safe_click((By.XPATH, '//*[@id="cmdSalva"]'))
    time.sleep(3)

print("Processamento concluído!")
driver.quit()
