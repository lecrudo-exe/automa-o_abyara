import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time

# Dicionário de corretores e seus respectivos gerentes (completo)
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
    "Sandy": "Azury",
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
    "SANDY": "Azury",
    "SHIRLEI": "Sandrafrancino",
    "WAGNER": "Dircesantos",
    "ZUPPO": "Deli",
    
    # Novos corretores da lista
    "Zoe": "Toro",
    "ZOE": "Toro",  # Versão em maiúsculas
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
    "Baronesa": "Camila Ferreira",
    "BARONESA": "Camila Ferreira",
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
    "DRIKA": "Azury",
    "ELLEN": "Emy",
    "ELOHIM": "Silas",
    "LELLO": "Serafina",
    "MANJON": "Dircesantos",
    "NECHI": "Storm",
    "PAES": "Azury",
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
}

# ====== INÍCIO – Leitura da nova planilha ======
arquivo_excel = './ABY-CARPEGIANE(uso).xlsx'          # ← nome do arquivo
excel_file = pd.ExcelFile(arquivo_excel)

print("Planilhas no arquivo:", excel_file.sheet_names)
nome_planilha = excel_file.sheet_names[0]        # usa a 1ª aba; altere se quiser
print(f"Usando a planilha: {nome_planilha}")

df = pd.read_excel(arquivo_excel, sheet_name=nome_planilha)
print("Colunas disponíveis:", df.columns.tolist())

# Renomeia colunas da nova planilha para manter compatibilidade
df = df.rename(columns={
    'CORRETOR ORIGEM': 'CORRETOR DE ORIGEM',
    'TELEFONE'       : 'FONE2',
    'NOME COMPLETO'  : 'NOME'
})
# ====== FIM – Leitura da nova planilha ======

# Inicializa o navegador
driver = webdriver.Chrome()
driver.maximize_window()

# LOGIN
driver.get("https://abyara.sigavi360.com.br/Acesso/Login?ReturnUrl=%2F")
time.sleep(2)

email_elem = driver.find_element(By.XPATH, "/html/body/div[2]/section/div[1]/div/div/div/form/div[1]/div[1]/div/input")
email_elem.send_keys("pegomessouza@gmail.com")

senha_elem = driver.find_element(By.XPATH, "/html/body/div[2]/section/div[1]/div/div/div/form/div[1]/div[2]/div/input")
senha_elem.send_keys("123456")

btn_elem = driver.find_element(By.XPATH, "/html/body/div[2]/section/div[1]/div/div/div/form/div[1]/div[3]/div/button")
btn_elem.click()
time.sleep(2)

# CAMINHO PARA A PÁGINA CERTA
driver.get('https://abyara.sigavi360.com.br/CRM/Fac')
time.sleep(2)

# ====== LOOP PARA FAZER A AUTOMAÇÃO ======
for index, row in df.iterrows():
    nome = str(row['NOME'])
    telefone = str(row['FONE2'])
    corretor_original = str(row['CORRETOR DE ORIGEM']).strip()
    
    if corretor_original not in corretores_gerentes:
        print(f"Corretor '{corretor_original}' não encontrado. Pulando {nome}.")
        continue
    
    gerente = corretores_gerentes[corretor_original]
    corretor = corretor_original
    
    print(f"Processando: {nome} - {telefone} | Corretor: {corretor} | Gerente: {gerente}")
    
    driver.get('https://abyara.sigavi360.com.br/CRM/Fac')
    time.sleep(3)
    
    telefone_elem = driver.find_element(By.XPATH, '/html/body/section/section/div/div/div[2]/div/div[1]/form/div[2]/div/div/div/div[10]/div[4]/input')
    ActionChains(driver).click(on_element=telefone_elem)\
        .key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL)\
        .send_keys(Keys.DELETE)\
        .send_keys(telefone)\
        .perform()
    
    ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
    time.sleep(0.5)
    ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
    time.sleep(0.5)
    
    driver.find_element(By.XPATH, '//*[@id="cmdBusca"]').click()
    time.sleep(3)
    
    cliente_existe = False
    try:
        outra_pagina = WebDriverWait(driver, 9).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/section/section/div/div/div[2]/div/div[2]'))
        )
        cliente_existe = outra_pagina.value_of_css_property("display") == "block"
    except (TimeoutException, NoSuchElementException):
        cliente_existe = False

    if cliente_existe:
        print(f"Cliente {nome} já existe. Nova busca.")
        time.sleep(2)
        try:
            driver.find_element(By.XPATH, '/html/body/section/section/div/div/div[2]/div/div[2]/div[1]/div[1]/button').click()
            time.sleep(2)
        except Exception as e:
            print(f"Erro ao clicar nova busca: {e}")
    
    else:
        print(f"Cadastrando novo cliente: {nome}")
        driver.get('https://abyara.sigavi360.com.br/CRM/Fac/Cadastro')
        time.sleep(2)
        
        driver.find_element(By.ID, 'Nome').send_keys(nome)
        time.sleep(2)
        
        driver.find_element(By.XPATH, '/html/body/div[2]/form/div[2]/div/div/div[1]/div[2]/div[1]/div/div/a').click()
        time.sleep(2)
        
        celular_elem = driver.find_element(By.XPATH, '/html/body/div[2]/form/div[2]/div/div/div[1]/div[2]/div[1]/div/table/tbody/tr/td[1]/span[1]/span/span[1]')
        celular_elem.click()
        ActionChains(driver).send_keys(Keys.ARROW_DOWN, Keys.ARROW_DOWN, Keys.ENTER).perform()
        
        telefone_elem = driver.find_element(By.XPATH,'/html/body/div[2]/form/div[2]/div/div/div[1]/div[2]/div[1]/div/table/tbody/tr/td[3]/input')
        ActionChains(driver).click(on_element=telefone_elem)\
            .key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL)\
            .send_keys(Keys.DELETE).send_keys(telefone).perform()
        time.sleep(2)
        
        driver.find_element(By.XPATH, '/html/body/div[2]/form/div[2]/div/div/div[1]/div[2]/div[1]/div/table/tbody/tr/td[4]/a[1]/span').click()
        time.sleep(2)
        
        sms_elem = driver.find_element(By.XPATH, '/html/body/div[2]/form/div[3]/div/div/div[1]/div[1]/div[1]/div[1]/div[1]/span[2]/span/span[1]')
        sms_elem.click()
        ActionChains(driver).send_keys(Keys.ARROW_DOWN, Keys.ENTER).perform()
        time.sleep(2)
        
        midia_elem = driver.find_element(By.XPATH, '/html/body/div[2]/form/div[3]/div/div/div[1]/div[1]/div[1]/div[1]/div[2]/span[2]/span/span[1]')
        midia_elem.click()
        ActionChains(driver).send_keys(Keys.ARROW_DOWN, Keys.ARROW_DOWN, Keys.ARROW_DOWN, Keys.ARROW_DOWN, Keys.ENTER).perform()
        time.sleep(2)
        
        equipe_elem = driver.find_element(By.XPATH, '/html/body/div[2]/form/div[3]/div/div/div[1]/div[1]/div[1]/div[2]/div[1]/span[1]/span/span[1]')
        equipe_elem.click()
        ActionChains(driver).send_keys(gerente, Keys.ENTER).perform()
        time.sleep(1)
        
        corretor_elem = driver.find_element(By.XPATH, '/html/body/div[2]/form/div[3]/div/div/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/span[1]/span/span[1]')
        corretor_elem.click()
        ActionChains(driver).send_keys(corretor, Keys.ENTER).perform()
        time.sleep(1)
        
        driver.find_element(By.XPATH, "/html/body/div[2]/form/div[3]/div/div/div[1]/div[2]/div[2]/a/span").click()
        time.sleep(2)
        
        driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div/div/div[2]/div[1]/div/div/label[2]").click()
        time.sleep(2)
        
        your_code = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div/div/div[2]/div[3]/div[1]/input")
        your_code.clear()
        your_code.send_keys("reserva paraiso")
        time.sleep(2)
        
        driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div/div/div[2]/div[3]/div[2]/button").click()
        time.sleep(2)
        
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#dvImovelOrigemComando a"))
        ).click()
        time.sleep(2)
        
        driver.find_element(By.XPATH, "/html/body/div[2]/form/div[1]/div/div[1]/button[2]").click()
        time.sleep(5)
        
        try:
            dup = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[8]'))
            )
            if dup.value_of_css_property("display") == "none":
                driver.find_element(By.XPATH, "/html/body/div[2]/form/div[1]/div/div[1]/button[1]").click()
                print(f"Cliente {nome} cadastrado com sucesso!")
                time.sleep(5)
            else:
                print(f"Cliente {nome} já existe (duplicidade).")
        except:
            print("Erro na verificação de duplicidade.")
        
        finally:
            driver.get("https://abyara.sigavi360.com.br/CRM/Fac/")
            time.sleep(2)

print("Processamento concluído!")
driver.quit()