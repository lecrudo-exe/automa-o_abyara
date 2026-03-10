import re
import time
import json
import unicodedata
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
    "Alemanha": "Andremarques",
    "ALEMANHA": "Andremarques",
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

# sobrescreve com o mapeamento mais completo do JSON, se disponível
try:
    with open('corretores.json', 'r', encoding='utf-8') as f:
        corretores_json = json.load(f)
        if isinstance(corretores_json, list) and corretores_json:
            corretores_gerentes = corretores_json[0].get('corretoresEquipes', corretores_gerentes)
        elif isinstance(corretores_json, dict):
            corretores_gerentes = corretores_json.get('corretoresEquipes', corretores_gerentes)
except Exception as e:
    print(f"Não foi possível carregar corretores.json, usando dicionário interno. Erro: {e}")

# =========================
# LEITURA DA PLANILHA
# =========================
arquivo_excel = './PLANTAO_HOME_STORE2.xlsx'
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

def normalizar_texto(valor: str) -> str:
    if valor is None:
        return ''
    nfkd = unicodedata.normalize('NFKD', str(valor))
    sem_acento = ''.join(c for c in nfkd if not unicodedata.combining(c))
    upper = sem_acento.upper()
    apenas_alfa = re.sub(r'[^A-Z0-9]+', ' ', upper)
    return re.sub(r'\s+', ' ', apenas_alfa).strip()

def selecionar_midia_por_posicao(posicao: int):
    ac = ActionChains(driver)
    for _ in range(posicao):
        ac.send_keys(Keys.ARROW_DOWN)
    ac.send_keys(Keys.ENTER)
    ac.perform()

midia_mapeamentos_brutos = [
    ("123I", 1),
    ("APTO.VC", 2),
    ("CARTEIRA", 3),
    ("CHAVES NA MAO", 4),
    ("CHAVES NA MÃO", 4),
    ("FACEBOOK", 5),
    ("FORMULARIO GOOGLE", 6),
    ("FORMULÁRIO GOOGLE", 6),
    ("GOOGLE DISPLAY", 7),
    ("GOOGLE SEARCH", 8),
    ("IMOVEL WEB", 9),
    ("IMÓVEL WEB", 9),
    ("INCORPORADOR", 11),
    ("INDICACAO", 12),
    ("INDICAÇÃO", 12),
    ("IND. CORRETOR", 12),
    ("IND CORRETOR", 12),
    ("INSTAGRAM", 13),
    ("LINKEDIN", 14),
    ("OLX", 15),
    ("OUTROS", 16),
    ("PADARIA", 17),
    ("PLACA", 18),
    ("RD STATION", 19),
    ("RETORNO", 20),
    ("SITE", 21),
    ("VISITACAO", 22),
    ("VISITAÇÃO", 22),
    ("STAND", 22),
    ("ACAO DE RUA", 22),
    ("AÇÃO DE RUA", 22),
    ("STAND/ACAO DE RUA", 22),
    ("STAND/AÇÃO DE RUA", 22),
    ("TWITTER", 23),
    ("VIVA REAL", 24),
    ("VIZINHO", 25),
    ("WHATSAPP", 26),
    ("YOUTUBE", 27),
    ("ZAP", 28),
]

midia_por_tipo_plantao = {}
for chave, posicao in midia_mapeamentos_brutos:
    chave_norm = normalizar_texto(chave)
    if chave_norm not in midia_por_tipo_plantao:
        midia_por_tipo_plantao[chave_norm] = posicao

canal_plantao_tipos = {
    normalizar_texto("VISITACAO"),
    normalizar_texto("VISITAÇÃO"),
    normalizar_texto("RETORNO"),
}

canal_carteira_tipos = {
    normalizar_texto("INDICACAO"),
    normalizar_texto("INDICAÇÃO"),
    normalizar_texto("IND. CORRETOR"),
    normalizar_texto("IND CORRETOR"),
    normalizar_texto("CARTEIRA"),
}

# =========================
# LOGIN
# =========================
driver.get("https://abyara.sigavi360.com.br/Acesso/Login?ReturnUrl=%2F")
time.sleep(2)

wait_visible((By.XPATH, "/html/body/div[2]/section/div[1]/div/div/div/form/div[1]/div[1]/div/input"))\
    .send_keys("pegomessouza@gmail.com")
wait_visible((By.XPATH, "/html/body/div[2]/section/div[1]/div/div/div/form/div[1]/div[2]/div/input"))\
    .send_keys("12345678910")
safe_click((By.XPATH, "/html/body/div[2]/section/div[1]/div/div/div/form/div[1]/div[3]/div/button"))
time.sleep(2)

# =========================
# LOOP DE CADASTRO
# =========================
driver.get('https://abyara.sigavi360.com.br/CRM/Fac')
time.sleep(2)

# mapa para tolerar variações de caixa no dicionário
mapa_corretores = {str(k).upper(): v for k, v in corretores_gerentes.items()}

# mapa de fallback: só o nome base antes de sufixos como "- Inc", "- Pagadoria", "- IND", etc.
# ex: "DIMI - INC" → chave base "DIMI"
mapa_corretores_base = {}
for k, v in mapa_corretores.items():
    base = re.split(r'\s*[-|]\s*', k)[0].strip()
    if base and base not in mapa_corretores_base:
        mapa_corretores_base[base] = v

for index, row in df.iterrows():
    nome = str(row.get('NOME') or '').strip()

    # sanitiza telefone (só dígitos)
    telefone_raw = str(row.get('FONE2') or '')
    telefone = re.sub(r'\D', '', telefone_raw)

    # precisa ter pelo menos DDD (2) + n?mero (9)
    if len(telefone) < 11:
        print(f"Telefone '{telefone_raw}' inv?lido (menos de 11 d?gitos). Pulando {nome}.")
        continue

    corretor_original_raw = str(row.get('CORRETOR DE ORIGEM') or '')
    corretor_original_norm = re.sub(r'\s+', ' ', corretor_original_raw).strip().upper()

    tipo_plantao_raw = str(row.get('TIPO PLANTAO') or '').strip()
    tipo_plantao_norm = normalizar_texto(tipo_plantao_raw)
    posicao_midia = midia_por_tipo_plantao.get(tipo_plantao_norm)

    if posicao_midia is None:
        print(f"Tipo de plantao '{tipo_plantao_raw}' nÃ£o mapeado. Pulando {nome}.")
        continue

    if corretor_original_norm in mapa_corretores:
        gerente  = mapa_corretores[corretor_original_norm]
        corretor = corretor_original_raw.strip()
    elif corretor_original_norm in mapa_corretores_base:
        gerente  = mapa_corretores_base[corretor_original_norm]
        corretor = corretor_original_raw.strip()
        print(f"Corretor '{corretor_original_raw}' encontrado via nome base (sem sufixo). Gerente: {gerente}")
    else:
        print(f"Corretor '{corretor_original_raw}' não encontrado no Sigavi (inativo). Cadastrando como equipe Tabatanascimento / inativo.")
        gerente  = "Tabatanascimento"
        corretor = "Corretor Inativo"

    canal_setas = 4 if tipo_plantao_norm in canal_plantao_tipos else 1
    if tipo_plantao_norm in canal_carteira_tipos:
        canal_setas = 1

    print(f"Processando: {nome} - {telefone} | Corretor: {corretor} | Gerente: {gerente}")

    # Página de busca (apenas preenche telefone; não valida duplicidade aqui)
    telefone_busca_locator = (By.XPATH, '/html/body/section/section/div/div/div[2]/div/div[1]/form/div[2]/div/div/div/div[10]/div[4]/input')
    telefone_elem_busca = None
    for tentativa in range(3):
        driver.get('https://abyara.sigavi360.com.br/CRM/Fac')
        time.sleep(3)
        try:
            telefone_elem_busca = wait_visible(telefone_busca_locator, timeout=20)
            break
        except TimeoutException:
            print(f"Página /CRM/Fac não carregou (tentativa {tentativa+1}/3). Tentando novamente...")
    if telefone_elem_busca is None:
        print(f"Não foi possível carregar a página de busca após 3 tentativas. Pulando {nome}.")
        continue
    scroll_into_view(telefone_elem_busca)

    # limpa e digita
    ActionChains(driver)\
        .click(on_element=telefone_elem_busca)\
        .key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL)\
        .send_keys(Keys.DELETE)\
        .send_keys(telefone)\
        .perform()

    # dispara busca e verifica duplicidade antes de ir para cadastro
    ActionChains(driver).send_keys(Keys.ENTER).perform()
    time.sleep(2.5)  # aguarda carregamento da grade de resultados
    resultado_busca_locator = (By.XPATH, '/html/body/section/section/div/div/div[2]/div/div[2]/div[2]/div[2]/div/div[4]/table/tbody/tr')
    duplicado = False
    try:
        linhas = WebDriverWait(driver, 6).until(EC.presence_of_all_elements_located(resultado_busca_locator))
        if linhas:
            primeira_td = linhas[0].find_element(By.XPATH, './td[1]')
            texto = (primeira_td.text or '').strip()
            texto_linha = (linhas[0].text or '').strip()
            print(f"Resultado da busca para {telefone}: primeira coluna='{texto}', linha='{texto_linha}'")
            if (texto and not texto.upper().startswith('NENHUM')) or (texto_linha and not texto_linha.upper().startswith('NENHUM')):
                print(f"Lead já existe para telefone {telefone}. Pulando {nome}.")
                duplicado = True
    except TimeoutException:
        print(f"Busca por telefone {telefone} não retornou linhas no tempo limite.")
    except Exception as e:
        print(f"Falha ao verificar duplicidade para {telefone}: {e}")

    if duplicado:
        continue

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
    time.sleep(2.5)

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
    tel_input.click()
    tel_input.send_keys(telefone)
    time.sleep(1)

    # Adiciona telefone (ícone de +/confirmar)
    safe_click((By.XPATH, '/html/body/div[2]/form/div[2]/div/div/div[1]/div[2]/div[1]/div/table/tbody/tr/td[4]/a[1]/span'))
    time.sleep(1)

    # Canal (SMS)
    sms_combo_locator = (By.XPATH, '/html/body/div[2]/form/div[3]/div/div/div[1]/div[1]/div[1]/div[1]/div[1]/span[2]/span/span[1]')
    safe_click(sms_combo_locator)
    ac_canal = ActionChains(driver)
    for _ in range(canal_setas):
        ac_canal.send_keys(Keys.ARROW_DOWN)
    ac_canal.send_keys(Keys.ENTER).perform()
    time.sleep(1)

    # Mídia
    midia_combo_locator = (By.XPATH, '/html/body/div[2]/form/div[3]/div/div/div[1]/div[1]/div[1]/div[1]/div[2]/span[2]/span/span[1]')
    safe_click(midia_combo_locator)
    # seleciona conforme TIPO PLANTAO
    selecionar_midia_por_posicao(posicao_midia)
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
    your_code.send_keys("alt studios")
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
        print(f"Lead duplicado encontrado para telefone {telefone}. Pulando {nome}.")
        continue
    except Exception:
        pass

    # 3) Salvar
    safe_click((By.XPATH, '//*[@id="cmdSalva"]'))
    time.sleep(3)

print("Processamento concluído!")
driver.quit()
