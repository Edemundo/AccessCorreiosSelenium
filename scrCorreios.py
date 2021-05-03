from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook

driver = webdriver.Chrome(executable_path='chromedriver.exe')
# driver = webdriver.PhantomJS()

# Abrindo planilha com CEPS
workbook = load_workbook(filename="Planilha - CEPS.xlsx")
sheet = workbook.active
ceps = []
i = 2

# Abrindo o Google Chrome na página dos correios
driver.get("https://buscacepinter.correios.com.br/app/endereco/index.php?t")

# Populando vetor ceps com planilha
while(True):
  ceps.append(sheet.cell(row = i, column=1).value)

  if(ceps[i-2] == None):
    ceps.pop()
    break

  ceps[i-2] = str(ceps[i-2])

  if(len(ceps[i-2]) == 6):
    ceps[i-2] = "0" + ceps[i-2] + "0"
  elif(len(ceps[i-2]) == 7):
    ceps[i-2] = "0" + ceps[i-2]
  i = i + 1

# logradouros = []
# bairros = []
# localidades = []
# cepsEncontrados = []
# cepsNaoEncontrados = []

# Criando nova planilha com os CEPS encontrados
workbookNova = Workbook()
sheetNova  = workbookNova.active
sheetNova["A1"] = "CEP"
sheetNova["B1"] = "Logradouro"
sheetNova["C1"] = "Bairro"
sheetNova["D1"] = "Localidade"
workbookNova.save(filename="Ceps Encontrados.xlsx")
# Criando planilha com os ceps não encontrados
workbookNotFind = Workbook()
sheetNotFind  = workbookNotFind.active
sheetNotFind["A1"] = "CEP"
workbookNotFind.save(filename="Ceps Não Encontrados.xlsx")
# Acessando o site do correios com cada um dos CEPS
contEncontrados = 2
contNaoEncontrados = 2
for i in range(len(ceps)):
  cep_end_elem = WebDriverWait(driver, 1000).until(
    EC.presence_of_element_located((By.XPATH, "//input[@name='endereco']"))
  )
  cep_end_elem.clear()
  cep_end_elem.send_keys(ceps[i])

  btn_pesquisar_elem = WebDriverWait(driver, 1000).until(
    EC.presence_of_element_located((By.XPATH, "//button[@name='btn_pesquisar']"))
  )
  btn_pesquisar_elem.click()

  # Caso o CEP exista serão extraidas essas informações
  try:
    td_logradouro_elem = WebDriverWait(driver, 1000).until(
      EC.presence_of_element_located((By.XPATH, "//td[@data-th='Logradouro/Nome']"))
    )

    td_bairro_elem = WebDriverWait(driver, 1000).until(
      EC.presence_of_element_located((By.XPATH, "//td[@data-th='Bairro/Distrito']"))
    )

    td_localidade_elem = WebDriverWait(driver, 1000).until(
      EC.presence_of_element_located((By.XPATH, "//td[@data-th='Localidade/UF']"))
    )

    td_cep_elem = WebDriverWait(driver, 1000).until(
      EC.presence_of_element_located((By.XPATH, "//td[@data-th='CEP']"))
    )

    # Armazenando essas informações em seus respectivos vetores
    # cepsEncontrados.append(td_cep_elem.text)
    # logradouros.append(td_logradouro_elem.text)
    # bairros.append(td_bairro_elem.text)
    # localidades.append(td_localidade_elem.text)

    # Populando nova planilha com informações do cep encontrado
    linhaColunaCEP = "A" + str(contEncontrados)
    sheetNova[linhaColunaCEP] = td_cep_elem.text

    linhaColunaLogra = "B" + str(contEncontrados)
    sheetNova[linhaColunaLogra] = td_logradouro_elem.text

    linhaColunaBairro = "C" + str(contEncontrados)
    sheetNova[linhaColunaBairro] = td_bairro_elem.text

    linhaColunaLocal = "D" + str(contEncontrados)
    sheetNova[linhaColunaLocal] = td_localidade_elem.text
    contEncontrados = contEncontrados + 1
    workbookNova.save(filename="Ceps Encontrados.xlsx")
    print("CEP:", td_cep_elem.text, "adicionado.", "CEPS SALVOS:", i)

  except:
    # Populando nova planilha com informações do cep não encontrado
    linhaColunaCEP = "A" + str(contNaoEncontrados)
    sheetNotFind[linhaColunaCEP] = ceps[i]
    workbookNotFind.save(filename="Ceps Não Encontrados.xlsx")
    print("CEP: ", td_cep_elem.text, " não foi encontrado.", "CEPS SALVOS:", i)

  btn_voltar_elem = WebDriverWait(driver, 1000).until(
    EC.presence_of_element_located((By.XPATH, "//button[@name='btn_voltar']"))
  )
  btn_voltar_elem.click()
  time.sleep(1)

driver.close()
