import pyautogui as pa
from selenium import webdriver
from selenium.webdriver.common.by import By
import time

# Configuração de pausa entre comandos do PyAutoGUI
pa.PAUSE = 1.3
time.sleep(1.0)  # Pausa inicial para garantir que todas as bibliotecas estejam prontas

# Inicializa o navegador
navegador = webdriver.Chrome()

# Etapa 1: Login no site "Suas Vendas"
navegador.get("https://www.suasvendas.com/login")
navegador.find_element(By.ID, "MainPageContent_txtEmail").click()
pa.typewrite('mateus.papieri@gmail.com')  # Insere o e-mail
navegador.find_element(By.XPATH, '//*[@id="MainPageContent_txtSenha"]').click()
pa.typewrite('mateus21')  # Insere a senha
navegador.find_element(By.ID, "btnAcessar").click()
time.sleep(2.5)  # Aguarda o carregamento da página

# Etapa 2: Acessa a seção de Lançamentos
navegador.find_element(By.XPATH, '//*[@id="menuCollapse"]/ul/li[4]/a').click()
time.sleep(1.0)
navegador.find_element(By.XPATH, '//*[@id="ctl00_ctl00_C2_cphAbaixoTitulo_SubMenu1_LnkLancamento"]').click()
time.sleep(2.0)

# Etapa 3: Filtros de "Tipo"
navegador.find_element(By.XPATH, '//*[@id="btnFiltrar"]').click()
time.sleep(0.7)
navegador.find_element(By.XPATH, '//*[@id="cbbFiltrar"]/div/a[1]').click()
time.sleep(0.4)
navegador.find_element(By.XPATH, '//*[@id="CblFilter_lanc_tipo"]/span[2]/label').click()  # Seleciona o tipo
navegador.find_element(By.XPATH, '//*[@id="ctl00_ctl00_C2_GridViewFilters_Fillanc_tipo_BtnFilter"]').click()
time.sleep(1.0)

# Etapa 4: Aplicando filtro "Grupo de Conta"
navegador.find_element(By.XPATH, '//*[@id="btnFiltrar"]').click()
time.sleep(0.7)
navegador.find_element(By.XPATH, '//*[@id="cbbFiltrar"]/div/a[2]').click()
time.sleep(0.4)

# Seleciona várias opções de "Grupo de Conta"
grupo_conta_opcoes = [7, 4, 6, 8]
for opcao in grupo_conta_opcoes:
    navegador.find_element(By.XPATH, f'//*[@id="CblFilter_cate_nome"]/span[{opcao}]/label').click()
    time.sleep(0.2)

navegador.find_element(By.XPATH, '//*[@id="ctl00_ctl00_C2_GridViewFilters_Filcate_nome_BtnFilter"]').click()
time.sleep(1.0)

# Etapa 5: Filtro de "Liquidado"
navegador.find_element(By.XPATH, '//*[@id="btnFiltrar"]').click()
time.sleep(0.7)
navegador.find_element(By.XPATH, '//*[@id="cbbFiltrar"]/div/a[3]').click()
time.sleep(0.4)
navegador.find_element(By.XPATH, '//*[@id="CblFilter_lanc_liquidado"]/span[1]/label').click()
navegador.find_element(By.XPATH, '//*[@id="ctl00_ctl00_C2_GridViewFilters_Fillanc_liquidado_BtnFilter"]').click()
time.sleep(1.0)

# Etapa 6: Filtro de "SV.PAG"
navegador.find_element(By.XPATH, '//*[@id="btnFiltrar"]').click()
time.sleep(0.7)
navegador.find_element(By.XPATH, '//*[@id="cbbFiltrar"]/div/a[4]').click()
time.sleep(0.4)

# Seleciona várias opções de "SV.PAG"
sv_pag_opcoes = [1, 5, 7]
for opcao in sv_pag_opcoes:
    navegador.find_element(By.XPATH, f'//*[@id="CblFilter_lanc_svpag_situacao"]/span[{opcao}]/label').click()
    time.sleep(0.4)

navegador.find_element(By.XPATH, '//*[@id="ctl00_ctl00_C2_GridViewFilters_Fillanc_svpag_situacao_BtnFilter"]').click()
time.sleep(1.0)

# Etapa 7: Filtro de "Status"
navegador.find_element(By.XPATH, '//*[@id="btnFiltrar"]').click()
time.sleep(0.7)
navegador.find_element(By.XPATH, '//*[@id="cbbFiltrar"]/div/a[5]').click()
time.sleep(0.4)
navegador.find_element(By.XPATH, '//*[@id="CblFilter_lanc_periodo"]/span[1]/label').click()
navegador.find_element(By.XPATH, '//*[@id="ctl00_ctl00_C2_GridViewFilters_Fillanc_periodo_BtnFilter"]').click()
time.sleep(0.7)

# Etapa 8: Download do relatório
navegador.find_element(By.XPATH, '//*[@id="ctl00_ctl00_C2_LnkAutoReport"]/b').click()
time.sleep(2.0)
# Navega no relatório para salvar ou baixar
for _ in range(4):
    pa.hotkey('Tab')
pa.hotkey('Enter')
time.sleep(2.5)
pa.hotkey('ctrl', 'w')

# Etapa 9: Acessa seção "Solicitações"
navegador.get("https://app6.suasvendas.com/Modulo/YourSales/Financeiro/SVPag/SolicitacoesSVPag.aspx")
time.sleep(2.5)

# Filtros de "Solicitações" (repetição similar aos passos anteriores)
navegador.find_element(By.XPATH, '//*[@id="btnFiltrar"]').click()
time.sleep(0.7)
navegador.find_element(By.XPATH, '//*[@id="cbbFiltrar"]/div/a[2]').click()
time.sleep(0.4)
navegador.find_element(By.XPATH, '//*[@id="CblFilter_vencida"]/span[2]/label').click()
navegador.find_element(By.XPATH, '//*[@id="ctl00_ctl00_C2_GridViewFilters_Filvencida_BtnFilter"]').click()
time.sleep(0.7)

# Filtro de Status
navegador.find_element(By.XPATH, '//*[@id="btnFiltrar"]').click()
time.sleep(0.7)
navegador.find_element(By.XPATH, '//*[@id="cbbFiltrar"]/div/a[1]').click()
time.sleep(0.4)
navegador.find_element(By.XPATH, '//*[@id="CblFilter_lanc_svpag_situacao"]/span[4]/label').click()
navegador.find_element(By.XPATH, '//*[@id="CblFilter_lanc_svpag_situacao"]/span[6]').click()
navegador.find_element(By.XPATH, '//*[@id="ctl00_ctl00_C2_GridViewFilters_Fillanc_svpag_situacao_BtnFilter"]').click()
time.sleep(0.7)

# Download do segundo relatório
navegador.find_element(By.XPATH, '//*[@id="ctl00_ctl00_C2_LnkAutoReport"]/b').click()
time.sleep(2.0)
for _ in range(4):
    pa.hotkey('Tab')
pa.hotkey('Enter')

# Etapa 10: Organiza o arquivo baixado em pastas específicas
pa.hotkey('win', 'e')
pa.press('right')
pa.press('Enter')
pa.hotkey('ctrl', 'shift', 'n')
pa.typewrite('SV att')
pa.press('Enter')
pa.press('left')
pa.hotkey('ctrl', 'x')
pa.press('right')
pa.press('enter')
pa.hotkey('ctrl', 'v')
pa.press('backspace')

# Etapa 11: Integração com Google Sheets para análise de dados
pa.press('Win')
pa.typewrite('Chrome')
pa.press('Enter')
pa.typewrite('https://docs.google.com/spreadsheets/d/1uBoZ4WTBm4fCB8hYP6i81Vg3ns7SW8h87nlpZvfGBYw/edit#gid=0')
pa.press('Enter')
time.sleep(2.5)

# Carrega dados do relatório para o Google Sheets
pa.hotkey('Alt', 'Shift', 'F')
time.sleep(0.5)
for _ in range(3):
    pa.press('down')
pa.press('Enter')
time.sleep(0.5)
pa.moveTo(660, 285)
pa.click(x=660, y=285)
pa.press('Enter')
time.sleep(2.0)
pa.typewrite('SV att')
pa.press('enter')

# Finalização
pa.hotkey('ctrl', 'Alt', 'Shift', '8')  # Envia o relatório
time.sleep(5.0)

# Limpeza dos arquivos antigos
pa.hotkey('win', 'e')
time.sleep(0.5)
pa.typewrite('SV')
time.sleep(0.5)
pa.press('delete')
