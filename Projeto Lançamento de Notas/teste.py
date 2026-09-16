from automations.automation_functions import *
from automations.automation_devolucao import *
from selenium.common.exceptions import NoSuchElementException


driver = conectar_chrome()
mudar_aba(driver, "Portal de importação de XML")

seletor = (
    "button"
    "[ng-click='ctrl.abrirDocumento()']"
)

clicar_campo(driver, seletor)
time.sleep(1.5)

mudar_aba(driver,"Sankhya Om")
time.sleep(1)
mudar_aba(driver,"Sankhya Om")

time.sleep(20)

seletor = "td > div.Taskbar-icon.icon-user-photo"
clicar_campo(driver, seletor)

seletor = (
    "//td[@class='gwt-MenuItem' "
    "and @role='menuitem' "
    "and @colspan='2']"
    "[.//span[@class='gwt-MenuItem-icon icon-photo-stack']]"
    "[.//span[normalize-space()='Layout da tela']]"
)

clicar_campo(driver, seletor)


seletor = "input[type='radio'][name='radioHTML5'][value='on']"

clicar_campo(driver,seletor)

seletor = "//button[@type='button' and normalize-space()='OK']"

clicar_campo(driver, seletor)

time.sleep(5)

mudar_aba(driver, "Central de Vendas")

time.sleep(10)

# PREENCHER 202160200

seletor = (
    "//label[.//span[@title='Natureza']]"
    "/following::input[1]"
)

clicar_campo(driver, seletor)
digitar_no_foco(driver, "202160200")
time.sleep(2)
# DATA DE NEGOCIAÇÃO

seletor = (
    "//label[.//span[@title='Dt. Neg.']]"
    "/following::input[1]"
)

clicar_campo(driver, seletor)
digitar_no_foco(driver, "14092026")
time.sleep(2)


# # DATA ENTRADA E SAÍDA

seletor = (
    "//label[.//span[@title='Dt. Entrada/Saída']]"
    "/following::input[1]"
)

clicar_campo(driver, seletor)
digitar_no_foco(driver, "14092026")
time.sleep(2)

# # DATA DE FATURAMENTO

seletor = (
    "//label[.//span[@title='Dt. do Faturamento']]"
    "/following::input[1]"
)

clicar_campo(driver, seletor)
digitar_no_foco(driver, "14092026")
time.sleep(2)


ActionChains(driver).send_keys(Keys.F7).perform()

time.sleep(8)

######################### PREENCHENDO O CODIGO DO MOTIVO

time.sleep(3)

seletor_tabela = (
    "div[role='row'][row-index='0'] "
    "div[role='gridcell'][col-id='AD_CODMOT']"
)

campos = driver.find_elements(
    By.CSS_SELECTOR,
    seletor_tabela
)

print(f"Quantidade de linhas para Cód. Motivo: {len(campos)}")

quantidade_linhas = len(campos)

clicar_campo(driver, seletor_tabela)


for campo in range(quantidade_linhas):
    digitar_no_foco(driver, "13")
    time.sleep(2)
    ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
    time.sleep(8)

fechar_aba(driver, "Central de Vendas")
