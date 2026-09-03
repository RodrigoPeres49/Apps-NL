import time

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ============================================================

# CONFIGURAÇÃO

# ============================================================

DEBUGGER_ADDRESS = "127.0.0.1:9222"

# ============================================================

# CONECTAR AO CHROME

# ============================================================

def conectar_chrome():


    chrome_options = Options()
    
    chrome_options.add_experimental_option(
        "debuggerAddress",
        DEBUGGER_ADDRESS
    )
    
    driver = webdriver.Chrome(
        options=chrome_options
    )
    
    return driver


# ============================================================

# LOCALIZAR PORTAL

# ============================================================

def localizar_portal(driver):


    print()
    print("================================================")
    print("PROCURANDO JANELA DO PORTAL")
    print("================================================")
    
    handle_portal = None
    
    for handle in driver.window_handles:
    
        try:
    
            driver.switch_to.window(handle)
    
            titulo_original = driver.title
            url_original = driver.current_url
    
            titulo = titulo_original.lower()
            url = url_original.lower()
    
            print("------------------------------------------------")
            print("HANDLE:", handle)
            print("TÍTULO:", titulo_original)
            print("URL:", url_original)
    
            if (
                "Portal de importação de Xml" in titulo
                or "portal de importacao de xml" in titulo
                or "tabcontent.jsp" in url
                and "importacaoxml" in url
            ):
    
                print()
                print("================================================")
                print("PORTAL SANKHYA ENCONTRADO")
                print("================================================")
    
                print("HANDLE DO PORTAL:")
                print(handle)
    
                handle_portal = handle
    
                break
    
        except Exception as erro:
    
            print(
                "Erro ao verificar janela:",
                erro
            )

    # Volta explicitamente para o Portal
    # antes de retornar.
    
    if handle_portal:
    
        driver.switch_to.window(
            handle_portal
        )
    
        print()
        print("JANELA ATUAL APÓS SELEÇÃO:")
        print(driver.title)
    
        return handle_portal
    
    return None




