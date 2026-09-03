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



# ============================================================

# MOSTRAR IFRAMES

# ============================================================

def listar_iframes(driver):


    driver.switch_to.default_content()
    
    iframes = driver.find_elements(
        By.TAG_NAME,
        "iframe"
    )
    
    print()
    print("================================================")
    print("IFRAMES ENCONTRADOS")
    print("================================================")
    
    print(
        f"TOTAL DE IFRAMES: {len(iframes)}"
    )
    
    for indice, iframe in enumerate(iframes):
    
        try:
    
            print()
            print(f"IFRAME {indice}")
    
            print(
                "ID:",
                iframe.get_attribute("id")
            )
    
            print(
                "NAME:",
                iframe.get_attribute("name")
            )
    
            print(
                "SRC:",
                iframe.get_attribute("src")
            )
    
        except Exception:
    
            pass
    
    return iframes


# ============================================================

# LISTAR INPUTS DA TELA ATUAL

# ============================================================

def listar_inputs(driver, local="DOCUMENTO PRINCIPAL"):


    inputs = driver.find_elements(
        By.TAG_NAME,
        "input"
    )
    
    print()
    print("================================================")
    print(f"INPUTS ENCONTRADOS - {local}")
    print("================================================")
    
    print(
        f"TOTAL DE INPUTS: {len(inputs)}"
    )
    
    for indice, input_element in enumerate(inputs):
    
        try:
    
            print()
            print(f"INPUT {indice}")
    
            print(
                "ID:",
                input_element.get_attribute("id")
            )
    
            print(
                "NAME:",
                input_element.get_attribute("name")
            )
    
            print(
                "PLACEHOLDER:",
                input_element.get_attribute("placeholder")
            )
    
            print(
                "ARIA-LABEL:",
                input_element.get_attribute("aria-label")
            )
    
            print(
                "TYPE:",
                input_element.get_attribute("type")
            )
    
            print(
                "VALUE:",
                input_element.get_attribute("value")
            )
    
            print(
                "CLASS:",
                input_element.get_attribute("class")
            )
    
        except Exception as erro:
    
            print(
                "ERRO AO LER INPUT:",
                erro
            )


# ============================================================

# LISTAR BOTÕES DA TELA ATUAL

# ============================================================

def listar_botoes(driver, local="DOCUMENTO PRINCIPAL"):


    botoes = driver.find_elements(
        By.TAG_NAME,
        "button"
    )
    
    print()
    print("================================================")
    print(f"BOTÕES ENCONTRADOS - {local}")
    print("================================================")
    
    print(
        f"TOTAL DE BOTÕES: {len(botoes)}"
    )
    
    for indice, botao in enumerate(botoes):
    
        try:
    
            texto = botao.text.strip()
    
            print()
            print(f"BOTÃO {indice}")
    
            print(
                "TEXTO:",
                texto
            )
    
            print(
                "ID:",
                botao.get_attribute("id")
            )
    
            print(
                "NG-CLICK:",
                botao.get_attribute("ng-click")
            )
    
            print(
                "TITLE:",
                botao.get_attribute("title")
            )
    
            print(
                "ARIA-LABEL:",
                botao.get_attribute("aria-label")
            )
    
            print(
                "CLASS:",
                botao.get_attribute("class")
            )
    
        except Exception:
    
            pass


# ============================================================

# PROCURAR BOTÃO APLICAR

# ============================================================

def localizar_botao_aplicar(driver):

    try:

        botao = WebDriverWait(
            driver,
            10
        ).until(
            EC.element_to_be_clickable(
                (
                    By.CSS_SELECTOR,
                    "div.modal-dialog button.btn-popup-ok.btn-primary"
                )
            )
        )

        print()
        print("BOTÃO APLICAR DO MODAL ENCONTRADO")

        return botao

    except Exception as erro:

        print(
            "BOTÃO APLICAR NÃO ENCONTRADO:"
        )

        print(
            str(erro)
        )

        return None


# ============================================================

# PROCURAR CAMPO DA NOTA

# ============================================================


def localizar_campo_nota(driver):

    print()
    print("================================================")
    print("PROCURANDO CAMPO NRO NOTA")
    print("================================================")
    
    
    # --------------------------------------------------------
    # 1 - LOCALIZAR O MODAL "INFORME OS PARÂMETROS"
    # --------------------------------------------------------
    
    try:
    
        modal = WebDriverWait(
            driver,
            10
        ).until(
            EC.visibility_of_element_located(
                (
                    By.XPATH,
                    "//div[contains(@class, 'modal-dialog') and .//h5[contains(normalize-space(.), 'Informe os parâmetros')]]"
                )
            )
        )
    
        print(
            "MODAL 'INFORME OS PARÂMETROS' ENCONTRADO"
        )
    
    except Exception as erro:
    
        print(
            "MODAL NÃO ENCONTRADO:"
        )
    
        print(
            str(erro)
        )
    
        return None
    
    
    # --------------------------------------------------------
    # 2 - PROCURAR O INPUT DENTRO DO SK-NUMBER-INPUT
    # --------------------------------------------------------
    
    try:
    
        campo = modal.find_element(
            By.CSS_SELECTOR,
            "sk-number-input input"
        )
    
        if campo.is_displayed():
    
            print()
            print(
                "CAMPO DA NOTA ENCONTRADO"
            )
    
            print(
                "TAG:",
                campo.tag_name
            )
    
            print(
                "TYPE:",
                campo.get_attribute("type")
            )
    
            print(
                "CLASS:",
                campo.get_attribute("class")
            )
    
            return campo
    
    
    except Exception as erro:
    
        print(
            "NÃO FOI POSSÍVEL LOCALIZAR O INPUT:"
        )
    
        print(
            str(erro)
        )
    
    
    return None


def localizar_campo_nota(driver):


    print()
    print("================================================")
    print("PROCURANDO CAMPO NRO NOTA")
    print("================================================")
    
    
    try:
    
        # ----------------------------------------------------
        # AGUARDAR O MODAL
        # ----------------------------------------------------
    
        modal = WebDriverWait(
            driver,
            10
        ).until(
            EC.visibility_of_element_located(
                (
                    By.XPATH,
                    "//div[contains(@class, 'modal-dialog')]"
                    "[.//h5[contains(normalize-space(.), 'Informe os parâmetros')]]"
                )
            )
        )
    
    
        # ----------------------------------------------------
        # PROCURAR O CAMPO ASSOCIADO AO TEXTO NRO NOTA
        # ----------------------------------------------------
    
        campo = modal.find_element(
            By.XPATH,
            ".//label[.//span[contains(normalize-space(.), 'Nro Nota')]]"
            "/following-sibling::div//sk-number-input//input"
        )
    
    
        WebDriverWait(
            driver,
            10
        ).until(
            lambda d: campo.is_displayed()
            and campo.is_enabled()
        )
    
    
        print()
        print("================================================")
        print("CAMPO NRO NOTA ENCONTRADO")
        print("================================================")
    
        print(
            "VALOR ATUAL:",
            campo.get_attribute("value")
        )
    
        print(
            "TYPE:",
            campo.get_attribute("type")
        )
    
        print(
            "CLASS:",
            campo.get_attribute("class")
        )
    
    
        return campo
    
    
    except Exception as erro:
    
        print()
        print("ERRO AO LOCALIZAR CAMPO NRO NOTA")
    
        print(
            str(erro)
        )
    
        return None


# ============================================================
# PREENCHER NÚMERO DA NOTA
# ============================================================

def preencher_numero_nota(driver, campo, numero_nota):

    try:

        print()
        print("================================================")
        print("PREENCHENDO NÚMERO DA NOTA")
        print("================================================")


        # Garantir que o campo esteja visível
        driver.execute_script(
            """
            arguments[0].scrollIntoView({
                block: 'center'
            });
            """,
            campo
        )


        time.sleep(0.5)


        # Clicar no campo
        campo.click()


        # Limpar valor existente
        campo.clear()


        # Digitar o número da nota
        campo.send_keys(
            str(numero_nota)
        )


        # Disparar eventos para o Angular/Sankhya
        driver.execute_script(
            """
            arguments[0].dispatchEvent(
                new Event('input', {
                    bubbles: true
                })
            );

            arguments[0].dispatchEvent(
                new Event('change', {
                    bubbles: true
                })
            );

            arguments[0].dispatchEvent(
                new Event('blur', {
                    bubbles: true
                })
            );
            """,
            campo
        )


        time.sleep(1)


        # Conferir o valor
        valor = campo.get_attribute(
            "value"
        )


        print(
            f"VALOR ATUAL DO CAMPO: {valor}"
        )


        if str(valor).strip() != str(numero_nota):

            print(
                "ERRO: O VALOR PREENCHIDO "
                "NÃO CORRESPONDE À NOTA."
            )

            return False


        print(
            "NOTA PREENCHIDA COM SUCESSO"
        )

        return True


    except Exception as erro:

        print()
        print("================================================")
        print("ERRO AO PREENCHER A NOTA")
        print("================================================")

        print(
            str(erro)
        )

        return False


# ============================================================

# EXPLORAR TODOS OS IFRAMES

# ============================================================

def explorar_iframes(driver):


    driver.switch_to.default_content()
    
    # ========================================================
    # PRIMEIRO TENTAR NO DOCUMENTO PRINCIPAL
    # ========================================================
    
    botao = localizar_botao_aplicar(driver)
    
    if botao:
    
        print()
        print("BOTÃO APLICAR ENCONTRADO NO DOCUMENTO PRINCIPAL")
    
        return {
            "local": "principal",
            "iframe": None,
            "botao": botao
        }
    
    
    # ========================================================
    # PROCURAR NOS IFRAMES
    # ========================================================
    
    iframes = driver.find_elements(
        By.TAG_NAME,
        "iframe"
    )
    
    print()
    print(
        f"TOTAL DE IFRAMES PARA PESQUISA: {len(iframes)}"
    )
    
    
    for indice in range(len(iframes)):
    
        try:
    
            driver.switch_to.default_content()
    
            iframes_atualizados = driver.find_elements(
                By.TAG_NAME,
                "iframe"
            )
    
            driver.switch_to.frame(
                iframes_atualizados[indice]
            )
    
            print()
            print(
                f"PROCURANDO APLICAR NO IFRAME {indice}"
            )
    
    
            botao = localizar_botao_aplicar(
                driver
            )
    
    
            if botao:
    
                print()
                print(
                    f"BOTÃO APLICAR ENCONTRADO NO IFRAME {indice}"
                )
    
                return {
                    "local": "iframe",
                    "iframe": indice,
                    "botao": botao
                }
    
    
        except Exception as erro:
    
            print(
                f"Erro no iframe {indice}:",
                erro
            )
    
    
    driver.switch_to.default_content()
    
    return None


# ============================================================

# CLICAR NO BOTAO

# ============================================================


def clicar_botao_aplicar(driver, botao):


    try:
    
        print()
        print("================================================")
        print("CLICANDO NO BOTÃO APLICAR")
        print("================================================")
    
    
        # Rola o botão até ficar visível
        driver.execute_script(
            """
            arguments[0].scrollIntoView({
                behavior: 'instant',
                block: 'center'
            });
            """,
            botao
        )
    
    
        time.sleep(0.5)
    
    
        # Tenta o clique normal primeiro
        try:
    
            botao.click()
    
            print(
                "CLIQUE NORMAL REALIZADO COM SUCESSO"
            )
    
            return True
    
        except Exception as erro:
    
            print(
                "Clique normal falhou:"
            )
    
            print(
                str(erro)
            )
    
    
        # Se o clique normal falhar,
        # tenta clique via JavaScript.
    
        driver.execute_script(
            "arguments[0].click();",
            botao
        )
    
        print(
            "CLIQUE VIA JAVASCRIPT REALIZADO COM SUCESSO"
        )
    
        return True
    
    
    except Exception as erro:
    
        print()
        print(
            "ERRO AO CLICAR NO BOTÃO APLICAR:"
        )
    
        print(
            str(erro)
        )
    
        return False





# ============================================================

# EXECUTAR CONSULTA

# ============================================================

def consultar_nota(numero_nota):


    try:
    
        print()
        print("================================================")
        print("INICIANDO AUTOMAÇÃO")
        print("================================================")
    
        print(f"NOTA PARA CONSULTA: {numero_nota}")
    
    
        # ----------------------------------------------------
        # CONECTAR AO CHROME
        # ----------------------------------------------------
    
        driver = conectar_chrome()
    
        print("CONECTADO AO GOOGLE CHROME")
    
    
        # ----------------------------------------------------
        # LOCALIZAR PORTAL
        # ----------------------------------------------------
    
        handle_portal = localizar_portal(
            driver
        )
    
        if not handle_portal:
    
            return {
                "sucesso": False,
                "mensagem": (
                    "Não foi encontrada a aba "
                    "Portal de importação de XML."
                )
            }

        # Localizar campo da nota
        campo_nota = localizar_campo_nota(driver)
        
        if not campo_nota:
            return {
                "sucesso": False,
                "mensagem": "Campo 'Nro Nota' não foi encontrado."
            }
        
        
        # Preencher a nota
        nota_preenchida = preencher_numero_nota(
            driver,
            campo_nota,
            numero_nota
        )
        
        if not nota_preenchida:
            return {
                "sucesso": False,
                "mensagem": "Não foi possível preencher o número da nota."
            }
        
        
        # Localizar o botão Aplicar
        resultado_aplicar = explorar_iframes(driver)


        # ============================================================
        # GARANTIR QUE ESTAMOS NO POP-UP DO PORTAL
        # ============================================================
        
        driver.switch_to.window(handle_portal)
        
        print()
        print("================================================")
        print("CONFIRMANDO JANELA ATUAL")
        print("================================================")
        
        print("TÍTULO:",driver.title)
        
        print(
            "URL:",
            driver.current_url
        )
    
        # ----------------------------------------------------
        # AGUARDAR CARREGAMENTO
        # ----------------------------------------------------
    
        time.sleep(2)
    
        driver.switch_to.window(
            handle_portal
        )
        
        driver.switch_to.default_content()
    
        # ----------------------------------------------------
        
        # LOCALIZAR O BOTÃO APLICAR
        
        # ----------------------------------------------------
        
        resultado_aplicar = explorar_iframes(
        driver
        )
        
        if not resultado_aplicar:
        
        
            return {
                "sucesso": False,
                "mensagem": (
                    "O botão Aplicar não foi encontrado "
                    "no Portal de importação de XML."
                )
            }

        
        # ----------------------------------------------------
        
        # CLICAR NO BOTÃO APLICAR
        
        # ----------------------------------------------------
        
        botao_aplicar = resultado_aplicar["botao"]
        
        clicou = clicar_botao_aplicar(
        driver,
        botao_aplicar
        )
        
        if not clicou:
        
        
            return {
                "sucesso": False,
                "mensagem": (
                    "O botão Aplicar foi encontrado, "
                    "mas não foi possível clicar nele."
                )
            }

        
        # ----------------------------------------------------
        
        # RESULTADO
        
        # ----------------------------------------------------
        
        time.sleep(2)
        
        return {
        "sucesso": True,
        "mensagem": (
        f"O botão Aplicar foi acionado com sucesso "
        f"para a nota {numero_nota}."
        )
        }

    
    
    except Exception as erro:
    
        print()
        print("================================================")
        print("ERRO NA AUTOMAÇÃO")
        print("================================================")
    
        print(str(erro))
    
        return {
            "sucesso": False,
            "mensagem": str(erro)
        }

