import time

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException

# ============================================================

# CONFIGURAÇÃO

# ============================================================

DEBUGGER_ADDRESS = "127.0.0.1:9222"

TEMPO_ESPERA = 10

# ============================================================

# CONECTAR AO GOOGLE CHROME

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
# FUNÇÃO PADRÃO PARA MUDAR DE ABA
# ============================================================

def mudar_aba(driver, nome):

    print()
    print("================================================")
    print(f"PROCURANDO ABA: {nome}")
    print("================================================")

    try:

        for handle in driver.window_handles:

            driver.switch_to.window(handle)

            titulo = driver.title.strip()

            print()
            print("HANDLE:", handle)
            print("TÍTULO:", titulo)
            print("URL:", driver.current_url)

            if nome.lower() in titulo.lower():

                print()
                print("ABA ENCONTRADA")
                print("MUDANÇA DE ABA REALIZADA COM SUCESSO")

                return True

        print()
        print("ABA NÃO ENCONTRADA")

        return False

    except Exception as erro:

        print()
        print("ERRO AO MUDAR DE ABA:")
        print(erro)

        return False


# ============================================================
# FUNÇÃO PADRÃO PARA FECHAR ABA
# ============================================================

def fechar_aba(driver, nome):

    print()
    print("================================================")
    print(f"PROCURANDO ABA PARA FECHAR: {nome}")
    print("================================================")

    try:

        for handle in driver.window_handles:

            driver.switch_to.window(handle)

            titulo = driver.title.strip()

            print()
            print("HANDLE:", handle)
            print("TÍTULO:", titulo)
            print("URL:", driver.current_url)

            if nome.lower() in titulo.lower():

                print()
                print("ABA ENCONTRADA")
                print("FECHANDO ABA...")

                driver.close()

                print()
                print("ABA FECHADA COM SUCESSO")

                # Voltar para a primeira aba disponível
                if driver.window_handles:

                    driver.switch_to.window(
                        driver.window_handles[0]
                    )

                    print()
                    print("ABA ATUAL:")
                    print(driver.title)

                return True

        print()
        print("ABA NÃO ENCONTRADA")

        return False

    except Exception as erro:

        print()
        print("ERRO AO FECHAR ABA:")
        print(erro)

        return False



# ============================================================

# PROCURAR ELEMENTO

# ============================================================

def procurar_elemento(
driver,
seletor
):


    # --------------------------------------------------------
    # PROCURAR NO DOCUMENTO ATUAL
    # --------------------------------------------------------
    
    try:
    
        elementos = driver.find_elements(
            By.CSS_SELECTOR,
            seletor
        )
    
        for elemento in elementos:
    
            try:
    
                if (
                    elemento.is_displayed()
                    and elemento.is_enabled()
                ):
    
                    return elemento
    
            except Exception:
    
                continue
    
    except Exception:
    
        pass
    
    
    # --------------------------------------------------------
    # PROCURAR NOS IFRAMES
    # --------------------------------------------------------
    
    try:
    
        iframes = driver.find_elements(
            By.TAG_NAME,
            "iframe"
        )
    
    except Exception:
    
        return None
    
    
    for indice in range(len(iframes)):
    
        try:
    
            iframes_atualizados = driver.find_elements(
                By.TAG_NAME,
                "iframe"
            )
    
            if indice >= len(iframes_atualizados):
    
                continue
    
    
            iframe = iframes_atualizados[indice]
    
            driver.switch_to.frame(
                iframe
            )
    
    
            elemento = procurar_elemento(
                driver,
                seletor
            )
    
    
            if elemento:
    
                return elemento
    
    
            driver.switch_to.parent_frame()
    
    
        except Exception:
    
            try:
    
                driver.switch_to.parent_frame()
    
            except Exception:
    
                pass
    
    
    return None

# ============================================================

# FUNÇÃO PADRÃO PARA CLICAR EM UM CAMPO/BOTÃO

# ============================================================


def clicar_campo(driver, seletor):
    print()
    print("================================================")
    print("PROCURANDO ELEMENTO PARA CLICAR")
    print("================================================")
    print("SELETOR:", seletor)

    inicio = time.time()

    while time.time() - inicio < TEMPO_ESPERA:

        try:
            driver.switch_to.default_content()

            # Procura o elemento NOVAMENTE
            elemento = procurar_elemento(
                driver,
                seletor
            )

            if not elemento:
                time.sleep(0.2)
                continue

            print("ELEMENTO ENCONTRADO")

            try:
                driver.execute_script(
                    """
                    arguments[0].scrollIntoView({
                        block: 'center',
                        inline: 'center'
                    });
                    """,
                    elemento
                )

                time.sleep(0.2)

                elemento.click()

                print("CLIQUE REALIZADO COM SUCESSO")
                return True

            except StaleElementReferenceException:

                print()
                print(
                    "ELEMENTO FICOU OBSOLETO "
                    "DURANTE O CLIQUE."
                )
                print(
                    "PROCURANDO O ELEMENTO NOVAMENTE..."
                )

                time.sleep(0.3)

                continue

            except Exception as erro:

                print("CLIQUE NORMAL FALHOU:")
                print(erro)

                # Procura NOVAMENTE antes do JavaScript
                try:
                    driver.switch_to.default_content()

                    elemento = procurar_elemento(
                        driver,
                        seletor
                    )

                    if not elemento:
                        continue

                    driver.execute_script(
                        """
                        arguments[0].click();
                        """,
                        elemento
                    )

                    print(
                        "CLIQUE VIA JAVASCRIPT "
                        "REALIZADO COM SUCESSO"
                    )

                    return True

                except StaleElementReferenceException:

                    print()
                    print(
                        "ELEMENTO FICOU OBSOLETO "
                        "NOVAMENTE."
                    )
                    print(
                        "TENTANDO NOVAMENTE..."
                    )

                    time.sleep(0.3)
                    continue

        except Exception as erro:

            print()
            print("ERRO AO PROCURAR ELEMENTO:")
            print(erro)

            time.sleep(0.3)

    print()
    print("TEMPO ESGOTADO.")
    print("ELEMENTO NÃO FOI CLICADO:")
    print(seletor)

    return False


# ============================================================

# FUNÇÃO PADRÃO PARA PREENCHER UM CAMPO

# ============================================================

def preencher_campo(driver, seletor, valor):
    print()
    print("================================================")
    print("PROCURANDO CAMPO PARA PREENCHER")
    print("================================================")
    print("SELETOR:", seletor)
    print("VALOR:", valor)

    inicio = time.time()

    while time.time() - inicio < TEMPO_ESPERA:

        try:
            driver.switch_to.default_content()

            # Procurar novamente o elemento
            campo = procurar_elemento(
                driver,
                seletor
            )

            if not campo:
                time.sleep(0.2)
                continue

            print("CAMPO ENCONTRADO")

            try:
                driver.execute_script(
                    """
                    arguments[0].scrollIntoView({
                        block: 'center',
                        inline: 'center'
                    });
                    """,
                    campo
                )

                time.sleep(0.2)

                campo.click()

                campo.send_keys(
                    Keys.CONTROL,
                    "a"
                )

                campo.send_keys(
                    Keys.BACKSPACE
                )

                campo.send_keys(
                    str(valor)
                )

                campo.send_keys(
                    Keys.TAB
                )

                time.sleep(0.5)

                valor_atual = campo.get_attribute(
                    "value"
                )

                print()
                print(
                    "VALOR ATUAL:",
                    valor_atual
                )

                if (
                    str(valor_atual).strip()
                    ==
                    str(valor).strip()
                ):
                    print()
                    print(
                        "CAMPO PREENCHIDO "
                        "COM SUCESSO"
                    )

                    return True

                print()
                print(
                    "VALOR NÃO CONFERIU."
                )
                print(
                    "TENTANDO NOVAMENTE..."
                )

            except StaleElementReferenceException:

                print()
                print(
                    "CAMPO FICOU OBSOLETO "
                    "DURANTE O PREENCHIMENTO."
                )

                print(
                    "PROCURANDO NOVAMENTE..."
                )

                time.sleep(0.3)

                continue

        except Exception as erro:

            print()
            print(
                "ERRO AO PREENCHER CAMPO:"
            )
            print(erro)

            time.sleep(0.3)

    print()
    print("TEMPO ESGOTADO.")
    print("CAMPO NÃO FOI PREENCHIDO:")
    print(seletor)

    return False


def digitar_no_foco(driver, valor):
    print()
    print("================================================")
    print("DIGITANDO NO CAMPO ATUAL")
    print("================================================")
    print("VALOR:", valor)

    try:
        campo = driver.switch_to.active_element

        campo.send_keys(str(valor))
        time.sleep(0.3)

        print("VALOR DIGITADO COM SUCESSO")

        # Pressiona F7 para salvar
        campo.send_keys(Keys.F7)

        print("F7 ENVIADO PARA SALVAR")
        return True

    except Exception as erro:
        print()
        print("ERRO AO DIGITAR OU SALVAR:")
        print(erro)
        return False

