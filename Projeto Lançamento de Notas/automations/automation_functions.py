import time

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import StaleElementReferenceException

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

            if nome.lower() in titulo.lower():

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

def procurar_elemento(driver, seletor):

    # ========================================================
    # IDENTIFICAR TIPO DE SELETOR
    # ========================================================

    if seletor.startswith("//") or seletor.startswith(".//"):

        tipo_seletor = By.XPATH

    else:

        tipo_seletor = By.CSS_SELECTOR


    # ========================================================
    # PROCURAR NO DOCUMENTO ATUAL
    # ========================================================

    try:

        elementos = driver.find_elements(
            tipo_seletor,
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


    # ========================================================
    # PROCURAR NOS IFRAMES
    # ========================================================

    try:

        iframes = driver.find_elements(
            By.TAG_NAME,
            "iframe"
        )

    except Exception:

        return None


    for indice in range(len(iframes)):

        try:

            # Atualizar lista de iframes
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
# VERIFICAR SE EXISTE UM ELEMENTO
# ============================================================

def verificar_elemento(driver, seletor, texto=None):

    try:

        driver.switch_to.default_content()

        elemento = procurar_elemento(
            driver,
            seletor
        )

        if not elemento:
            return False

        # Se não foi informado texto,
        # basta o elemento existir
        if texto is None:
            return True

        # Verificar o texto do elemento
        texto_elemento = elemento.text.strip()

        if texto_elemento == texto:
            return True

        return False

    except Exception:

        return False


# ============================================================
# VERIFICAR TELA
# ============================================================

def verificar_tela(driver):

    grade_xml = (
        "//div[@flex and @layout='row' and @ng-click='select()']"
    )

    tela_grade_xml = verificar_elemento(driver,grade_xml)

    if tela_grade_xml:
        return "Grade XML"

    inicial_xml = (
            "div"
            "[col-id='CODTIPOPER']"
        )

    tela_inicial_xml = verificar_elemento(driver,inicial_xml)

    if tela_inicial_xml:
        return "Inicial XML"

    return "Nenhuma tela enconrtada"


# ============================================================
# VERIFICAR TEXTO DIGITADO NO FOCO
# ============================================================

def verificar_digitado_no_foco(driver):

    try:

        campo = driver.switch_to.active_element

        valor = campo.get_attribute("value")

        print()
        print("================================================")
        print("VALOR NO FOCO")
        print("================================================")
        print("VALOR:", valor)

        return valor

    except Exception as erro:

        print()
        print("ERRO AO VERIFICAR VALOR NO FOCO:")
        print(erro)

        return None



# ============================================================

# FUNÇÃO PADRÃO PARA CLICAR EM UM CAMPO/BOTÃO

# ============================================================


def clicar_campo(driver, seletor):

    try:

        # ====================================================
        # IDENTIFICAR TIPO DO SELETOR
        # ====================================================

        if (
            seletor.startswith("//")
            or seletor.startswith("(//")
            or seletor.startswith(".//")
        ):
            tipo = By.XPATH

        else:
            tipo = By.CSS_SELECTOR


        # ====================================================
        # PROCURAR NO DOCUMENTO ATUAL
        # ====================================================

        def procurar_no_documento():

            elementos = driver.find_elements(
                tipo,
                seletor
            )

            for elemento in elementos:

                try:

                    if (
                        elemento.is_displayed()
                        and elemento.is_enabled()
                    ):

                        return elemento

                except (
                    StaleElementReferenceException
                ):

                    continue

            return None


        # ====================================================
        # PRIMEIRA TENTATIVA
        # ====================================================

        driver.switch_to.default_content()

        elemento = procurar_no_documento()


        # ====================================================
        # PROCURAR NOS IFRAMES
        # ====================================================

        if elemento is None:

            frames = driver.find_elements(
                By.TAG_NAME,
                "iframe"
            )

            for frame in frames:

                try:

                    driver.switch_to.default_content()

                    driver.switch_to.frame(frame)

                    elemento = procurar_no_documento()

                    if elemento is not None:
                        break

                except Exception:

                    continue


        # ====================================================
        # ELEMENTO NÃO ENCONTRADO
        # ====================================================

        if elemento is None:

            driver.switch_to.default_content()

            print()
            print("================================================")
            print("ELEMENTO NÃO ENCONTRADO")
            print("================================================")
            print("SELETOR:", seletor)
            print("TIPO:", tipo)

            return False


        # ====================================================
        # ROLAR ATÉ O ELEMENTO
        # ====================================================

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


        # ====================================================
        # CLICAR
        # ====================================================

        try:

            elemento.click()

        except Exception:

            driver.execute_script(
                "arguments[0].click();",
                elemento
            )


        print()
        print("CLIQUE REALIZADO")
        print("SELETOR:", seletor)

        return True


    except Exception as erro:

        driver.switch_to.default_content()

        print()
        print("================================================")
        print("ERRO AO CLICAR")
        print("================================================")
        print("SELETOR:", seletor)
        print("ERRO:", erro)

        return False

# ============================================================

# FUNÇÃO PADRÃO PARA PREENCHER UM CAMPO

# ============================================================

def preencher_campo(driver, seletor, valor):

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

    try:
        campo = driver.switch_to.active_element

        campo.send_keys(str(valor))
        time.sleep(0.3)

        print("VALOR DIGITADO COM SUCESSO")
        
        return True

    except Exception as erro:
        print()
        print("ERRO AO DIGITAR OU SALVAR:")
        print(erro)
        return False



def procurar_elemento_por_id(driver, id_elemento):

    try:
        elemento = driver.find_element(By.ID, id_elemento)

        if elemento.is_displayed():
            return elemento

    except Exception:
        pass

    # Procurar dentro dos iframes
    iframes = driver.find_elements(By.TAG_NAME, "iframe")

    for iframe in iframes:

        try:
            driver.switch_to.frame(iframe)

            elemento = procurar_elemento_por_id(
                driver,
                id_elemento
            )

            if elemento:
                return elemento

            driver.switch_to.parent_frame()

        except Exception:
            driver.switch_to.parent_frame()

    return None

def verificar_divergencias_devolucao(driver):

    divergencias = []

    elementos = [
        ("divImp", "Impostos"),
        ("divFin", "Financeiro"),
        ("divProd", "Produtos não encontrados"),
        ("divLotes", "Lotes de produtos não informados"),
    ]

    for id_elemento, tipo in elementos:

        driver.switch_to.default_content()

        elemento = procurar_elemento_por_id(
            driver,
            id_elemento
        )

        if elemento:

            texto = elemento.text.strip()

            if texto:

                divergencias.append({
                    "tipo": tipo,
                    "mensagem": texto
                })

                print()
                print(f"DIVERGÊNCIA: {tipo}")
                print(f"MENSAGEM: {texto}")

    driver.switch_to.default_content()

    return {
        "existe": len(divergencias) > 0,
        "divergencias": divergencias
    }

