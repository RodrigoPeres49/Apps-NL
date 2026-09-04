import time

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException

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

# LOCALIZAR A ABA DO PORTAL

# ============================================================

def localizar_portal(driver):


    print()
    print("================================================")
    print("PROCURANDO ABA DO PORTAL")
    print("================================================")
    
    for handle in driver.window_handles:
    
        try:
    
            # Mudar para a aba atual
            driver.switch_to.window(handle)
    
            titulo = driver.title.strip()
    
            print()
            print("------------------------------------------------")
            print("HANDLE:", handle)
            print("TÍTULO:", titulo)
            print("URL:", driver.current_url)
    
            # Verificar se é o Portal
            if "portal de importação de xml" in titulo.lower():
    
                print()
                print("================================================")
                print("PORTAL DE IMPORTAÇÃO DE XML ENCONTRADO")
                print("================================================")
    
                return handle
    
    
        except Exception as erro:
    
            print()
            print("ERRO AO VERIFICAR ABA:")
    
            print(
                str(erro)
            )
    
    
    print()
    print("================================================")
    print("PORTAL DE IMPORTAÇÃO DE XML NÃO ENCONTRADO")
    print("================================================")
    
    return None

# ============================================================

# PROCURAR ELEMENTO NO CONTEXTO ATUAL

# ============================================================

def procurar_elemento_contexto_atual(driver,seletor):


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
    
    
    return None

# ============================================================

# PROCURAR ELEMENTO NO DOCUMENTO E IFRAMES

# ============================================================

def procurar_elemento(driver,seletor):


    # --------------------------------------------------------
    # PRIMEIRO PROCURAR NO CONTEXTO ATUAL
    # --------------------------------------------------------
    
    elemento = procurar_elemento_contexto_atual(
        driver,
        seletor
    )
    
    if elemento:
    
        return elemento
    
    
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
    
            # Atualizar a lista de iframes
            iframes_atualizados = driver.find_elements(
                By.TAG_NAME,
                "iframe"
            )
    
            if indice >= len(iframes_atualizados):
    
                continue
    
    
            iframe = iframes_atualizados[indice]
    
    
            # Entrar no iframe
            driver.switch_to.frame(
                iframe
            )
    
    
            print(
                f"PROCURANDO NO IFRAME {indice}"
            )
    
    
            # Procurar dentro do iframe
            elemento = procurar_elemento(
                driver,
                seletor
            )
    
    
            # Se encontrou, manter o contexto atual
            if elemento:
    
                return elemento
    
    
            # Se não encontrou, voltar
            driver.switch_to.parent_frame()
    
    
        except Exception as erro:
    
            print(
                f"ERRO AO VERIFICAR IFRAME {indice}:"
            )
    
            print(
                str(erro)
            )
    
    
            try:
    
                driver.switch_to.parent_frame()
    
            except Exception:
    
                pass
    
    
    return None

# ============================================================

# LOCALIZAR E PREENCHER O NÚMERO DA NOTA

# ============================================================

def preencher_numero_nota(driver,numero_nota):


    print()
    print("================================================")
    print("PROCURANDO CAMPO DO NÚMERO DA NOTA")
    print("================================================")
    
    
    # Corresponde ao input:
    
    # <input
    # type="text"
    # ng-model="value"
    # ng-keydown="onInputKeyDown($event)"
    # ng-change="changeHandler()"
    # >
    
    seletor_nota = (
        "input"
        "[type='text']"
        "[ng-model='value']"
        "[ng-keydown='onInputKeyDown($event)']"
        "[ng-change='changeHandler()']"
    )
    
    
    try:
    
        # Sempre iniciar pelo documento principal
        driver.switch_to.default_content()
    
    
        # Aguardar até localizar o campo
        campo = WebDriverWait(
            driver,
            TEMPO_ESPERA
        ).until(
            lambda d: procurar_elemento(
                d,
                seletor_nota
            )
        )
    
    
        print()
        print("================================================")
        print("CAMPO DA NOTA ENCONTRADO")
        print("================================================")
    
    
        # Centralizar o campo na tela
        driver.execute_script(
            """
            arguments[0].scrollIntoView({
                block: 'center',
                inline: 'center'
            });
            """,
            campo
        )
    
    
        time.sleep(
            0.5
        )
    
    
        # ----------------------------------------------------
        # CLICAR NO CAMPO
        # ----------------------------------------------------
    
        campo.click()
    
    
        # ----------------------------------------------------
        # LIMPAR O CAMPO
        # ----------------------------------------------------
    
        campo.send_keys(
            Keys.CONTROL,
            "a"
        )
    
    
        campo.send_keys(
            Keys.BACKSPACE
        )
    
    
        # ----------------------------------------------------
        # DIGITAR O NÚMERO DA NOTA
        # ----------------------------------------------------
    
        campo.send_keys(
            str(numero_nota)
        )
    
    
        # ----------------------------------------------------
        # DISPARAR EVENTOS
        # ----------------------------------------------------
    
        driver.execute_script(
            """
            arguments[0].dispatchEvent(
                new Event(
                    'input',
                    {
                        bubbles: true
                    }
                )
            );
    
            arguments[0].dispatchEvent(
                new Event(
                    'change',
                    {
                        bubbles: true
                    }
                )
            );
    
            arguments[0].blur();
            """,
            campo
        )
    
    
        # Aguardar atualização
        time.sleep(
            1
        )
    
    
        # ----------------------------------------------------
        # CONFERIR O VALOR
        # ----------------------------------------------------
    
        valor = campo.get_attribute(
            "value"
        )
    
    
        print()
        print(
            "VALOR ATUAL DO CAMPO:",
            valor
        )
    
    
        if (
            str(valor).strip()
            !=
            str(numero_nota).strip()
        ):
    
            print()
            print("================================================")
            print("ERRO AO PREENCHER A NOTA")
            print("================================================")
    
            return False
    
    
        print()
        print("================================================")
        print("NOTA PREENCHIDA COM SUCESSO")
        print("================================================")
    
    
        return True
    
    
    except TimeoutException:
    
        print()
        print("================================================")
        print("TEMPO ESGOTADO")
        print("CAMPO DA NOTA NÃO ENCONTRADO")
        print("================================================")
    
        return False
    
    
    except Exception as erro:
    
        print()
        print("================================================")
        print("ERRO AO PREENCHER NÚMERO DA NOTA")
        print("================================================")
    
        print(
            str(erro)
        )
    
        return False

# ============================================================

# LOCALIZAR O BOTÃO APLICAR

# ============================================================

def localizar_botao_aplicar(driver):

    print()
    print("================================================")
    print("PROCURANDO BOTÃO APLICAR")
    print("================================================")
    
    
    # Corresponde ao botão:
    
    # <button
    # ng-click="applyFilter();"
    # primary=""
    # block=""
    # >
    # Aplicar
    # </button>
    
    seletor_aplicar = (
        "button"
        "[ng-click='applyFilter();']"
    )
    
    
    try:
    
        # Sempre começar no documento principal
        driver.switch_to.default_content()
    
    
        # Procurar botão no documento e iframes
        botao = WebDriverWait(
            driver,
            TEMPO_ESPERA
        ).until(
            lambda d: procurar_elemento(
                d,
                seletor_aplicar
            )
        )
    
    
        texto = botao.text.strip()
    
    
        print(
            "TEXTO DO BOTÃO:",
            texto
        )
    
    
        # Confirmar que é realmente Aplicar
        if texto.lower() != "aplicar":
    
            print()
            print(
                "BOTÃO ENCONTRADO, "
                "MAS NÃO É O BOTÃO APLICAR."
            )
    
            return None
    
    
        print()
        print("================================================")
        print("BOTÃO APLICAR ENCONTRADO")
        print("================================================")
    
    
        return botao
    
    
    except TimeoutException:
    
        print()
        print("================================================")
        print("TEMPO ESGOTADO")
        print("BOTÃO APLICAR NÃO ENCONTRADO")
        print("================================================")
    
        return None
    
    
    except Exception as erro:
    
        print()
        print("================================================")
        print("ERRO AO LOCALIZAR BOTÃO APLICAR")
        print("================================================")
    
        print(
            str(erro)
        )
    
        return None

# ============================================================

# CLICAR NO BOTÃO APLICAR

# ============================================================

def clicar_botao_aplicar(driver,botao):


    print()
    print("================================================")
    print("CLICANDO NO BOTÃO APLICAR")
    print("================================================")
    
    
    try:
    
        # Centralizar botão
        driver.execute_script(
            """
            arguments[0].scrollIntoView({
                block: 'center',
                inline: 'center'
            });
            """,
            botao
        )
    
    
        time.sleep(
            0.5
        )
    
    
        # ----------------------------------------------------
        # TENTAR CLIQUE NORMAL
        # ----------------------------------------------------
    
        try:
    
            botao.click()
    
    
            print()
            print(
                "CLIQUE NORMAL REALIZADO "
                "COM SUCESSO"
            )
    
    
            return True
    
    
        except Exception as erro:
    
            print()
            print(
                "CLIQUE NORMAL FALHOU:"
            )
    
            print(
                str(erro)
            )
    
    
        # ----------------------------------------------------
        # TENTAR CLIQUE VIA JAVASCRIPT
        # ----------------------------------------------------
    
        driver.execute_script(
            """
            arguments[0].click();
            """,
            botao
        )
    
    
        print()
        print(
            "CLIQUE VIA JAVASCRIPT "
            "REALIZADO COM SUCESSO"
        )
    
    
        return True
    
    
    except Exception as erro:
    
        print()
        print("================================================")
        print("ERRO AO CLICAR NO BOTÃO APLICAR")
        print("================================================")
    
        print(
            str(erro)
        )
    
        return False
    
    
# ============================================================

# FUNÇÃO PRINCIPAL

# ============================================================

def consultar_nota(numero_nota):


    print()
    print("================================================")
    print("INICIANDO AUTOMAÇÃO")
    print("================================================")
    
    print(
        f"NOTA PARA CONSULTA: {numero_nota}"
    )
    
    
    try:
    
        # ----------------------------------------------------
        # CONECTAR AO GOOGLE CHROME
        # ----------------------------------------------------
    
        driver = conectar_chrome()
    
    
        print()
        print(
            "CONECTADO AO GOOGLE CHROME"
        )
    
    
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
    
    
        # ----------------------------------------------------
        # ENTRAR NA ABA CORRETA
        # ----------------------------------------------------
    
        driver.switch_to.window(
            handle_portal
        )
    
    
        driver.switch_to.default_content()
    
    
        # ----------------------------------------------------
        # PREENCHER NÚMERO DA NOTA
        # ----------------------------------------------------
    
        nota_preenchida = preencher_numero_nota(
            driver,
            numero_nota
        )
    
    
        if not nota_preenchida:
    
            return {
                "sucesso": False,
                "mensagem": (
                    "Não foi possível preencher "
                    "o número da nota."
                )
            }
    
    
        # ----------------------------------------------------
        # LOCALIZAR BOTÃO APLICAR
        # ----------------------------------------------------
    
        botao_aplicar = localizar_botao_aplicar(
            driver
        )
    
    
        if not botao_aplicar:
    
            return {
                "sucesso": False,
                "mensagem": (
                    "O botão Aplicar "
                    "não foi encontrado."
                )
            }
    
    
        # ----------------------------------------------------
        # CLICAR NO BOTÃO APLICAR
        # ----------------------------------------------------
    
        clicou = clicar_botao_aplicar(
            driver,
            botao_aplicar
        )
    
    
        if not clicou:
    
            return {
                "sucesso": False,
                "mensagem": (
                    "O botão Aplicar foi encontrado, "
                    "mas não foi possível clicar."
                )
            }
    
    
        # ----------------------------------------------------
        # RESULTADO FINAL
        # ----------------------------------------------------
    
        print()
        print("================================================")
        print("AUTOMAÇÃO FINALIZADA COM SUCESSO")
        print("================================================")
    
    
        return {
            "sucesso": True,
            "mensagem": (
                f"Nota {numero_nota} "
                "processada com sucesso."
            )
        }
    
    
    except Exception as erro:
    
        print()
        print("================================================")
        print("ERRO NA AUTOMAÇÃO")
        print("================================================")
    
        print(
            str(erro)
        )
    
    
        return {
            "sucesso": False,
            "mensagem": str(erro)
        }
    
